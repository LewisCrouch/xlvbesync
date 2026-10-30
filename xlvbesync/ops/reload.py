"""Hot-reload strategies for workbooks and add-ins.

Two modes, per project design:

``live``
    Inject into the already-loaded VBProject and save. Instant, preserves
    global state, and is what you want 95% of the time.

``swap``
    Build the updated file, uninstall the add-in, replace the file on disk,
    then reinstall. Guarantees the on-disk file is authoritative and gives the
    add-in a clean ``Workbook_Open`` / ``AddinInstall`` cycle, at the cost of
    losing module-level state.
"""

from __future__ import annotations

import contextlib
from dataclasses import dataclass
from pathlib import Path

from xlvbesync.core import console
from xlvbesync.core.addin import (
    discard_backup,
    ensure_addin_entry,
    find_addin,
    is_addin,
    reload_hosts,
    replace_file,
    settle,
    unload_addin,
)
from xlvbesync.core.config import MODE_LIVE, MODE_SWAP, ProjectConfig, TargetConfig
from xlvbesync.core.customui import find_customui_dir, sync_customui
from xlvbesync.core.errors import TargetNotFoundError, XlvbesyncError
from xlvbesync.core.excel import get_excel_for, quiet_excel, same_path
from xlvbesync.core.state import SyncState
from xlvbesync.ops.inject import InjectReport, inject_target, print_report
from xlvbesync.ops.session import open_target


@dataclass
class ReloadReport:
    """Outcome of a reload."""

    target: str
    mode: str
    inject: InjectReport | None = None
    reloaded: bool = False
    note: str = ""

    @property
    def ok(self) -> bool:
        return self.inject is None or self.inject.ok


def reload_live(
    config: ProjectConfig,
    target: TargetConfig,
    *,
    state: SyncState | None = None,
    only: list[Path] | None = None,
    force: bool = False,
) -> ReloadReport:
    """Inject straight into the loaded project, then save."""
    report = ReloadReport(target=target.resolved_name(), mode=MODE_LIVE)

    with open_target(config, target, quiet=True) as session:
        report.inject = inject_target(
            session, state=state, only=only, force=force
        )
    return report


def reload_swap(
    config: ProjectConfig,
    target: TargetConfig,
    *,
    state: SyncState | None = None,
    only: list[Path] | None = None,
    force: bool = False,
) -> ReloadReport:
    """Full unload / replace-on-disk / reload cycle for an add-in.

    Sequence
    --------
    1. Open a *hidden second copy* of the add-in from a staging path and inject
       into that, so the live add-in is untouched while we build the new file.
    2. Save the staged copy and close it.
    3. Uninstall the live add-in and wait for the file handle to drop.
    4. ``os.replace`` the staged file over the original.
    5. Reinstall.

    Staging matters: injecting into the live add-in and *then* swapping would
    make step 3 rewrite the file we just edited, undoing the work.
    """
    name = target.resolved_name()
    report = ReloadReport(target=name, mode=MODE_SWAP)

    path = config.target_path(target)
    if not path.is_file():
        raise TargetNotFoundError(f"Target file not found: {path}")

    if not is_addin(path):
        # Swap only makes sense for add-ins; a workbook can just be saved.
        console.warn(
            f"{path.name} is not an add-in; using live mode instead of swap."
        )
        return reload_live(config, target, state=state, only=only, force=force)

    app, started = get_excel_for(path, launch=config.launch_excel)

    # Keep the real extension last: Excel refuses to open a file whose
    # extension it does not recognise, so "Demo.xlvbesync-stage.xlam" works while
    # "Demo.xlam.xlvbesync-stage" is fragile.
    staged = path.with_name(f"{path.stem}.xlvbesync-stage{path.suffix}")
    with contextlib.suppress(OSError):
        staged.unlink()

    backups: list[tuple[Path, Path | None]] = []
    hosts: list = []

    try:
        with quiet_excel(app):
            # --- 1/2. Build the updated file in isolation ----------------
            import shutil

            shutil.copy2(path, staged)

            staged_target = TargetConfig(
                path=str(staged),
                name=name,
                mode=MODE_LIVE,
                save_after_inject=True,
                exclude=list(target.exclude),
            )

            with open_target(
                config,
                staged_target,
                quiet=True,
                close_when_done=True,
                app=app,
            ) as session:
                # Point the session at the real source folder, not one derived
                # from the staging filename.
                session.src_dir = config.target_dir(target)
                report.inject = inject_target(
                    session, state=state, only=only, force=force, save=True
                )

            if not report.inject.ok:
                raise XlvbesyncError(
                    f"Injection into the staged copy of {path.name} failed; "
                    "the live add-in was left untouched."
                )

            # Sync customUI files (Ribbon XML and images) into the staged copy
            customui_dir = find_customui_dir(config.target_dir(target))
            if customui_dir is not None:
                sync_customui(staged, customui_dir)

            settle()

            # --- 3. Unload from every Excel instance ---------------------
            hosts = unload_addin(app, path)
            settle()

            # --- 4. Replace on disk --------------------------------------
            # Replace the repo target file and any instance's installed add-in
            # path (e.g. %AppData%\Microsoft\AddIns\name.xlam).
            destinations: list[Path] = [path]
            for _, handle in hosts:
                if handle is not None and handle.path is not None:
                    if not any(same_path(handle.path, d) for d in destinations):
                        destinations.append(handle.path)

            for dest in destinations:
                bak = replace_file(staged, dest)
                backups.append((dest, bak))

            # --- 5. Reload everywhere it was installed -------------------
            restored = reload_hosts(path, hosts)
            report.reloaded = True
            if restored > 1:
                report.note = f"re-enabled in {restored} Excel instances"

        for _, bak in backups:
            discard_backup(bak)
        backups.clear()

    except Exception:
        # Best effort: get the user's add-in back into a working state.
        import shutil

        for dest, bak in backups:
            if bak is not None and bak.is_file():
                with contextlib.suppress(Exception):
                    shutil.copy2(bak, dest)
        if hosts:
            with contextlib.suppress(Exception):
                reload_hosts(path, hosts)
        raise
    finally:
        with contextlib.suppress(OSError):
            staged.unlink()
        if started:
            with contextlib.suppress(Exception):
                app.Quit()

    return report


def reload_target(
    config: ProjectConfig,
    target: TargetConfig,
    *,
    mode: str | None = None,
    state: SyncState | None = None,
    only: list[Path] | None = None,
    force: bool = False,
) -> ReloadReport:
    """Reload ``target`` using the requested (or configured) mode."""
    target_path = config.target_path(target)
    chosen = mode or target.mode

    # If the user did not explicitly request live mode and this add-in has
    # customUI, it cannot be injected live so swap mode is required.
    if (
        mode is None
        and is_addin(target_path)
        and find_customui_dir(config.target_dir(target)) is not None
    ):
        chosen = MODE_SWAP

    if chosen == MODE_SWAP:
        return reload_swap(config, target, state=state, only=only, force=force)
    return reload_live(config, target, state=state, only=only, force=force)


def ensure_installed(config: ProjectConfig, target: TargetConfig) -> bool:
    """Make sure an add-in target is registered and installed in Excel."""
    path = config.target_path(target)
    if not is_addin(path):
        return False
    app, _ = get_excel_for(path, launch=config.launch_excel)
    handle = find_addin(app, path) or ensure_addin_entry(app, path)
    try:
        if not handle.com.Installed:
            handle.com.Installed = True
        return True
    except Exception as exc:  # noqa: BLE001
        console.warn(f"Could not install {path.name}: {console.describe_exception(exc)}")
        return False


def print_reload(report: ReloadReport, *, verbose: bool = False) -> None:
    """Render a reload report."""
    if report.inject is not None:
        print_report(report.inject, verbose=verbose)
    if report.reloaded:
        console.detail(f"add-in reloaded ({report.mode} mode)")
    if report.note:
        console.detail(report.note)
