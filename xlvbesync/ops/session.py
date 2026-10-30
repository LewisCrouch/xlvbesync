"""A resolved target plus its live Excel/VBProject handles.

Both the CLI and the watcher need the same preamble: figure out which file we
are operating on, attach to Excel, make sure the file is open, grab its
VBProject, and silence prompts. ``TargetSession`` centralises that so the two
entry points cannot drift apart.
"""

from __future__ import annotations

import contextlib
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterator

from xlvbesync.core import console
from xlvbesync.core.addin import is_addin
from xlvbesync.core.config import ProjectConfig, TargetConfig
from xlvbesync.core.errors import TargetNotFoundError
from xlvbesync.core.excel import (
    ensure_workbook,
    find_workbook,
    get_excel_for,
    quiet_excel,
    same_path,
)
from xlvbesync.core.vbproject import get_vbproject


@dataclass
class TargetSession:
    """Everything needed to read or write one target's VBA."""

    config: ProjectConfig
    target: TargetConfig
    path: Path
    src_dir: Path
    #: Every folder injected into this target: shared includes first, then the
    #: target's own folder (which is also ``src_dir``).
    source_dirs: list[Path]
    app: Any
    workbook: Any
    project: Any
    opened_by_us: bool
    started_excel: bool

    @property
    def name(self) -> str:
        return self.target.resolved_name()

    @property
    def is_addin(self) -> bool:
        return is_addin(self.path)

    @property
    def exclude(self) -> set[str]:
        return set(self.target.exclude)

    def save(self) -> bool:
        """Save the workbook so the on-disk file matches the loaded project."""
        try:
            self.workbook.Save()
        except Exception as exc:  # noqa: BLE001
            console.warn(
                f"Could not save {self.path.name}: {console.describe_exception(exc)}"
            )
            return False

        # If the loaded workbook is in a different location than self.path
        # (e.g. an add-in loaded from %AppData%\Microsoft\AddIns), sync the
        # saved file back to the configured target path.
        try:
            loaded = Path(self.workbook.FullName)
            if not same_path(loaded, self.path):
                import shutil

                shutil.copy2(loaded, self.path)
        except Exception as exc:  # noqa: BLE001
            console.warn(
                f"Could not sync saved file to {self.path}: {console.describe_exception(exc)}"
            )

        return True


@contextlib.contextmanager
def open_target(
    config: ProjectConfig,
    target: TargetConfig,
    *,
    require_open: bool = False,
    read_only: bool = False,
    launch: bool | None = None,
    quiet: bool = True,
    close_when_done: bool | None = None,
    app: Any | None = None,
) -> Iterator[TargetSession]:
    """Yield a :class:`TargetSession` for ``target``.

    Parameters
    ----------
    require_open:
        If True, refuse to open the file ourselves - used by hot-reload, where
        silently opening a closed add-in would be surprising.
    close_when_done:
        Close the workbook on exit. Defaults to "only if we opened it".
    app:
        Use this Excel instance instead of resolving one. Swap mode passes the
        instance it already holds so staging cannot drift to another process.
    """
    path = config.target_path(target)
    if not path.is_file():
        raise TargetNotFoundError(
            f"Target file not found: {path}",
            "Check the 'path' entry in xlvbesync.json, or pass --target.",
        )

    started_excel = False
    if app is None:
        # Prefer whichever Excel instance already hosts this file, so a
        # hot-reload lands in the session the user is actually looking at.
        app, started_excel = get_excel_for(
            path, launch=config.launch_excel if launch is None else launch
        )

    workbook = find_workbook(app, path)
    opened_by_us = False

    if workbook is None:
        if require_open:
            from xlvbesync.core.errors import TargetNotOpenError

            raise TargetNotOpenError(
                f"{path.name} is not currently open in Excel.",
                "Open the workbook (or install the add-in) and retry.",
            )
        workbook, opened_by_us = ensure_workbook(app, path, read_only=read_only)

    try:
        loaded = str(workbook.FullName)
    except Exception:  # noqa: BLE001
        loaded = ""
    if loaded and not same_path(loaded, path):
        console.warn(
            f"{path.name} is loaded from {loaded}, not {path}; "
            "injecting into the loaded copy."
        )

    project = get_vbproject(workbook)

    should_close = opened_by_us if close_when_done is None else close_when_done

    with quiet_excel(app, enabled=quiet):
        session = TargetSession(
            config=config,
            target=target,
            path=path,
            src_dir=config.target_dir(target),
            source_dirs=config.source_dirs(target),
            app=app,
            workbook=workbook,
            project=project,
            opened_by_us=opened_by_us,
            started_excel=started_excel,
        )
        try:
            yield session
        finally:
            if should_close:
                with contextlib.suppress(Exception):
                    workbook.Close(SaveChanges=False)
            if started_excel:
                with contextlib.suppress(Exception):
                    app.Quit()


def resolve_targets(
    config: ProjectConfig,
    selector: str | None,
) -> list[TargetConfig]:
    """Return the targets a command should act on.

    ``None`` means "all configured targets"; otherwise the selector is matched
    by name, filename or path.
    """
    if selector is None:
        return list(config.targets)

    found = config.find_target(selector)
    if found is None:
        known = ", ".join(t.resolved_name() for t in config.targets) or "(none)"
        raise TargetNotFoundError(
            f"No configured target matches {selector!r}.",
            f"Configured targets: {known}",
        )
    return [found]
