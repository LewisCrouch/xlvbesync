"""Extract VBA components from a workbook into the repo."""

from __future__ import annotations

import contextlib
from dataclasses import dataclass, field
from pathlib import Path

from xlvba.core import console
from xlvba.core.components import FORM_BINARY_EXT, is_source_file
from xlvba.core.state import SyncState, hash_text
from xlvba.core.vbproject import document_module_map, export_any, list_components
from xlvba.ops.inject import (
    InjectReport,
    inject_target,
    print_report as print_inject,
    shared_sources,
)
from xlvba.ops.session import TargetSession


@dataclass
class ExtractReport:
    """What an extraction did."""

    target: str
    written: list[str] = field(default_factory=list)
    unchanged: list[str] = field(default_factory=list)
    skipped: list[tuple[str, str]] = field(default_factory=list)
    pruned: list[str] = field(default_factory=list)

    #: Shared-library components left alone because the repo copy is
    #: authoritative, as ``(filename, shared folder)``.
    shared: list[tuple[str, str]] = field(default_factory=list)

    #: Result of pushing the shared library back into the target.
    pushed: InjectReport | None = None

    @property
    def counts(self) -> dict[str, int]:
        out = {
            "written": len(self.written),
            "unchanged": len(self.unchanged),
            "skipped": len(self.skipped),
            "pruned": len(self.pruned),
        }
        if self.shared:
            out["shared"] = len(self.shared)
        return out

    @property
    def changed(self) -> bool:
        return bool(self.written or self.pruned)


def extract_target(
    session: TargetSession,
    *,
    state: SyncState | None = None,
    prune: bool = False,
    include_empty: bool = True,
    dry_run: bool = False,
    extract_shared: bool = False,
) -> ExtractReport:
    """Export every supported component of ``session`` to its source folder.

    Shared-library components (those provided by an ``include`` folder) are
    *not* extracted by default. Writing them into the target's own folder
    would silently turn them into a local override, so instead the repo copy
    stays authoritative and is pushed back into the target.

    Parameters
    ----------
    prune:
        Delete local source files whose component no longer exists in Excel.
    include_empty:
        Write files for components with no code. Sheets usually have none, and
        skipping them keeps the tree tidy, but they are needed if you want to
        start adding sheet code locally.
    extract_shared:
        Pull shared components out of Excel after all, writing each one back
        over its file in the shared folder it came from. This is the explicit
        "Excel wins" direction and skips the push-back.
    """
    report = ExtractReport(target=session.name)
    components = list_components(session.project, exclude=session.exclude)
    doc_names = document_module_map(session.workbook)
    shared = shared_sources(session)

    if not dry_run:
        session.src_dir.mkdir(parents=True, exist_ok=True)

    seen_files: set[str] = set()

    for info in components:
        shared_path = shared.get(info.ref.filename.casefold())

        if shared_path is not None and not extract_shared:
            # The repo's shared copy wins; leave the target folder clean.
            report.shared.append(
                (info.ref.filename, _folder_label(session, shared_path))
            )
            continue

        # An explicit shared extract updates the shared file in place, never
        # a duplicate in the target's own folder.
        if shared_path is not None:
            dest = shared_path
        else:
            dest = session.src_dir / info.ref.filename
        seen_files.add(dest.name.casefold())
        if info.ref.ext == ".frm":
            seen_files.add(dest.with_suffix(FORM_BINARY_EXT).name.casefold())

        if info.is_empty and not include_empty:
            hint = doc_names.get(info.name, "")
            report.skipped.append((info.ref.filename, f"empty{f' - {hint}' if hint else ''}"))
            continue

        if dry_run:
            report.written.append(info.ref.filename)
            continue

        try:
            text = export_any(
                info, dest, ansi_encoding=session.config.ansi_encoding
            )
        except Exception as exc:  # noqa: BLE001 - one bad module must not stop the rest
            report.skipped.append((info.ref.filename, console.describe_exception(exc)))
            continue

        digest = hash_text(text)
        rel = dest.relative_to(session.config.src_root)
        if state is not None:
            if state.is_unchanged(session.name, rel, digest):
                report.unchanged.append(info.ref.filename)
            else:
                report.written.append(info.ref.filename)
            state.record(session.name, rel, digest)
        else:
            report.written.append(info.ref.filename)

        # Track the .frx separately so the watcher does not treat the binary
        # sibling as an unknown new file.
        if info.ref.ext == ".frm" and state is not None:
            frx = dest.with_suffix(FORM_BINARY_EXT)
            if frx.is_file():
                from xlvba.core.state import hash_file

                state.record(
                    session.name, frx.relative_to(session.config.src_root), hash_file(frx)
                )

    if prune:
        report.pruned.extend(
            _prune_orphans(session, seen_files, state=state, dry_run=dry_run)
        )

    # The shared library is repo-authoritative, so extracting a target is also
    # the moment to make sure the target is running the repo's copy. Pushing
    # every shared file (not just the ones already present) means a target that
    # has just opted in picks the library up here.
    if shared and not extract_shared and not dry_run:
        # force=True is essential, not an optimisation to remove: the hash
        # cache tracks the *file*, so an untouched shared file looks unchanged
        # even when someone edited the component in the VBE. Skipping it there
        # would leave Excel's version running and quietly lose the guarantee
        # that the repo copy wins.
        report.pushed = inject_target(
            session, state=state, only=list(shared.values()), force=True
        )

    return report


def _folder_label(session: TargetSession, path: Path) -> str:
    """Name of the src_dir folder ``path`` lives in, for reporting."""
    try:
        return path.parent.relative_to(session.config.src_root).as_posix()
    except ValueError:
        return path.parent.name


def _prune_orphans(
    session: TargetSession,
    keep: set[str],
    *,
    state: SyncState | None,
    dry_run: bool,
) -> list[str]:
    """Delete local sources with no matching component in Excel."""
    removed: list[str] = []
    if not session.src_dir.is_dir():
        return removed

    for path in sorted(session.src_dir.iterdir()):
        if not path.is_file():
            continue
        if not (is_source_file(path) or path.suffix.lower() == FORM_BINARY_EXT):
            continue
        if path.name.casefold() in keep:
            continue

        removed.append(path.name)
        if dry_run:
            continue
        with contextlib.suppress(OSError):
            path.unlink()
        if state is not None:
            state.forget(session.name, path.relative_to(session.config.src_root))

    return removed


def print_report(report: ExtractReport, *, verbose: bool = False) -> None:
    """Render an extraction report to the console."""
    console.info(f"{report.target}: {console.summarise(report.counts)}")

    if verbose:
        for name in report.written:
            console.detail(f"written   {name}")
        for name in report.unchanged:
            console.detail(f"unchanged {name}")

    for name in report.pruned:
        console.detail(f"pruned    {name}")

    for name, folder in report.shared:
        console.detail(f"shared    {name} (kept {folder}/{name})")

    for name, reason in report.skipped:
        # Empty sheet modules are normal and expected; do not alarm the user.
        if reason.startswith("empty"):
            if verbose:
                console.detail(f"skipped   {name} ({reason})")
        else:
            console.warn(f"skipped {name}: {reason}")

    if report.pushed is not None and (report.pushed.changed or verbose):
        console.detail("pushed shared library back into the target:")
        print_inject(report.pushed, verbose=verbose)
