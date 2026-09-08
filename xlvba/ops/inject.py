"""Inject repo source files into a workbook's VBProject."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from xlvba.core import console
from xlvba.core.codec import read_text
from xlvba.core.components import (
    FORM_BINARY_EXT,
    ComponentRef,
    is_source_file,
)
from xlvba.core.errors import XlvbaError
from xlvba.core.mutate import InjectionResult, inject_component, remove_component
from xlvba.core.state import SyncState, hash_file, hash_text
from xlvba.core.vbproject import list_components
from xlvba.ops.session import TargetSession


@dataclass
class InjectReport:
    """What an injection did."""

    target: str
    results: list[InjectionResult] = field(default_factory=list)
    failures: list[tuple[str, str]] = field(default_factory=list)
    saved: bool = False

    @property
    def counts(self) -> dict[str, int]:
        out: dict[str, int] = {}
        for r in self.results:
            out[r.action] = out.get(r.action, 0) + 1
        if self.failures:
            out["failed"] = len(self.failures)
        return out

    @property
    def changed(self) -> bool:
        return any(r.action in ("created", "updated", "removed") for r in self.results)

    @property
    def ok(self) -> bool:
        return not self.failures


def discover_sources(session: TargetSession) -> list[Path]:
    """All managed source files in the target's folder, sorted.

    Standard modules and classes are injected before userforms, and document
    modules last: a form's code may reference helpers, and importing in
    dependency-friendly order avoids spurious compile errors mid-sync.
    """
    if not session.src_dir.is_dir():
        return []

    order = {".bas": 0, ".cls": 1, ".frm": 2, ".dsr": 3, ".dcm": 4}
    files = [
        p
        for p in session.src_dir.rglob("*")
        if p.is_file() and is_source_file(p) and not _ignored(session, p)
    ]
    files.sort(key=lambda p: (order.get(p.suffix.lower(), 9), p.stem.casefold()))
    return files


def _ignored(session: TargetSession, path: Path) -> bool:
    """True if the file matches a configured ignore pattern."""
    try:
        rel = path.relative_to(session.config.root).as_posix()
    except ValueError:
        rel = path.name
    from fnmatch import fnmatch

    for pattern in session.config.ignore:
        if fnmatch(rel, pattern) or fnmatch(path.name, pattern):
            return True
    if path.name.startswith("~$"):
        return True
    return False


def inject_file(
    session: TargetSession,
    path: Path,
    *,
    state: SyncState | None = None,
    force: bool = False,
) -> InjectionResult:
    """Inject a single source file into the target's VBProject."""
    ref = ComponentRef.from_path(path)
    text = read_text(path)
    digest = hash_text(text)
    rel = path.relative_to(session.config.src_root)

    # Skip identical content unless forced. This is what makes the watcher
    # cheap: an editor autosave that changed nothing costs one hash.
    if not force and state is not None and state.is_unchanged(session.name, rel, digest):
        # A userform's binary sibling can change without the .frm changing.
        if ref.ext != ".frm" or _frx_unchanged(session, path, state):
            return InjectionResult(ref, "unchanged")

    frx = path.with_suffix(FORM_BINARY_EXT) if ref.ext == ".frm" else None

    result = inject_component(
        session.project,
        ref,
        text,
        ansi_encoding=session.config.ansi_encoding,
        frx_source=frx,
    )

    if state is not None and result.action in ("created", "updated", "unchanged"):
        state.record(session.name, rel, digest)
        if frx is not None and frx.is_file():
            state.record(
                session.name, frx.relative_to(session.config.src_root), hash_file(frx)
            )

    return result


def _frx_unchanged(session: TargetSession, frm: Path, state: SyncState) -> bool:
    frx = frm.with_suffix(FORM_BINARY_EXT)
    if not frx.is_file():
        return True
    rel = frx.relative_to(session.config.src_root)
    return state.is_unchanged(session.name, rel, hash_file(frx))


def inject_target(
    session: TargetSession,
    *,
    state: SyncState | None = None,
    only: list[Path] | None = None,
    force: bool = False,
    save: bool | None = None,
    sync_deletions: bool | None = None,
    dry_run: bool = False,
) -> InjectReport:
    """Inject some or all source files into the target.

    Newly created local files are detected implicitly: any file present locally
    but absent from the VBProject is *created* by
    :func:`xlvba.core.mutate.inject_component`.
    """
    report = InjectReport(target=session.name)
    files = only if only is not None else discover_sources(session)

    excluded = {e.casefold() for e in session.exclude}

    for path in files:
        if path.stem.casefold() in excluded:
            continue
        if dry_run:
            ref = ComponentRef.from_path(path)
            report.results.append(InjectionResult(ref, "would-sync"))
            continue
        try:
            report.results.append(inject_file(session, path, state=state, force=force))
        except XlvbaError as exc:
            report.failures.append((path.name, str(exc)))
        except Exception as exc:  # noqa: BLE001
            report.failures.append((path.name, console.describe_exception(exc)))

    # Optionally remove components whose local file has disappeared.
    should_prune = (
        session.config.sync_deletions if sync_deletions is None else sync_deletions
    )
    if should_prune and only is None and not dry_run:
        report.results.extend(_remove_orphans(session, files, state=state))

    # Persist so the on-disk file matches the live project.
    want_save = session.target.save_after_inject if save is None else save
    if want_save and report.changed and not dry_run:
        report.saved = session.save()

    return report


def _remove_orphans(
    session: TargetSession,
    local_files: list[Path],
    *,
    state: SyncState | None,
) -> list[InjectionResult]:
    """Delete VBProject components that no longer have a local source file."""
    local_names = {p.stem.casefold() for p in local_files}
    excluded = {e.casefold() for e in session.exclude}
    out: list[InjectionResult] = []

    for info in list_components(session.project, exclude=session.exclude):
        if info.is_document:
            continue  # cannot be removed
        key = info.name.casefold()
        if key in local_names or key in excluded:
            continue
        out.append(remove_component(session.project, info.name))
        if state is not None:
            state.forget(session.name, Path(session.name) / info.ref.filename)

    return out


def remove_source(
    session: TargetSession,
    path: Path,
    *,
    state: SyncState | None = None,
) -> InjectionResult:
    """Handle a locally deleted source file by removing its component."""
    ref = ComponentRef.from_path(path)
    result = remove_component(session.project, ref.name)
    if state is not None:
        with __import__("contextlib").suppress(ValueError):
            state.forget(session.name, path.relative_to(session.config.src_root))
    return result


def print_report(report: InjectReport, *, verbose: bool = False) -> None:
    """Render an injection report to the console."""
    suffix = " (saved)" if report.saved else ""
    console.info(f"{report.target}: {console.summarise(report.counts)}{suffix}")

    for r in report.results:
        if r.action in ("created", "removed") or verbose:
            console.detail(str(r))
        if r.lossy:
            chars = ", ".join(f"line {n}: U+{ord(c):04X}" for n, c in r.lossy[:5])
            console.warn(
                f"{r.ref.name}: {len(r.lossy)} character(s) are not "
                f"representable in {report.target}'s ANSI code page and were "
                f"replaced with '?' ({chars})"
            )
        if r.action == "skipped" and r.detail:
            console.warn(f"skipped {r.ref.name}: {r.detail}")

    for name, reason in report.failures:
        console.error(f"{name}: {reason}")
