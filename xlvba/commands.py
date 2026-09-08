"""Command handlers. One function per subcommand, all returning an exit code."""

from __future__ import annotations

import argparse
from pathlib import Path

from xlvba.core import console
from xlvba.core.config import (
    CONFIG_FILENAME,
    DEFAULT_SRC_DIR,
    ProjectConfig,
    TargetConfig,
    load_config,
)
from xlvba.core.errors import ConfigError
from xlvba.core.state import SyncState
from xlvba.core.vbproject import document_module_map, list_components
from xlvba.ops import doctor as doctor_ops
from xlvba.ops.extract import extract_target, print_report as print_extract
from xlvba.ops.init import init_project
from xlvba.ops.inject import discover_sources, inject_target, print_report as print_inject
from xlvba.ops.reload import print_reload, reload_target
from xlvba.ops.session import open_target, resolve_targets


def _load(args: argparse.Namespace) -> ProjectConfig:
    """Load config, honouring --config/--root and ad-hoc --target."""
    root = Path(args.root).resolve() if getattr(args, "root", None) else Path.cwd()
    explicit = Path(args.config).resolve() if getattr(args, "config", None) else None

    # An explicit file path as --target lets you skip xlvba.json entirely.
    raw_target = getattr(args, "target", None)
    if raw_target and _looks_like_path(raw_target):
        path = Path(raw_target)
        resolved = path if path.is_absolute() else (root / path)
        entry = TargetConfig(
            path=str(resolved.resolve()),
            mode=getattr(args, "mode", None) or "live",
        )
        entry.validate()
        cfg = ProjectConfig(
            root=root,
            src_dir=getattr(args, "src_dir", None) or DEFAULT_SRC_DIR,
            targets=[entry],
        )
        cfg.validate()
        return cfg

    cfg = load_config(explicit, root=root)
    if getattr(args, "src_dir", None):
        cfg.src_dir = args.src_dir
    if getattr(args, "launch", False):
        cfg.launch_excel = True
    return cfg


def _looks_like_path(value: str) -> bool:
    """True if --target names a file rather than a configured target."""
    suffix = Path(value).suffix.lower()
    if suffix not in (".xlsm", ".xlam", ".xlsb", ".xls"):
        return False
    # A bare filename could be either; treat it as a path only if it exists.
    p = Path(value)
    return p.is_absolute() or len(p.parts) > 1 or p.exists()


def _selected(args: argparse.Namespace, cfg: ProjectConfig) -> list[TargetConfig]:
    raw = getattr(args, "target", None)
    if raw and _looks_like_path(raw):
        return list(cfg.targets)
    return resolve_targets(cfg, raw)


# ---------------------------------------------------------------------------
# init
# ---------------------------------------------------------------------------


def cmd_init(args: argparse.Namespace) -> int:
    root = Path(args.root).resolve() if args.root else Path.cwd()
    root.mkdir(parents=True, exist_ok=True)
    init_project(
        root,
        list(args.target or []),
        src_dir=args.src_dir or DEFAULT_SRC_DIR,
        mode=args.mode or "live",
        force=args.force,
    )
    return 0


# ---------------------------------------------------------------------------
# extract
# ---------------------------------------------------------------------------


def cmd_extract(args: argparse.Namespace) -> int:
    cfg = _load(args)
    state = SyncState.load(cfg.root)
    failed = False

    for target in _selected(args, cfg):
        with open_target(
            cfg, target, read_only=False, quiet=True
        ) as session:
            report = extract_target(
                session,
                state=state,
                prune=args.prune,
                include_empty=not args.skip_empty,
                dry_run=args.dry_run,
            )
        print_extract(report, verbose=args.verbose)
        if report.skipped and any(
            not reason.startswith("empty") for _, reason in report.skipped
        ):
            failed = True

    if not args.dry_run:
        state.save()
    return 1 if failed else 0


# ---------------------------------------------------------------------------
# inject
# ---------------------------------------------------------------------------


def cmd_inject(args: argparse.Namespace) -> int:
    cfg = _load(args)
    state = SyncState.load(cfg.root)
    ok = True

    only = [Path(f).resolve() for f in (args.file or [])] or None

    for target in _selected(args, cfg):
        with open_target(cfg, target, quiet=True) as session:
            report = inject_target(
                session,
                state=state,
                only=only,
                force=args.force,
                save=None if args.no_save else True,
                sync_deletions=args.sync_deletions or None,
                dry_run=args.dry_run,
            )
        print_inject(report, verbose=args.verbose)
        ok = ok and report.ok

    if not args.dry_run:
        state.save()
    return 0 if ok else 1


# ---------------------------------------------------------------------------
# reload
# ---------------------------------------------------------------------------


def cmd_reload(args: argparse.Namespace) -> int:
    cfg = _load(args)
    state = SyncState.load(cfg.root)
    ok = True

    for target in _selected(args, cfg):
        report = reload_target(
            cfg,
            target,
            mode=args.mode,
            state=state,
            force=args.force,
        )
        print_reload(report, verbose=args.verbose)
        ok = ok and report.ok

    state.save()
    return 0 if ok else 1


# ---------------------------------------------------------------------------
# watch
# ---------------------------------------------------------------------------


def cmd_watch(args: argparse.Namespace) -> int:
    from xlvba.watch import watch

    cfg = _load(args)
    if args.debounce is not None:
        cfg.debounce_ms = args.debounce
    if args.sync_deletions:
        cfg.sync_deletions = True

    return watch(
        cfg,
        targets=_selected(args, cfg),
        mode=args.mode,
        verbose=args.verbose,
        initial_sync=args.initial_sync,
    )


# ---------------------------------------------------------------------------
# list
# ---------------------------------------------------------------------------


def cmd_list(args: argparse.Namespace) -> int:
    cfg = _load(args)

    for target in _selected(args, cfg):
        console.heading(f"{target.resolved_name()} ({cfg.target_path(target).name})")

        with open_target(cfg, target, quiet=True) as session:
            components = list_components(session.project, exclude=session.exclude)
            doc_names = document_module_map(session.workbook)
            # Keep original casing for display, casefolded keys for matching.
            local_files = {p.name.casefold(): p.name for p in discover_sources(session)}
            local = set(local_files)

            rows: list[tuple[str, ...]] = []
            for info in components:
                note = doc_names.get(info.name, "")
                rows.append(
                    (
                        info.ref.filename,
                        info.ref.label,
                        str(info.line_count),
                        "yes" if info.ref.filename.casefold() in local else "no",
                        note,
                    )
                )

            excel_names = {c.ref.filename.casefold() for c in components}
            for key in sorted(local - excel_names):
                rows.append(
                    (local_files[key], "local only", "-", "yes", "will be created")
                )

        console.info(
            console.format_table(
                rows, ("file", "type", "lines", "local", "notes")
            )
            or "  (no components)"
        )

    return 0


# ---------------------------------------------------------------------------
# doctor
# ---------------------------------------------------------------------------


def cmd_doctor(args: argparse.Namespace) -> int:
    cfg: ProjectConfig | None
    try:
        cfg = _load(args)
    except ConfigError:
        cfg = None

    checks, healthy = doctor_ops.run_doctor(cfg)
    doctor_ops.print_checks(checks)

    if healthy:
        console.line("\nAll critical checks passed.")
        return 0
    console.line("\nOne or more critical checks failed; see remedies above.", "error")
    return 1
