"""Command handlers. One function per subcommand, all returning an exit code."""

from __future__ import annotations

import argparse
from pathlib import Path

from xlvbesync.core import console
from xlvbesync.core.components import is_source_file
from xlvbesync.core.config import (
    CONFIG_FILENAME,
    DEFAULT_SRC_DIR,
    ProjectConfig,
    TargetConfig,
    load_config,
)
from xlvbesync.core.errors import ConfigError
from xlvbesync.core.format import format_source_file
from xlvbesync.core.state import SyncState
from xlvbesync.core.vbproject import document_module_map, list_components
from xlvbesync.ops import doctor as doctor_ops
from xlvbesync.ops.extract import extract_target, print_report as print_extract
from xlvbesync.ops.init import init_project
from xlvbesync.ops.inject import discover_sources, inject_target, print_report as print_inject
from xlvbesync.ops.reload import print_reload, reload_target
from xlvbesync.ops.session import open_target, resolve_targets


def _load(args: argparse.Namespace) -> ProjectConfig:
    """Load config, honouring --config/--root and ad-hoc --target."""
    root = Path(args.root).resolve() if getattr(args, "root", None) else Path.cwd()
    explicit = Path(args.config).resolve() if getattr(args, "config", None) else None

    # An explicit file path as --target lets you skip xlvbesync.json entirely.
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


def _origin_label(cfg: ProjectConfig, path: Path) -> str:
    """Note naming the shared folder a source file came from, else empty."""
    if not cfg.is_shared_source(path):
        return ""
    try:
        folder = path.parent.resolve().relative_to(cfg.src_root).as_posix()
    except ValueError:
        return "shared"
    return f"shared: {folder}"


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
        mode=args.mode,
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
                extract_shared=args.extract_shared,
            )
        print_extract(report, verbose=args.verbose)
        if report.skipped and any(
            not reason.startswith("empty") for _, reason in report.skipped
        ):
            failed = True
        if report.pushed is not None and not report.pushed.ok:
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
    from xlvbesync.watch import watch

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
            # Shared includes are discovered before the target's own folder, so
            # an own-folder file of the same name overwrites the shared entry -
            # exactly the precedence the injector uses.
            local_files: dict[str, str] = {}
            origin: dict[str, str] = {}
            for path in discover_sources(session):
                key = path.name.casefold()
                local_files[key] = path.name
                origin[key] = _origin_label(cfg, path)
            local = set(local_files)

            rows: list[tuple[str, ...]] = []
            for info in components:
                key = info.ref.filename.casefold()
                notes = [doc_names.get(info.name, ""), origin.get(key, "")]
                rows.append(
                    (
                        info.ref.filename,
                        info.ref.label,
                        str(info.line_count),
                        "yes" if key in local else "no",
                        ", ".join(n for n in notes if n),
                    )
                )

            excel_names = {c.ref.filename.casefold() for c in components}
            for key in sorted(local - excel_names):
                notes = [n for n in ("will be created", origin[key]) if n]
                rows.append(
                    (local_files[key], "local only", "-", "yes", ", ".join(notes))
                )

        console.info(
            console.format_table(
                rows, ("file", "type", "lines", "local", "notes")
            )
            or "  (no components)"
        )

    return 0


# ---------------------------------------------------------------------------
# format
# ---------------------------------------------------------------------------


def cmd_format(args: argparse.Namespace) -> int:
    """Pretty-print local source files using the project's indent setting."""
    cfg = _load(args)
    seen: set[Path] = set()
    changed = 0
    total = 0

    for target in _selected(args, cfg):
        for folder in cfg.source_dirs(target):
            if not folder.is_dir():
                continue
            for path in folder.rglob("*"):
                if not path.is_file() or not is_source_file(path):
                    continue
                key = path.resolve()
                if key in seen:
                    continue
                seen.add(key)
                total += 1
                if format_source_file(path, indent=cfg.indent):
                    changed += 1
                    console.detail(str(path.relative_to(cfg.root)))

    console.info(f"formatted {changed} of {total} files")
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


def cmd_form_host(_args: argparse.Namespace) -> int:
    """Speak JSON with the UserForm designer. Replies are the only stdout."""
    from xlvbesync.ops.forms import serve

    return serve()
