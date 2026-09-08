"""Argument parsing and entry point for the ``xlvba`` command."""

from __future__ import annotations

import argparse
import sys

from xlvba import __version__, commands
from xlvba.core import console
from xlvba.core.config import DEFAULT_SRC_DIR, VALID_MODES
from xlvba.core.errors import XlvbaError

EPILOG = """\
examples:
  xlvba init --target MyBook.xlsm      create xlvba.json for a workbook
  xlvba extract                        pull all VBA out of every target
  xlvba extract --target Tools --prune mirror one target, deleting orphans
  xlvba inject                         push all local source back into Excel
  xlvba inject --file src/Tools/mod.bas  push a single module
  xlvba watch                          hot-reload on every save
  xlvba reload --mode swap             full add-in unload/replace/reload
  xlvba list                           show components on both sides
  xlvba doctor                         diagnose setup problems
"""


def _add_common(parser: argparse.ArgumentParser) -> None:
    """Options shared by every subcommand."""
    parser.add_argument(
        "--root",
        metavar="DIR",
        help="project root (default: current directory)",
    )
    parser.add_argument(
        "--config",
        metavar="FILE",
        help="path to xlvba.json (default: search upwards)",
    )
    parser.add_argument(
        "-v", "--verbose", action="store_true", help="show per-file detail"
    )
    parser.add_argument(
        "-q", "--quiet", action="store_true", help="only warnings and errors"
    )


def _add_target(parser: argparse.ArgumentParser, *, help_text: str) -> None:
    parser.add_argument("--target", metavar="NAME|FILE", help=help_text)
    parser.add_argument(
        "--src-dir",
        metavar="DIR",
        help=f"source root relative to project (default: {DEFAULT_SRC_DIR})",
    )
    parser.add_argument(
        "--launch",
        action="store_true",
        help="start Excel if it is not already running",
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="xlvba",
        description=(
            "Extract, inject and hot-reload VBA code for Excel .xlsm "
            "workbooks and .xlam add-ins."
        ),
        epilog=EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--version", action="version", version=f"xlvba {__version__}")

    sub = parser.add_subparsers(dest="command", metavar="<command>")

    # -- init -------------------------------------------------------------
    p_init = sub.add_parser(
        "init", help="create xlvba.json for one or more Excel files"
    )
    _add_common(p_init)
    p_init.add_argument(
        "--target",
        action="append",
        metavar="FILE",
        help="path to a .xlsm/.xlam (repeatable)",
    )
    p_init.add_argument(
        "--src-dir", metavar="DIR", help=f"source root (default: {DEFAULT_SRC_DIR})"
    )
    p_init.add_argument(
        "--mode", choices=VALID_MODES, help="default reload mode (default: live)"
    )
    p_init.add_argument(
        "--force", action="store_true", help="overwrite an existing xlvba.json"
    )
    p_init.set_defaults(func=commands.cmd_init)

    # -- extract ----------------------------------------------------------
    p_ex = sub.add_parser("extract", help="export VBA from Excel into the repo")
    _add_common(p_ex)
    _add_target(p_ex, help_text="target name or file (default: all targets)")
    p_ex.add_argument(
        "--prune",
        action="store_true",
        help="delete local files whose component no longer exists in Excel",
    )
    p_ex.add_argument(
        "--skip-empty",
        action="store_true",
        help="do not write files for components with no code",
    )
    p_ex.add_argument(
        "--dry-run", action="store_true", help="report without writing files"
    )
    p_ex.set_defaults(func=commands.cmd_extract)

    # -- inject -----------------------------------------------------------
    p_in = sub.add_parser("inject", help="push repo source into Excel")
    _add_common(p_in)
    _add_target(p_in, help_text="target name or file (default: all targets)")
    p_in.add_argument(
        "--file",
        action="append",
        metavar="PATH",
        help="inject only this file (repeatable)",
    )
    p_in.add_argument(
        "--force",
        action="store_true",
        help="inject even if the content hash is unchanged",
    )
    p_in.add_argument(
        "--no-save",
        action="store_true",
        help="do not save the workbook after injecting",
    )
    p_in.add_argument(
        "--sync-deletions",
        action="store_true",
        help="remove components that have no local source file",
    )
    p_in.add_argument(
        "--dry-run", action="store_true", help="report without changing Excel"
    )
    p_in.set_defaults(func=commands.cmd_inject)

    # -- reload -----------------------------------------------------------
    p_rl = sub.add_parser("reload", help="inject and hot-reload (live or swap)")
    _add_common(p_rl)
    _add_target(p_rl, help_text="target name or file (default: all targets)")
    p_rl.add_argument(
        "--mode",
        choices=VALID_MODES,
        help=(
            "live: inject into the loaded project and save (default). "
            "swap: uninstall the add-in, replace the file on disk, reinstall."
        ),
    )
    p_rl.add_argument(
        "--force", action="store_true", help="ignore content hashes"
    )
    p_rl.set_defaults(func=commands.cmd_reload)

    # -- watch ------------------------------------------------------------
    p_w = sub.add_parser("watch", help="hot-reload whenever a source file changes")
    _add_common(p_w)
    _add_target(p_w, help_text="target name or file (default: all targets)")
    p_w.add_argument("--mode", choices=VALID_MODES, help="override reload mode")
    p_w.add_argument(
        "--debounce",
        type=int,
        metavar="MS",
        help="coalesce window in milliseconds (default: 250)",
    )
    p_w.add_argument(
        "--sync-deletions",
        action="store_true",
        help="remove components when their local file is deleted",
    )
    p_w.add_argument(
        "--initial-sync",
        action="store_true",
        help="inject everything once at startup before watching",
    )
    p_w.set_defaults(func=commands.cmd_watch)

    # -- list -------------------------------------------------------------
    p_ls = sub.add_parser("list", help="list components in Excel and locally")
    _add_common(p_ls)
    _add_target(p_ls, help_text="target name or file (default: all targets)")
    p_ls.set_defaults(func=commands.cmd_list)

    # -- doctor -----------------------------------------------------------
    p_dr = sub.add_parser("doctor", help="check prerequisites and diagnose problems")
    _add_common(p_dr)
    _add_target(p_dr, help_text="target name or file (default: all targets)")
    p_dr.set_defaults(func=commands.cmd_doctor)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if not getattr(args, "command", None):
        parser.print_help()
        return 0

    console.configure(
        verbose=getattr(args, "verbose", False),
        quiet=getattr(args, "quiet", False),
        timestamps=args.command == "watch",
    )

    try:
        return int(args.func(args))
    except XlvbaError as exc:
        console.error(str(exc))
        return 2
    except KeyboardInterrupt:
        console.info("\ninterrupted")
        return 130
    except Exception as exc:  # noqa: BLE001 - top-level guard
        console.error(console.describe_exception(exc))
        if getattr(args, "verbose", False):
            import traceback

            traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
