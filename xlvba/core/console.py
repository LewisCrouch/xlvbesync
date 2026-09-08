"""Console output that cannot crash on a legacy code page.

A Windows console often runs cp1252. Printing a character it cannot represent
raises ``UnicodeEncodeError`` *from inside print*, which would abort a sync
halfway through. Everything user-facing goes through here instead.
"""

from __future__ import annotations

import sys
import time
from typing import Any, TextIO

_LEVEL_ORDER = {"debug": 10, "info": 20, "warn": 30, "error": 40, "quiet": 100}

_state = {"level": 20, "timestamps": False}


def configure(*, verbose: bool = False, quiet: bool = False, timestamps: bool = False) -> None:
    """Set global verbosity."""
    if quiet:
        _state["level"] = _LEVEL_ORDER["warn"]
    elif verbose:
        _state["level"] = _LEVEL_ORDER["debug"]
    else:
        _state["level"] = _LEVEL_ORDER["info"]
    _state["timestamps"] = timestamps


def _sanitise(text: str, stream: TextIO) -> str:
    """Replace characters the stream's encoding cannot represent."""
    encoding = getattr(stream, "encoding", None) or "utf-8"
    try:
        text.encode(encoding)
        return text
    except (UnicodeEncodeError, LookupError):
        return text.encode(encoding, errors="backslashreplace").decode(encoding, errors="replace")


def _emit(
    level: str,
    message: str,
    stream: TextIO | None = None,
    *,
    prefix_level: bool = True,
) -> None:
    if _LEVEL_ORDER[level] < _state["level"]:
        return
    out = stream or (sys.stderr if level in ("warn", "error") else sys.stdout)
    prefix = ""
    if prefix_level and level in ("warn", "error"):
        prefix = f"{level.upper()} "
    stamp = f"{time.strftime('%H:%M:%S')} " if _state["timestamps"] else ""
    try:
        print(f"{stamp}{prefix}{_sanitise(message, out)}", file=out, flush=True)
    except Exception:  # noqa: BLE001 - never let logging break the run
        pass


def debug(message: str) -> None:
    _emit("debug", message)


def info(message: str) -> None:
    _emit("info", message)


def warn(message: str) -> None:
    _emit("warn", message)


def error(message: str) -> None:
    _emit("error", message)


def detail(message: str) -> None:
    """Indented supplementary line."""
    _emit("info", f"  {message}")


def line(message: str, level: str = "info") -> None:
    """Emit at ``level`` but always on stdout.

    Used for reports (doctor, list) where interleaving stdout and stderr would
    scramble the ordering of a check and its remedy. The level still governs
    visibility, but no WARN/ERROR prefix is added since the caller supplies
    its own status marker.
    """
    _emit(level, message, stream=sys.stdout, prefix_level=False)


def heading(message: str) -> None:
    _emit("info", f"\n{message}")


def describe_exception(exc: BaseException) -> str:
    """Render an exception for the console, including any remedy."""
    from xlvba.core.errors import XlvbaError, com_description, hresult_of

    if isinstance(exc, XlvbaError):
        return str(exc)

    hres = hresult_of(exc)
    if hres is not None:
        return f"{com_description(exc)} (HRESULT 0x{hres & 0xFFFFFFFF:08X})"
    return f"{type(exc).__name__}: {exc}"


def format_table(rows: list[tuple[str, ...]], headers: tuple[str, ...]) -> str:
    """Render a simple fixed-width table."""
    if not rows:
        return ""
    widths = [len(h) for h in headers]
    for row in rows:
        for i, cell in enumerate(row):
            if i < len(widths):
                widths[i] = max(widths[i], len(str(cell)))
    lines = [
        "  ".join(h.ljust(widths[i]) for i, h in enumerate(headers)),
        "  ".join("-" * w for w in widths),
    ]
    for row in rows:
        lines.append("  ".join(str(c).ljust(widths[i]) for i, c in enumerate(row)))
    return "\n".join(lines)


def plural(count: int, singular: str, suffix: str = "s") -> str:
    return f"{count} {singular}{'' if count == 1 else suffix}"


def summarise(counts: dict[str, int]) -> str:
    """Turn ``{'created': 2, 'updated': 1}`` into readable text."""
    parts = [f"{v} {k}" for k, v in counts.items() if v]
    return ", ".join(parts) if parts else "nothing to do"


def unused(*_args: Any) -> None:  # pragma: no cover
    return None
