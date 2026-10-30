"""Content-hash cache so the watcher can skip no-op saves.

VS Code (and most editors) fire multiple filesystem events for a single Ctrl+S:
a truncate, one or more writes, and often an atomic rename. Debouncing alone
still leaves us re-injecting identical bytes, which needlessly resets the VBA
project. Hashing the content we last pushed makes the sync idempotent.
"""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
from typing import Any

STATE_FILENAME = ".xlvbesync-state.json"


def hash_text(text: str) -> str:
    """Stable hash of source text, insensitive to line-ending style.

    Line endings are normalised first because the VBE always hands code back
    with CRLF, while the repo copy may have been written with LF by another
    tool. Without normalising, every file would look permanently dirty.
    """
    unified = text.replace("\r\n", "\n").replace("\r", "\n")
    return hashlib.sha256(unified.encode("utf-8")).hexdigest()


def hash_bytes(data: bytes) -> str:
    """Stable hash of binary content (used for .frx form resources)."""
    return hashlib.sha256(data).hexdigest()


def hash_file(path: Path) -> str | None:
    """Hash a file on disk, or ``None`` if it cannot be read."""
    try:
        raw = path.read_bytes()
    except OSError:
        return None
    if path.suffix.lower() == ".frx":
        return hash_bytes(raw)
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        return hash_bytes(raw)
    return hash_text(text)


class SyncState:
    """Persistent map of ``target -> relative path -> content hash``."""

    def __init__(self, path: Path, data: dict[str, Any] | None = None) -> None:
        self.path = path
        self._data: dict[str, dict[str, str]] = {}
        if data:
            for target, files in data.items():
                if isinstance(files, dict):
                    self._data[target] = {
                        str(k): str(v) for k, v in files.items() if isinstance(v, str)
                    }

    # -- persistence ------------------------------------------------------

    @classmethod
    def load(cls, root: Path) -> SyncState:
        path = root / STATE_FILENAME
        if not path.is_file():
            return cls(path)
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            # A corrupt cache is not fatal; worst case we re-inject once.
            return cls(path)
        if not isinstance(data, dict):
            return cls(path)
        return cls(path, data.get("files") if "files" in data else data)

    def save(self) -> None:
        """Persist the cache, tolerating a read-only or locked filesystem."""
        payload = {"version": 1, "files": self._data}
        tmp = self.path.with_suffix(self.path.suffix + ".tmp")
        try:
            tmp.write_text(
                json.dumps(payload, indent=1, ensure_ascii=False),
                encoding="utf-8",
            )
            os.replace(tmp, self.path)
        except OSError:
            with __import__("contextlib").suppress(OSError):
                tmp.unlink()

    # -- queries ----------------------------------------------------------

    @staticmethod
    def _key(rel: Path | str) -> str:
        return str(rel).replace("\\", "/").casefold()

    def get(self, target: str, rel: Path | str) -> str | None:
        return self._data.get(target.casefold(), {}).get(self._key(rel))

    def is_unchanged(self, target: str, rel: Path | str, digest: str | None) -> bool:
        """True if ``digest`` matches what we last synced for this file."""
        if digest is None:
            return False
        return self.get(target, rel) == digest

    def record(self, target: str, rel: Path | str, digest: str | None) -> None:
        if digest is None:
            return
        self._data.setdefault(target.casefold(), {})[self._key(rel)] = digest

    def forget(self, target: str, rel: Path | str) -> None:
        self._data.get(target.casefold(), {}).pop(self._key(rel), None)

    def forget_target(self, target: str) -> None:
        self._data.pop(target.casefold(), None)

    def tracked(self, target: str) -> set[str]:
        return set(self._data.get(target.casefold(), {}))

    def clear(self) -> None:
        self._data.clear()
