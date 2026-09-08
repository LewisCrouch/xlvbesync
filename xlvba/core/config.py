"""Project configuration (``xlvba.json``).

A project file lets you run bare ``xlvba extract`` / ``xlvba watch`` with no
arguments. Every field can still be overridden per invocation on the CLI.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any

from xlvba.core.codec import DEFAULT_ANSI_ENCODING
from xlvba.core.errors import ConfigError

CONFIG_FILENAME = "xlvba.json"

#: Injection strategies.
MODE_LIVE = "live"
MODE_SWAP = "swap"
VALID_MODES = (MODE_LIVE, MODE_SWAP)

DEFAULT_SRC_DIR = "src"
DEFAULT_DEBOUNCE_MS = 250

DEFAULT_IGNORE = (
    ".git/**",
    ".venv/**",
    "__pycache__/**",
    "**/*.frx",
    "**/~$*",
    "**/*.tmp",
)


@dataclass
class TargetConfig:
    """One Excel file managed by this project."""

    #: Path to the .xlsm/.xlam, absolute or relative to the project root.
    path: str

    #: Folder holding extracted source, relative to ``src_dir``.
    #: Defaults to the target's stem.
    name: str = ""

    #: "live" (inject into the loaded VBProject) or "swap" (unload, replace
    #: the file on disk, reload). Only meaningful for .xlam.
    mode: str = MODE_LIVE

    #: Save the workbook after a successful live injection so disk matches.
    save_after_inject: bool = True

    #: Skip these component names entirely (exact match, case-insensitive).
    exclude: list[str] = field(default_factory=list)

    def resolved_name(self) -> str:
        return self.name or Path(self.path).stem

    def validate(self) -> None:
        if not self.path:
            raise ConfigError("A target entry is missing its 'path'.")
        suffix = Path(self.path).suffix.lower()
        if suffix not in (".xlsm", ".xlam", ".xlsb", ".xls"):
            raise ConfigError(
                f"Target {self.path!r} has unsupported extension {suffix!r}.",
                "Supported: .xlsm, .xlam, .xlsb, .xls",
            )
        if self.mode not in VALID_MODES:
            raise ConfigError(
                f"Target {self.path!r} has invalid mode {self.mode!r}.",
                f"Valid modes: {', '.join(VALID_MODES)}",
            )


@dataclass
class ProjectConfig:
    """Top-level ``xlvba.json`` contents."""

    root: Path
    src_dir: str = DEFAULT_SRC_DIR
    targets: list[TargetConfig] = field(default_factory=list)
    ansi_encoding: str = DEFAULT_ANSI_ENCODING
    debounce_ms: int = DEFAULT_DEBOUNCE_MS
    ignore: list[str] = field(default_factory=lambda: list(DEFAULT_IGNORE))

    #: Launch Excel automatically if it is not already running.
    launch_excel: bool = False

    #: Delete components from Excel when their source file is deleted.
    #: Off by default: an accidental file deletion should not nuke live code.
    sync_deletions: bool = False

    # -- paths ------------------------------------------------------------

    @property
    def src_root(self) -> Path:
        return (self.root / self.src_dir).resolve()

    def target_dir(self, target: TargetConfig) -> Path:
        return self.src_root / target.resolved_name()

    def target_path(self, target: TargetConfig) -> Path:
        p = Path(target.path)
        return p.resolve() if p.is_absolute() else (self.root / p).resolve()

    # -- lookup -----------------------------------------------------------

    def find_target(self, key: str) -> TargetConfig | None:
        """Find a target by name, filename, or full path."""
        needle = key.casefold()
        for t in self.targets:
            if t.resolved_name().casefold() == needle:
                return t
            if Path(t.path).name.casefold() == needle:
                return t
            if str(self.target_path(t)).casefold() == needle:
                return t
        return None

    def target_for_source(self, path: Path) -> TargetConfig | None:
        """Which target owns the source file at ``path``?"""
        try:
            rel = path.resolve().relative_to(self.src_root)
        except ValueError:
            return None
        if not rel.parts:
            return None
        return self.find_target(rel.parts[0])

    # -- (de)serialisation -------------------------------------------------

    def validate(self) -> None:
        if not self.targets:
            raise ConfigError(
                "No targets configured.",
                f"Add at least one entry to 'targets' in {CONFIG_FILENAME}, or "
                "pass --target on the command line.",
            )
        if self.debounce_ms < 0:
            raise ConfigError("'debounce_ms' must be >= 0.")
        seen: set[str] = set()
        for t in self.targets:
            t.validate()
            key = t.resolved_name().casefold()
            if key in seen:
                raise ConfigError(
                    f"Duplicate target name {t.resolved_name()!r}.",
                    "Give one of them an explicit distinct 'name'.",
                )
            seen.add(key)

    def to_dict(self) -> dict[str, Any]:
        return {
            "src_dir": self.src_dir,
            "ansi_encoding": self.ansi_encoding,
            "debounce_ms": self.debounce_ms,
            "launch_excel": self.launch_excel,
            "sync_deletions": self.sync_deletions,
            "ignore": list(self.ignore),
            "targets": [asdict(t) for t in self.targets],
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any], root: Path) -> ProjectConfig:
        if not isinstance(data, dict):
            raise ConfigError(f"{CONFIG_FILENAME} must contain a JSON object.")

        raw_targets = data.get("targets", [])
        if isinstance(raw_targets, (str, dict)):
            raw_targets = [raw_targets]
        if not isinstance(raw_targets, list):
            raise ConfigError("'targets' must be a list.")

        targets: list[TargetConfig] = []
        for entry in raw_targets:
            if isinstance(entry, str):
                targets.append(TargetConfig(path=entry))
            elif isinstance(entry, dict):
                known = {f for f in TargetConfig.__dataclass_fields__}
                unknown = set(entry) - known
                if unknown:
                    raise ConfigError(
                        f"Unknown target key(s): {', '.join(sorted(unknown))}",
                        f"Valid keys: {', '.join(sorted(known))}",
                    )
                targets.append(TargetConfig(**entry))
            else:
                raise ConfigError(f"Invalid target entry: {entry!r}")

        cfg = cls(
            root=root,
            src_dir=str(data.get("src_dir", DEFAULT_SRC_DIR)),
            targets=targets,
            ansi_encoding=str(data.get("ansi_encoding", DEFAULT_ANSI_ENCODING)),
            debounce_ms=int(data.get("debounce_ms", DEFAULT_DEBOUNCE_MS)),
            ignore=list(data.get("ignore", DEFAULT_IGNORE)),
            launch_excel=bool(data.get("launch_excel", False)),
            sync_deletions=bool(data.get("sync_deletions", False)),
        )
        return cfg


def find_config_file(start: Path) -> Path | None:
    """Search ``start`` and its parents for ``xlvba.json``."""
    current = start.resolve()
    for candidate in [current, *current.parents]:
        cfg = candidate / CONFIG_FILENAME
        if cfg.is_file():
            return cfg
    return None


def load_config(path: Path | None = None, *, root: Path | None = None) -> ProjectConfig:
    """Load configuration, searching upwards when ``path`` is omitted."""
    cfg_path = path or find_config_file(root or Path.cwd())
    if cfg_path is None:
        raise ConfigError(
            f"No {CONFIG_FILENAME} found in this directory or any parent.",
            "Run `xlvba init --target <file.xlsm>` to create one, or pass "
            "--target explicitly.",
        )
    try:
        raw = cfg_path.read_text(encoding="utf-8-sig")
    except OSError as exc:
        raise ConfigError(f"Cannot read {cfg_path}: {exc}") from exc
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ConfigError(
            f"{cfg_path} is not valid JSON (line {exc.lineno}, column {exc.colno}): {exc.msg}"
        ) from exc

    cfg = ProjectConfig.from_dict(data, cfg_path.parent.resolve())
    cfg.validate()
    return cfg


def save_config(cfg: ProjectConfig, path: Path | None = None) -> Path:
    """Write configuration to disk as UTF-8 JSON."""
    dest = path or (cfg.root / CONFIG_FILENAME)
    dest.write_text(
        json.dumps(cfg.to_dict(), indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    return dest
