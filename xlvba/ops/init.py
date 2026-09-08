"""Project bootstrap: create ``xlvba.json`` for a target."""

from __future__ import annotations

from pathlib import Path

from xlvba.core import console
from xlvba.core.addin import is_addin
from xlvba.core.config import (
    CONFIG_FILENAME,
    MODE_LIVE,
    DEFAULT_SRC_DIR,
    ProjectConfig,
    TargetConfig,
    load_config,
    save_config,
)
from xlvba.core.errors import ConfigError, TargetNotFoundError


def init_project(
    root: Path,
    target_paths: list[str],
    *,
    src_dir: str = DEFAULT_SRC_DIR,
    mode: str = MODE_LIVE,
    force: bool = False,
) -> Path:
    """Create or extend ``xlvba.json`` in ``root``.

    Relative target paths are stored relative to the project root so the config
    stays portable between machines.
    """
    config_path = root / CONFIG_FILENAME

    if config_path.exists() and not force:
        existing = load_config(config_path)
        added = 0
        for raw in target_paths:
            entry = _make_target(root, raw, mode)
            if existing.find_target(Path(entry.path).name) is not None:
                console.detail(f"already configured: {Path(entry.path).name}")
                continue
            existing.targets.append(entry)
            # Resolve clashes against targets already in the file.
            _disambiguate(existing.targets)
            added += 1
        if not added:
            console.info(f"{CONFIG_FILENAME} unchanged")
            return config_path
        existing.validate()
        save_config(existing, config_path)
        console.info(f"updated {CONFIG_FILENAME} (+{added} target)")
        return config_path

    if not target_paths:
        raise ConfigError(
            "At least one --target is required to initialise a project.",
            "Example: xlvba init --target MyBook.xlsm",
        )

    entries = [_make_target(root, raw, mode) for raw in target_paths]
    _disambiguate(entries)

    config = ProjectConfig(
        root=root,
        src_dir=src_dir,
        targets=entries,
    )
    config.validate()
    save_config(config, config_path)
    console.info(f"created {config_path}")
    return config_path


def _disambiguate(entries: list[TargetConfig]) -> None:
    """Give distinct source-folder names to targets sharing a filename stem.

    ``Book.xlsm`` alongside ``Book.xlam`` is a very common pairing (a workbook
    and the add-in built from it), and both would default to the folder name
    ``Book``. We suffix with the extension instead of failing.
    """
    counts: dict[str, int] = {}
    for entry in entries:
        key = entry.resolved_name().casefold()
        counts[key] = counts.get(key, 0) + 1

    for entry in entries:
        if counts[entry.resolved_name().casefold()] > 1 and not entry.name:
            suffix = Path(entry.path).suffix.lstrip(".").lower()
            entry.name = f"{Path(entry.path).stem}-{suffix}"
            console.detail(f"named source folder '{entry.name}' to avoid a clash")


def _make_target(root: Path, raw: str, mode: str) -> TargetConfig:
    """Build a target entry, validating that the file exists."""
    path = Path(raw)
    resolved = path if path.is_absolute() else (root / path)
    resolved = resolved.resolve()

    if not resolved.is_file():
        raise TargetNotFoundError(
            f"Target file not found: {resolved}",
            "Create the .xlsm/.xlam first, or pass an existing path.",
        )

    # Store a relative path when the file lives inside the project.
    try:
        stored = resolved.relative_to(root.resolve()).as_posix()
    except ValueError:
        stored = str(resolved)

    # An add-in defaults to live mode too, but swap is available per-target.
    entry = TargetConfig(path=stored, mode=mode)
    entry.validate()

    if is_addin(resolved):
        console.detail(f"{resolved.name}: add-in detected (mode={mode})")
    return entry
