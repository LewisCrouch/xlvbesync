"""Write operations against a VBProject: import, replace, create, remove.

Replacement strategy
--------------------
For **standard modules, classes and userforms** the reliable approach is
remove-then-import: the VBE's ``Import`` always *adds* a component, renaming to
``Module1_1`` if the name is taken, so we must delete the old one first. We
snapshot the original code before removing it, and restore that snapshot if the
import fails, so a syntax-error-free-but-COM-rejected import cannot destroy
work.

For **document modules** (sheets, ThisWorkbook) removal is impossible. Those
are updated in place with ``DeleteLines`` + ``AddFromString``.
"""

from __future__ import annotations

import contextlib
import tempfile
from pathlib import Path
from typing import Any

from xlvbesync.core.codec import (
    DEFAULT_ANSI_ENCODING,
    encode_ansi,
    ensure_vb_name,
    normalise_newlines,
    strip_attributes,
)
from xlvbesync.core.components import (
    FORM_BINARY_EXT,
    ComponentRef,
    is_document_module,
    validate_vba_name,
)
from xlvbesync.core.errors import XlvbesyncError, com_description
from xlvbesync.core.retry import com_retry
from xlvbesync.core.vbproject import find_component, read_code


class InjectionResult:
    """Outcome of a single component sync."""

    __slots__ = ("ref", "action", "detail", "lossy")

    def __init__(
        self,
        ref: ComponentRef,
        action: str,
        detail: str = "",
        lossy: list[tuple[int, str]] | None = None,
    ) -> None:
        self.ref = ref
        self.action = action  # created | updated | unchanged | removed | skipped
        self.detail = detail
        self.lossy = lossy or []

    def __str__(self) -> str:  # pragma: no cover - display only
        suffix = f" ({self.detail})" if self.detail else ""
        return f"{self.action:<9} {self.ref.name}{suffix}"


# ---------------------------------------------------------------------------
# Document modules (sheets / ThisWorkbook)
# ---------------------------------------------------------------------------


@com_retry()
def replace_document_code(component: Any, text: str) -> bool:
    """Overwrite a document module's code in place.

    Returns True if the code actually changed. The header (if the file still
    carries one from an old export) is stripped, because inserting
    ``Attribute VB_Name`` as code is a compile error.
    """
    body = normalise_newlines(strip_attributes(text) or text).strip("\r\n")
    module = component.CodeModule

    current = ""
    count = int(module.CountOfLines)
    if count > 0:
        current = normalise_newlines(str(module.Lines(1, count))).strip("\r\n")

    if current == body:
        return False

    if count > 0:
        module.DeleteLines(1, count)
    if body:
        module.AddFromString(body)
    return True


# ---------------------------------------------------------------------------
# Importable components (.bas / .cls / .frm)
# ---------------------------------------------------------------------------


@com_retry()
def _import_file(project: Any, path: Path) -> Any:
    """Import a staged ANSI file and return the new VBComponent."""
    return project.VBComponents.Import(str(path))


@com_retry()
def _remove_component(project: Any, component: Any) -> None:
    project.VBComponents.Remove(component)


def _stage_ansi(
    text: str,
    name: str,
    ext: str,
    tmpdir: Path,
    ansi_encoding: str,
    frx_source: Path | None,
) -> tuple[Path, list[tuple[int, str]]]:
    """Write ``text`` as an ANSI file the VBE can import."""
    # The VB_Name attribute is what actually names the imported component, so
    # force it to agree with the filename - the filename is authoritative.
    prepared = ensure_vb_name(text, name) if ext in (".bas", ".cls", ".frm") else text
    data, lossy = encode_ansi(prepared, ansi_encoding)

    staged = tmpdir / f"{name}{ext}"
    staged.write_bytes(data)

    # Userforms need their binary sibling next to the .frm at import time.
    if ext == ".frm" and frx_source is not None and frx_source.is_file():
        (tmpdir / f"{name}{FORM_BINARY_EXT}").write_bytes(frx_source.read_bytes())

    return staged, lossy


def inject_component(
    project: Any,
    ref: ComponentRef,
    text: str,
    *,
    ansi_encoding: str = DEFAULT_ANSI_ENCODING,
    frx_source: Path | None = None,
) -> InjectionResult:
    """Create or replace one component in ``project`` from ``text``.

    Document modules are line-synced; everything else is remove-then-import
    with rollback on failure.
    """
    validate_vba_name(ref.name)
    existing = find_component(project, ref.name)

    # --- document modules: in-place line sync ---------------------------
    if ref.is_document:
        if existing is None:
            return InjectionResult(
                ref,
                "skipped",
                "no matching sheet/ThisWorkbook module in the target "
                "(document modules cannot be created)",
            )
        changed = replace_document_code(existing, text)
        return InjectionResult(ref, "updated" if changed else "unchanged")

    # A file whose extension says .bas but which exists in Excel as a sheet
    # module would be destroyed by remove-then-import; refuse instead.
    if existing is not None:
        with contextlib.suppress(Exception):
            if is_document_module(int(existing.Type)) != ref.is_document:
                return InjectionResult(
                    ref,
                    "skipped",
                    f"type mismatch: exists in Excel as a document module but "
                    f"the local file is {ref.ext}",
                )

    backup: str | None = None
    if existing is not None:
        with contextlib.suppress(Exception):
            backup = read_code(existing)

    with tempfile.TemporaryDirectory(prefix="xlvbesync-import-") as raw_tmp:
        tmpdir = Path(raw_tmp)
        staged, lossy = _stage_ansi(
            text, ref.name, ref.ext, tmpdir, ansi_encoding, frx_source
        )

        if existing is not None:
            _remove_component(project, existing)

        try:
            _import_file(project, staged)
        except Exception as exc:  # noqa: BLE001
            # Roll the old code back so a rejected import is not destructive.
            if backup is not None:
                with contextlib.suppress(Exception):
                    restored = project.VBComponents.Add(ref.component_type)
                    restored.Name = ref.name
                    restored.CodeModule.AddFromString(strip_attributes(backup) or backup)
            raise XlvbesyncError(
                f"Failed to import {ref.name}{ref.ext}: {com_description(exc)}",
                "Check the file for VBA syntax errors, and make sure no code "
                "from this project is currently running.",
            ) from exc

    action = "updated" if existing is not None else "created"
    return InjectionResult(ref, action, lossy=lossy)


def remove_component(project: Any, name: str) -> InjectionResult:
    """Delete a component from the project by name."""
    component = find_component(project, name)
    if component is None:
        return InjectionResult(ComponentRef(name, 1), "unchanged", "not present")

    try:
        ctype = int(component.Type)
    except Exception:  # noqa: BLE001
        ctype = 1

    ref = ComponentRef(name, ctype)
    if ref.is_document:
        return InjectionResult(ref, "skipped", "document modules cannot be removed")

    _remove_component(project, component)
    return InjectionResult(ref, "removed")


@com_retry()
def create_empty_component(project: Any, ref: ComponentRef) -> Any:
    """Add a new, empty component of the requested type."""
    validate_vba_name(ref.name)
    component = project.VBComponents.Add(ref.component_type)
    component.Name = ref.name
    return component
