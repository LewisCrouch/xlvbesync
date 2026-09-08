"""Read/enumerate/export operations against a workbook's VBProject."""

from __future__ import annotations

import contextlib
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterator

from xlvba.core.codec import (
    DEFAULT_ANSI_ENCODING,
    convert_export_to_utf8,
    decode_ansi,
    normalise_newlines,
    write_text,
)
from xlvba.core.components import (
    FORM_BINARY_EXT,
    ComponentRef,
    ext_for_type,
    is_document_module,
)
from xlvba.core.errors import (
    ProjectProtectedError,
    VBOMAccessDeniedError,
    com_description,
    is_vbom_denied,
)
from xlvba.core.retry import com_retry

#: vbext_pp_none - project is not protected.
VBEXT_PP_NONE = 0


def get_vbproject(workbook: Any) -> Any:
    """Return the workbook's VBProject, translating the two failure modes.

    Excel reports a *trust* failure and a *password protection* failure with
    similar-looking automation errors, but the fixes are completely different,
    so we disambiguate here rather than leaving the user guessing.
    """
    try:
        project = workbook.VBProject
    except Exception as exc:  # noqa: BLE001
        if is_vbom_denied(exc):
            raise VBOMAccessDeniedError(
                f"Excel denied access to the VBA project of "
                f"{_safe_name(workbook)}: {com_description(exc)}"
            ) from exc
        raise

    try:
        if int(project.Protection) != VBEXT_PP_NONE:
            raise ProjectProtectedError(
                f"The VBA project in {_safe_name(workbook)} is locked for viewing."
            )
    except ProjectProtectedError:
        raise
    except Exception:  # noqa: BLE001 - Protection unreadable; press on
        pass

    return project


def _safe_name(workbook: Any) -> str:
    try:
        return str(workbook.Name)
    except Exception:  # noqa: BLE001
        return "<workbook>"


@dataclass
class ComponentInfo:
    """A snapshot of one VBComponent."""

    ref: ComponentRef
    line_count: int
    com: Any

    @property
    def name(self) -> str:
        return self.ref.name

    @property
    def component_type(self) -> int:
        return self.ref.component_type

    @property
    def is_document(self) -> bool:
        return self.ref.is_document

    @property
    def is_empty(self) -> bool:
        return self.line_count == 0


def iter_components(project: Any) -> Iterator[Any]:
    """Yield raw VBComponent objects by index.

    Indexing rather than using a Python ``for`` over the COM collection avoids
    an invalidated enumerator if a component is added or removed mid-loop.
    """
    try:
        count = project.VBComponents.Count
    except Exception as exc:  # noqa: BLE001
        if is_vbom_denied(exc):
            raise VBOMAccessDeniedError(
                f"Cannot enumerate VBA components: {com_description(exc)}"
            ) from exc
        raise
    for i in range(1, count + 1):
        try:
            yield project.VBComponents.Item(i)
        except Exception:  # noqa: BLE001
            continue


def describe_component(component: Any) -> ComponentInfo | None:
    """Build a :class:`ComponentInfo`, or ``None`` for unsupported types."""
    try:
        name = str(component.Name)
        ctype = int(component.Type)
    except Exception:  # noqa: BLE001
        return None

    try:
        ext_for_type(ctype)
    except Exception:  # noqa: BLE001 - unsupported component type
        return None

    try:
        lines = int(component.CodeModule.CountOfLines)
    except Exception:  # noqa: BLE001 - forms may refuse CodeModule briefly
        lines = 0

    return ComponentInfo(ComponentRef(name, ctype), lines, component)


def list_components(project: Any, *, exclude: set[str] | None = None) -> list[ComponentInfo]:
    """All supported components, sorted by type then name."""
    skip = {e.casefold() for e in (exclude or set())}
    found: list[ComponentInfo] = []
    for raw in iter_components(project):
        info = describe_component(raw)
        if info is None:
            continue
        if info.name.casefold() in skip:
            continue
        found.append(info)
    found.sort(key=lambda c: (c.component_type, c.name.casefold()))
    return found


def find_component(project: Any, name: str) -> Any | None:
    """Locate a VBComponent by name, case-insensitively."""
    wanted = name.casefold()
    for raw in iter_components(project):
        try:
            if str(raw.Name).casefold() == wanted:
                return raw
        except Exception:  # noqa: BLE001
            continue
    return None


def component_exists(project: Any, name: str) -> bool:
    return find_component(project, name) is not None


# ---------------------------------------------------------------------------
# Reading code
# ---------------------------------------------------------------------------


@com_retry()
def read_code(component: Any) -> str:
    """Return a component's full code body as text (no VBE header).

    ``Lines(1, CountOfLines)`` is the only way to read code without going
    through a file export, and is what we use for document modules.
    """
    module = component.CodeModule
    count = int(module.CountOfLines)
    if count <= 0:
        return ""
    return normalise_newlines(str(module.Lines(1, count)))


@com_retry()
def export_component(component: Any, destination: Path, *, ansi_encoding: str = DEFAULT_ANSI_ENCODING) -> str:
    """Export a component to ``destination`` as UTF-8.

    The VBE writes ANSI to a temp location first; we then transcode. Userform
    binaries (.frx) are copied verbatim because they are not text.
    """
    destination.parent.mkdir(parents=True, exist_ok=True)
    ext = destination.suffix

    with tempfile.TemporaryDirectory(prefix="xlvba-export-") as tmpdir:
        staged = Path(tmpdir) / f"{component.Name}{ext}"
        component.Export(str(staged))

        text = convert_export_to_utf8(staged, destination, ansi_encoding=ansi_encoding)

        # Userforms: bring the binary resource sibling along untouched.
        staged_frx = staged.with_suffix(FORM_BINARY_EXT)
        if staged_frx.is_file():
            destination.with_suffix(FORM_BINARY_EXT).write_bytes(staged_frx.read_bytes())

    return text


def export_document_module(
    component: Any,
    destination: Path,
) -> str:
    """Write a document module's code out as UTF-8.

    Document modules *can* be exported by the VBE, but the resulting .cls
    cannot be re-imported, and the export includes a misleading class header.
    We therefore store only the code body.
    """
    text = read_code(component)
    write_text(destination, text)
    return text


def export_any(
    info: ComponentInfo,
    destination: Path,
    *,
    ansi_encoding: str = DEFAULT_ANSI_ENCODING,
) -> str:
    """Export a component, dispatching on whether it is a document module."""
    if info.is_document:
        return export_document_module(info.com, destination)
    return export_component(info.com, destination, ansi_encoding=ansi_encoding)


# ---------------------------------------------------------------------------
# Sheet / document-module mapping
# ---------------------------------------------------------------------------


def document_module_map(workbook: Any) -> dict[str, str]:
    """Map ``CodeName -> friendly description`` for the workbook's documents.

    Useful for reporting, since ``Sheet1`` tells the user nothing about which
    tab it is.
    """
    mapping: dict[str, str] = {}
    with contextlib.suppress(Exception):
        mapping[str(workbook.CodeName)] = "ThisWorkbook"
    with contextlib.suppress(Exception):
        for i in range(1, int(workbook.Sheets.Count) + 1):
            sheet = workbook.Sheets.Item(i)
            with contextlib.suppress(Exception):
                mapping[str(sheet.CodeName)] = f"Sheet: {sheet.Name}"
    return mapping


def read_export_text(path: Path, ansi_encoding: str = DEFAULT_ANSI_ENCODING) -> str:
    """Read a VBE export file that may be ANSI or UTF-8."""
    return normalise_newlines(decode_ansi(path.read_bytes(), ansi_encoding))
