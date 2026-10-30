"""VBComponent type metadata and filename <-> component mapping."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from xlvbesync.core.errors import ComponentTypeError

# ---------------------------------------------------------------------------
# vbext_ComponentType constants (from the VBIDE type library)
# ---------------------------------------------------------------------------

VBEXT_CT_STD_MODULE = 1
VBEXT_CT_CLASS_MODULE = 2
VBEXT_CT_MS_FORM = 3
VBEXT_CT_ACTIVEX_DESIGNER = 11
VBEXT_CT_DOCUMENT = 100

#: Extension used for sheet / ThisWorkbook code. These cannot be imported, so
#: we deliberately use a non-standard extension: it signals to the user (and to
#: our own injector) that the file is line-synced rather than replaced. ".cls"
#: would be ambiguous with real class modules.
DOC_MODULE_EXT = ".dcm"

EXT_BY_TYPE: dict[int, str] = {
    VBEXT_CT_STD_MODULE: ".bas",
    VBEXT_CT_CLASS_MODULE: ".cls",
    VBEXT_CT_MS_FORM: ".frm",
    VBEXT_CT_ACTIVEX_DESIGNER: ".dsr",
    VBEXT_CT_DOCUMENT: DOC_MODULE_EXT,
}

TYPE_BY_EXT: dict[str, int] = {
    ".bas": VBEXT_CT_STD_MODULE,
    ".cls": VBEXT_CT_CLASS_MODULE,
    ".frm": VBEXT_CT_MS_FORM,
    ".dsr": VBEXT_CT_ACTIVEX_DESIGNER,
    DOC_MODULE_EXT: VBEXT_CT_DOCUMENT,
}

TYPE_LABELS: dict[int, str] = {
    VBEXT_CT_STD_MODULE: "standard module",
    VBEXT_CT_CLASS_MODULE: "class module",
    VBEXT_CT_MS_FORM: "userform",
    VBEXT_CT_ACTIVEX_DESIGNER: "ActiveX designer",
    VBEXT_CT_DOCUMENT: "document module",
}

#: Extensions the watcher and injector treat as VBA source.
SOURCE_EXTENSIONS: frozenset[str] = frozenset(TYPE_BY_EXT)

#: Userforms carry a binary sibling holding the form layout and any embedded
#: resources. It must travel with the .frm or the form loses its controls.
FORM_BINARY_EXT = ".frx"


def ext_for_type(component_type: int) -> str:
    """Return the file extension used for a VBComponent type."""
    try:
        return EXT_BY_TYPE[int(component_type)]
    except (KeyError, TypeError, ValueError) as exc:
        raise ComponentTypeError(
            f"Unsupported VBComponent type: {component_type!r}",
            "Only standard modules, class modules, userforms and document "
            "modules are supported.",
        ) from exc


def type_for_ext(ext: str) -> int:
    """Return the VBComponent type implied by a file extension."""
    key = ext.lower()
    if not key.startswith("."):
        key = f".{key}"
    try:
        return TYPE_BY_EXT[key]
    except KeyError as exc:
        raise ComponentTypeError(
            f"Unsupported source extension: {ext!r}",
            f"Supported extensions: {', '.join(sorted(SOURCE_EXTENSIONS))}",
        ) from exc


def label_for_type(component_type: int) -> str:
    """Human-readable name for a component type."""
    return TYPE_LABELS.get(int(component_type), f"type {component_type}")


def is_document_module(component_type: int) -> bool:
    """True for sheet and ThisWorkbook modules."""
    return int(component_type) == VBEXT_CT_DOCUMENT


def is_source_file(path: Path) -> bool:
    """True if ``path`` looks like a VBA source file we manage."""
    return path.suffix.lower() in SOURCE_EXTENSIONS


@dataclass(frozen=True)
class ComponentRef:
    """A VBA component identified by name and type.

    ``name`` is the VBA identifier (``Attribute VB_Name``), which for document
    modules is the sheet's *CodeName* rather than its tab caption.
    """

    name: str
    component_type: int

    @property
    def ext(self) -> str:
        return ext_for_type(self.component_type)

    @property
    def filename(self) -> str:
        return f"{self.name}{self.ext}"

    @property
    def label(self) -> str:
        return label_for_type(self.component_type)

    @property
    def is_document(self) -> bool:
        return is_document_module(self.component_type)

    @classmethod
    def from_path(cls, path: Path) -> ComponentRef:
        """Derive a component reference from a source filename."""
        return cls(name=path.stem, component_type=type_for_ext(path.suffix))

    def __str__(self) -> str:  # pragma: no cover - display only
        return f"{self.name} ({self.label})"


def validate_vba_name(name: str) -> None:
    """Raise if ``name`` is not a legal VBA module identifier.

    VBA identifiers: start with a letter, then letters/digits/underscore, max
    31 characters. Getting this wrong produces a uselessly generic COM error at
    import time, so we check up front.
    """
    if not name:
        raise ComponentTypeError("Component name is empty.")
    if len(name) > 31:
        raise ComponentTypeError(
            f"Component name {name!r} is {len(name)} characters long.",
            "VBA module names are limited to 31 characters.",
        )
    if not name[0].isalpha():
        raise ComponentTypeError(
            f"Component name {name!r} does not start with a letter.",
            "VBA module names must begin with a letter.",
        )
    bad = [ch for ch in name if not (ch.isalnum() or ch == "_")]
    if bad:
        raise ComponentTypeError(
            f"Component name {name!r} contains invalid character(s): {''.join(sorted(set(bad)))}",
            "VBA module names may only contain letters, digits and underscores.",
        )
