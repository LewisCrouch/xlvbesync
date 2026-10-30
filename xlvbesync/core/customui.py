"""CustomUI (Ribbon XML and images) packaging and extraction.

Excel workbooks (.xlsm) and add-ins (.xlam) store Ribbon customizations in the
OpenXML package under ``customUI/``. This module handles syncing local customUI
source files (XML definitions, relationship files, and icons) into the zip
archive, as well as extracting them from an existing archive.
"""

from __future__ import annotations

import contextlib
import os
import shutil
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

CUSTOMUI14_REL_TYPE = (
    "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
)
CUSTOMUI_REL_TYPE = (
    "http://schemas.microsoft.com/office/2006/01/relationships/ui/extensibility"
)

CUSTOMUI_EXTENSIONS = {
    ".xml",
    ".rels",
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".ico",
    ".bmp",
    ".svg",
}

CONTENT_TYPE_MAP = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "ico": "image/x-icon",
    "bmp": "image/bmp",
    "svg": "image/svg+xml",
    "xml": "application/xml",
    "rels": "application/vnd.openxmlformats-package.relationships+xml",
}


def is_customui_file(path: Path | str) -> bool:
    """True if ``path`` looks like a Ribbon customUI source file."""
    p = Path(path)
    name = p.name
    if name.startswith("~$") or name.endswith("~") or p.suffix.lower() == ".tmp":
        return False

    suffix = p.suffix.lower()
    if suffix not in CUSTOMUI_EXTENSIONS:
        return False

    parts_lower = [part.casefold() for part in p.parts]
    return "customui" in parts_lower or name.casefold().startswith("customui")


def find_customui_dir(target_dir: Path) -> Path | None:
    """Return the customUI directory under ``target_dir`` if one exists."""
    for cand in ("customUI", "customui", "CustomUI"):
        d = target_dir / cand
        if d.is_dir():
            return d
    return None


def extract_customui(archive_path: Path, dest_dir: Path) -> list[str]:
    """Extract customUI files from an archive into ``dest_dir / 'customUI'``.

    Returns a list of extracted file names relative to ``dest_dir``.
    """
    if not archive_path.is_file():
        return []

    extracted: list[str] = []
    try:
        with zipfile.ZipFile(archive_path, "r") as z:
            for item in z.infolist():
                parts = Path(item.filename).parts
                if not parts or parts[0].casefold() != "customui":
                    continue
                if item.is_dir():
                    continue

                rel_parts = parts[1:]
                out_path = dest_dir / "customUI" / Path(*rel_parts)
                out_path.parent.mkdir(parents=True, exist_ok=True)
                with open(out_path, "wb") as f:
                    f.write(z.read(item))
                extracted.append(str(out_path.relative_to(dest_dir).as_posix()))
    except Exception:  # noqa: BLE001
        return []

    return extracted


def _ensure_customui_relationship(rels_xml: bytes, main_xml_name: str) -> bytes:
    """Ensure ``_rels/.rels`` contains a relationship for the customUI XML part."""
    try:
        ET.register_namespace("", REL_NS)
        root = ET.fromstring(rels_xml)
    except Exception:  # noqa: BLE001
        return rels_xml

    # Check if a relationship already targets any customUI file.
    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
        target = rel.attrib.get("Target", "")
        if "customui" in target.casefold():
            return rels_xml

    rel_type = (
        CUSTOMUI14_REL_TYPE
        if "14" in main_xml_name.casefold()
        else CUSTOMUI_REL_TYPE
    )

    existing_ids = {
        rel.attrib.get("Id", "") for rel in root.findall(f"{{{REL_NS}}}Relationship")
    }
    rid = "rIdCustomUI"
    counter = 1
    while rid in existing_ids:
        rid = f"rIdCustomUI{counter}"
        counter += 1

    target_path = (
        main_xml_name
        if main_xml_name.casefold().startswith("customui/")
        else f"customUI/{main_xml_name}"
    )

    new_rel = ET.Element(
        f"{{{REL_NS}}}Relationship",
        attrib={"Id": rid, "Type": rel_type, "Target": target_path},
    )
    root.append(new_rel)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _ensure_content_types(ct_xml: bytes, extensions: set[str]) -> bytes:
    """Ensure ``[Content_Types].xml`` has <Default> entries for used extensions."""
    try:
        ET.register_namespace("", CONTENT_TYPES_NS)
        root = ET.fromstring(ct_xml)
    except Exception:  # noqa: BLE001
        return ct_xml

    existing_exts = {
        d.attrib.get("Extension", "").lower()
        for d in root.findall(f"{{{CONTENT_TYPES_NS}}}Default")
    }

    for ext in extensions:
        clean_ext = ext.lstrip(".").lower()
        if clean_ext not in existing_exts and clean_ext in CONTENT_TYPE_MAP:
            new_def = ET.Element(
                f"{{{CONTENT_TYPES_NS}}}Default",
                attrib={
                    "Extension": clean_ext,
                    "ContentType": CONTENT_TYPE_MAP[clean_ext],
                },
            )
            root.append(new_def)
            existing_exts.add(clean_ext)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def sync_customui(archive_path: Path, customui_dir: Path) -> bool:
    """Sync all customUI files from ``customui_dir`` into ``archive_path``.

    Returns True if the archive was updated, False if no customUI files
    existed or the archive could not be modified.
    """
    if not archive_path.is_file() or not customui_dir.is_dir():
        return False

    files_to_sync: dict[str, Path] = {}
    extensions: set[str] = set()
    main_xml_name: str | None = None

    for p in customui_dir.rglob("*"):
        if not p.is_file() or not is_customui_file(p):
            continue
        rel = p.relative_to(customui_dir)
        if rel.parts and rel.parts[0].casefold() == "customui":
            entry_name = rel.as_posix()
        else:
            entry_name = f"customUI/{rel.as_posix()}"

        files_to_sync[entry_name] = p
        extensions.add(p.suffix.lower())

        if p.name.casefold() in ("customui14.xml", "customui.xml"):
            main_xml_name = entry_name

    if not files_to_sync:
        return False

    if main_xml_name is None:
        main_xml_name = "customUI/customUI14.xml"

    # Stage the replacement in a temporary file beside the archive.
    temp_fd, temp_path_str = tempfile.mkstemp(
        suffix=archive_path.suffix,
        prefix=f"{archive_path.stem}.xlvbesync-cui-",
        dir=archive_path.parent,
    )
    os.close(temp_fd)
    temp_path = Path(temp_path_str)

    try:
        norm_map = {k.casefold(): (k, v) for k, v in files_to_sync.items()}
        written_norm: set[str] = set()

        with (
            zipfile.ZipFile(archive_path, "r") as src_zip,
            zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as dst_zip,
        ):
            for item in src_zip.infolist():
                norm = item.filename.casefold()
                if norm in norm_map:
                    arcname, src_file = norm_map[norm]
                    dst_zip.write(src_file, arcname=arcname)
                    written_norm.add(norm)
                elif norm == "_rels/.rels":
                    data = _ensure_customui_relationship(
                        src_zip.read(item), main_xml_name
                    )
                    dst_zip.writestr(item, data)
                elif norm == "[content_types].xml":
                    data = _ensure_content_types(src_zip.read(item), extensions)
                    dst_zip.writestr(item, data)
                else:
                    dst_zip.writestr(item, src_zip.read(item))

            # Add any new files that weren't present in the source archive.
            for norm, (arcname, src_file) in norm_map.items():
                if norm not in written_norm:
                    dst_zip.write(src_file, arcname=arcname)
                    written_norm.add(norm)

        # Atomic replacement on the same volume.
        os.replace(temp_path, archive_path)
        return True
    except Exception:
        if temp_path.exists():
            with contextlib.suppress(OSError):
                temp_path.unlink()
        raise
