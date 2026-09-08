"""Encoding bridge between the VBE (ANSI/cp1252) and the repo (UTF-8).

Why this module exists
----------------------
``VBComponent.Export`` writes files in the host's ANSI code page (cp1252 on a
Western Windows install), and ``VBComponents.Import`` expects the same. Git,
VS Code and every other sane tool want UTF-8. So:

* on **export** we read the VBE's ANSI output and rewrite it as UTF-8;
* on **inject** we read UTF-8 from the repo and stage an ANSI temp file to
  hand to ``Import``.

Characters that genuinely cannot survive the ANSI round trip (e.g. emoji, CJK)
are reported rather than silently mangled, because a corrupted string literal
in VBA is a nasty bug to chase later.
"""

from __future__ import annotations

import re
from pathlib import Path

#: The encoding the VBE reads and writes. Technically this is the system ANSI
#: code page; cp1252 covers Western installs. Override via
#: ``xlvba.json -> "ansi_encoding"`` if you run a different locale.
DEFAULT_ANSI_ENCODING = "cp1252"

#: Repo-side encoding. Non-negotiable per project requirements.
TEXT_ENCODING = "utf-8"

#: VBA source is CRLF-terminated. The VBE normalises this internally but we
#: keep files CRLF on disk so diffs do not churn.
VBA_NEWLINE = "\r\n"

_BOM = "\ufeff"


def system_ansi_encoding() -> str:
    """Return the host ANSI code page as a Python codec name.

    Falls back to :data:`DEFAULT_ANSI_ENCODING` when it cannot be determined.
    """
    try:
        import ctypes

        cp = ctypes.windll.kernel32.GetACP()  # type: ignore[attr-defined]
    except Exception:
        return DEFAULT_ANSI_ENCODING
    if not cp:
        return DEFAULT_ANSI_ENCODING
    return f"cp{cp}"


def decode_ansi(raw: bytes, ansi_encoding: str = DEFAULT_ANSI_ENCODING) -> str:
    """Decode VBE-exported bytes to ``str``.

    Some files (notably ones previously touched by other tooling) may already
    be UTF-8 or carry a BOM, so we sniff before committing to ANSI.
    """
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig")

    # If it decodes cleanly as UTF-8 *and* contains a multi-byte sequence, it
    # is almost certainly genuine UTF-8 rather than ANSI.
    if any(b >= 0x80 for b in raw):
        try:
            text = raw.decode("utf-8")
        except UnicodeDecodeError:
            pass
        else:
            return text

    # cp1252 has undefined slots (0x81, 0x8D, 0x8F, 0x90, 0x9D); replace rather
    # than explode so a stray byte cannot block an extraction.
    return raw.decode(ansi_encoding, errors="replace")


def unencodable_chars(text: str, ansi_encoding: str = DEFAULT_ANSI_ENCODING) -> list[tuple[int, str]]:
    """Return ``(line_number, char)`` pairs that cannot be encoded to ANSI.

    Line numbers are 1-based to match editor gutters.
    """
    bad: list[tuple[int, str]] = []
    seen: set[str] = set()
    for lineno, line in enumerate(text.splitlines(), start=1):
        for ch in line:
            if ch in seen:
                continue
            try:
                ch.encode(ansi_encoding)
            except UnicodeEncodeError:
                seen.add(ch)
                bad.append((lineno, ch))
    return bad


def encode_ansi(
    text: str,
    ansi_encoding: str = DEFAULT_ANSI_ENCODING,
    *,
    strict: bool = False,
) -> tuple[bytes, list[tuple[int, str]]]:
    """Encode repo text to VBE-compatible ANSI bytes.

    Returns the bytes plus a list of characters that had to be substituted.
    With ``strict=True`` a :class:`UnicodeEncodeError` propagates instead.
    """
    lossy = unencodable_chars(text, ansi_encoding)
    if strict and lossy:
        raise UnicodeEncodeError(
            ansi_encoding,
            text,
            0,
            1,
            f"{len(lossy)} character(s) cannot be represented in {ansi_encoding}",
        )
    # "replace" emits '?' which at least keeps the file syntactically valid.
    return text.encode(ansi_encoding, errors="replace"), lossy


def normalise_newlines(text: str, newline: str = VBA_NEWLINE) -> str:
    """Force consistent line endings, collapsing CRLF/CR/LF mixes."""
    unified = text.replace("\r\n", "\n").replace("\r", "\n")
    if newline == "\n":
        return unified
    return unified.replace("\n", newline)


def strip_bom(text: str) -> str:
    """Remove a leading BOM if present."""
    return text[1:] if text.startswith(_BOM) else text


def read_text(path: Path) -> str:
    """Read a repo-side source file as UTF-8 with LF-normalised content.

    Uses ``utf-8-sig`` so a BOM written by another editor is transparently
    removed. Falls back to the ANSI code page if the file is not valid UTF-8,
    which happens when a user drops a raw VBE export into the folder by hand.
    """
    raw = path.read_bytes()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode(system_ansi_encoding(), errors="replace")
    return normalise_newlines(strip_bom(text))


def write_text(path: Path, text: str, *, newline: str = VBA_NEWLINE) -> None:
    """Write ``text`` to ``path`` as UTF-8 (no BOM) with CRLF endings."""
    path.parent.mkdir(parents=True, exist_ok=True)
    data = normalise_newlines(strip_bom(text), newline).encode(TEXT_ENCODING)
    path.write_bytes(data)


def convert_export_to_utf8(
    exported: Path,
    destination: Path,
    *,
    ansi_encoding: str = DEFAULT_ANSI_ENCODING,
) -> str:
    """Convert a VBE ANSI export in-place to a UTF-8 file at ``destination``.

    Returns the decoded text so callers can hash or inspect it without a
    second read.
    """
    text = decode_ansi(exported.read_bytes(), ansi_encoding)
    write_text(destination, text)
    return normalise_newlines(strip_bom(text))


# ---------------------------------------------------------------------------
# VBA source header handling
# ---------------------------------------------------------------------------

#: Matches the ``VERSION 1.0 CLASS`` banner that opens .cls/.frm exports.
_VERSION_RE = re.compile(r"^VERSION\s+[\d.]+\s+CLASS\s*$", re.IGNORECASE)

#: Matches a bare ``Attribute VB_Name = "Foo"`` line.
_ATTR_NAME_RE = re.compile(r'^Attribute\s+VB_Name\s*=\s*"(?P<name>[^"]*)"\s*$', re.IGNORECASE)


def split_header(text: str) -> tuple[str, str]:
    """Split a VBE export into ``(header, body)``.

    The *header* is the metadata preamble the VBE requires on import: the
    ``VERSION`` banner, the ``BEGIN...END`` block for classes/forms, and the
    ``Attribute VB_*`` lines. The *body* is the actual code.

    For a plain standard module the header is usually just the ``Attribute
    VB_Name`` line. Returning them separately lets us push only the body
    through ``CodeModule.AddFromString`` (used for document modules, where the
    header must never be inserted as code).
    """
    lines = normalise_newlines(text, "\n").split("\n")
    idx = 0
    n = len(lines)

    if idx < n and _VERSION_RE.match(lines[idx].strip()):
        idx += 1
        # Optional BEGIN ... END block (present for classes and userforms).
        if idx < n and lines[idx].strip().upper().startswith("BEGIN"):
            depth = 0
            while idx < n:
                stripped = lines[idx].strip().upper()
                if stripped.startswith("BEGIN"):
                    depth += 1
                elif stripped == "END":
                    depth -= 1
                    if depth == 0:
                        idx += 1
                        break
                idx += 1

    # Consume the contiguous run of Attribute lines that follows.
    while idx < n and lines[idx].strip().lower().startswith("attribute "):
        idx += 1

    header = VBA_NEWLINE.join(lines[:idx])
    body = VBA_NEWLINE.join(lines[idx:])
    return header, body


def extract_vb_name(text: str) -> str | None:
    """Return the ``Attribute VB_Name`` value from a VBE export, if present."""
    for line in normalise_newlines(text, "\n").split("\n")[:80]:
        match = _ATTR_NAME_RE.match(line.strip())
        if match:
            return match.group("name")
        if line.strip().lower().startswith(("sub ", "function ", "public ", "private ", "option ")):
            break
    return None


def ensure_vb_name(text: str, name: str) -> str:
    """Ensure the export text declares ``Attribute VB_Name = "<name>"``.

    Used when injecting a hand-written file whose author omitted the header, or
    where the filename is authoritative and disagrees with a stale attribute.
    """
    lines = normalise_newlines(text, "\n").split("\n")
    attr = f'Attribute VB_Name = "{name}"'

    for i, line in enumerate(lines):
        if _ATTR_NAME_RE.match(line.strip()):
            lines[i] = attr
            return VBA_NEWLINE.join(lines)

    # No VB_Name present. Insert it after any VERSION/BEGIN-END preamble.
    header, body = split_header(text)
    if header:
        return VBA_NEWLINE.join([header, attr, body]) if body else f"{header}{VBA_NEWLINE}{attr}"
    return VBA_NEWLINE.join([attr, *lines])


def strip_attributes(text: str) -> str:
    """Return only the code body, with the VBE metadata header removed."""
    return split_header(text)[1]
