"""Pretty-print VBA source without talking to Excel.

The VBE header (``VERSION`` / ``BEGIN...END`` / ``Attribute``) is left alone.
Only the code body is indented and keyword-cased. String literals, comments
and line-continuation tails are not rewritten beyond indent.
"""

from __future__ import annotations

import re
from pathlib import Path

from xlvbesync.core.codec import (
    TEXT_ENCODING,
    VBA_NEWLINE,
    normalise_newlines,
    read_text,
    split_header,
    strip_bom,
    write_text,
)

#: Default indent when the project does not set ``indent`` in xlvbesync.json.
DEFAULT_INDENT = "    "

# Language keywords and built-in types only. Functions that are commonly
# used as names (Left, Right, Name, ...) are left to the author.
_KEYWORDS = (
    "AddressOf",
    "Alias",
    "And",
    "Any",
    "As",
    "Attribute",
    "Base",
    "Binary",
    "Boolean",
    "ByRef",
    "ByVal",
    "Byte",
    "Call",
    "Case",
    "Compare",
    "Const",
    "Currency",
    "Database",
    "Date",
    "Debug",
    "Declare",
    "Decimal",
    "Dim",
    "Do",
    "Double",
    "Each",
    "Else",
    "ElseIf",
    "Empty",
    "End",
    "Enum",
    "Eqv",
    "Error",
    "Event",
    "Exit",
    "Explicit",
    "False",
    "For",
    "Friend",
    "Function",
    "Get",
    "Global",
    "GoSub",
    "GoTo",
    "If",
    "Imp",
    "Implements",
    "In",
    "Integer",
    "Is",
    "Let",
    "Lib",
    "Like",
    "Long",
    "LongLong",
    "LongPtr",
    "Loop",
    "Me",
    "Mod",
    "Module",
    "New",
    "Next",
    "Not",
    "Nothing",
    "Null",
    "Object",
    "On",
    "Option",
    "Optional",
    "Or",
    "ParamArray",
    "Preserve",
    "Print",
    "Private",
    "Property",
    "PtrSafe",
    "Public",
    "RaiseEvent",
    "ReDim",
    "Rem",
    "Resume",
    "Select",
    "Set",
    "Single",
    "Static",
    "Step",
    "Stop",
    "String",
    "Sub",
    "Then",
    "To",
    "True",
    "Type",
    "Until",
    "Variant",
    "Wend",
    "While",
    "With",
    "WithEvents",
    "Xor",
)

_KEYWORD_CASE = {k.lower(): k for k in _KEYWORDS}

_PREPROCESSOR_CASE = {
    "#if": "#If",
    "#elseif": "#ElseIf",
    "#else": "#Else",
    "#end": "#End",
    "#const": "#Const",
}

_ACCESS = re.compile(r"^(Public|Private|Friend|Static)\s+", re.IGNORECASE)
_END_BLOCK = re.compile(
    r"^End\s+(Sub|Function|Property|If|With|Select|Type|Enum)\b",
    re.IGNORECASE,
)
_PROPERTY = re.compile(r"^Property\s+(Get|Let|Set)\b", re.IGNORECASE)
_PROC = re.compile(r"^(Sub|Function)\b", re.IGNORECASE)
_TYPE_BLOCK = re.compile(r"^Type\s+\w+", re.IGNORECASE)
_ENUM_BLOCK = re.compile(r"^Enum\s+\w+", re.IGNORECASE)
_LABEL = re.compile(r"^([A-Za-z][A-Za-z0-9_]*)\s*:\s*$")
_LABEL_SKIP = frozenset({"else", "case"})


def format_vba(text: str, *, indent: str | None = None) -> str:
    """Return ``text`` pretty-printed with CRLF endings and no BOM.

    ``indent`` is the unit repeated per indent level (``"\\t"`` or spaces).
    """
    unit = DEFAULT_INDENT if indent is None else indent
    header, body = split_header(normalise_newlines(strip_bom(text), VBA_NEWLINE))
    formatted_body = _format_body(body, unit)
    if header and formatted_body:
        result = header + VBA_NEWLINE + formatted_body
    elif header:
        result = header
    else:
        result = formatted_body
    if result and not result.endswith(("\r\n", "\n")):
        result += VBA_NEWLINE
    return normalise_newlines(result, VBA_NEWLINE)


def format_source_file(path: Path, *, indent: str | None = None) -> bool:
    """Format ``path`` in place. Return True when the on-disk bytes changed."""
    formatted = format_vba(read_text(path), indent=indent)
    expected = normalise_newlines(strip_bom(formatted), VBA_NEWLINE).encode(TEXT_ENCODING)
    if path.read_bytes() == expected:
        return False
    write_text(path, formatted)
    return True


def _format_body(body: str, indent_unit: str) -> str:
    if not body:
        return ""
    unified = normalise_newlines(body, "\n")
    raw_lines = unified.split("\n")
    if raw_lines and raw_lines[-1] == "":
        raw_lines = raw_lines[:-1]

    formatted: list[str] = []
    indent = 0
    for group in _logical_groups(raw_lines):
        if len(group) == 1 and not group[0].strip():
            formatted.append("")
            continue

        logical = _logical_text(group)
        if _is_label(logical):
            formatted.extend(_emit_group(group, 0, indent_unit))
            continue

        before, after = _classify(logical)
        indent = max(0, indent + before)
        formatted.extend(_emit_group(group, indent, indent_unit))
        indent = max(0, indent + after)

    if not formatted:
        return ""
    return VBA_NEWLINE.join(formatted) + VBA_NEWLINE


def _logical_groups(lines: list[str]) -> list[list[str]]:
    groups: list[list[str]] = []
    current: list[str] = []
    for line in lines:
        if not line.strip():
            if current:
                groups.append(current)
                current = []
            groups.append([line])
            continue
        current.append(line)
        if not _is_continued(line):
            groups.append(current)
            current = []
    if current:
        groups.append(current)
    return groups


def _logical_text(group: list[str]) -> str:
    parts: list[str] = []
    for phys in group:
        code = _strip_comment(phys).rstrip()
        if _is_continued(phys):
            code = code[:-1].rstrip()
        if code.strip():
            parts.append(code.strip())
    return " ".join(parts)


def _emit_group(group: list[str], indent: int, indent_unit: str) -> list[str]:
    out: list[str] = []
    for i, phys in enumerate(group):
        extra = 1 if i > 0 else 0
        out.append(_emit_line(phys, indent + extra, indent_unit, continued=i > 0))
    return out


def _emit_line(line: str, indent: int, indent_unit: str, *, continued: bool) -> str:
    if not line.strip():
        return ""
    recased = _recase_line(line)
    if continued:
        trailing = recased[len(recased.rstrip()) :]
        return indent_unit * indent + recased.strip() + trailing
    return indent_unit * indent + recased.strip()


def _classify(logical: str) -> tuple[int, int]:
    s = logical.strip()
    if not s:
        return 0, 0

    if re.match(r"^#End\s+If\b", s, re.IGNORECASE):
        return -1, 0
    if re.match(r"^#ElseIf\b", s, re.IGNORECASE):
        return (-1, 1) if _block_then(s) else (-1, 0)
    if re.match(r"^#Else\b", s, re.IGNORECASE):
        return -1, 1
    if re.match(r"^#If\b", s, re.IGNORECASE):
        return (0, 1) if _block_then(s) else (0, 0)
    if re.match(r"^#Const\b", s, re.IGNORECASE):
        return 0, 0

    if _END_BLOCK.match(s):
        return -1, 0
    if re.match(r"^ElseIf\b", s, re.IGNORECASE):
        return -1, 1
    if re.match(r"^Else\b", s, re.IGNORECASE):
        return -1, 1
    if re.match(r"^Case\b", s, re.IGNORECASE):
        return -1, 1
    if re.match(r"^Next\b", s, re.IGNORECASE):
        rest = re.sub(r"^Next\b", "", s, count=1, flags=re.IGNORECASE).strip()
        return (-(1 + rest.count(",")) if rest else -1), 0
    if re.match(r"^Loop\b", s, re.IGNORECASE):
        return -1, 0
    if re.match(r"^Wend\b", s, re.IGNORECASE):
        return -1, 0
    if re.match(r"^Exit\b", s, re.IGNORECASE):
        return 0, 0

    body = _strip_access(s)
    if _PROPERTY.match(body):
        return (0, 0) if _contains_end_block(s, "Property") else (0, 1)
    if _PROC.match(body):
        kind = "Function" if re.match(r"^Function\b", body, re.IGNORECASE) else "Sub"
        return (0, 0) if _contains_end_block(s, kind) else (0, 1)
    if _TYPE_BLOCK.match(body) or _ENUM_BLOCK.match(body):
        kind = "Type" if _TYPE_BLOCK.match(body) else "Enum"
        return (0, 0) if _contains_end_block(s, kind) else (0, 1)
    if re.match(r"^If\b", s, re.IGNORECASE):
        if _contains_end_block(s, "If"):
            return 0, 0
        return (0, 1) if _block_then(s) else (0, 0)
    if re.match(r"^For\b", s, re.IGNORECASE):
        return 0, 1
    if re.match(r"^Do\b", s, re.IGNORECASE):
        return 0, 1
    if re.match(r"^While\b", s, re.IGNORECASE):
        return 0, 1
    if re.match(r"^With\b", s, re.IGNORECASE):
        return 0, 1
    if re.match(r"^Select\s+Case\b", s, re.IGNORECASE):
        return 0, 1
    return 0, 0


def _strip_access(s: str) -> str:
    while True:
        match = _ACCESS.match(s)
        if match is None:
            return s
        s = s[match.end() :]


def _block_then(s: str) -> bool:
    idx = _find_keyword(s, "then")
    if idx is None:
        return True
    return s[idx + 4 :].strip() == ""


def _contains_end_block(s: str, kind: str) -> bool:
    """True if ``s`` contains ``End <kind>`` as keywords (not the line start-only case)."""
    needle = kind.lower()
    start = 0
    while True:
        idx = _find_keyword(s[start:], "end")
        if idx is None:
            return False
        abs_idx = start + idx
        rest = s[abs_idx + 3 :].lstrip()
        if rest[: len(kind)].lower() == needle:
            after = rest[len(kind) : len(kind) + 1] or " "
            if not _is_ident_char(after):
                return True
        start = abs_idx + 1


def _is_label(s: str) -> bool:
    match = _LABEL.match(s.strip())
    if match is None:
        return False
    return match.group(1).lower() not in _LABEL_SKIP


def _is_continued(line: str) -> bool:
    code = _strip_comment(line)
    stripped = code.rstrip()
    if len(stripped) < 2 or stripped[-1] != "_" or not stripped[-2].isspace():
        return False
    return not _index_in_string(code, len(stripped) - 1)


def _strip_comment(line: str) -> str:
    i = 0
    n = len(line)
    in_str = False
    while i < n:
        ch = line[i]
        if ch == '"':
            if in_str and i + 1 < n and line[i + 1] == '"':
                i += 2
                continue
            in_str = not in_str
            i += 1
            continue
        if not in_str:
            if ch == "'":
                return line[:i]
            if _is_rem_at(line, i):
                return line[:i]
        i += 1
    return line


def _is_rem_at(line: str, i: int) -> bool:
    if line[i : i + 3].lower() != "rem":
        return False
    after = line[i + 3] if i + 3 < len(line) else " "
    if after.isalnum() or after == "_":
        return False
    before = line[i - 1] if i else " "
    return before.isspace() or before == ":" or i == 0


def _index_in_string(line: str, index: int) -> bool:
    in_str = False
    i = 0
    while i < index and i < len(line):
        ch = line[i]
        if ch == '"':
            if in_str and i + 1 < len(line) and line[i + 1] == '"':
                i += 2
                continue
            in_str = not in_str
        i += 1
    return in_str


def _find_keyword(s: str, word: str) -> int | None:
    needle = word.lower()
    length = len(needle)
    i = 0
    n = len(s)
    in_str = False
    while i < n:
        ch = s[i]
        if ch == '"':
            if in_str and i + 1 < n and s[i + 1] == '"':
                i += 2
                continue
            in_str = not in_str
            i += 1
            continue
        if not in_str and s[i : i + length].lower() == needle:
            before = s[i - 1] if i else " "
            after = s[i + length] if i + length < n else " "
            if not _is_ident_char(before) and not _is_ident_char(after):
                return i
        i += 1
    return None


def _is_ident_char(ch: str) -> bool:
    return ch.isalnum() or ch == "_"


def _recase_line(line: str) -> str:
    out: list[str] = []
    i = 0
    n = len(line)
    while i < n:
        ch = line[i]
        if ch == '"':
            j = i + 1
            while j < n:
                if line[j] == '"':
                    if j + 1 < n and line[j + 1] == '"':
                        j += 2
                        continue
                    j += 1
                    break
                j += 1
            out.append(line[i:j])
            i = j
            continue
        if ch == "'":
            out.append(line[i:])
            break
        if _is_rem_at(line, i):
            out.append(line[i:])
            break
        if ch == "#" and i + 1 < n and line[i + 1].isalpha():
            j = i + 1
            while j < n and _is_ident_char(line[j]):
                j += 1
            token = line[i:j]
            mapped = _PREPROCESSOR_CASE.get(token.lower())
            if mapped is not None:
                out.append(mapped)
                i = j
                continue
        if ch.isalpha() or ch == "_":
            j = i + 1
            while j < n and _is_ident_char(line[j]):
                j += 1
            token = line[i:j]
            out.append(_KEYWORD_CASE.get(token.lower(), token))
            i = j
            # VBE rejects "Mid $" / "Chr $" / "Environ $"; the suffix must
            # sit on the identifier. Glue optional space before $.
            k = i
            while k < n and line[k] in " \t":
                k += 1
            if k < n and line[k] == "$":
                out.append("$")
                i = k + 1
            continue
        out.append(ch)
        i += 1
    return "".join(out)
