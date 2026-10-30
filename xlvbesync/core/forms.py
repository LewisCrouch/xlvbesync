"""Read and write a UserForm through Excel's designer.

Exported ``.frm`` files store only the form shell. The controls live in the
sibling ``.frx`` (an MS-OFORMS stream). Building that stream by hand is brittle;
Excel already persists it. This module imports the pair into a private
workbook, walks ``Designer`` / ``Controls``, and exports again.

MSForms collections (``Controls``, ``Pages``, ``Tabs``) are zero-based.
Excel's own collections are one-based. The iterators below try zero-based
first.
"""

from __future__ import annotations

import os
import tempfile
from pathlib import Path
from typing import Any

from xlvbesync.core.codec import (
    DEFAULT_ANSI_ENCODING,
    encode_ansi,
    ensure_vb_name,
    normalise_newlines,
    read_text,
    split_header,
    write_text,
)
from xlvbesync.core.components import FORM_BINARY_EXT, VBEXT_CT_MS_FORM, validate_vba_name
from xlvbesync.core.errors import ComponentTypeError, VBOMAccessDeniedError, XlvbesyncError, com_description, is_vbom_denied
from xlvbesync.core.retry import com_retry
from xlvbesync.core.vbproject import export_component, get_vbproject

_TYPE_BY_KEY = {
    "label": "Label",
    "textbox": "TextBox",
    "combobox": "ComboBox",
    "listbox": "ListBox",
    "checkbox": "CheckBox",
    "optionbutton": "OptionButton",
    "togglebutton": "ToggleButton",
    "frame": "Frame",
    "commandbutton": "CommandButton",
    "tabstrip": "TabStrip",
    "multipage": "MultiPage",
    "scrollbar": "ScrollBar",
    "spinbutton": "SpinButton",
    "image": "Image",
    "page": "Page",
    "userform": "UserForm",
}

_PICTURES = {"Picture", "MouseIcon", "DisabledPicture"}
_SKIP_SET = _PICTURES | {"Name"}
_FLOAT_PROPS = {
    "Left",
    "Top",
    "Width",
    "Height",
    "GridX",
    "GridY",
    "ScrollLeft",
    "ScrollTop",
    "ScrollWidth",
    "ScrollHeight",
}
_CONTAINERS = {"Frame", "Page"}


def canonical_type(raw: str) -> str:
    """Map a COM type-info name onto an MSForms control type, or ''."""
    key = (raw or "").strip().lower()
    for prefix in ("msforms.", "forms."):
        if key.startswith(prefix):
            key = key[len(prefix) :]
    if key.startswith("_"):
        key = key[1:]
    if key.startswith("i") and key[1:] in _TYPE_BY_KEY:
        key = key[1:]
    return _TYPE_BY_KEY.get(key, "")


def prog_id_for(type_name: str) -> str:
    """Return the ``Controls.Add`` progid for a control type."""
    if type_name in _TYPE_BY_KEY.values() and type_name not in {"Page", "UserForm"}:
        return f"Forms.{type_name}.1"
    if type_name == "Page":
        return "Forms.Page.1"
    if type_name == "UserForm":
        return "Forms.UserForm.1"
    return ""


def code_body(text: str) -> str:
    """Return the procedure text, whether ``text`` is a full export or just code."""
    stripped = text.lstrip("\ufeff").lstrip()
    if stripped.upper().startswith("VERSION"):
        return split_header(text)[1]
    return text


def splice_code(exported: str, code: str) -> str:
    """Keep Excel's exported header and replace the procedure text."""
    header, _old = split_header(exported)
    body = code_body(code).strip("\r\n")
    header = header.rstrip("\r\n")
    if header and body:
        return header + "\r\n\r\n" + body + "\r\n"
    if header:
        return header + "\r\n"
    return body + ("\r\n" if body else "")


def iter_indexed(collection: Any) -> list[Any]:
    """Yield a COM collection, honouring MSForms' zero-based indexes."""
    try:
        count = int(collection.Count)
    except Exception as exc:  # noqa: BLE001
        raise XlvbesyncError(
            f"Could not read a form collection: {com_description(exc)}",
            "Close any Excel dialog and try the designer again.",
        ) from exc
    if count <= 0:
        return []
    zero: list[Any] = []
    for index in range(count):
        try:
            zero.append(collection.Item(index))
        except Exception:  # noqa: BLE001 - one-based collection
            zero = []
            break
    if len(zero) == count:
        return zero
    items: list[Any] = []
    for index in range(1, count + 1):
        items.append(collection.Item(index))
    return items


def _com_type_name(obj: Any) -> str:
    try:
        return str(obj._oleobj_.GetTypeInfo().GetDocumentation(-1)[0])
    except Exception:  # noqa: BLE001
        return ""


def _safe_str(obj: Any, attr: str, default: str = "") -> str:
    try:
        value = getattr(obj, attr)
        return "" if value is None else str(value)
    except Exception:  # noqa: BLE001
        return default


def _json_scalar(value: Any) -> Any | None:
    if isinstance(value, bool):
        return bool(value)
    if isinstance(value, int) and not isinstance(value, bool):
        return int(value)
    if isinstance(value, float):
        if value != value:  # NaN
            return None
        return float(value)
    if isinstance(value, str):
        return value
    return None


def _read_font(font: Any) -> dict[str, Any]:
    def flag(name: str) -> bool:
        try:
            return bool(getattr(font, name))
        except Exception:  # noqa: BLE001
            return False

    size = 8.0
    try:
        size = float(font.Size)
    except Exception:  # noqa: BLE001
        size = 8.0
    return {
        "name": _safe_str(font, "Name", "Tahoma") or "Tahoma",
        "size": size,
        "bold": flag("Bold"),
        "italic": flag("Italic"),
        "underline": flag("Underline"),
    }


# The controls returned by Designer.Controls do not expose a Properties
# collection over IDispatch. The names come from the type info instead, and
# values are read and written as ordinary attributes (``control.Caption``).
_SKIP_READ = {
    "Parent",
    "Controls",
    "Object",
    "Pages",
    "Tabs",
    "SelectedItem",
    "Column",
    "List",
    "Properties",
    "Application",
    "ActiveControl",
}


def _property_names(target: Any) -> list[str]:
    names: list[str] = []
    seen: set[str] = set()
    try:
        import pythoncom

        info = target._oleobj_.GetTypeInfo()
        attr = info.GetTypeAttr()
        for index in range(int(attr.cFuncs)):
            desc = info.GetFuncDesc(index)
            if int(desc.invkind) != int(pythoncom.INVOKE_PROPERTYGET):
                continue
            if int(desc.cParams) > 0:
                continue
            name = str(info.GetNames(desc.memid)[0])
            key = name.lower()
            if not name or name.startswith("_") or key in seen:
                continue
            seen.add(key)
            names.append(name)
    except Exception:  # noqa: BLE001
        return ["Caption", "Text", "Value", "Left", "Top", "Width", "Height", "Enabled", "Visible", "Font"]
    return names


def read_properties(target: Any) -> dict[str, Any]:
    """Read scalar designer properties. Pictures stay as a marker so apply leaves them alone."""
    out: dict[str, Any] = {}
    for name in _property_names(target):
        if name in _SKIP_READ:
            continue
        if name in _PICTURES:
            out[name] = {"picture": True}
            continue
        if name == "Font":
            try:
                out[name] = _read_font(target.Font)
            except Exception:  # noqa: BLE001
                continue
            continue
        try:
            value = getattr(target, name)
        except Exception:  # noqa: BLE001
            continue
        scalar = _json_scalar(value)
        if scalar is not None:
            out[name] = scalar
    return out


def _write_font(target: Any, spec: dict[str, Any]) -> None:
    font = target.Font
    if spec.get("name"):
        font.Name = str(spec["name"])
    if "size" in spec and spec["size"] is not None:
        font.Size = float(spec["size"])
    if "bold" in spec:
        font.Bold = bool(spec["bold"])
    if "italic" in spec:
        font.Italic = bool(spec["italic"])
    if "underline" in spec:
        font.Underline = bool(spec["underline"])


def _coerce(name: str, value: Any) -> Any:
    if isinstance(value, float) and name not in _FLOAT_PROPS and value.is_integer():
        return int(value)
    return value


def _set_property(target: Any, name: str, value: Any) -> None:
    if isinstance(value, dict):
        if value.get("picture"):
            return
        if name == "Font" or "size" in value or "bold" in value:
            _write_font(target, value)
            return
        return
    if name in _SKIP_SET:
        return
    setattr(target, name, _coerce(name, value))


def _bind(bindings: dict[str, dict[str, Any]], key: str, obj: Any, kind: str, opaque: bool, parent: str) -> None:
    bindings[key] = {"obj": obj, "kind": kind, "opaque": bool(opaque), "parent": parent}


def _read_controls(container: Any, bindings: dict[str, dict[str, Any]], parent: str) -> list[dict[str, Any]]:
    nodes: list[dict[str, Any]] = []
    try:
        controls = iter_indexed(container.Controls)
    except XlvbesyncError:
        return nodes
    except Exception:  # noqa: BLE001
        return nodes
    for z, ctrl in enumerate(controls):
        nodes.append(_read_control(ctrl, z, bindings, parent))
    return nodes


def _read_control(ctrl: Any, z: int, bindings: dict[str, dict[str, Any]], parent: str) -> dict[str, Any]:
    type_name = canonical_type(_com_type_name(ctrl))
    opaque = not type_name
    name = _safe_str(ctrl, "Name", f"Control{z + 1}")
    key = name or f"opaque{parent}{z}"
    node: dict[str, Any] = {
        "key": key,
        "name": name,
        "progId": prog_id_for(type_name),
        "type": type_name or "Unknown",
        "opaque": opaque,
        "z": z,
        "properties": {} if opaque else read_properties(ctrl),
        "children": [],
    }
    _bind(bindings, key, ctrl, "control", opaque, parent)
    if type_name == "MultiPage":
        node["children"] = _read_pages(ctrl, bindings, key)
    elif type_name == "TabStrip":
        node["tabs"] = _read_tabs(ctrl)
    elif type_name in _CONTAINERS:
        node["children"] = _read_controls(ctrl, bindings, key)
    return node


def _read_pages(multipage: Any, bindings: dict[str, dict[str, Any]], parent: str) -> list[dict[str, Any]]:
    pages: list[dict[str, Any]] = []
    try:
        items = iter_indexed(multipage.Pages)
    except Exception:  # noqa: BLE001
        return pages
    for index, page in enumerate(items):
        name = _safe_str(page, "Name", f"Page{index + 1}")
        caption = _safe_str(page, "Caption", name)
        key = name
        node = {
            "key": key,
            "name": name,
            "progId": "Forms.Page.1",
            "type": "Page",
            "opaque": False,
            "z": index,
            "properties": {"Caption": caption, "Name": name},
            "children": _read_controls(page, bindings, key),
        }
        _bind(bindings, key, page, "page", False, parent)
        pages.append(node)
    return pages


def _read_tabs(tabstrip: Any) -> list[dict[str, str]]:
    tabs: list[dict[str, str]] = []
    try:
        items = iter_indexed(tabstrip.Tabs)
    except Exception:  # noqa: BLE001
        return tabs
    for index, tab in enumerate(items):
        name = _safe_str(tab, "Name", f"Tab{index + 1}")
        tabs.append({"name": name, "caption": _safe_str(tab, "Caption", name)})
    return tabs


# Designer-host state. These are not form properties and writing them back
# toggles the hidden Excel window's toolbox rather than the UserForm.
_FORM_SKIP = {
    "ActiveControl",
    "Controls",
    "Selected",
    "CanPaste",
    "CanRedo",
    "CanUndo",
    "DesignMode",
    "ShowToolbox",
}


def read_component_properties(component: Any) -> dict[str, Any]:
    """Read the VBE property grid for a UserForm.

    ``component.Properties`` is the one-based VBIDE collection. It carries
    Width, Caption, StartUpPosition and the grid flags. The Designer object's
    own type info only exposes a handful of them.
    """
    out: dict[str, Any] = {}
    try:
        props = component.Properties
        count = int(props.Count)
    except Exception:  # noqa: BLE001
        return out
    for index in range(1, count + 1):
        try:
            prop = props.Item(index)
            name = str(prop.Name)
        except Exception:  # noqa: BLE001
            continue
        if name in _FORM_SKIP or name in _SKIP_READ:
            continue
        if name in _PICTURES:
            out[name] = {"picture": True}
            continue
        if name == "Font":
            try:
                out[name] = _read_font(component.Designer.Font)
            except Exception:  # noqa: BLE001
                continue
            continue
        try:
            value = prop.Value
        except Exception:  # noqa: BLE001
            continue
        scalar = _json_scalar(value)
        if scalar is not None:
            out[name] = scalar
    return out


def _component_property(component: Any, name: str) -> Any:
    try:
        if name == "Font":
            return _read_font(component.Designer.Font)
        return _json_scalar(component.Properties(name).Value)
    except Exception:  # noqa: BLE001
        return None


def _set_component_property(component: Any, name: str, value: Any) -> None:
    if isinstance(value, dict):
        if value.get("picture"):
            return
        if name == "Font":
            _write_font(component.Designer, value)
            return
        return
    if name in _SKIP_SET or name in _FORM_SKIP:
        return
    component.Properties(name).Value = _coerce(name, value)


def _apply_form_properties(component: Any, properties: dict[str, Any], warnings: list[str]) -> None:
    label = _safe_str(component, "Name", "form")
    for name, value in properties.items():
        if name in _SKIP_SET or name in _FORM_SKIP:
            continue
        if _values_equal(_component_property(component, name), value):
            continue
        try:
            _set_component_property(component, name, value)
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"{label}.{name}: {com_description(exc)}")


def read_designer(component: Any) -> tuple[dict[str, Any], dict[str, dict[str, Any]]]:
    """Return ``(model, bindings)`` for a UserForm component."""
    try:
        designer = component.Designer
    except Exception as exc:  # noqa: BLE001
        if is_vbom_denied(exc):
            raise VBOMAccessDeniedError(
                f"Excel denied access to the form designer: {com_description(exc)}"
            ) from exc
        raise XlvbesyncError(
            f"Could not open the form designer: {com_description(exc)}",
            "Confirm the file is a UserForm and that Trust access to the VBA project object model is on.",
        ) from exc
    bindings: dict[str, dict[str, Any]] = {}
    model = {
        "name": str(component.Name),
        "progId": "Forms.UserForm.1",
        "type": "UserForm",
        "properties": read_component_properties(component),
        "children": _read_controls(designer, bindings, ""),
    }
    return model, bindings


def _flatten(node: dict[str, Any], parent: str, out: list[tuple[str, dict[str, Any]]]) -> None:
    for child in node.get("children") or []:
        out.append((parent, child))
        _flatten(child, str(child.get("key") or ""), out)


def _temp_name(index: int) -> str:
    return f"xTmp{index:04d}"


def _remove_binding(binding: dict[str, Any]) -> None:
    obj = binding["obj"]
    if binding.get("kind") == "page":
        parent = obj.Parent
        index = int(obj.Index)
        try:
            parent.Pages.Remove(index)
        except Exception:  # noqa: BLE001
            parent.Pages.Remove(max(0, index - 1))
        return
    parent = obj.Parent
    parent.Controls.Remove(str(obj.Name))


def _add_control(container: Any, prog_id: str, name: str) -> Any:
    try:
        return container.Controls.Add(prog_id, name, True)
    except Exception:  # noqa: BLE001
        created = container.Controls.Add(prog_id)
        created.Name = name
        return created


def _container(parent_key: str, bindings: dict[str, dict[str, Any]], designer: Any) -> Any:
    if not parent_key:
        return designer
    binding = bindings.get(parent_key)
    if binding is None:
        raise XlvbesyncError(
            f"The form refers to a missing container {parent_key!r}.",
            "Reload the designer and try the edit again.",
        )
    return binding["obj"]


def _values_equal(current: Any, new: Any) -> bool:
    if isinstance(new, dict):
        if new.get("picture"):
            return True
        if not isinstance(current, dict):
            return False
        return all(_values_equal(current.get(key), value) for key, value in new.items())
    if isinstance(new, bool) or isinstance(current, bool):
        return bool(current) == bool(new)
    if isinstance(current, (int, float)) and isinstance(new, (int, float)):
        return abs(float(current) - float(new)) < 0.05
    return current == new


def _current_property(target: Any, name: str) -> Any:
    try:
        if name == "Font":
            return _read_font(target.Font)
        return _json_scalar(getattr(target, name))
    except Exception:  # noqa: BLE001
        return None


def _apply_properties(target: Any, properties: dict[str, Any], warnings: list[str], label: str) -> None:
    for name, value in properties.items():
        if name in _SKIP_SET:
            continue
        if _values_equal(_current_property(target, name), value):
            continue
        try:
            _set_property(target, name, value)
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"{label}.{name}: {com_description(exc)}")


def _sync_tabs(ctrl: Any, tabs: list[dict[str, Any]], warnings: list[str]) -> None:
    try:
        live = iter_indexed(ctrl.Tabs)
    except Exception as exc:  # noqa: BLE001
        warnings.append(f"tabs: {com_description(exc)}")
        return
    for index, spec in enumerate(tabs):
        if index >= len(live):
            break
        tab = live[index]
        try:
            if spec.get("caption") is not None:
                tab.Caption = str(spec["caption"])
            if spec.get("name"):
                tab.Name = str(spec["name"])
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"tab {index}: {com_description(exc)}")


def apply_model(
    component: Any,
    model: dict[str, Any],
    bindings: dict[str, dict[str, Any]],
) -> tuple[dict[str, dict[str, Any]], list[str]]:
    """Mutate the live designer so it matches ``model``.

    Identity is the control ``key`` captured when the form was opened (the
    original name). Renames update ``Name`` without recreating the control, so
    pictures and any property we do not model survive.
    """
    warnings: list[str] = []
    designer = component.Designer
    desired_pairs = []
    _flatten(model, "", desired_pairs)
    desired = {str(child.get("key")): (parent, child) for parent, child in desired_pairs}
    opaque_keys = {key for key, binding in bindings.items() if binding.get("opaque")}

    # Free names before assigning the final ones, or two controls cannot swap.
    for index, (key, (_parent, child)) in enumerate(desired.items()):
        binding = bindings.get(key)
        if binding is None or binding.get("opaque") or binding.get("kind") == "page":
            continue
        new_name = str(child.get("name") or "")
        if new_name and _safe_str(binding["obj"], "Name") != new_name:
            try:
                binding["obj"].Name = _temp_name(index)
            except Exception as exc:  # noqa: BLE001
                warnings.append(f"rename {key}: {com_description(exc)}")

    for key, (_parent, child) in desired.items():
        binding = bindings.get(key)
        if binding is None or binding.get("opaque"):
            continue
        new_name = str(child.get("name") or "")
        if not new_name or binding.get("kind") == "page":
            if binding.get("kind") == "page":
                try:
                    if child.get("properties", {}).get("Caption") is not None:
                        binding["obj"].Caption = str(child["properties"]["Caption"])
                    if new_name:
                        binding["obj"].Name = new_name
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"page {key}: {com_description(exc)}")
            continue
        try:
            if _safe_str(binding["obj"], "Name") != new_name:
                validate_vba_name(new_name)
                binding["obj"].Name = new_name
        except (ComponentTypeError, XlvbesyncError) as exc:
            warnings.append(f"rename {key}: {exc}")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"rename {key}: {com_description(exc)}")

    # Delete deepest first so a frame goes after the controls it still owns.
    stale = [key for key in list(bindings) if key not in desired and key not in opaque_keys]
    depths = {key: 0 for key in stale}

    def depth_of(key: str) -> int:
        binding = bindings.get(key)
        if not binding:
            return 0
        parent = binding.get("parent") or ""
        if not parent or parent not in depths:
            return 1
        return 1 + depth_of(parent)

    stale.sort(key=depth_of, reverse=True)
    for key in stale:
        binding = bindings.get(key)
        if binding is None:
            continue
        try:
            _remove_binding(binding)
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"delete {key}: {com_description(exc)}")
            continue
        del bindings[key]

    additions = [(parent, child) for parent, child in desired_pairs if str(child.get("key")) not in bindings]
    additions.sort(key=lambda item: _model_depth(model, str(item[1].get("key"))))
    for parent_key, child in additions:
        if child.get("opaque"):
            continue
        type_name = str(child.get("type") or canonical_type(str(child.get("progId") or "")))
        name = str(child.get("name") or "")
        key = str(child.get("key") or name)
        if key in bindings:
            continue
        if not name:
            warnings.append(f"skip unnamed control {key}")
            continue
        try:
            validate_vba_name(name)
            container = _container(parent_key, bindings, designer)
            if type_name == "Page":
                page = container.Pages.Add(name)
                try:
                    page.Caption = str((child.get("properties") or {}).get("Caption") or name)
                except Exception:  # noqa: BLE001
                    pass
                _bind(bindings, key, page, "page", False, parent_key)
            else:
                prog_id = str(child.get("progId") or prog_id_for(type_name))
                if not prog_id:
                    warnings.append(f"no progid for {name}")
                    continue
                created = _add_control(container, prog_id, name)
                _bind(bindings, key, created, "control", False, parent_key)
                if type_name == "MultiPage":
                    _bind_new_pages(created, child, bindings, key, warnings)
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"add {name}: {com_description(exc)}")

    _apply_form_properties(component, model.get("properties") or {}, warnings)
    for parent_key, child in desired_pairs:
        key = str(child.get("key") or "")
        binding = bindings.get(key)
        if binding is None or binding.get("opaque"):
            continue
        label = str(child.get("name") or key)
        _apply_properties(binding["obj"], child.get("properties") or {}, warnings, label)
        if child.get("type") == "TabStrip" and child.get("tabs"):
            _sync_tabs(binding["obj"], child["tabs"], warnings)
        binding["parent"] = parent_key

    _apply_z_order(model, bindings, warnings)
    return bindings, warnings


def _bind_new_pages(
    multipage: Any,
    node: dict[str, Any],
    bindings: dict[str, dict[str, Any]],
    parent_key: str,
    warnings: list[str],
) -> None:
    """A new MultiPage already has pages. Bind those instead of adding a second set."""
    try:
        live = iter_indexed(multipage.Pages)
    except Exception as exc:  # noqa: BLE001
        warnings.append(f"pages: {com_description(exc)}")
        return
    wanted = [child for child in node.get("children") or [] if child.get("type") == "Page"]
    while len(live) > len(wanted) and live:
        try:
            multipage.Pages.Remove(len(live) - 1)
        except Exception:  # noqa: BLE001
            try:
                multipage.Pages.Remove(len(live))
            except Exception as exc:  # noqa: BLE001
                warnings.append(f"remove page: {com_description(exc)}")
                break
        live = iter_indexed(multipage.Pages)
    for page_node, page_obj in zip(wanted, live):
        page_key = str(page_node.get("key") or page_node.get("name") or "")
        if not page_key:
            continue
        try:
            page_name = str(page_node.get("name") or "")
            if page_name:
                page_obj.Name = page_name
            caption = (page_node.get("properties") or {}).get("Caption")
            if caption is not None:
                page_obj.Caption = str(caption)
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"page {page_key}: {com_description(exc)}")
        _bind(bindings, page_key, page_obj, "page", False, parent_key)


def _model_depth(model: dict[str, Any], key: str) -> int:
    found = 0

    def walk(node: dict[str, Any], depth: int) -> None:
        nonlocal found
        for child in node.get("children") or []:
            if str(child.get("key")) == key:
                found = depth
            walk(child, depth + 1)

    walk(model, 1)
    return found


def _apply_z_order(model: dict[str, Any], bindings: dict[str, dict[str, Any]], warnings: list[str]) -> None:
    """Children are stored front-to-back. Sending each to the back in that order reproduces it."""

    def walk(node: dict[str, Any]) -> None:
        kids = [child for child in node.get("children") or [] if child.get("type") != "Page"]
        for child in kids:
            binding = bindings.get(str(child.get("key")))
            if binding is not None and binding.get("kind") != "page":
                try:
                    binding["obj"].ZOrder(1)
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"z-order {child.get('name')}: {com_description(exc)}")
        for child in node.get("children") or []:
            if child.get("children"):
                walk(child)

    walk(model)


def _stage_import(frm: Path, ansi_encoding: str, tmpdir: Path) -> Path:
    text = ensure_vb_name(read_text(frm), frm.stem)
    data, _lossy = encode_ansi(normalise_newlines(text), ansi_encoding)
    staged = tmpdir / frm.name
    staged.write_bytes(data)
    frx = frm.with_suffix(FORM_BINARY_EXT)
    if frx.is_file():
        (tmpdir / frx.name).write_bytes(frx.read_bytes())
    return staged


@com_retry()
def import_form(project: Any, frm: Path, *, ansi_encoding: str = DEFAULT_ANSI_ENCODING) -> Any:
    """Import a ``.frm``/``.frx`` pair. The binary must sit beside the staged text."""
    if not frm.is_file():
        raise XlvbesyncError(
            f"Form file not found: {frm}",
            "Create the UserForm first, or check the path.",
        )
    with tempfile.TemporaryDirectory(prefix="xlvbesync-form-") as tmp:
        staged = _stage_import(frm, ansi_encoding, Path(tmp))
        try:
            return project.VBComponents.Import(str(staged))
        except Exception as exc:  # noqa: BLE001
            if is_vbom_denied(exc):
                raise VBOMAccessDeniedError(
                    f"Excel denied the form import: {com_description(exc)}"
                ) from exc
            raise XlvbesyncError(
                f"Excel could not import {frm.name}: {com_description(exc)}",
                "If the form lost its controls, keep the .frx beside the .frm and extract again.",
            ) from exc


@com_retry()
def write_form_files(
    component: Any,
    dest: Path,
    code: str,
    *,
    ansi_encoding: str = DEFAULT_ANSI_ENCODING,
) -> None:
    """Export the designer, splice ``code`` over the procedure section, write ``.frx`` then ``.frm``."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    if str(component.Name) != dest.stem:
        validate_vba_name(dest.stem)
        component.Name = dest.stem
    with tempfile.TemporaryDirectory(prefix="xlvbesync-form-") as tmp:
        staged = Path(tmp) / dest.name
        text = export_component(component, staged, ansi_encoding=ansi_encoding)
        spliced = splice_code(text, code)
        frx = staged.with_suffix(FORM_BINARY_EXT)
        target_frx = dest.with_suffix(FORM_BINARY_EXT)
        # The watcher keys off the .frm and remaps a .frx change onto it.
        # Publishing the binary first means the debounced sync sees both.
        if frx.is_file():
            stage_frx = target_frx.with_name(target_frx.name + ".xlvbesync-stage")
            stage_frx.write_bytes(frx.read_bytes())
            os.replace(stage_frx, target_frx)
        write_text(dest, spliced)


def add_blank_form(project: Any, name: str) -> Any:
    """Create an empty UserForm component named ``name``."""
    validate_vba_name(name)
    try:
        component = project.VBComponents.Add(VBEXT_CT_MS_FORM)
        component.Name = name
    except Exception as exc:  # noqa: BLE001
        if is_vbom_denied(exc):
            raise VBOMAccessDeniedError(
                f"Excel denied creating the UserForm: {com_description(exc)}"
            ) from exc
        raise XlvbesyncError(
            f"Could not create UserForm {name}: {com_description(exc)}",
            "Check Trust access to the VBA project object model, then try again.",
        ) from exc
    try:
        component.Properties("Caption").Value = name
    except Exception:  # noqa: BLE001
        pass
    return component
