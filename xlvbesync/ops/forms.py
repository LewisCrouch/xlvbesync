"""Private Excel session for the UserForm designer.

``DispatchEx`` is deliberate: the designer must not attach to the user's
Excel. A hidden instance never loads their add-ins, and quitting it cannot
close the workbook they are editing. One process, one STA thread, one
command at a time.
"""

from __future__ import annotations

import json
import sys
import uuid
from pathlib import Path
from typing import Any

from xlvbesync.core.codec import DEFAULT_ANSI_ENCODING
from xlvbesync.core.errors import XlvbesyncError, com_description
from xlvbesync.core.excel import com_apartment, launch_excel, quiet_excel
from xlvbesync.core.forms import (
    add_blank_form,
    apply_model,
    import_form,
    read_designer,
    write_form_files,
)
from xlvbesync.core.vbproject import get_vbproject

_DEFAULT_CODE = "Option Explicit\r\n"


class FormHost:
    """Owns the hidden Excel process and one workbook per open form."""

    def __init__(self) -> None:
        self.app: Any = None
        self.sessions: dict[str, dict[str, Any]] = {}

    def _application(self) -> Any:
        if self.app is None:
            # A dedicated process. GetActiveObject would bind a random running
            # instance and could write into the user's project.
            self.app = launch_excel(visible=False)
            try:
                self.app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            except Exception:  # noqa: BLE001
                pass
        return self.app

    def _workbook(self) -> Any:
        app = self._application()
        workbook = app.Workbooks.Add()
        get_vbproject(workbook)
        return workbook

    def create(self, frm: Path, name: str, *, ansi_encoding: str = DEFAULT_ANSI_ENCODING) -> dict[str, Any]:
        frm = frm.resolve()
        if frm.exists() or frm.with_suffix(".frx").exists():
            raise XlvbesyncError(
                f"{frm.name} already exists.",
                "Pick a different UserForm name.",
            )
        app = self._application()
        with quiet_excel(app):
            workbook = self._workbook()
            try:
                component = add_blank_form(workbook.VBProject, name)
                write_form_files(component, frm, _DEFAULT_CODE, ansi_encoding=ansi_encoding)
            finally:
                workbook.Close(SaveChanges=False)
        return {"frm": str(frm)}

    def open(self, frm: Path, *, ansi_encoding: str = DEFAULT_ANSI_ENCODING) -> dict[str, Any]:
        frm = frm.resolve()
        app = self._application()
        with quiet_excel(app):
            workbook = self._workbook()
            try:
                component = import_form(workbook.VBProject, frm, ansi_encoding=ansi_encoding)
                model, bindings = read_designer(component)
            except Exception:
                workbook.Close(SaveChanges=False)
                raise
        session_id = uuid.uuid4().hex
        self.sessions[session_id] = {
            "path": frm,
            "workbook": workbook,
            "component": component,
            "bindings": bindings,
            "ansi": ansi_encoding,
        }
        return {"session": session_id, "model": model}

    def apply(self, session_id: str, model: dict[str, Any]) -> dict[str, Any]:
        session = self._session(session_id)
        app = self._application()
        with quiet_excel(app):
            bindings, warnings = apply_model(session["component"], model, session["bindings"])
        session["bindings"] = bindings
        return {"warnings": warnings}

    def save(
        self,
        session_id: str,
        model: dict[str, Any],
        code: str,
        *,
        dest: Path | None = None,
    ) -> dict[str, Any]:
        session = self._session(session_id)
        path = dest.resolve() if dest is not None else session["path"]
        app = self._application()
        with quiet_excel(app):
            bindings, warnings = apply_model(session["component"], model, session["bindings"])
            session["bindings"] = bindings
            if str(session["component"].Name) != path.stem:
                session["component"].Name = path.stem
            write_form_files(
                session["component"],
                path,
                code,
                ansi_encoding=session["ansi"],
            )
        session["path"] = path
        return {"frm": str(path), "warnings": warnings}

    def close(self, session_id: str) -> dict[str, Any]:
        session = self.sessions.pop(session_id, None)
        if session is None:
            return {}
        app = self.app
        if app is not None:
            with quiet_excel(app):
                try:
                    session["workbook"].Close(SaveChanges=False)
                except Exception:  # noqa: BLE001
                    pass
        session["component"] = None
        session["bindings"] = {}
        return {}

    def shutdown(self) -> None:
        for session_id in list(self.sessions):
            try:
                self.close(session_id)
            except Exception:  # noqa: BLE001
                pass
        app = self.app
        self.app = None
        if app is not None:
            try:
                app.Quit()
            except Exception:  # noqa: BLE001
                pass

    def _session(self, session_id: str) -> dict[str, Any]:
        try:
            return self.sessions[session_id]
        except KeyError as exc:
            raise XlvbesyncError(
                "That form designer session is no longer open.",
                "Close the designer tab and open the form again.",
            ) from exc


def _emit(payload: dict[str, Any]) -> None:
    line = json.dumps(payload, ensure_ascii=False) + "\n"
    sys.stdout.buffer.write(line.encode("utf-8"))
    sys.stdout.buffer.flush()


def _handle(host: FormHost, message: dict[str, Any]) -> dict[str, Any]:
    cmd = str(message.get("cmd") or "")
    ansi = str(message.get("ansi") or DEFAULT_ANSI_ENCODING)
    if cmd == "ping":
        return {"pong": True}
    if cmd == "create":
        return host.create(Path(str(message.get("frm") or "")), str(message.get("name") or ""), ansi_encoding=ansi)
    if cmd == "open":
        return host.open(Path(str(message.get("frm") or "")), ansi_encoding=ansi)
    if cmd == "apply":
        return host.apply(str(message.get("session") or ""), message.get("model") or {})
    if cmd == "save":
        raw_dest = message.get("frm")
        return host.save(
            str(message.get("session") or ""),
            message.get("model") or {},
            str(message.get("code") or ""),
            dest=Path(str(raw_dest)) if raw_dest else None,
        )
    if cmd == "close":
        return host.close(str(message.get("session") or ""))
    if cmd == "shutdown":
        return {"shutdown": True}
    raise XlvbesyncError(
        f"Unknown form-host command {cmd!r}.",
        "Update the xlvbesync extension so it matches this xlvbesync install.",
    )


def serve() -> int:
    """Read one JSON command per line from stdin and write one JSON reply."""
    host = FormHost()
    try:
        with com_apartment():
            while True:
                raw = sys.stdin.buffer.readline()
                if not raw:
                    break
                text = raw.decode("utf-8", errors="replace").strip()
                if not text:
                    continue
                req_id: Any = None
                try:
                    message = json.loads(text)
                    req_id = message.get("id")
                    if message.get("cmd") == "shutdown":
                        _emit({"id": req_id, "ok": True, "shutdown": True})
                        break
                    result = _handle(host, message)
                    _emit({"id": req_id, "ok": True, **result})
                except XlvbesyncError as exc:
                    _emit({"id": req_id, "ok": False, "error": exc.message, "remedy": exc.remedy or ""})
                except Exception as exc:  # noqa: BLE001 - keep the pipe alive
                    _emit({"id": req_id, "ok": False, "error": com_description(exc), "remedy": ""})
    finally:
        host.shutdown()
    return 0
