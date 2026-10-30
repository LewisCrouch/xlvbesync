"""Exception hierarchy and COM HRESULT decoding for xlvbesync.

Every failure that a user can plausibly fix is surfaced as an XlvbesyncError
subclass carrying a human-readable remedy, rather than a raw pywin32
`com_error` tuple which is close to unreadable.
"""

from __future__ import annotations

# ---------------------------------------------------------------------------
# HRESULTs we care about. pywin32 surfaces these as the first element of
# com_error.args (signed 32-bit).
# ---------------------------------------------------------------------------

#: Server busy / "Call was rejected by callee." Excel is mid-calculation, has a
#: modal dialog open, or the user is dragging a cell. Always transient.
RPC_E_CALL_REJECTED = -2147418111

#: "The message filter indicated that the application is busy."
RPC_E_SERVERCALL_RETRYLATER = -2147417846

#: RPC server unavailable - Excel died or was closed mid-operation.
RPC_E_DISCONNECTED = -2147417848
RPC_S_SERVER_UNAVAILABLE = -2147023174

#: Programmatic access to the VBA project is denied (Trust Center setting off).
VBA_E_ACCESS_DENIED = -2146822220

#: Generic automation error raised by the VBE when access to VBProject is
#: blocked. Excel is inconsistent about which of these it returns.
E_FAIL_AUTOMATION = -2146777998

#: Common "cannot access" / locked-file HRESULTs.
E_ACCESSDENIED = -2147024891  # 0x80070005
ERROR_SHARING_VIOLATION = -2147024864  # 0x80070020

#: HRESULTs that justify an automatic retry.
TRANSIENT_HRESULTS = frozenset(
    {
        RPC_E_CALL_REJECTED,
        RPC_E_SERVERCALL_RETRYLATER,
    }
)

#: HRESULTs that mean "the VBA project object model is not trusted".
VBOM_DENIED_HRESULTS = frozenset(
    {
        VBA_E_ACCESS_DENIED,
        E_FAIL_AUTOMATION,
        E_ACCESSDENIED,
    }
)


class XlvbesyncError(Exception):
    """Base class for all xlvbesync failures.

    Parameters
    ----------
    message:
        What went wrong, in plain language.
    remedy:
        Optional actionable next step shown to the user.
    """

    def __init__(self, message: str, remedy: str | None = None) -> None:
        super().__init__(message)
        self.message = message
        self.remedy = remedy

    def __str__(self) -> str:
        if self.remedy:
            return f"{self.message}\n  -> {self.remedy}"
        return self.message


class ExcelNotRunningError(XlvbesyncError):
    """No Excel instance could be attached to or started."""


class VBOMAccessDeniedError(XlvbesyncError):
    """Excel refused programmatic access to the VBA project object model."""

    DEFAULT_REMEDY = (
        "Enable Excel > File > Options > Trust Center > Trust Center Settings > "
        "Macro Settings > 'Trust access to the VBA project object model', then "
        "fully restart Excel. Run `xlvbesync doctor` to verify."
    )

    def __init__(self, message: str, remedy: str | None = None) -> None:
        super().__init__(message, remedy or self.DEFAULT_REMEDY)


class TargetNotFoundError(XlvbesyncError):
    """The configured .xlsm/.xlam target does not exist on disk."""


class TargetNotOpenError(XlvbesyncError):
    """The target workbook/add-in is not currently open in Excel."""


class ComponentTypeError(XlvbesyncError):
    """A file extension or VBComponent type is unsupported."""


class FileLockedError(XlvbesyncError):
    """A file could not be replaced because Excel still holds a handle."""


class VbeBusyError(XlvbesyncError):
    """The VBE is in break mode or VBA code is currently executing."""

    DEFAULT_REMEDY = (
        "Press Stop (Reset) in the VBA editor to leave break mode, or wait for "
        "the running macro to finish, then retry."
    )

    def __init__(self, message: str, remedy: str | None = None) -> None:
        super().__init__(message, remedy or self.DEFAULT_REMEDY)


class ProjectProtectedError(XlvbesyncError):
    """The VBProject is locked for viewing (password protected)."""

    DEFAULT_REMEDY = (
        "Unlock the VBA project: in the VBE, Tools > VBAProject Properties > "
        "Protection, clear 'Lock project for viewing', save and reopen the file."
    )

    def __init__(self, message: str, remedy: str | None = None) -> None:
        super().__init__(message, remedy or self.DEFAULT_REMEDY)


class ConfigError(XlvbesyncError):
    """xlvbesync.json is missing, malformed, or internally inconsistent."""


def hresult_of(exc: BaseException) -> int | None:
    """Return the HRESULT from a pywin32 ``com_error``, or ``None``.

    pywin32 raises ``pywintypes.com_error`` whose ``.args`` are
    ``(hresult, description, excepinfo, argerror)``. ``excepinfo`` itself may
    carry a nested *scode* which is often the more specific error, so we fall
    back to it when the outer HRESULT is the generic DISP_E_EXCEPTION.
    """
    args = getattr(exc, "args", None)
    if not args:
        return None

    outer = args[0] if isinstance(args[0], int) else None

    # DISP_E_EXCEPTION: the real error hides in excepinfo[5] (scode).
    if outer == -2147352567 and len(args) > 2:
        excepinfo = args[2]
        if isinstance(excepinfo, tuple) and len(excepinfo) >= 6:
            scode = excepinfo[5]
            if isinstance(scode, int) and scode != 0:
                return scode
    return outer


def com_description(exc: BaseException) -> str:
    """Best-effort human-readable text out of a ``com_error``."""
    args = getattr(exc, "args", None)
    if not args:
        return str(exc)

    # excepinfo description is usually the most specific.
    if len(args) > 2 and isinstance(args[2], tuple) and len(args[2]) >= 3:
        desc = args[2][2]
        if desc:
            return str(desc).strip()

    if len(args) > 1 and args[1]:
        return str(args[1]).strip()

    return str(exc)


def is_transient(exc: BaseException) -> bool:
    """True if ``exc`` is a COM error worth retrying."""
    return hresult_of(exc) in TRANSIENT_HRESULTS


def is_vbom_denied(exc: BaseException) -> bool:
    """True if ``exc`` looks like a VBA-project-trust rejection."""
    if hresult_of(exc) not in VBOM_DENIED_HRESULTS:
        return False
    text = com_description(exc).lower()
    # E_ACCESSDENIED/E_FAIL are generic; require corroborating wording unless
    # the HRESULT is the unambiguous VBA-specific one.
    if hresult_of(exc) == VBA_E_ACCESS_DENIED:
        return True
    return any(
        kw in text
        for kw in ("programmatic access", "visual basic project", "vbproject", "not trusted")
    )
