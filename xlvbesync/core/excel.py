"""Excel COM connection management.

Responsibilities
----------------
* attach to a running Excel, or start one on demand;
* initialise COM correctly on whatever thread we happen to be on (the
  watchdog observer runs handlers on a worker thread, which *must* call
  ``CoInitialize`` itself);
* install an OLE message filter so Excel's "server busy" rejections turn into
  automatic waits instead of hard failures;
* suppress and reliably restore ``DisplayAlerts`` / ``ScreenUpdating`` /
  ``EnableEvents`` so injection never triggers a dialog and never leaves
  Excel in a broken state if we crash.
"""

from __future__ import annotations

import contextlib
import os
from pathlib import Path
from typing import Any, Iterator

import pythoncom
import win32com.client

from xlvbesync.core.errors import (
    ExcelNotRunningError,
    TargetNotOpenError,
    VBOMAccessDeniedError,
    com_description,
    is_vbom_denied,
)
from xlvbesync.core.retry import com_retry, wait_until

EXCEL_PROGID = "Excel.Application"


# ---------------------------------------------------------------------------
# Thread / COM apartment plumbing
# ---------------------------------------------------------------------------


@contextlib.contextmanager
def com_apartment() -> Iterator[None]:
    """Initialise COM for the current thread and uninitialise on exit.

    Safe to nest: ``CoInitialize`` refcounts, and RPC_E_CHANGED_MODE simply
    means the thread already lives in a compatible apartment.
    """
    initialised = False
    try:
        pythoncom.CoInitialize()
        initialised = True
    except pythoncom.com_error as exc:  # pragma: no cover - environment specific
        # RPC_E_CHANGED_MODE (-2147417850): already initialised as MTA.
        if exc.args and exc.args[0] != -2147417850:
            raise
    try:
        yield
    finally:
        if initialised:
            with contextlib.suppress(Exception):
                pythoncom.CoUninitialize()


@contextlib.contextmanager
def message_filter() -> Iterator[None]:
    """Install pywin32's OLE message filter for the current thread.

    Without this, a call arriving while Excel is busy is rejected outright.
    With it, COM retries transparently for a short window. We still keep
    :func:`xlvbesync.core.retry.com_retry` as a second line of defence.
    """
    installed = False
    try:
        pythoncom.CoRegisterMessageFilter(None)  # type: ignore[attr-defined]
        installed = True
    except Exception:  # noqa: BLE001 - not available on every pywin32 build
        installed = False
    try:
        yield
    finally:
        if installed:
            with contextlib.suppress(Exception):
                pythoncom.CoRegisterMessageFilter(None)  # type: ignore[attr-defined]


# ---------------------------------------------------------------------------
# Attaching to Excel
# ---------------------------------------------------------------------------


@com_retry()
def attach_excel() -> Any:
    """Return the running Excel Application object, or raise.

    ``GetActiveObject`` binds to the instance in the ROT, which is the one the
    user is actually looking at - important, because injecting into a fresh
    hidden instance would appear to do nothing.
    """
    try:
        return win32com.client.GetActiveObject(EXCEL_PROGID)
    except Exception as exc:  # noqa: BLE001
        raise ExcelNotRunningError(
            "No running Excel instance found.",
            "Open Excel (with your workbook or add-in loaded) and retry, or "
            "pass --launch to let xlvbesync start Excel itself.",
        ) from exc


@com_retry()
def launch_excel(*, visible: bool = True) -> Any:
    """Start a dedicated Excel instance."""
    app = win32com.client.DispatchEx(EXCEL_PROGID)
    app.Visible = visible
    return app


def get_excel(*, launch: bool = False, visible: bool = True) -> tuple[Any, bool]:
    """Attach to Excel, optionally launching it.

    Returns ``(app, started_by_us)`` so callers know whether they own the
    instance and should quit it when finished.
    """
    try:
        return attach_excel(), False
    except ExcelNotRunningError:
        if not launch:
            raise
    return launch_excel(visible=visible), True


def app_key(app: Any) -> Any:
    """Stable identity for an Excel Application, for de-duplication.

    ``Hwnd`` is unique per instance and cheap to read. Falls back to the COM
    wrapper's id if the property is unavailable.
    """
    try:
        return int(app.Hwnd)
    except Exception:  # noqa: BLE001
        return id(app)


def iter_excel_instances() -> Iterator[Any]:
    """Yield every running Excel Application object.

    ``GetActiveObject`` only ever returns the *first* instance registered in
    the Running Object Table, but Excel can legitimately run several
    processes at once - and an add-in loaded in instance B keeps its file
    locked even after we uninstall it from instance A. Enumerating the ROT is
    the only way to see them all.
    """
    try:
        rot = pythoncom.GetRunningObjectTable()
        ctx = pythoncom.CreateBindCtx(0)
    except Exception:  # noqa: BLE001 - no ROT available
        return

    seen: set[Any] = set()

    try:
        monikers = list(rot.EnumRunning())
    except Exception:  # noqa: BLE001
        return

    for moniker in monikers:
        # The ROT holds all sorts of objects; only Excel documents and the
        # Application itself expose an .Application we can use.
        try:
            name = moniker.GetDisplayName(ctx, None)
        except Exception:  # noqa: BLE001
            name = ""

        lowered = str(name).lower()
        if not (
            lowered.endswith((".xls", ".xlsx", ".xlsm", ".xlsb", ".xlam", ".xla"))
            or "excel" in lowered
        ):
            continue

        try:
            obj = rot.GetObject(moniker)
            dispatch = win32com.client.Dispatch(
                obj.QueryInterface(pythoncom.IID_IDispatch)
            )
            app = dispatch.Application
            _ = app.Version  # prove it really is Excel and still alive
        except Exception:  # noqa: BLE001 - stale or foreign ROT entry
            continue

        key = app_key(app)
        if key in seen:
            continue
        seen.add(key)
        yield app


def all_excel_instances() -> list[Any]:
    """Every running Excel Application, with the ROT-active one first."""
    instances = list(iter_excel_instances())

    try:
        primary = win32com.client.GetActiveObject(EXCEL_PROGID)
    except Exception:  # noqa: BLE001
        return instances

    primary_key = app_key(primary)
    ordered = [primary]
    ordered.extend(a for a in instances if app_key(a) != primary_key)
    return ordered


# ---------------------------------------------------------------------------
# Quiet-mode context manager
# ---------------------------------------------------------------------------


class _QuietState:
    """Snapshot of the Excel UI flags we mutate."""

    FLAGS = ("DisplayAlerts", "ScreenUpdating", "EnableEvents", "AskToUpdateLinks")

    def __init__(self, app: Any) -> None:
        self.app = app
        self.saved: dict[str, Any] = {}

    def capture(self) -> None:
        for flag in self.FLAGS:
            try:
                self.saved[flag] = getattr(self.app, flag)
            except Exception:  # noqa: BLE001 - flag unsupported on this build
                pass

    def apply_quiet(self) -> None:
        for flag in ("DisplayAlerts", "ScreenUpdating", "EnableEvents", "AskToUpdateLinks"):
            if flag not in self.saved:
                continue
            # EnableEvents must be False so Workbook_BeforeSave and friends do
            # not fire mid-injection; ScreenUpdating False keeps the VBE from
            # flickering and speeds bulk imports up noticeably.
            with contextlib.suppress(Exception):
                setattr(self.app, flag, False)

    def restore(self) -> None:
        for flag, value in self.saved.items():
            with contextlib.suppress(Exception):
                setattr(self.app, flag, value)


@contextlib.contextmanager
def quiet_excel(app: Any, *, enabled: bool = True) -> Iterator[None]:
    """Silence Excel prompts for the duration of the block.

    Always restores the previous values, including on exception, so a failed
    injection cannot leave the user with events permanently disabled.
    """
    if not enabled:
        yield
        return

    state = _QuietState(app)
    state.capture()
    state.apply_quiet()
    try:
        yield
    finally:
        state.restore()


# ---------------------------------------------------------------------------
# Workbook helpers
# ---------------------------------------------------------------------------


def same_path(a: str | os.PathLike[str], b: str | os.PathLike[str]) -> bool:
    """Case-insensitive path comparison suitable for Windows."""
    try:
        if os.path.normcase(os.path.abspath(str(a))) == os.path.normcase(
            os.path.abspath(str(b))
        ):
            return True
    except Exception:  # noqa: BLE001
        pass
    try:
        pa, pb = Path(a), Path(b)
        if pa.exists() and pb.exists():
            return os.path.samefile(pa, pb)
    except Exception:  # noqa: BLE001
        pass
    return False


def is_addin_path(path: Path | str) -> bool:
    """True if ``path`` has an Excel add-in extension."""
    return Path(path).suffix.lower() in (".xlam", ".xla")


def workbook_or_none(workbook: Any) -> Any | None:
    """Return ``workbook`` only if it is a live Excel Workbook COM object.

    ``Workbooks.Open`` of an already-loaded add-in can yield ``None`` / Empty
    without raising, which previously became ``None.VBProject``.
    """
    if workbook is None:
        return None
    try:
        _ = workbook.Name
    except Exception:  # noqa: BLE001 - Empty, disconnected, or not a Workbook
        return None
    return workbook


def iter_workbooks(app: Any) -> Iterator[Any]:
    """Yield every open workbook, including loaded add-ins.

    Add-ins appear in ``Application.Workbooks`` once installed, but iterating
    by index is safer than a generator over the COM collection because the
    collection can mutate under us.
    """
    try:
        count = app.Workbooks.Count
    except Exception:  # noqa: BLE001
        return
    for i in range(1, count + 1):
        try:
            yield app.Workbooks.Item(i)
        except Exception:  # noqa: BLE001 - workbook closed mid-iteration
            continue


def lookup_workbook(app: Any, name: str) -> Any | None:
    """Fetch a workbook by name via the collection's own indexer.

    This is essential for add-ins: a loaded .xlam is *not* returned by
    ``Workbooks.Item(i)`` iteration (nor counted in ``Workbooks.Count``), but
    ``Workbooks("Name.xlam")`` resolves it perfectly well. Iterating alone
    would make us conclude an installed add-in was closed.
    """
    seen: set[str] = set()
    candidates: list[str] = []
    for cand in (name, Path(name).name):
        key = cand.casefold()
        if cand and key not in seen:
            seen.add(key)
            candidates.append(cand)

    for cand in candidates:
        try:
            wb = workbook_or_none(app.Workbooks.Item(cand))
            if wb is not None:
                return wb
        except Exception:  # noqa: BLE001 - not open
            pass
        try:
            wb = workbook_or_none(app.Workbooks(cand))
            if wb is not None:
                return wb
        except Exception:  # noqa: BLE001 - not open
            pass
    return None


def find_workbook(app: Any, target: Path) -> Any | None:
    """Locate an open workbook/add-in by full path, then by file name.

    Installed add-ins are hidden from ``Workbooks`` iteration. They are also
    often loaded from a different folder than the repo target (Excel's AddIns
    directory, a mapped share, a previous copy). Live inject has to hit the
    loaded project, so a same-filename add-in is accepted when the paths
    differ.
    """
    target = Path(target)
    for wb in iter_workbooks(app):
        try:
            full = wb.FullName
        except Exception:  # noqa: BLE001
            continue
        if same_path(full, target):
            return wb

    wb = lookup_workbook(app, target.name)
    if wb is None:
        return None
    try:
        if same_path(wb.FullName, target):
            return wb
    except Exception:  # noqa: BLE001 - FullName unreadable
        if is_addin_path(target):
            return wb
        return None
    if is_addin_path(target):
        return wb
    return None


def find_workbook_by_name(app: Any, name: str) -> Any | None:
    """Locate an open workbook/add-in by file name (case-insensitive)."""
    wanted = name.casefold()
    for wb in iter_workbooks(app):
        try:
            if str(wb.Name).casefold() == wanted:
                return wb
        except Exception:  # noqa: BLE001
            continue
    return lookup_workbook(app, name)


def find_host(path: Path) -> Any | None:
    """Return the Excel instance that already has ``path`` open, if any.

    With several Excel processes running, ``GetActiveObject`` may hand back an
    instance that knows nothing about our target. Injecting there would report
    success while the user's actual session kept running the old code, so we
    always prefer the instance that genuinely hosts the file.
    """
    for app in all_excel_instances():
        try:
            if find_workbook(app, path) is not None:
                return app
            if is_addin_path(path) and find_workbook_by_name(app, path.name) is not None:
                return app
        except Exception:  # noqa: BLE001 - instance died mid-scan
            continue
    return None


def get_excel_for(
    path: Path,
    *,
    launch: bool = False,
    visible: bool = True,
) -> tuple[Any, bool]:
    """Attach to the best Excel instance for working on ``path``.

    Preference order: an instance already hosting the file, then the
    ROT-active instance, then (optionally) a newly launched one.
    """
    host = find_host(path)
    if host is not None:
        return host, False

    instances = all_excel_instances()
    if instances:
        return instances[0], False

    if launch:
        return launch_excel(visible=visible), True

    raise ExcelNotRunningError(
        "No running Excel instance found.",
        "Open Excel (with your workbook or add-in loaded) and retry, or pass "
        "--launch to let xlvbesync start Excel itself.",
    )


@com_retry()
def open_workbook(app: Any, path: Path, *, read_only: bool = False) -> Any:
    """Open a workbook or add-in, returning the Workbook object.

    ``AddToMru=False`` keeps the user's recent-files list clean;
    ``UpdateLinks=0`` and ``CorruptLoad`` defaults avoid prompts.
    """
    return workbook_or_none(
        app.Workbooks.Open(
            str(path),
            UpdateLinks=0,
            ReadOnly=read_only,
            AddToMru=False,
            Notify=False,
        )
    )


def ensure_workbook(app: Any, path: Path, *, read_only: bool = False) -> tuple[Any, bool]:
    """Return ``(workbook, opened_by_us)`` for ``path``.

    Opening an already-loaded add-in often returns Empty instead of the
    Workbook, so we fall back to a by-name lookup (and, for add-ins, a
    forced AddIns install) rather than handing ``None`` to ``VBProject``.
    """
    existing = find_workbook(app, path)
    if existing is not None:
        return existing, False

    if is_addin_path(path):
        from xlvbesync.core.addin import force_install

        if force_install(app, path):
            existing = find_workbook(app, path) or find_workbook_by_name(
                app, path.name
            )
            if existing is not None:
                return existing, False

    opened = open_workbook(app, path, read_only=read_only)
    if opened is not None:
        return opened, True

    existing = find_workbook(app, path) or find_workbook_by_name(app, path.name)
    if existing is not None:
        return existing, False

    raise TargetNotOpenError(
        f"Excel did not open {path.name}.",
        "If this is an add-in, enable it under File > Options > Add-ins "
        "(or open the .xlam as a workbook), then retry.",
    )


def is_vbe_busy(app: Any) -> bool:
    """True if VBA is executing or the VBE sits in break mode.

    Importing or replacing a module in either state either fails outright or
    silently resets the running project, so callers should refuse to proceed.
    """
    try:
        # Mode 1 = design, 2 = run, 3 = break. Not exposed directly, so we
        # probe via the VBE's active window state instead.
        vbe = app.VBE
    except Exception as exc:  # noqa: BLE001
        if is_vbom_denied(exc):
            raise VBOMAccessDeniedError(
                f"Cannot reach the VBA editor: {com_description(exc)}"
            ) from exc
        return False

    try:
        # When code is suspended the VBE exposes a non-empty call stack.
        return bool(vbe.ActiveCodePane) and bool(
            getattr(vbe.MainWindow, "Caption", "").lower().find("[break]") >= 0
        )
    except Exception:  # noqa: BLE001
        return False


def wait_for_idle(app: Any, timeout: float = 10.0) -> bool:
    """Wait until Excel reports it is ready to accept automation."""

    def ready() -> bool:
        return str(app.Ready) not in ("False", "0") if hasattr(app, "Ready") else True

    return wait_until(ready, timeout)
