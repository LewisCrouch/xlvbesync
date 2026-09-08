"""Excel add-in (.xlam) lifecycle management.

Excel holds an exclusive handle on a loaded .xlam, so overwriting the file on
disk while it is installed fails with a sharing violation. The bulletproof
sequence is:

1. record the add-in's ``Installed`` state and its ``AddIns`` collection entry;
2. set ``Installed = False`` (fires ``Workbook_AddinUninstall``);
3. close the underlying workbook if Excel still has it open;
4. wait for the OS handle to actually drop - this is *not* instantaneous, and
   is the step most naive implementations omit;
5. replace the file;
6. set ``Installed = True`` again (fires ``Workbook_AddinInstall`` and
   ``Workbook_Open``).

Step 4 matters because ``Installed = False`` returns before Excel has finished
releasing the file. We poll with an exclusive-open probe rather than sleeping a
fixed amount.
"""

from __future__ import annotations

import contextlib
import os
import shutil
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from xlvba.core.errors import FileLockedError, XlvbaError, com_description
from xlvba.core.excel import (
    find_workbook,
    find_workbook_by_name,
    lookup_workbook,
    same_path,
)
from xlvba.core.retry import com_retry, wait_until

#: How long to wait for Excel to release its handle on the .xlam.
DEFAULT_UNLOCK_TIMEOUT = 15.0

#: How long to wait for the add-in to report itself installed again.
DEFAULT_RELOAD_TIMEOUT = 20.0


def is_addin(path: Path | str) -> bool:
    """True if the path looks like an Excel add-in."""
    return Path(path).suffix.lower() in (".xlam", ".xla")


# ---------------------------------------------------------------------------
# File lock probing
# ---------------------------------------------------------------------------


def is_file_locked(path: Path) -> bool:
    """True if the file cannot currently be opened for exclusive writing.

    Opening in ``r+b`` mode is the cheapest reliable probe on Windows: it
    fails with ``PermissionError`` while another process holds a write or deny
    handle, which is exactly the condition that would make our replace fail.
    """
    if not path.exists():
        return False
    try:
        with open(path, "r+b"):
            return False
    except PermissionError:
        return True
    except OSError:
        return True


def wait_for_unlock(path: Path, timeout: float = DEFAULT_UNLOCK_TIMEOUT) -> bool:
    """Block until ``path`` is writable, returning False on timeout."""
    return wait_until(lambda: not is_file_locked(path), timeout, interval=0.15)


def _excel_owner_file(path: Path) -> Path:
    """Path of the ``~$name.xlam`` owner file Excel creates alongside."""
    return path.with_name(f"~${path.name}")


# ---------------------------------------------------------------------------
# Add-in discovery
# ---------------------------------------------------------------------------


@dataclass
class AddinHandle:
    """A located ``AddIn`` collection entry plus its prior state."""

    com: Any
    name: str
    was_installed: bool


def find_addin(app: Any, path: Path) -> AddinHandle | None:
    """Find the ``AddIns`` entry matching ``path``.

    Matching is by full path where available, falling back to filename: an
    add-in registered from a different folder than the one we are editing
    would otherwise be silently mistaken for ours.
    """
    filename = path.name.casefold()

    try:
        count = int(app.AddIns.Count)
    except Exception:  # noqa: BLE001
        return None

    fallback: AddinHandle | None = None

    for i in range(1, count + 1):
        try:
            entry = app.AddIns.Item(i)
        except Exception:  # noqa: BLE001
            continue

        try:
            installed = bool(entry.Installed)
        except Exception:  # noqa: BLE001
            installed = False

        try:
            full = str(entry.FullName)
        except Exception:  # noqa: BLE001
            full = ""

        try:
            entry_name = str(entry.Name)
        except Exception:  # noqa: BLE001
            entry_name = ""

        if full and same_path(full, path):
            return AddinHandle(entry, entry_name or path.name, installed)

        if entry_name.casefold() == filename and fallback is None:
            fallback = AddinHandle(entry, entry_name, installed)

    return fallback


@com_retry()
def register_addin(app: Any, path: Path) -> Any:
    """Add ``path`` to the AddIns collection (without installing it)."""
    return app.AddIns.Add(str(path), False)


def is_loaded(app: Any, path: Path) -> bool:
    """True if the add-in's workbook is actually present in ``app``.

    Note that a loaded add-in does not appear in ``Workbooks`` iteration, so
    this relies on the by-name indexer (see
    :func:`xlvba.core.excel.lookup_workbook`).
    """
    wb = find_workbook_by_name(app, path.name)
    if wb is None:
        return False
    try:
        return same_path(wb.FullName, path)
    except Exception:  # noqa: BLE001
        return False


def force_install(app: Any, path: Path, *, timeout: float = DEFAULT_RELOAD_TIMEOUT) -> bool:
    """Make sure the add-in is genuinely loaded in ``app``.

    Needed because an automation-launched Excel does not load add-ins at
    startup even though ``AddIn.Installed`` still reads True from the
    registry. Toggling the flag forces a real load.
    """
    handle = find_addin(app, path)
    if handle is None:
        handle = ensure_addin_entry(app, path)

    if is_loaded(app, path):
        return True

    try:
        if bool(handle.com.Installed):
            set_installed(handle, False)
            settle(0.2)
        set_installed(handle, True)
    except Exception:  # noqa: BLE001
        return False

    return wait_until(lambda: is_loaded(app, path), timeout, interval=0.2)


def uninstall_and_deregister(app: Any, path: Path) -> None:
    """Fully remove an add-in registration - used to clean up after tests.

    Leaving a test add-in flagged installed would make it auto-load in the
    user's real Excel sessions.
    """
    handle = find_addin(app, path)
    if handle is None:
        return
    with contextlib.suppress(Exception):
        set_installed(handle, False)
    close_addin_workbook(app, path)


def ensure_addin_entry(app: Any, path: Path) -> AddinHandle:
    """Find, or register then find, the AddIns entry for ``path``."""
    handle = find_addin(app, path)
    if handle is not None:
        return handle

    register_addin(app, path)
    handle = find_addin(app, path)
    if handle is None:
        raise XlvbaError(
            f"Could not register {path.name} in Excel's add-ins collection.",
            "Install it manually once via File > Options > Add-ins > Browse, "
            "then retry.",
        )
    return handle


# ---------------------------------------------------------------------------
# Uninstall / reinstall
# ---------------------------------------------------------------------------


@com_retry()
def set_installed(handle: AddinHandle, installed: bool) -> None:
    handle.com.Installed = installed


def close_addin_workbook(app: Any, path: Path) -> bool:
    """Close the add-in's workbook if Excel still holds it open.

    After ``Installed = False`` Excel *usually* closes the workbook, but a
    workbook opened directly (rather than installed) stays open, and so does
    one whose ``AddinUninstall`` handler raised.
    """
    wb = find_workbook(app, path) or find_workbook_by_name(app, path.name)
    if wb is None:
        return False
    try:
        wb.Close(SaveChanges=False)
        return True
    except Exception:  # noqa: BLE001 - already closing, or refused
        return False


def unload_from_instance(app: Any, path: Path) -> AddinHandle | None:
    """Uninstall and close the add-in within a single Excel instance."""
    handle = find_addin(app, path)

    if handle is not None and handle.was_installed:
        try:
            set_installed(handle, False)
        except Exception as exc:  # noqa: BLE001
            raise XlvbaError(
                f"Could not uninstall add-in {path.name}: {com_description(exc)}",
                "Check the add-in's Workbook_AddinUninstall handler for errors.",
            ) from exc

    # Even when not installed as an add-in, the file may be open as a plain
    # workbook - which locks it just as effectively.
    close_addin_workbook(app, path)
    return handle


def unload_addin(
    app: Any,
    path: Path,
    *,
    timeout: float = DEFAULT_UNLOCK_TIMEOUT,
    all_instances: bool = True,
) -> list[tuple[Any, AddinHandle | None]]:
    """Unload the add-in from every Excel instance holding it.

    Returns ``[(app, handle), ...]`` for each instance, so the caller can
    restore exactly the instances that had it installed.

    Excel can run multiple processes, and an add-in loaded in a second
    instance keeps the file locked no matter what we do to the first. We
    therefore uninstall everywhere before touching the file.
    """
    from xlvba.core.excel import all_excel_instances, app_key

    if all_instances:
        instances = all_excel_instances()
        # Guarantee the caller's own instance is included even if the ROT
        # enumeration missed it (it can lag for a freshly started process).
        if not any(app_key(a) == app_key(app) for a in instances):
            instances.insert(0, app)
    else:
        instances = [app]

    hosts: list[tuple[Any, AddinHandle | None]] = []
    for instance in instances:
        try:
            hosts.append((instance, unload_from_instance(instance, path)))
        except XlvbaError:
            raise
        except Exception:  # noqa: BLE001 - instance died mid-call
            continue

    if not wait_for_unlock(path, timeout):
        holder = _excel_owner_file(path)
        hint = (
            f" An Excel owner file is still present ({holder.name})."
            if holder.exists()
            else ""
        )
        raise FileLockedError(
            f"{path.name} is still locked {timeout:.0f}s after unloading it from "
            f"{len(hosts)} Excel instance(s).{hint}",
            "Close any Excel window showing this file, make sure no macro from "
            "it is running, then retry. If it persists, close Excel entirely.",
        )

    return hosts


def reload_addin(
    app: Any,
    path: Path,
    handle: AddinHandle | None,
    *,
    timeout: float = DEFAULT_RELOAD_TIMEOUT,
) -> None:
    """Re-install the add-in in one Excel instance after the file was replaced.

    The stored ``AddIn`` COM reference can go stale across the file swap
    (Excel rebuilds the collection entry), so we re-resolve it by path.
    """
    should_install = handle.was_installed if handle is not None else True
    if not should_install:
        return

    fresh = find_addin(app, path)
    if fresh is None:
        fresh = ensure_addin_entry(app, path)

    def loaded() -> bool:
        return is_loaded(app, path)

    # Excel persists the installed flag in the registry, so it can report
    # Installed=True while the workbook is not actually loaded - this is the
    # normal state of an automation-launched instance, which deliberately
    # skips add-ins at startup. Assigning True again is then a no-op, so we
    # must toggle through False to force a genuine load.
    try:
        already_flagged = bool(fresh.com.Installed)
    except Exception:  # noqa: BLE001
        already_flagged = False

    if already_flagged and not loaded():
        with contextlib.suppress(Exception):
            set_installed(fresh, False)
        settle(0.2)

    try:
        set_installed(fresh, True)
    except Exception as exc:  # noqa: BLE001
        raise XlvbaError(
            f"Replaced {path.name} on disk but could not re-enable it: "
            f"{com_description(exc)}",
            "Re-tick it under File > Options > Add-ins > Excel Add-ins > Go.",
        ) from exc

    if not wait_until(loaded, timeout, interval=0.2):
        # The flag is set but Excel never opened it. Surface this rather than
        # pretending the reload succeeded.
        raise XlvbaError(
            f"{path.name} is marked installed but Excel did not load it.",
            "Check the add-in's Workbook_Open / Workbook_AddinInstall handlers "
            "for a runtime error, then reload manually.",
        )


def reload_hosts(
    path: Path,
    hosts: list[tuple[Any, AddinHandle | None]],
    *,
    timeout: float = DEFAULT_RELOAD_TIMEOUT,
) -> int:
    """Re-install the add-in in each instance that previously had it.

    Every instance is attempted even if one fails, so a single misbehaving
    Excel process cannot leave the others without their add-in.
    """
    restored = 0
    errors: list[str] = []

    for app, handle in hosts:
        if handle is None or not handle.was_installed:
            continue
        try:
            reload_addin(app, path, handle, timeout=timeout)
            restored += 1
        except Exception as exc:  # noqa: BLE001
            errors.append(com_description(exc))

    if errors and restored == 0:
        raise XlvbaError(
            f"Replaced {path.name} but could not re-enable it: {errors[0]}",
            "Re-tick it under File > Options > Add-ins > Excel Add-ins > Go.",
        )
    return restored


# ---------------------------------------------------------------------------
# Atomic file replacement
# ---------------------------------------------------------------------------


def replace_file(source: Path, destination: Path, *, keep_backup: bool = True) -> Path | None:
    """Replace ``destination`` with ``source``, keeping a rollback copy.

    ``os.replace`` is atomic on the same volume. When the source sits on a
    different volume we fall back to copy-then-replace via a sibling temp file
    so the destination is never left half-written.
    """
    backup: Path | None = None
    if keep_backup and destination.exists():
        backup = destination.with_suffix(destination.suffix + ".xlvba-bak")
        with contextlib.suppress(OSError):
            backup.unlink()
        shutil.copy2(destination, backup)

    try:
        staged = destination.with_suffix(destination.suffix + ".xlvba-tmp")
        with contextlib.suppress(OSError):
            staged.unlink()
        shutil.copy2(source, staged)
        os.replace(staged, destination)
    except OSError as exc:
        # Roll back so the user is never left without a working add-in.
        if backup is not None and backup.exists():
            with contextlib.suppress(OSError):
                shutil.copy2(backup, destination)
        raise FileLockedError(
            f"Could not replace {destination.name}: {exc}",
            "The file is still held open. Close Excel windows referencing it "
            "and retry.",
        ) from exc

    return backup


def discard_backup(backup: Path | None) -> None:
    """Delete a rollback copy created by :func:`replace_file`."""
    if backup is None:
        return
    with contextlib.suppress(OSError):
        backup.unlink()


@contextlib.contextmanager
def unloaded_addin(
    app: Any,
    path: Path,
    *,
    unlock_timeout: float = DEFAULT_UNLOCK_TIMEOUT,
    reload_timeout: float = DEFAULT_RELOAD_TIMEOUT,
):
    """Context manager wrapping the full unload / replace / reload dance.

    The ``finally`` block always attempts a reload, so an exception raised
    while the file is being swapped cannot leave the user's Excel session
    without its add-in.
    """
    hosts = unload_addin(app, path, timeout=unlock_timeout)
    try:
        yield hosts
    finally:
        with contextlib.suppress(Exception):
            reload_hosts(path, hosts, timeout=reload_timeout)


def settle(seconds: float = 0.15) -> None:
    """Brief pause letting Excel finish processing pending COM events."""
    time.sleep(seconds)
