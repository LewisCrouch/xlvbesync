"""Pre-flight diagnostics.

``xlvbesync doctor`` answers the question "why isn't this working?" before the
user has to guess. Every check reports pass/fail/warn plus a concrete fix.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path

from xlvbesync.core import console
from xlvbesync.core.addin import is_addin, is_file_locked
from xlvbesync.core.codec import system_ansi_encoding
from xlvbesync.core.config import ProjectConfig
from xlvbesync.core.errors import XlvbesyncError

OK = "ok"
WARN = "warn"
FAIL = "fail"

_SYMBOL = {OK: "[ ok ]", WARN: "[warn]", FAIL: "[fail]"}


@dataclass
class Check:
    name: str
    status: str
    detail: str = ""
    remedy: str = ""


def check_platform() -> Check:
    if sys.platform != "win32":
        return Check(
            "Platform",
            FAIL,
            f"{sys.platform} is not supported",
            "xlvbesync drives Excel through Windows COM and requires Windows.",
        )
    return Check("Platform", OK, f"Windows, Python {sys.version.split()[0]}")


def check_pywin32() -> Check:
    try:
        import win32com.client  # noqa: F401
        import pythoncom

        build = getattr(pythoncom, "__gen_path__", "")
        return Check("pywin32", OK, f"available{f' (gen: {build})' if build else ''}")
    except ImportError as exc:
        return Check(
            "pywin32",
            FAIL,
            str(exc),
            "pip install pywin32, then run "
            "`python .venv/Scripts/pywin32_postinstall.py -install`.",
        )


def check_watchdog() -> Check:
    try:
        import watchdog

        return Check("watchdog", OK, f"version {watchdog.version.VERSION}")
    except Exception:  # noqa: BLE001
        try:
            import watchdog  # noqa: F401

            return Check("watchdog", OK, "available")
        except ImportError as exc:
            return Check(
                "watchdog",
                WARN,
                str(exc),
                "pip install watchdog (only needed for `xlvbesync watch`).",
            )


def check_vbom_registry() -> Check:
    """Read AccessVBOM from the registry for every installed Office version."""
    try:
        import winreg
    except ImportError:  # pragma: no cover
        return Check("Trust access to VBA project", WARN, "registry unavailable")

    findings: list[str] = []
    enabled_anywhere = False
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office") as office:
            i = 0
            while True:
                try:
                    version = winreg.EnumKey(office, i)
                except OSError:
                    break
                i += 1
                if not version[0].isdigit():
                    continue
                sub = rf"Software\Microsoft\Office\{version}\Excel\Security"
                try:
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub) as key:
                        value, _ = winreg.QueryValueEx(key, "AccessVBOM")
                except OSError:
                    continue
                findings.append(f"Office {version}: AccessVBOM={value}")
                if int(value) == 1:
                    enabled_anywhere = True
    except OSError as exc:
        return Check("Trust access to VBA project", WARN, f"cannot read registry: {exc}")

    if not findings:
        return Check(
            "Trust access to VBA project",
            WARN,
            "setting not found in the registry",
            "Enable Excel > File > Options > Trust Center > Trust Center "
            "Settings > Macro Settings > 'Trust access to the VBA project "
            "object model'.",
        )

    if enabled_anywhere:
        return Check("Trust access to VBA project", OK, "; ".join(findings))

    return Check(
        "Trust access to VBA project",
        FAIL,
        "; ".join(findings),
        "Enable Excel > File > Options > Trust Center > Trust Center Settings "
        "> Macro Settings > 'Trust access to the VBA project object model', "
        "then fully restart Excel.",
    )


def check_excel_running() -> tuple[Check, object | None]:
    from xlvbesync.core.errors import ExcelNotRunningError
    from xlvbesync.core.excel import attach_excel

    try:
        app = attach_excel()
    except ExcelNotRunningError:
        return (
            Check(
                "Excel instance",
                WARN,
                "not running",
                "Start Excel, or pass --launch to let xlvbesync open it.",
            ),
            None,
        )
    except Exception as exc:  # noqa: BLE001
        return Check("Excel instance", FAIL, console.describe_exception(exc)), None

    try:
        version = str(app.Version)
        count = int(app.Workbooks.Count)
        detail = f"Excel {version}, {console.plural(count, 'workbook')} open"
    except Exception:  # noqa: BLE001
        detail = "attached"
    return Check("Excel instance", OK, detail), app


def check_vbom_live(app: object | None) -> Check:
    """Actually try to touch a VBProject - the only definitive test."""
    if app is None:
        return Check("VBProject access (live)", WARN, "skipped, Excel not running")

    from xlvbesync.core.errors import VBOMAccessDeniedError, com_description, is_vbom_denied
    from xlvbesync.core.excel import iter_workbooks
    from xlvbesync.core.vbproject import get_vbproject

    for wb in iter_workbooks(app):  # type: ignore[arg-type]
        try:
            project = get_vbproject(wb)
            _ = project.VBComponents.Count
            return Check(
                "VBProject access (live)",
                OK,
                f"verified against {wb.Name}",
            )
        except VBOMAccessDeniedError as exc:
            return Check("VBProject access (live)", FAIL, exc.message, exc.remedy or "")
        except Exception as exc:  # noqa: BLE001
            if is_vbom_denied(exc):
                return Check(
                    "VBProject access (live)",
                    FAIL,
                    com_description(exc),
                    VBOMAccessDeniedError.DEFAULT_REMEDY,
                )
            continue

    return Check(
        "VBProject access (live)",
        WARN,
        "no open workbook to test against",
        "Open a workbook in Excel and rerun doctor for a definitive check.",
    )


def check_encoding() -> Check:
    ansi = system_ansi_encoding()
    return Check(
        "Encoding",
        OK,
        f"repo files UTF-8; VBE code page {ansi}",
    )


def check_targets(config: ProjectConfig | None, app: object | None) -> list[Check]:
    """Verify each configured target exists, and report its live state."""
    if config is None:
        return [
            Check(
                "Project config",
                WARN,
                "no xlvbesync.json found",
                "Run `xlvbesync init --target <file.xlsm>` to create one.",
            )
        ]

    checks = [Check("Project config", OK, f"{len(config.targets)} target(s)")]

    from xlvbesync.core.excel import find_workbook

    for target in config.targets:
        path = config.target_path(target)
        label = f"Target '{target.resolved_name()}'"

        if not path.is_file():
            checks.append(
                Check(label, FAIL, f"missing: {path}", "Fix 'path' in xlvbesync.json.")
            )
            continue

        bits = [f"{path.name}", f"mode={target.mode}"]

        if app is not None:
            wb = find_workbook(app, path)  # type: ignore[arg-type]
            bits.append("open in Excel" if wb is not None else "not open")

        if is_addin(path):
            bits.append("add-in")
            if is_file_locked(path):
                bits.append("file locked (expected while loaded)")

        src = config.target_dir(target)
        bits.append(f"src={'present' if src.is_dir() else 'not extracted yet'}")

        checks.append(Check(label, OK, ", ".join(bits)))

    return checks


def run_doctor(config: ProjectConfig | None) -> tuple[list[Check], bool]:
    """Execute all checks. Returns ``(checks, healthy)``."""
    checks: list[Check] = [
        check_platform(),
        check_pywin32(),
        check_watchdog(),
        check_encoding(),
        check_vbom_registry(),
    ]

    excel_check, app = check_excel_running()
    checks.append(excel_check)
    checks.append(check_vbom_live(app))
    checks.extend(check_targets(config, app))

    healthy = not any(c.status == FAIL for c in checks)
    return checks, healthy


def print_checks(checks: list[Check]) -> None:
    """Render diagnostics."""
    for check in checks:
        text = f"{_SYMBOL[check.status]} {check.name}"
        if check.detail:
            text += f": {check.detail}"
        level = {FAIL: "error", WARN: "warn", OK: "info"}[check.status]
        # All on stdout so each remedy stays attached to its own check.
        console.line(text, level)
        if check.remedy and check.status != OK:
            console.line(f"       -> {check.remedy}", level)
