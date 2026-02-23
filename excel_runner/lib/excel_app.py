from __future__ import annotations

import gc
import os
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Optional


class ExcelAutomationError(RuntimeError):
    pass


XL_CALC_AUTOMATIC = -4105
XL_CALC_MANUAL = -4135

MSO_AUTOMATION_SECURITY_LOW = 1


@dataclass
class ExcelSession:
    """
    Thin wrapper around Excel COM application and workbook with robust cleanup.
    """

    excel: object
    workbook: object
    pid: Optional[int] = None

    def close(self, save_changes: bool = False) -> None:
        try:
            self.workbook.Close(SaveChanges=save_changes)
        except Exception:
            pass

    def quit(self) -> None:
        try:
            self.excel.Quit()
        except Exception:
            pass


def _require_windows() -> None:
    if os.name != "nt":
        raise ExcelAutomationError("This runner works only on Windows (requires Microsoft Excel COM).")


def _get_excel_pid(excel_app: object) -> Optional[int]:
    """
    Best-effort: map Excel.Application.Hwnd -> process id.
    """
    try:
        import win32process  # type: ignore
    except Exception:
        return None

    try:
        hwnd = excel_app.Hwnd
        _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
        return int(pid)
    except Exception:
        return None


def _taskkill(pid: int) -> None:
    """
    Best-effort kill of a single Excel instance by PID (and its children).
    """
    try:
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            check=False,
            capture_output=True,
            text=True,
        )
    except Exception:
        pass


def _pid_exists(pid: int) -> bool:
    try:
        res = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}"],
            check=False,
            capture_output=True,
            text=True,
        )
        out = (res.stdout or "") + (res.stderr or "")
        return str(pid) in out
    except Exception:
        return False


def open_excel_workbook(
    workbook_path: Path,
    *,
    visible: bool = False,
    display_alerts: bool = False,
    screen_updating: bool = False,
    enable_events: bool = False,
    ask_to_update_links: bool = False,
    calculation: int = XL_CALC_AUTOMATIC,
) -> ExcelSession:
    """
    Create isolated Excel instance (DispatchEx) and open workbook.
    """
    _require_windows()

    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except Exception as e:
        raise ExcelAutomationError(
            "pywin32 is not installed. Install: pip install -r excel_runner/requirements.txt"
        ) from e

    pythoncom.CoInitialize()

    excel = None
    wb = None
    pid: Optional[int] = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")

        pid = _get_excel_pid(excel)

        # Try to lower macro security for this automation session (best-effort).
        try:
            excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
        except Exception:
            pass

        excel.Visible = bool(visible)
        excel.DisplayAlerts = bool(display_alerts)
        excel.ScreenUpdating = bool(screen_updating)
        excel.EnableEvents = bool(enable_events)
        try:
            excel.AskToUpdateLinks = bool(ask_to_update_links)
        except Exception:
            pass
        try:
            excel.Calculation = int(calculation)
        except Exception:
            pass

        wb = excel.Workbooks.Open(str(workbook_path), UpdateLinks=0, ReadOnly=False)
        return ExcelSession(excel=excel, workbook=wb, pid=pid)
    except Exception as e:
        # If creation/open fails, still attempt to quit and uninit COM.
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        if pid and _pid_exists(pid):
            _taskkill(pid)
        raise ExcelAutomationError(f"Failed to start Excel/open workbook: {e}") from e


def cleanup_excel_session(session: Optional[ExcelSession]) -> None:
    """
    Close workbook and quit Excel, guaranteeing COM cleanup.
    If Excel hangs, kill only the created Excel PID (best-effort).
    """
    if session is None:
        return

    pid = session.pid
    excel = session.excel
    wb = session.workbook

    try:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass
    finally:
        # Release references and uninitialize COM
        session.workbook = None  # type: ignore[assignment]
        session.excel = None  # type: ignore[assignment]
        wb = None
        excel = None
        gc.collect()
        try:
            import pythoncom  # type: ignore

            pythoncom.CoUninitialize()
        except Exception:
            pass

        if pid and _pid_exists(pid):
            # Last resort: kill only our Excel instance (if still alive).
            _taskkill(pid)

