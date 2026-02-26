from __future__ import annotations

import argparse
import datetime as dt
import traceback
from pathlib import Path
from typing import Any, Dict, Optional

from lib.excel_app import (
    ExcelAutomationError,
    ExcelSession,
    XL_CALC_AUTOMATIC,
    open_excel_workbook,
    cleanup_excel_session,
)
from lib.io import RunnerIOError, ensure_dir, read_payload, short_exc, write_json_stdout, copy_file, is_windows
from lib.pdf_export import PdfExportError, export_sheets_to_pdf
from lib.png_assets import AssetsError, PreparedAssets, prepare_assets


class PayloadError(RuntimeError):
    pass


XL_UP = -4162  # xlUp


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Windows Excel Runner (pywin32): fill xlsm, run macro, export PDFs")
    p.add_argument("--payload", required=True, help="Path to JSON payload, or '-' for stdin")
    return p.parse_args(argv)


def _require(obj: Dict[str, Any], key: str) -> Any:
    if key not in obj:
        raise PayloadError(f"Missing key in payload: {key}")
    return obj[key]


def _req_str(obj: Dict[str, Any], key: str) -> str:
    v = _require(obj, key)
    if not isinstance(v, str) or not v:
        raise PayloadError(f"Expected non-empty string for payload.{key}")
    return v


def _req_num(obj: Dict[str, Any], key: str) -> float:
    v = _require(obj, key)
    if isinstance(v, (int, float)):
        return float(v)
    raise PayloadError(f"Expected number for payload.{key}")


def _parse_iso_date(s: str, *, field: str) -> dt.datetime:
    try:
        d = dt.date.fromisoformat(s)
        return dt.datetime(d.year, d.month, d.day)
    except Exception as e:
        raise PayloadError(f"Invalid ISO date for {field}: {s!r} (expected YYYY-MM-DD)") from e


def prepare_work_dir(template_path: Path, work_dir: Path) -> Dict[str, Path]:
    ensure_dir(work_dir)
    output_dir = work_dir / "output"
    assets_dir = work_dir / "assets"
    ensure_dir(output_dir)
    ensure_dir(assets_dir)

    workbook_path = work_dir / "template.xlsm"
    copy_file(template_path, workbook_path)

    return {
        "work_dir": work_dir.resolve(),
        "workbook_path": workbook_path.resolve(),
        "output_dir": output_dir.resolve(),
        "assets_dir": assets_dir.resolve(),
    }


def fill_input_sheet(workbook: object, payload: Dict[str, Any]) -> None:
    """
    Fill sheet 'ввод данных' with payload values.
    """
    ws = workbook.Worksheets("ввод данных")

    client = _require(payload, "client")
    job = _require(payload, "job")
    company = _require(payload, "company")

    full_name = _req_str(client, "full_name")
    passport_no = _req_str(client, "passport_no")
    dob = _parse_iso_date(_req_str(client, "dob"), field="client.dob")
    address = _req_str(client, "address")
    country_display = _req_str(client, "country_display")

    currency_symbol = _req_str(job, "currency_symbol")
    fx_rate = _req_num(job, "fx_rate")
    salary_rub = _req_num(job, "salary_rub")
    position = _req_str(job, "position")
    contract_start_date = _parse_iso_date(_req_str(job, "contract_start_date"), field="job.contract_start_date")
    contract_number = _req_str(job, "contract_number")

    selected_company_name = _req_str(company, "selected_company_name")

    # Required cells per spec
    ws.Range("C2").Value = full_name
    ws.Range("C3").Value = passport_no
    ws.Range("C4").Value = dob
    ws.Range("C5").Value = address
    ws.Range("C6").Value = country_display
    ws.Range("C7").Value = fx_rate
    ws.Range("B8").Value = currency_symbol
    ws.Range("C11").Value = selected_company_name
    ws.Range("C15").Value = salary_rub
    ws.Range("C16").Value = position
    ws.Range("C17").Value = contract_start_date
    ws.Range("C20").Value = contract_number

    # VBA (Module1.bas) reads company name from 'ввод данных'!B11.
    # Set both B11 and C11 to be safe.
    ws.Range("B11").Value = selected_company_name

    # Optional: if C18 is not a formula, set it to contract_number.
    try:
        f = str(ws.Range("C18").Formula or "")
        if not f.startswith("="):
            ws.Range("C18").Value = contract_number
    except Exception:
        pass


def update_company_assets_paths(workbook: object, company_name: str, assets: PreparedAssets) -> None:
    """
    Update sheet 'компании' row for selected company with absolute asset paths.

    IMPORTANT: This mapping follows the extracted VBA in `extracted_vba/modules/Module1.bas`:
      - col D (4): stamp/seal path
      - col E (5): logo path
      - col F (6): director sign path
      - col G (7): client sign path
    """
    ws = workbook.Worksheets("компании")

    # Find last row in col A
    last = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    target_row = None
    for r in range(2, int(last) + 1):
        v = ws.Cells(r, 1).Value
        if v is None:
            continue
        if str(v).strip().lower() == company_name.strip().lower():
            target_row = r
            break

    if target_row is None:
        raise PayloadError(f"Company not found in sheet 'компании' column A: {company_name!r}")

    ws.Cells(target_row, 4).Value = str(assets.seal_path)  # D: stamp/seal
    ws.Cells(target_row, 5).Value = str(assets.logo_path)  # E: logo
    ws.Cells(target_row, 6).Value = str(assets.director_sign_path)  # F
    ws.Cells(target_row, 7).Value = str(assets.client_sign_path)  # G


def run_vba_entrypoint(excel_app: object, macro_name: str) -> None:
    try:
        excel_app.Run(macro_name)
    except Exception as e:
        raise ExcelAutomationError(f"Failed to run macro {macro_name!r}: {e}") from e


def build_export_plan(payload: Dict[str, Any]) -> Dict[str, str]:
    exp = _require(payload, "export")
    output_files = _require(exp, "output_files")

    contract_sheet = _req_str(exp, "contract_template")
    bank_sheet = _req_str(exp, "bank_template")
    insurance_sheet = _req_str(exp, "insurance_template")
    salary_sheet = _req_str(exp, "salary_template")

    contract_pdf = _req_str(output_files, "contract")
    bank_pdf = _req_str(output_files, "bank")
    insurance_pdf = _req_str(output_files, "insurance")
    salary_pdf = _req_str(output_files, "salary")

    return {
        contract_sheet: contract_pdf,
        bank_sheet: bank_pdf,
        insurance_sheet: insurance_pdf,
        salary_sheet: salary_pdf,
    }


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)

    if not is_windows():
        write_json_stdout(
            {
                "status": "error",
                "message": "This runner works only on Windows (requires Microsoft Excel installed).",
                "details": "",
            }
        )
        return 2

    session: Optional[ExcelSession] = None
    work_dir_path: Optional[Path] = None
    try:
        payload = read_payload(args.payload)

        template_path = Path(_req_str(payload, "template_path"))
        work_dir = Path(_req_str(payload, "work_dir"))
        work_dir_path = work_dir
        if not template_path.exists():
            raise PayloadError(f"template_path not found: {template_path}")

        paths = prepare_work_dir(template_path, work_dir)

        company = _require(payload, "company")
        company_name = _req_str(company, "selected_company_name")
        assets_payload = _require(company, "assets")
        assets = prepare_assets(paths["work_dir"], assets_payload)

        session = open_excel_workbook(
            paths["workbook_path"],
            visible=False,
            display_alerts=False,
            screen_updating=False,
            enable_events=False,
            ask_to_update_links=False,
            calculation=XL_CALC_AUTOMATIC,
        )
        # Expose Excel PID for external watchdog/timeout handling.
        try:
            if session.pid:
                (paths["work_dir"] / "excel_pid.txt").write_text(str(session.pid), encoding="utf-8")
        except Exception:
            pass

        wb = session.workbook
        excel = session.excel

        fill_input_sheet(wb, payload)
        update_company_assets_paths(wb, company_name, assets)

        # Ensure formulas/links are in a consistent state before macro run.
        try:
            wb.RefreshAll()
        except Exception:
            pass
        try:
            excel.CalculateFull()
        except Exception:
            pass

        run_vba_entrypoint(excel, "Module1.UpdateAllStampsFromConfig")

        try:
            wb.RefreshAll()
        except Exception:
            pass
        try:
            excel.CalculateFull()
        except Exception:
            pass

        export_plan = build_export_plan(payload)
        res = export_sheets_to_pdf(workbook=wb, output_dir=paths["output_dir"], sheet_to_filename=export_plan)

        write_json_stdout(
            {
                "status": "ok",
                "output_dir": str(res.output_dir),
                "pdf_files": res.pdf_files,
            }
        )
        return 0
    except (RunnerIOError, PayloadError, AssetsError, ExcelAutomationError, PdfExportError) as e:
        write_json_stdout(
            {
                "status": "error",
                "message": str(e),
                "details": short_exc(e),
            }
        )
        return 2
    except BaseException as e:
        tb = traceback.format_exc(limit=50)
        if len(tb) > 8000:
            tb = tb[:8000] + "\n...(truncated)"
        write_json_stdout(
            {
                "status": "error",
                "message": "Unhandled error",
                "details": tb,
            }
        )
        return 2
    finally:
        cleanup_excel_session(session)
        try:
            if work_dir_path is not None:
                pid_file = (work_dir_path / "excel_pid.txt")
                if pid_file.exists():
                    pid_file.unlink()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())

