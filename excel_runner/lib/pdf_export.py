from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple

from .io import ensure_dir


class PdfExportError(RuntimeError):
    pass


XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0


@dataclass(frozen=True)
class PdfExportResult:
    output_dir: Path
    pdf_files: List[str]


def export_sheets_to_pdf(
    *,
    workbook: object,
    output_dir: Path,
    sheet_to_filename: Dict[str, str],
) -> PdfExportResult:
    """
    Export individual worksheets to PDF without using VBA.

    `sheet_to_filename`: {worksheet_name: output_filename.pdf}
    """
    ensure_dir(output_dir)

    exported: List[str] = []
    missing: List[str] = []

    for sheet_name, out_name in sheet_to_filename.items():
        if not sheet_name:
            continue
        if not out_name:
            continue

        try:
            ws = workbook.Worksheets(sheet_name)
        except Exception:
            missing.append(sheet_name)
            continue

        pdf_path = (output_dir / out_name).resolve()
        try:
            ws.ExportAsFixedFormat(
                Type=XL_TYPE_PDF,
                Filename=str(pdf_path),
                Quality=XL_QUALITY_STANDARD,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
        except Exception as e:
            raise PdfExportError(f"Failed to export sheet '{sheet_name}' to PDF: {e}") from e

        exported.append(out_name)

    if missing:
        raise PdfExportError(f"Sheets not found in workbook: {', '.join(missing)}")

    return PdfExportResult(output_dir=output_dir.resolve(), pdf_files=exported)

