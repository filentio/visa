from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


class ExtractError(RuntimeError):
    pass


class EncryptedWorkbookError(ExtractError):
    pass


class NoVBAFoundError(ExtractError):
    pass


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


_INVALID_FILENAME_RE = re.compile(r'[<>:"/\\|?*\x00-\x1F]')


def safe_filename(name: str, *, fallback: str = "unnamed", max_len: int = 160) -> str:
    """
    Make a filesystem-safe filename component (no extension).
    Keeps Unicode (incl. Cyrillic), only replaces invalid/special characters.
    """
    n = _INVALID_FILENAME_RE.sub("_", name).strip()
    # Windows quirk: trailing dots/spaces are not allowed
    n = n.rstrip(" .")
    if not n:
        n = fallback
    if len(n) > max_len:
        n = n[:max_len].rstrip(" .")
        if not n:
            n = fallback
    return n


def normalize_newlines(text: str) -> str:
    # oletools may return CRLF; keep repo-friendly LF
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    if not text.endswith("\n"):
        text += "\n"
    return text


def _ensure_unique_path(path: Path, used: Dict[str, int]) -> Path:
    key = str(path).lower()
    if key not in used and not path.exists():
        used[key] = 1
        return path
    used[key] = used.get(key, 1) + 1
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    candidate = parent / f"{stem}__{used[key]}{suffix}"
    while str(candidate).lower() in used or candidate.exists():
        used[key] += 1
        candidate = parent / f"{stem}__{used[key]}{suffix}"
    used[str(candidate).lower()] = 1
    return candidate


def try_build_workbook_map_openpyxl(xlsm_path: Path) -> Dict[str, Any]:
    import openpyxl

    wb = openpyxl.load_workbook(
        filename=str(xlsm_path),
        read_only=True,
        data_only=True,
        keep_vba=False,
    )
    try:
        sheets = [{"index": i + 1, "name": name} for i, name in enumerate(wb.sheetnames)]

        named_ranges: List[Dict[str, Any]] = []
        try:
            defined_names = wb.defined_names  # type: ignore[attr-defined]
            # `defined_names` is iterable over names, indexable by name.
            for name in defined_names:
                try:
                    dn = defined_names[name]
                except Exception:
                    continue

                entry: Dict[str, Any] = {"name": name}
                for attr in ("localSheetId", "comment", "hidden"):
                    if hasattr(dn, attr):
                        entry[attr] = getattr(dn, attr)

                # openpyxl stores the raw formula/reference in attr_text
                if hasattr(dn, "attr_text"):
                    entry["attr_text"] = getattr(dn, "attr_text")

                dests: List[Dict[str, str]] = []
                if hasattr(dn, "destinations"):
                    try:
                        for sheet_title, ref in dn.destinations:
                            dests.append({"sheet": sheet_title, "ref": ref})
                    except Exception:
                        pass
                if dests:
                    entry["destinations"] = dests

                named_ranges.append(entry)
        except Exception:
            named_ranges = []

        return {
            "source_file": xlsm_path.name,
            "generated_at": _utc_now_iso(),
            "method": "openpyxl",
            "sheets": sheets,
            "named_ranges": named_ranges,
        }
    finally:
        try:
            wb.close()
        except Exception:
            pass


def try_build_workbook_map_xml(xlsm_path: Path) -> Dict[str, Any]:
    import zipfile

    from lxml import etree

    NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ns = {"m": NS_MAIN, "r": NS_R}

    with zipfile.ZipFile(xlsm_path) as zf:
        raw = zf.read("xl/workbook.xml")

    root = etree.fromstring(raw)  # noqa: S320 (trusted local file)

    sheets: List[Dict[str, Any]] = []
    for i, el in enumerate(root.xpath("./m:sheets/m:sheet", namespaces=ns)):
        sheets.append(
            {
                "index": i + 1,
                "name": el.get("name"),
                "sheetId": el.get("sheetId"),
                "rId": el.get(f"{{{NS_R}}}id"),
            }
        )

    named_ranges: List[Dict[str, Any]] = []
    for el in root.xpath("./m:definedNames/m:definedName", namespaces=ns):
        named_ranges.append(
            {
                "name": el.get("name"),
                "localSheetId": el.get("localSheetId"),
                "hidden": el.get("hidden"),
                "comment": el.get("comment"),
                # Keep exact text as in XML (no strip) to preserve spaces/special chars.
                "text": el.text if el.text is not None else "",
            }
        )

    return {
        "source_file": xlsm_path.name,
        "generated_at": _utc_now_iso(),
        "method": "xml(xl/workbook.xml)",
        "sheets": sheets,
        "named_ranges": named_ranges,
    }


def build_workbook_map(xlsm_path: Path) -> Dict[str, Any]:
    try:
        return try_build_workbook_map_openpyxl(xlsm_path)
    except Exception as e_openpyxl:
        try:
            data = try_build_workbook_map_xml(xlsm_path)
            data["openpyxl_error"] = repr(e_openpyxl)
            return data
        except Exception as e_xml:
            return {
                "source_file": xlsm_path.name,
                "generated_at": _utc_now_iso(),
                "method": "failed",
                "error": {
                    "openpyxl": repr(e_openpyxl),
                    "xml": repr(e_xml),
                },
                "sheets": [],
                "named_ranges": [],
            }


def workbook_map_to_md(data: Dict[str, Any]) -> str:
    lines: List[str] = []
    lines.append(f"## Workbook map: {data.get('source_file', '')}")
    lines.append("")
    lines.append(f"- Generated at (UTC): {data.get('generated_at', '')}")
    lines.append(f"- Method: {data.get('method', '')}")
    if data.get("method") == "xml(xl/workbook.xml)" and data.get("openpyxl_error"):
        lines.append(f"- Note: openpyxl failed, fallback used: `{data.get('openpyxl_error')}`")
    if data.get("method") == "failed":
        err = data.get("error") or {}
        lines.append("- ERROR: could not parse workbook structure")
        lines.append(f"  - openpyxl: `{err.get('openpyxl', '')}`")
        lines.append(f"  - xml: `{err.get('xml', '')}`")
    lines.append("")

    sheets = data.get("sheets") or []
    lines.append(f"### Sheets ({len(sheets)})")
    if not sheets:
        lines.append("")
        lines.append("_No sheet info available._")
    else:
        for s in sheets:
            name = s.get("name")
            if name is None:
                name = ""
            extra = []
            if s.get("sheetId") is not None:
                extra.append(f"sheetId={s.get('sheetId')}")
            if s.get("rId") is not None:
                extra.append(f"rId={s.get('rId')}")
            suffix = f" ({', '.join(extra)})" if extra else ""
            lines.append(f"- {s.get('index')}. `{name}`{suffix}")

    named_ranges = data.get("named_ranges") or []
    lines.append("")
    lines.append(f"### Named ranges ({len(named_ranges)})")
    if not named_ranges:
        lines.append("")
        lines.append("_No named ranges found or could not extract._")
    else:
        for nr in named_ranges:
            name = nr.get("name") or ""
            if "destinations" in nr and nr["destinations"]:
                dests = ", ".join(
                    f"{d.get('sheet','')}!{d.get('ref','')}" for d in nr.get("destinations", [])
                )
                lines.append(f"- `{name}` → {dests}")
            elif "attr_text" in nr and nr["attr_text"] is not None:
                lines.append(f"- `{name}` → `{nr.get('attr_text')}`")
            elif "text" in nr and nr["text"] is not None:
                text = nr.get("text")
                lines.append(f"- `{name}` → `{text}`")
            else:
                lines.append(f"- `{name}`")

    lines.append("")
    return "\n".join(lines)


def _olevba_is_encrypted(vba_parser: Any) -> bool:
    # oletools API differs slightly between versions; try a few known shapes.
    for attr in ("detect_is_encrypted", "is_encrypted"):
        if not hasattr(vba_parser, attr):
            continue
        val = getattr(vba_parser, attr)
        try:
            res = val() if callable(val) else val
            return bool(res)
        except Exception:
            continue
    return False


@dataclass(frozen=True)
class ExtractedFile:
    kind: str  # modules/classes/forms
    original_name: str
    output_path: Path


def extract_vba_modules(xlsm_path: Path, out_dir: Path) -> Tuple[List[ExtractedFile], List[Tuple[Any, ...]]]:
    try:
        from oletools.olevba import VBA_Parser  # type: ignore[import-not-found]
    except Exception as e:
        raise ExtractError(
            "oletools is not installed. Install dependencies: pip install -r tools/excel_extract/requirements.txt"
        ) from e

    try:
        parser = VBA_Parser(str(xlsm_path))
    except Exception as e:
        msg = str(e).lower()
        if "encrypt" in msg or "password" in msg:
            raise EncryptedWorkbookError(
                "Файл выглядит зашифрованным/парольным, oletools не может прочитать VBA. "
                "Нужна незашифрованная копия .xlsm."
            ) from e
        raise ExtractError(f"Не удалось открыть файл через oletools: {e}") from e

    extracted: List[ExtractedFile] = []
    analysis: List[Tuple[Any, ...]] = []
    used_paths: Dict[str, int] = {}
    try:
        if _olevba_is_encrypted(parser):
            raise EncryptedWorkbookError(
                "Файл зашифрован/парольный, oletools (olevba) не может извлечь VBA. "
                "Нужна незашифрованная копия .xlsm."
            )

        has_vba = False
        try:
            has_vba = bool(parser.detect_vba_macros())
        except Exception:
            # if detect fails, we'll still attempt extraction
            has_vba = False

        try:
            analysis = list(parser.analyze_macros())  # type: ignore[attr-defined]
        except Exception:
            analysis = []

        any_written = False
        for (_, _stream_path, vba_filename, vba_code) in parser.extract_macros():
            # vba_filename often includes extension (.bas/.cls/.frm)
            vba_filename = str(vba_filename)
            ext = Path(vba_filename).suffix.lower()
            base = Path(vba_filename).stem

            if ext == ".cls":
                kind = "classes"
            elif ext == ".frm":
                kind = "forms"
            else:
                # default to module
                kind = "modules"
                if ext not in (".bas", ".cls", ".frm"):
                    ext = ".bas"

            safe_base = safe_filename(base, fallback="module")
            out_path = out_dir / kind / f"{safe_base}{ext}"
            out_path = _ensure_unique_path(out_path, used_paths)

            text = vba_code.decode("utf-8", errors="replace") if isinstance(vba_code, bytes) else str(vba_code)
            out_path.write_text(normalize_newlines(text), encoding="utf-8")

            extracted.append(ExtractedFile(kind=kind, original_name=vba_filename, output_path=out_path))
            any_written = True

        if not any_written:
            if has_vba:
                raise ExtractError("VBA обнаружен, но извлечь модули не удалось (пустой результат).")
            raise NoVBAFoundError("VBA-макросы не найдены в файле (no VBA found).")

        return extracted, analysis
    finally:
        try:
            parser.close()
        except Exception:
            pass


def write_olevba_report(
    *,
    xlsm_path: Path,
    out_dir: Path,
    extracted: List[ExtractedFile],
    analysis: List[Tuple[Any, ...]],
    error: Optional[BaseException],
) -> None:
    report_dir = out_dir / "meta"
    report_dir.mkdir(parents=True, exist_ok=True)

    lines: List[str] = []
    lines.append("olevba_report (brief)")
    lines.append("")
    lines.append(f"Source: {xlsm_path.name}")
    lines.append(f"Generated at (UTC): {_utc_now_iso()}")
    lines.append("")

    if error is not None:
        lines.append("Status: ERROR")
        lines.append(f"Error: {type(error).__name__}: {error}")
    else:
        lines.append("Status: OK")
    lines.append("")

    lines.append(f"Extracted files: {len(extracted)}")
    for item in extracted:
        rel = item.output_path.relative_to(out_dir)
        lines.append(f"- {item.kind}: {item.original_name} -> {rel.as_posix()}")

    lines.append("")
    lines.append(f"Macro analysis rows: {len(analysis)}")
    # Common tuple layout: (type, keyword, description)
    for row in analysis[:200]:
        try:
            parts = [str(x) for x in row]
            lines.append("- " + " | ".join(parts))
        except Exception:
            continue
    if len(analysis) > 200:
        lines.append(f"... truncated ({len(analysis) - 200} more rows)")
    lines.append("")

    (report_dir / "olevba_report.txt").write_text("\n".join(lines), encoding="utf-8")


def prepare_output_dir(out_dir: Path) -> None:
    resolved = out_dir.resolve()
    cwd = Path.cwd().resolve()
    if resolved in (Path("/"), cwd, cwd.parent):
        raise ExtractError(
            f"Unsafe --out path: {out_dir!s}. Refusing to delete/create this directory. "
            "Use a dedicated folder, e.g. --out extracted_vba"
        )

    if out_dir.exists():
        shutil.rmtree(out_dir)
    (out_dir / "modules").mkdir(parents=True, exist_ok=True)
    (out_dir / "classes").mkdir(parents=True, exist_ok=True)
    (out_dir / "forms").mkdir(parents=True, exist_ok=True)
    (out_dir / "meta").mkdir(parents=True, exist_ok=True)


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Extract VBA modules from .xlsm into text files + workbook map.")
    p.add_argument("--input", required=True, help='Path to .xlsm (supports Cyrillic), e.g. "шаблон 12.09.xlsm"')
    p.add_argument("--out", required=True, help="Output directory, e.g. extracted_vba")
    return p.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> int:
    args = parse_args(argv)
    xlsm_path = Path(args.input)
    out_dir = Path(args.out)

    if not xlsm_path.exists():
        print(f"ERROR: input file not found: {xlsm_path}", file=sys.stderr)
        return 2

    prepare_output_dir(out_dir)

    # Workbook map is useful even if VBA is missing; write it first.
    workbook_map = build_workbook_map(xlsm_path)
    (out_dir / "workbook_map.json").write_text(
        json.dumps(workbook_map, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    (out_dir / "workbook_map.md").write_text(workbook_map_to_md(workbook_map), encoding="utf-8")

    extracted: List[ExtractedFile] = []
    analysis: List[Tuple[Any, ...]] = []
    err: Optional[BaseException] = None
    try:
        extracted, analysis = extract_vba_modules(xlsm_path, out_dir)
    except BaseException as e:
        err = e

    write_olevba_report(
        xlsm_path=xlsm_path,
        out_dir=out_dir,
        extracted=extracted,
        analysis=analysis,
        error=err,
    )

    if err is not None:
        print(f"ERROR: {err}", file=sys.stderr)
        if isinstance(err, EncryptedWorkbookError):
            print("Hint: Нужна незашифрованная/без пароля копия .xlsm.", file=sys.stderr)
        return 2

    print(f"OK: extracted {len(extracted)} VBA files into {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

