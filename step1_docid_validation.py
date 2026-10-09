"""
step2_docid_validation.py
=========================
STEP 2 — PDF Doc-ID Validation against Excel Master List

Responsibility
--------------
This script performs TWO checks on every PDF in the input folder:

  1. FILENAME FORMAT  — must be  <doc_id>_<MM-DD-YYYY>.pdf
  2. DOC-ID LOOKUP    — the extracted doc_id must exist in Column A
                        of  docid_list.xlsx  (placed in the project root)

It does NOT:
- Extract revision dates from PDF content
- Call the Gemini API
- Perform OCR
- Compare Excel revision dates

Validation rules
----------------
Both conditions must be satisfied for a PDF to be considered VALID:

  a) Filename strictly follows  <doc_id>_MM-DD-YYYY.pdf
       - Underscore separator between doc_id and date
       - Date in exactly MM-DD-YYYY format (two-digit month, two-digit day, four-digit year)
       - The date must be a valid calendar date (e.g. 02-30-2026 is rejected)
       - Extension must be exactly .pdf (case-insensitive, no double extensions)

  b) The doc_id extracted from the filename exists in Column A of docid_list.xlsx
       - Comparison is case-insensitive and strip-whitespace safe
       - The first row is skipped as a header (set SKIP_HEADER=False to disable)

Failure reasons (stored in the manifest)
-----------------------------------------
  NOT_A_PDF_FILE                             - file extension is not .pdf
  DOUBLE_PDF_EXTENSION                       - filename ends in .pdf.pdf
  EMPTY_FILENAME_STEM                        - nothing before the extension
  WRONG_SEPARATOR_DASH_INSTEAD_OF_UNDERSCORE
  SINGLE_DIGIT_MONTH_OR_DAY
  WRONG_DATE_ORDER_YYYY_MM_DD
  TWO_DIGIT_YEAR
  DATE_USES_UNDERSCORES_INSTEAD_OF_HYPHENS
  DATE_USES_DOTS_INSTEAD_OF_HYPHENS
  INVALID_FILENAME_FORMAT                    - generic pattern mismatch
  INVALID_CALENDAR_DATE                      - pattern matched but date is not real
  DOC_ID_NOT_IN_EXCEL                        - doc_id not found in docid_list.xlsx
  EXCEL_FILE_NOT_FOUND                       - docid_list.xlsx is missing
  EXCEL_READ_ERROR                           - could not read the Excel file

Output
------
  Valid PDFs   -> output/step_2/correct/
  Invalid PDFs -> output/step_2/wrong/
  Manifest     -> output/step_2/step2_manifest.json

Usage
-----
  1. Place  docid_list.xlsx  in the same folder as this script.
     Column A should contain Doc IDs (header row is optional).
  2. Place PDFs to validate inside  input/  (or update INPUT_FOLDER below).
  3. Run:
       python step2_docid_validation.py
"""

import logging
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

# ---------------------------------------------------------------------------
# CONFIGURATION  (edit these paths if your layout differs)
# ---------------------------------------------------------------------------

ROOT_DIR: Path = Path(__file__).resolve().parent

# Folder that contains the incoming PDFs to validate
INPUT_FOLDER: Path = ROOT_DIR / "input"

# Excel file that holds the authoritative Doc-ID list (Column A)
EXCEL_FILE: Path = ROOT_DIR / "docid_list.xlsx"

# Output destinations
OUTPUT_FOLDER: Path = ROOT_DIR / "output"
CORRECT_FOLDER: Path = OUTPUT_FOLDER / "correct"
WRONG_FOLDER: Path = OUTPUT_FOLDER / "wrong"

# Excel report path (replaces JSON manifest)
REPORT_PATH: Path = OUTPUT_FOLDER / "validation_report.xlsx"

# If True the very first row of Column A is skipped (treated as a header)
SKIP_HEADER: bool = True

# ---------------------------------------------------------------------------
# LOGGING SETUP
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  [%(levelname)-7s]  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("step2")

# ---------------------------------------------------------------------------
# FILENAME PATTERN  (identical rule as Step 1)
# doc_id = one or more chars that are NOT underscore
# date   = exactly MM-DD-YYYY
# ---------------------------------------------------------------------------

_FILENAME_PATTERN = re.compile(
    r"^(?P<doc_id>[^_]+)_(?P<date>\d{2}-\d{2}-\d{4})\.pdf$",
    re.IGNORECASE,
)

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------


def _parse_date(date_str: str) -> Optional[datetime]:
    """Return a datetime if date_str is a valid MM-DD-YYYY calendar date, else None."""
    try:
        return datetime.strptime(date_str, "%m-%d-%Y")
    except ValueError:
        return None


def _diagnose_format_failure(filename: str) -> str:
    """Return a specific error code explaining why the filename pattern failed."""
    stem = filename.lower().removesuffix(".pdf")

    if re.match(r"^[^_]+-\d{2}-\d{2}-\d{4}$", stem):
        return "WRONG_SEPARATOR_DASH_INSTEAD_OF_UNDERSCORE"
    if re.search(r"_\d-\d{2}-\d{4}$", stem) or re.search(r"_\d{2}-\d-\d{4}$", stem):
        return "SINGLE_DIGIT_MONTH_OR_DAY"
    if re.search(r"_\d{4}-\d{2}-\d{2}$", stem):
        return "WRONG_DATE_ORDER_YYYY_MM_DD"
    if re.search(r"_\d{2}-\d{2}-\d{2}$", stem):
        return "TWO_DIGIT_YEAR"
    if re.search(r"_\d{2}_\d{2}_\d{4}$", stem):
        return "DATE_USES_UNDERSCORES_INSTEAD_OF_HYPHENS"
    if re.search(r"_\d{2}\.\d{2}\.\d{4}$", stem):
        return "DATE_USES_DOTS_INSTEAD_OF_HYPHENS"
    return "INVALID_FILENAME_FORMAT"


def validate_filename(filename: str) -> tuple:
    """
    Validate the PDF filename format only (no Excel lookup here).

    Returns (is_valid, doc_id, date_str, reason)
    """
    name = filename.strip()

    if not name.lower().endswith(".pdf"):
        return False, None, None, "NOT_A_PDF_FILE"

    stem = name[:-4]
    if stem.lower().endswith(".pdf"):
        return False, None, None, "DOUBLE_PDF_EXTENSION"
    if not stem:
        return False, None, None, "EMPTY_FILENAME_STEM"

    match = _FILENAME_PATTERN.match(name)
    if not match:
        return False, None, None, _diagnose_format_failure(name)

    doc_id = match.group("doc_id")
    date_str = match.group("date")

    if _parse_date(date_str) is None:
        return False, doc_id, date_str, "INVALID_CALENDAR_DATE"

    return True, doc_id, date_str, ""


# ---------------------------------------------------------------------------
# EXCEL LOADER
# ---------------------------------------------------------------------------


def load_valid_doc_ids(excel_path: Path, skip_header: bool = SKIP_HEADER) -> tuple:
    """
    Read Column A (Doc ID) and the 'Agent Assigned' column from the Excel file.

    The agent column is found dynamically by scanning the header row for a cell
    whose text contains 'agent' (case-insensitive). Falls back to Column S (index 19)
    if no such header is found.

    Returns (doc_id_set, all_doc_ids_ordered, agent_map, error_reason)
      doc_id_set          - set of uppercase doc_ids for fast lookup
      all_doc_ids_ordered - list of original doc_ids in sheet order (for missing report)
      agent_map           - dict {doc_id_upper: agent_name} (empty string if no agent)
      error_reason        - empty string on success, error code string on failure
    """
    if not excel_path.exists():
        log.error("Excel file not found: %s", excel_path)
        return set(), [], {}, "EXCEL_FILE_NOT_FOUND"

    try:
        import openpyxl  # type: ignore
    except ImportError:
        log.error("openpyxl is not installed. Run:  pip install openpyxl")
        return set(), [], {}, "EXCEL_READ_ERROR"

    try:
        # Do NOT use read_only=True — in read-only mode openpyxl truncates row
        # tuples to the last non-empty cell, so a row with data only in col A
        # returns len==1 and row[18] raises IndexError (agent always blank).
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb.active

        doc_ids: set = set()
        all_doc_ids_ordered: list = []
        agent_map: dict = {}    # {DOC_ID_UPPER: "Agent Name"}

        # ── Step 1: find the agent column by scanning the header row ─────────
        # Default = column S (1-based col 19 → 0-based index 18)
        agent_col_idx: int = 18

        if skip_header:
            header_row = next(
                ws.iter_rows(min_row=1, max_row=1, values_only=True), None
            )
            if header_row:
                for i, cell_val in enumerate(header_row):
                    if cell_val and "agent" in str(cell_val).lower():
                        agent_col_idx = i
                        log.info(
                            "Found agent column at index %d  (header: '%s')",
                            i + 1, cell_val,
                        )
                        break
                else:
                    log.info(
                        "No 'Agent' header found — defaulting to column S (index 19)"
                    )

        # ── Step 2: read data rows ────────────────────────────────────────────
        start_row = 2 if skip_header else 1
        for row in ws.iter_rows(min_row=start_row, values_only=True):
            cell_value = row[0]          # Column A — Doc ID
            if cell_value is None:
                continue
            value = str(cell_value).strip()
            if not value:
                continue

            # Safely fetch agent cell; row may be shorter than agent_col_idx
            agent_cell = row[agent_col_idx] if len(row) > agent_col_idx else None
            agent_name = str(agent_cell).strip() if agent_cell is not None else ""

            doc_ids.add(value.upper())
            all_doc_ids_ordered.append(value)
            agent_map[value.upper()] = agent_name

        wb.close()
        log.info(
            "Loaded %d Doc ID(s) from Excel: %s", len(doc_ids), excel_path.name
        )
        return doc_ids, all_doc_ids_ordered, agent_map, ""

    except Exception as exc:
        log.error("Failed to read Excel file: %s", exc)
        return set(), [], {}, "EXCEL_READ_ERROR"




# ---------------------------------------------------------------------------
# FILE MOVER
# ---------------------------------------------------------------------------


def safe_move(src: Path, dest_folder: Path) -> Path:
    """
    Move *src* into *dest_folder*, creating the folder if necessary.
    If a file with the same name already exists, a numeric suffix is appended.
    Returns the final destination path.
    """
    dest_folder.mkdir(parents=True, exist_ok=True)
    dest = dest_folder / src.name

    if dest.exists():
        stem, suffix = src.stem, src.suffix
        counter = 1
        while dest.exists():
            dest = dest_folder / f"{stem}_{counter}{suffix}"
            counter += 1

    src.rename(dest)
    return dest


# ---------------------------------------------------------------------------
# ROUTING HELPER
# ---------------------------------------------------------------------------


def _route_file(
    pdf_path: Path,
    dest_folder: Path,
    filename: str,
    doc_id,
    date_str,
    status: str,
    reason: str,
    records: list,
    agent_name: str = "",
) -> None:
    """Move *pdf_path* to *dest_folder* and append a record to *records*."""
    try:
        dest = safe_move(pdf_path, dest_folder)
        records.append(
            {
                "original_filename": filename,
                "doc_id": doc_id or "",
                "filename_date": date_str or "",
                "agent_assigned": agent_name,
                "destination": str(dest),
                "status": status,
                "error_reason": reason,
            }
        )
    except (FileNotFoundError, PermissionError, OSError) as exc:
        log.error("%s | MOVE_FAILED: %s", filename, exc)
        records.append(
            {
                "original_filename": filename,
                "doc_id": doc_id or "",
                "filename_date": date_str or "",
                "agent_assigned": agent_name,
                "destination": "",
                "status": "MOVE_FAILED",
                "error_reason": f"FILE_MOVE_ERROR: {exc}",
            }
        )


# ---------------------------------------------------------------------------
# CORE PROCESSING
# ---------------------------------------------------------------------------


def _write_excel_report(
    valid_records: list,
    invalid_records: list,
    missing_doc_ids: list,
    report_path: Path,
) -> None:
    """
    Write an Excel workbook with three sheets:
      Sheet 1 - Correct  : all valid PDFs
      Sheet 2 - Wrong    : all invalid PDFs with error reasons
      Sheet 3 - Missing  : Doc IDs from Excel master list not found in input folder
    """
    try:
        import openpyxl  # type: ignore
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side  # type: ignore
    except ImportError:
        log.error("openpyxl is not installed. Run:  pip install openpyxl")
        return

    wb = openpyxl.Workbook()

    # ── Colour palette ────────────────────────────────────────────────────────
    GREEN_HEADER  = "1E7E34"   # dark green
    GREEN_ROW     = "D4EDDA"   # light green
    RED_HEADER    = "A71D2A"   # dark red
    RED_ROW       = "F8D7DA"   # light red
    ORANGE_HEADER = "9C4A00"   # dark orange
    ORANGE_ROW    = "FFE5CC"   # light orange
    WHITE         = "FFFFFF"

    thin_side = Side(style="thin", color="AAAAAA")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    def _header_font(color=WHITE):
        return Font(bold=True, color=color, name="Calibri", size=11)

    def _cell_font():
        return Font(name="Calibri", size=10)

    def _fill(hex_color):
        return PatternFill("solid", fgColor=hex_color)

    def _center():
        return Alignment(horizontal="center", vertical="center", wrap_text=True)

    def _left():
        return Alignment(horizontal="left", vertical="center", wrap_text=True)

    def _write_sheet(
        ws,
        headers: list,
        rows: list,
        header_color: str,
        row_color: str,
        col_widths: list,
    ):
        """Populate a worksheet with styled headers and data rows."""
        # Header row
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = _header_font()
            cell.fill = _fill(header_color)
            cell.alignment = _center()
            cell.border = thin_border

        ws.row_dimensions[1].height = 20

        # Data rows
        for row_idx, record in enumerate(rows, start=2):
            shade = _fill(row_color) if row_idx % 2 == 0 else _fill(WHITE)
            for col_idx, value in enumerate(record, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.font = _cell_font()
                cell.fill = shade
                cell.alignment = _left()
                cell.border = thin_border

        # Column widths
        for col_idx, width in enumerate(col_widths, start=1):
            ws.column_dimensions[
                openpyxl.utils.get_column_letter(col_idx)
            ].width = width

        # Freeze header row
        ws.freeze_panes = "A2"

    # ── Sheet 1: Correct ──────────────────────────────────────────────────────
    ws_correct = wb.active
    ws_correct.title = "Correct"
    correct_headers = ["#", "Original Filename", "Doc ID", "Date (MM-DD-YYYY)", "Agent Assigned", "Destination"]
    correct_rows = [
        (
            i + 1,
            r["original_filename"],
            r["doc_id"],
            r["filename_date"],
            r.get("agent_assigned", ""),
            r["destination"],
        )
        for i, r in enumerate(valid_records)
    ]
    _write_sheet(ws_correct, correct_headers, correct_rows, GREEN_HEADER, GREEN_ROW,
                 [5, 50, 25, 18, 22, 60])

    # ── Sheet 2: Wrong ────────────────────────────────────────────────────────
    ws_wrong = wb.create_sheet(title="Wrong")
    wrong_headers = ["#", "Original Filename", "Doc ID", "Date (MM-DD-YYYY)", "Agent Assigned", "Status", "Error Reason"]
    wrong_rows = [
        (
            i + 1,
            r["original_filename"],
            r["doc_id"],
            r["filename_date"],
            r.get("agent_assigned", ""),
            r["status"],
            r["error_reason"],
        )
        for i, r in enumerate(invalid_records)
    ]
    _write_sheet(ws_wrong, wrong_headers, wrong_rows, RED_HEADER, RED_ROW,
                 [5, 50, 25, 18, 22, 12, 40])

    # ── Sheet 3: Missing ──────────────────────────────────────────────────────
    ws_missing = wb.create_sheet(title="Missing")
    missing_headers = ["#", "Doc ID (from Master List)", "Agent Assigned", "Remark"]
    missing_rows = [
        (i + 1, item["doc_id"], item.get("agent_assigned", ""), "No matching file found in input folder")
        for i, item in enumerate(missing_doc_ids)
    ]
    _write_sheet(ws_missing, missing_headers, missing_rows, ORANGE_HEADER, ORANGE_ROW,
                 [5, 35, 22, 45])

    # ── Summary tab strip colours (tab colour) ───────────────────────────────
    ws_correct.sheet_properties.tabColor  = "1E7E34"
    ws_wrong.sheet_properties.tabColor    = "A71D2A"
    ws_missing.sheet_properties.tabColor  = "E07B00"

    # Save
    report_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(report_path)
    log.info("Excel report saved: %s", report_path)


def process_pdf_folder(
    input_folder: Path = INPUT_FOLDER,
    excel_file: Path = EXCEL_FILE,
    correct_folder: Path = CORRECT_FOLDER,
    wrong_folder: Path = WRONG_FOLDER,
) -> dict:
    """
    Scan *input_folder* for PDFs, validate each one, and route to the
    appropriate output folder.

    Validation order
    ----------------
    1. Filename format check  (same rules as Step 1)
    2. Doc-ID Excel lookup    (doc_id must be in Column A of docid_list.xlsx)

    A file must pass BOTH checks to land in *correct_folder*.

    Returns summary dict with keys: ``valid_count``, ``invalid_count``,
    ``missing_count``, ``total_discovered``
    """
    log.info("=" * 70)
    log.info("STEP 2 - Doc-ID Validation Started")
    log.info("Input  folder : %s", input_folder)
    log.info("Excel  file   : %s", excel_file)
    log.info("Correct folder: %s", correct_folder)
    log.info("Wrong   folder: %s", wrong_folder)
    log.info("=" * 70)

    # Load Excel Doc IDs (set for fast lookup + ordered list + agent map)
    valid_doc_ids, all_doc_ids_ordered, agent_map, excel_error = load_valid_doc_ids(excel_file)

    if excel_error:
        log.warning(
            "Cannot load Doc IDs from Excel (%s). All files will be marked invalid.",
            excel_error,
        )

    # Discover PDFs
    if not input_folder.exists():
        log.error("Input folder does not exist: %s", input_folder)
        return {"valid_count": 0, "invalid_count": 0, "missing_count": 0, "total_discovered": 0}

    pdf_files = sorted(input_folder.glob("*.pdf")) + sorted(input_folder.glob("*.PDF"))

    # Deduplicate (Windows is case-insensitive but glob may return both)
    seen: set = set()
    unique_pdfs: list = []
    for p in pdf_files:
        if p.name.lower() not in seen:
            seen.add(p.name.lower())
            unique_pdfs.append(p)

    total = len(unique_pdfs)
    log.info("Discovered %d PDF file(s) in input folder", total)

    valid_records: list = []
    invalid_records: list = []
    # Track which Doc IDs were actually matched by an input file
    matched_doc_ids: set = set()

    for pdf_path in unique_pdfs:
        filename = pdf_path.name
        log.info("--- Processing: %s", filename)

        # Check 1: filename format
        fmt_ok, doc_id, date_str, fmt_reason = validate_filename(filename)

        if not fmt_ok:
            log.info("  FILENAME_INVALID | reason=%s", fmt_reason)
            _route_file(
                pdf_path, wrong_folder,
                filename, doc_id, date_str, "INVALID", fmt_reason,
                invalid_records,
                agent_name="",
            )
            continue

        # Check 2: Doc-ID in Excel
        if excel_error:
            log.info("  DOC_ID_CHECK_SKIPPED | %s", excel_error)
            _route_file(
                pdf_path, wrong_folder,
                filename, doc_id, date_str, "INVALID", excel_error,
                invalid_records,
                agent_name="",
            )
            continue

        doc_id_upper = doc_id.upper()
        if doc_id_upper not in valid_doc_ids:
            log.info("  DOC_ID_NOT_IN_EXCEL | doc_id=%s", doc_id)
            _route_file(
                pdf_path, wrong_folder,
                filename, doc_id, date_str, "INVALID", "DOC_ID_NOT_IN_EXCEL",
                invalid_records,
                agent_name="",
            )
            continue

        # Both checks passed — record the matched doc_id
        matched_doc_ids.add(doc_id_upper)
        agent_name = agent_map.get(doc_id_upper, "")
        log.info("  VALID | doc_id=%s | date=%s | agent=%s", doc_id, date_str, agent_name)
        _route_file(
            pdf_path, correct_folder,
            filename, doc_id, date_str, "VALID", "",
            valid_records,
            agent_name=agent_name,
        )

    # ── Build Missing list ────────────────────────────────────────────────────
    # Doc IDs in Excel master list that had no corresponding file in input folder
    missing_doc_ids: list = [
        {"doc_id": doc_id, "agent_assigned": agent_map.get(doc_id.upper(), "")}
        for doc_id in all_doc_ids_ordered
        if doc_id.upper() not in matched_doc_ids
    ]
    log.info("Missing Doc IDs (in Excel but no file in input): %d", len(missing_doc_ids))

    # Summary
    valid_count   = len(valid_records)
    invalid_count = len(invalid_records)
    missing_count = len(missing_doc_ids)

    log.info("=" * 70)
    log.info(
        "STEP 2 COMPLETE - Total: %d | Valid: %d | Invalid: %d | Missing: %d",
        total, valid_count, invalid_count, missing_count,
    )
    log.info("=" * 70)

    # Write Excel report (3 sheets)
    _write_excel_report(valid_records, invalid_records, missing_doc_ids, REPORT_PATH)

    return {
        "total_discovered": total,
        "valid_count": valid_count,
        "invalid_count": invalid_count,
        "missing_count": missing_count,
    }


# ---------------------------------------------------------------------------
# ENTRY POINT
# ---------------------------------------------------------------------------


def main() -> None:
    """Run Step 2 Doc-ID validation."""
    result = process_pdf_folder()
    print(
        f"\n[STEP 2 COMPLETE]\n"
        f"  Total PDFs discovered : {result.get('total_discovered', 0)}\n"
        f"  ✅ Valid   (format OK + doc_id found in Excel) : {result.get('valid_count', 0)}"
        f"  ->  output/correct/\n"
        f"  ❌ Invalid (bad format OR doc_id NOT in Excel) : {result.get('invalid_count', 0)}"
        f"  ->  output/wrong/\n"
        f"  🔍 Missing (doc_id in Excel but no file found): {result.get('missing_count', 0)}\n"
        f"\n  📊 Excel report saved to: output/validation_report.xlsx\n"
        f"     (Sheets: 'Correct', 'Wrong', 'Missing')\n"
    )


if __name__ == "__main__":
    main()
