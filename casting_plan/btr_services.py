"""BTR upload inspection and CAMP workbook generation.

The service intentionally keeps the supplied macro-enabled CAMP template as the
output container.  Sheet1 is rebuilt with the selected planning material while
Sheet2 receives the uploaded source rows (minus ``Bal Bill Qty``) for audit.
"""

from __future__ import annotations

from collections import defaultdict
from copy import copy
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


REQUIRED_HEADERS = {
    "Planning Mat",
    "Sec Wt",
    "Thick",
    "Ext Grade",
    "Bal Roll Qty",
    "Length",
    "Min Len",
    "Max Len",
    "Embossing",
}

OUTPUT_HEADERS = [
    "Sales Office", "SOTy", "Age SO", "SO Date", "SO No.", "SOItm",
    "Sold To Party Name", "Ship to Party Name", "Ship to Party City", "Material",
    "Sec Wt", "Ext Grade", "Length", "Order Qty", "BTR", "Min Len",
    "Max Len", "Insp By", "Tole%", "FI Rel Sta", "SBU Released",
    "Planning Mat"
]

OUTPUT_SOURCE_MAP = {
    "Sales Office": "Sales Office", "SOTy": "SOTy", "Age SO": "Age SO",
    "SO Date": "SO Date", "SO No.": "SO No.", "SOItm": "SOItm",
    "Sold To Party Name": "Sold To Party Name", "Ship to Party Name": "Ship To Party Name",
    "Ship to Party City": "Ship To Party City", "Material": "Material",
    "Sec Wt": "Sec Wt", "Ext Grade": "Ext Grade", "Length": "Length",
    "Order Qty": "Order Qty", "BTR": "Bal Roll Qty", "Min Len": "Min Len",
    "Max Len": "Max Len", "Insp By": "Insp By", "Tole%": "Tole%",
    "FI Rel Sta": "FI Rel Sta", "SBU Released": "SBU Released",
    "Planning Mat": "Planning Mat"
}


@dataclass
class BTRInspection:
    sheet_name: str
    header_row: int
    row_count: int
    planning_materials: list[dict[str, Any]]
    preview_records: list[dict[str, Any]] = None


@dataclass
class CampaignResult:
    output_path: Path
    input_row_count: int
    output_row_count: int
    skipped_row_count: int
    zero_length_resolved_count: int
    total_btr_qty: float
    bucket_counts: dict[str, int]


def _clean_header(value: Any) -> str:
    return " ".join(str(value or "").replace("\n", " ").split())


def _number(value: Any, default: float = 0.0) -> float:
    if value in (None, ""):
        return default
    try:
        return float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return default


def _find_btr_sheet(workbook):
    for sheet in workbook.worksheets:
        for row_number in range(1, min(sheet.max_row, 20) + 1):
            headers = [_clean_header(sheet.cell(row_number, col).value) for col in range(1, sheet.max_column + 1)]
            if REQUIRED_HEADERS.issubset(set(headers)):
                return sheet, row_number, headers
    required = ", ".join(sorted(REQUIRED_HEADERS))
    raise ValueError(f"BTR header row not found. Required columns: {required}")


def _read_source(path: Path):
    workbook = load_workbook(path, data_only=True, read_only=True)
    sheet, header_row, headers = _find_btr_sheet(workbook)
    rows = []
    for values in sheet.iter_rows(min_row=header_row + 1, values_only=True):
        row = {header: values[index] if index < len(values) else None for index, header in enumerate(headers) if header}
        if any(value not in (None, "") for value in row.values()):
            rows.append(row)
    workbook.close()
    return sheet.title, header_row, headers, rows


def inspect_btr_workbook(path: Path) -> BTRInspection:
    sheet_name, header_row, _headers, rows = _read_source(Path(path))
    counts: dict[str, int] = defaultdict(int)
    quantities: dict[str, float] = defaultdict(float)
    for row in rows:
        material = str(row.get("Planning Mat") or "").strip()
        if material:
            counts[material] += 1
            quantities[material] += _number(row.get("Bal Roll Qty"))
    materials = [
        {"name": name, "row_count": counts[name], "btr_qty": round(quantities[name], 3)}
        for name in sorted(counts, key=str.casefold)
    ]
    if not materials:
        raise ValueError("No Planning Mat values were found in the uploaded BTR file.")
        
    preview_records = []
    for r in rows[:15]:
        so_date = str(r.get("SO Date") or r.get("SODate") or "").strip()
        if " " in so_date:
            so_date = so_date.split(" ")[0]
        req_date = str(r.get("Req. Deli. Date") or r.get("ReqDeliDate") or "").strip()
        if " " in req_date:
            req_date = req_date.split(" ")[0]
        preview_records.append({
            "Sales_Office": str(r.get("Sales Office") or r.get("SalesOffice") or "-").strip(),
            "SOTy": str(r.get("SOTy") or r.get("SO Type") or "-").strip(),
            "Age_SO": str(r.get("Age SO") or r.get("AgeSO") or "-").strip(),
            "Age_FI_2": str(r.get("Age FI(2)") or r.get("AgeFI2") or "-").strip(),
            "Type_of_Customer": str(r.get("Type of Customer") or "-").strip(),
            "Material_group_5": str(r.get("Material group 5") or "-").strip(),
            "SO_Date": so_date or "-",
            "Req_Deli_Date": req_date or "-",
            "CO": str(r.get("CO") or "-").strip(),
            "Remarks": str(r.get("Remarks") or "-").strip(),
        })
    return BTRInspection(sheet_name, header_row, len(rows), materials, preview_records)


def _grade_bucket(row: dict[str, Any]) -> str:
    marker = f"{row.get('Emb*') or ''} {row.get('Embossing') or ''}".upper()
    if "HT" in marker or "JSHT" in marker or "HIGH TENSILE" in marker:
        return "HT"
    if " MS" in f" {marker}" or "JSMS" in marker or "MILD STEEL" in marker:
        return "MS"
    return "PANTHER"


def _prepare_rows(source_rows: list[dict[str, Any]], planning_material: str):
    selected = []
    input_count = 0
    skipped = 0
    resolved = 0
    for row in source_rows:
        if str(row.get("Planning Mat") or "").strip().casefold() != planning_material.casefold():
            continue
        input_count += 1
        length = _number(row.get("Length"))
        if length <= 0:
            min_length = _number(row.get("Min Len"))
            max_length = _number(row.get("Max Len"))
            if min_length > 0 and max_length > 0:
                length = max_length
                resolved += 1
            else:
                skipped += 1
                continue
        btr_qty = _number(row.get("Bal Roll Qty"))
        if btr_qty <= 0:
            skipped += 1
            continue
            
        sec_weight = _number(row.get("Sec Wt"))
        weight_per_piece = sec_weight * length / 1000 if sec_weight > 0 else 0
        pieces = btr_qty / weight_per_piece if weight_per_piece > 0 else None
        
        # New Rule: Skip if No. of Pcs is less than 0.9
        if pieces is not None and pieces < 0.9:
            skipped += 1
            continue
            
        # Clean Emb* column (JINDAL STEEL PG HT -> HT, JINDAL STEEL PG MS -> MS)
        raw_emb_star = str(row.get("Emb*") or "").strip().upper()
        raw_emb_star = raw_emb_star.replace("PG HT", "HT").replace("PG MS", "MS")
        clean_emb_star = " ".join(raw_emb_star.split())
        
        prepared = dict(row)
        # Update the row with the cleaned Emb* value so it flows to the Excel output
        prepared["Emb*"] = clean_emb_star
        
        prepared.update({
            "_bucket": _grade_bucket(row),
            "_length": length,
            "_btr": btr_qty,
            "_weight_per_piece": weight_per_piece,
            "_pieces": pieces,
            "_thickness": _number(row.get("Thick"), float("inf")),
            "_sec_wt": sec_weight,
        })
        selected.append(prepared)

    # Determine primary sort: Thickness for Angles, Section Weight for others
    is_angle = planning_material.strip().upper().startswith("AL_")
    
    # Sorting: Decreasing order of Sec Wt, BTR, Length
    selected.sort(key=lambda row: (
        -row["_sec_wt"],
        -row["_btr"],
        -row["_length"],
        -(row["_pieces"] if row["_pieces"] is not None else -1),
        str(row.get("Embossing") or "").casefold(),
    ))
    return selected, input_count, skipped, resolved

def _copy_header_style(sheet, template_sheet=None):
    gray_fill = PatternFill("solid", fgColor="D9D9D9")
    dark_bold_font = Font(color="000000", bold=True, size=10, name="Calibri")
    thin = Border(*(Side(style="thin", color="A6A6A6") for _ in range(4)))
    for col, heading in enumerate(OUTPUT_HEADERS, 1):
        cell = sheet.cell(1, col, heading)
        cell.fill = gray_fill
        cell.font = dark_bold_font
        cell.border = thin
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.row_dimensions[1].height = 28


def _write_source_sheet(sheet, headers: list[str], rows: list[dict[str, Any]]):
    if sheet.max_row:
        sheet.delete_rows(1, sheet.max_row)
    audit_headers = [header for header in headers if header and header != "Bal Bill Qty"]
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    thin = Side(style="thin", color="B7C9D6")
    for col, header in enumerate(audit_headers, 1):
        cell = sheet.cell(1, col, header)
        cell.font = Font(bold=True, color="17365D")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row_number, row in enumerate(rows, 2):
        for col, header in enumerate(audit_headers, 1):
            sheet.cell(row_number, col, row.get(header))
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(audit_headers))}{max(1, len(rows) + 1)}"


def _formula_sum(column: str, row_numbers: list[int]) -> str:
    if not row_numbers:
        return "=0"
    return "=SUM(" + ",".join(f"{column}{row}" for row in row_numbers) + ")"


def generate_campaign_workbook(
    btr_path: Path,
    template_path: Path,
    output_path: Path,
    planning_material: str,
) -> CampaignResult:
    """Generate a sorted CAMP workbook for one planning material."""
    _sheet_name, _header_row, source_headers, source_rows = _read_source(Path(btr_path))
    rows, input_count, skipped, resolved = _prepare_rows(source_rows, planning_material.strip())
    if input_count == 0:
        raise ValueError(f"Planning material '{planning_material}' was not found in the BTR file.")
    if not rows:
        raise ValueError("No plannable rows remain after applying length and positive-BTR rules.")

    template_path = Path(template_path)
    if not template_path.exists():
        raise ValueError("CAMP PLAN FORMAT.xlsm template is missing from data/reference.")
    workbook = load_workbook(template_path, keep_vba=True)
    if "Sheet1" not in workbook.sheetnames:
        raise ValueError("CAMP template must contain Sheet1.")
    sheet = workbook["Sheet1"]

    # Preserve template column widths and header appearance, then replace its old plan.
    header_styles = [copy(sheet.cell(1, col)._style) for col in range(1, min(76, sheet.max_column) + 1)]
    if sheet.max_row:
        sheet.delete_rows(1, sheet.max_row)
    _copy_header_style(sheet)
    for col, style in enumerate(header_styles, 1):
        sheet.cell(1, col)._style = style
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(OUTPUT_HEADERS))}1"

    border_side = Side(style="thin", color="B8C2CC")
    data_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
    output_row = 2
    serial = 1
    bucket_counts: dict[str, int] = {}
    for row in rows:
        bucket_counts[row["_bucket"]] = bucket_counts.get(row["_bucket"], 0) + 1

    for row in rows:
        current = output_row
        for col, output_header in enumerate(OUTPUT_HEADERS, 1):
            if output_header == "Length":
                value = row.get("_length") or row.get("Length") or row.get("Cutting Length")
            elif output_header == "BTR":
                value = row.get("_btr") or row.get("Bal Roll Qty") or row.get("BTR")
            elif output_header == "Planning Mat":
                value = row.get("Planning Mat") or planning_material
            else:
                src_key = OUTPUT_SOURCE_MAP.get(output_header, output_header)
                value = row.get(src_key, "")
            cell = sheet.cell(current, col, value)
            cell.border = data_border
            if output_header in ("Sec Wt", "Length", "Order Qty", "BTR"):
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        output_row += 1

    sheet.auto_filter.ref = f"A1:{get_column_letter(len(OUTPUT_HEADERS))}{max(1, output_row - 1)}"

    audit_sheet = workbook["Sheet2"] if "Sheet2" in workbook.sheetnames else workbook.create_sheet("Sheet2")
    _write_source_sheet(audit_sheet, source_headers, source_rows)
    if "HOLD" in workbook.sheetnames:
        hold = workbook["HOLD"]
        if hold.max_row > 1:
            hold.delete_rows(2, hold.max_row - 1)

    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    workbook.close()

    return CampaignResult(
        output_path=output_path,
        input_row_count=input_count,
        output_row_count=len(rows),
        skipped_row_count=skipped,
        zero_length_resolved_count=resolved,
        total_btr_qty=round(sum(row["_btr"] for row in rows), 3),
        bucket_counts=bucket_counts,
    )
