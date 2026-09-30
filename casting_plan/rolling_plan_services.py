"""Generate a rolling-plan workbook from a generated CAMP workbook."""
from __future__ import annotations

import math
import re
from collections import defaultdict, deque
from copy import copy
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from django.conf import settings
from .casting_plan import generate_casting_plan_file

DEFAULT_BILLET_LENGTH_M = 11.65
DEFAULT_COOLING_BED_LIMIT_M = 86.0  
DEFAULT_RM_ALLOWANCE = 0.07

BILLET_SECTION_WEIGHTS = {
    "BL_165X165": 214.0,
    "BL_250X250": 490.0,
    "BL_130X130": 132.0,
    "BB_355X280X90": 480.0,
}

FINAL_HEADERS = [
    "Section", "Thick ness", "Sec Wt", "Ext Grade", "R.M. GRADE",
    "ROLLING QTY", "BILLET LENGTH (MTR)", "RM QTY", "BILLET SECTION WT.", "BLOOM (PCS)", 
    "BLOOM CUM", "ROLLED PCS", "ROLLED LENGTH", "COOLING BED LENGTH", 
    "Cutting plan", "CUTTING LENGTH", "EMBOSSING", "CUSTOMER", "DESTINATION", 
    "ORDER REF NO", "Insp By", "SO Type", "REMARKS", "SO No.", "SO Itm", 
    "MIN", "MAX", "Age SO",
]

RULE_HEADERS = [
    "Campaign Plan", "Campaign Plan", "Campaign Plan", "Campaign Plan", "Grade Mapping",
    "BTR", "Algorithmic Output", "Formula", "Assumption", "Formula", "Cumulative",
    "Formula", "Formula", "Formula: strictly <= 86", "Formula", "Cutting Length", 
    "Campaign Plan", "Sold To Party", "Ship To City", "Campaign Plan CO",
    "Inspection Agency", "SO Type", "S.O.Remark", "SO No.", "SO Item", 
    "Min Len", "Max Len", "Age",
]

@dataclass
class RollingPlanResult:
    output_path: Path
    row_count: int
    rolling_qty: float
    rm_qty: float
    bloom_count: int
    yield_pct: float
    max_cooling_bed_length: float
    inferred_grade_count: int


def _clean_header(value: Any) -> str:
    return " ".join(str(value or "").replace("\n", " ").split())


def _number(value: Any, default: float = 0.0) -> float:
    if value in (None, ""):
        return default
    try:
        return float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return default


def _normalise_grade(value: Any) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


RAW_GRADE_RANK = {
    _normalise_grade(grade): rank
    for rank, grade in enumerate((
        "C18HMn", "C18Mn1(EN)", "C18H1Mn(HT)", "HT(V)",
        "A572Gr50", "A572 Gr60", "E500STR08", "E460STR07",
        "Gr880", "55C11", "1080HH", "R260", "R260(IRS)",
        "E410SHP", "FR Fe-490", "WR-Fe 480B", "E450STR05",
    ), 1)
}


def _raw_grade_rank(value: Any) -> int:
    """Return the business ranking shown in the rolling-plan reference."""
    key = _normalise_grade(value)
    key = {
        "ASTMA572GR50": "A572GR50",
        "ASTMA572GR60": "A572GR60",
    }.get(key, key)
    if key in RAW_GRADE_RANK:
        return RAW_GRADE_RANK[key]
    for ranked_grade in sorted(RAW_GRADE_RANK, key=len, reverse=True):
        if ranked_grade in key or key in ranked_grade:
            return RAW_GRADE_RANK[ranked_grade]
    return 999


def _is_higher_section(billet_section_weight: float) -> bool:
    return _number(billet_section_weight) > 300


def _read_grade_map(path: Path):
    workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    workbook.close()
    mapping = {}
    for row in rows[1:]:
        if not row or row[0] in (None, ""):
            continue
        grade = _normalise_grade(row[0])
        std = _normalise_grade(row[3] if len(row) > 3 else "")
        mapping[(grade, std)] = {
            "higher": str(row[1] or "").strip(),
            "lighter": str(row[2] or "").strip(),
            "std": str(row[3] or "").strip(),
        }
    return rows, mapping


def _select_grade_mapping(mapping, ext_grade: str, std: str):
    grade_key = _normalise_grade(ext_grade)
    std_key = _normalise_grade(std)
    exact = mapping.get((grade_key, std_key))
    if exact:
        return exact, False
    candidates = [value for (grade, _), value in mapping.items() if grade == grade_key]
    unique = {(value["higher"], value["lighter"], value["std"]): value for value in candidates}
    if len(unique) == 1:
        return next(iter(unique.values())), False
    aliases = [
        (mapped_grade, value)
        for (mapped_grade, mapped_std), value in mapping.items()
        if (not std_key or mapped_std == std_key)
        and (mapped_grade in grade_key or grade_key in mapped_grade)
    ]
    if aliases:
        return max(aliases, key=lambda item: len(item[0]))[1], False
    return None, True


def _remarks_suffix(value: Any) -> str:
    text = str(value or "").strip()
    return text.rsplit("-", 1)[-1].strip() if text else "UNSPECIFIED"


def _rolling_suffix(value: Any, fallback_std: str = "") -> str:
    text = str(value or "").strip().upper()
    tokens = [re.sub(r"[^A-Z0-9/]", "", token) for token in text.split("-")]
    if "BSEN" in tokens:
        return "BSEN"
    if "BIS" in tokens:
        return "BIS"
    if "ASTM" in tokens:
        index = tokens.index("ASTM")
        return "".join(tokens[index:])
    return _remarks_suffix(value) if text else (fallback_std or "UNSPECIFIED")


def _normalise_section(value: Any) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper().replace("×", "X"))


def _display_raw_grade(value: Any) -> str:
    """Format a Rail Mill raw grade for the rolling-plan output."""
    grade = re.sub(r"\s+", "", str(value or "").strip().upper())
    if grade.startswith("A572"):
        grade = f"ASTM{grade}"
    return grade


def _raw_material_display(rm_size: Any, raw_grade: Any) -> str:
    """Build values such as BL_ASTMA572GR50_285X390."""
    size = str(rm_size or "").strip().upper().replace("×", "X")
    prefix_match = re.match(r"\s*(BL|BB)", size)
    prefix = prefix_match.group(1) if prefix_match else ""
    dimensions_source = size[prefix_match.end():] if prefix_match else size
    dimensions = re.sub(r"[^0-9X]", "", dimensions_source)
    grade = _display_raw_grade(raw_grade)

    if prefix and grade and dimensions:
        return f"{prefix}_{grade}_{dimensions}"
    if grade and dimensions:
        return f"{grade}_{dimensions}"
    return re.sub(r"\s+", "_", size)


def _select_calculation_master(master_rows, planning_material, material, ext_grade, sec_wt):
    """Resolve the most specific Rail Mill Calculation Master record."""
    planning_key = _normalise_section(planning_material)
    material_key = _normalise_section(material)
    grade_key = _normalise_grade(ext_grade)
    section_weight = _number(sec_wt, float("nan"))

    candidates = [
        master
        for master in master_rows
        if _normalise_section(master.planning_mat) == planning_key
    ]

    if material_key:
        exact_material = [
            master
            for master in candidates
            if _normalise_section(master.material) == material_key
        ]
        if exact_material:
            return exact_material[0]

    exact_grade = [
        master
        for master in candidates
        if _normalise_grade(master.ext_grade) == grade_key
    ]
    if math.isfinite(section_weight):
        exact_weight = [
            master
            for master in exact_grade
            if master.sec_wt is not None
            and abs(float(master.sec_wt) - section_weight) < 0.001
        ]
        if exact_weight:
            return exact_weight[0]

    return exact_grade[0] if exact_grade else None


def _is_angle_section(planning_material: Any) -> bool:
    return _normalise_section(planning_material).startswith("AL")


def _fourth_sort_value(planning_material: Any, thickness: Any, section_weight: Any) -> float:
    return _number(thickness, float("inf")) if _is_angle_section(planning_material) else _number(section_weight, float("inf"))


def _raw_sizes(value: Any) -> list[str]:
    text = re.sub(r"\s+", "", str(value or "").upper())
    return [raw_size for raw_size in BILLET_SECTION_WEIGHTS if raw_size in text]


def _read_raw_material_map(path: Path) -> tuple[list[tuple[str, str]], list[list[Any]]]:
    workbook = load_workbook(path, data_only=True, read_only=True)
    if "Product List" not in workbook.sheetnames:
        workbook.close()
        raise ValueError("Raw-material workbook must contain the Product List sheet.")
    sheet = workbook["Product List"]
    source_rows = list(sheet.iter_rows(values_only=True))
    workbook.close()
    mappings: list[tuple[str, str]] = []
    for ordinal_col, section_col in ((0, 1), (3, 4), (6, 7), (9, 10), (12, 13)):
        current_sizes: list[str] = []
        for row in source_rows:
            ordinal = row[ordinal_col] if ordinal_col < len(row) else None
            section = row[section_col] if section_col < len(row) else None
            size_cell = row[section_col + 1] if section_col + 1 < len(row) else None
            sizes = _raw_sizes(size_cell)
            if sizes:
                current_sizes = sizes
            if isinstance(ordinal, (int, float)) and section and current_sizes:
                mappings.append((_normalise_section(section), current_sizes[0]))
    return mappings, source_rows


def _resolve_raw_material(planning_material: str, mappings: list[tuple[str, str]]) -> str:
    section = _normalise_section(planning_material)
    matches = [(key, raw) for key, raw in mappings if section == key or section.startswith(key + "X")]
    if matches:
        return max(matches, key=lambda item: len(item[0]))[1]
    if any(token in section for token in ("305", "356", "406", "400")):
        return "BB_355X280X90"
    if any(token in section for token in ("254", "250", "203", "200", "180", "152")):
        return "BL_250X250"
    return "BL_165X165"


def _extract_std(remarks: Any, grade_mapping: dict[str, str] | None = None) -> str:
    text = _normalise_grade(remarks)
    for std in ("BSEN", "BIS"):
        if std in text:
            return std
    return str((grade_mapping or {}).get("std") or "").strip()


def _campaign_rows(sheet):
    headers = [_clean_header(cell.value) for cell in sheet[1]]
    index = {header: i + 1 for i, header in enumerate(headers) if header}
    required = {"Planning Mat", "Sec Wt", "Thick ness", "Grade", "Cutting Length", "BTR"}
    missing = sorted(required - set(index))
    if missing:
        raise ValueError(f"Campaign Sheet1 is missing required columns: {', '.join(missing)}")

    rows = []
    for row_number in range(2, sheet.max_row + 1):
        row = {header: sheet.cell(row_number, column).value for header, column in index.items()}
        if _number(row.get("BTR")) <= 0 or not str(row.get("Planning Mat") or "").strip():
            continue
        row["_source_row"] = row_number
        row["_thickness"] = _number(row.get("Thick ness"), float("inf"))
        rows.append(row)
    if not rows:
        raise ValueError("The campaign workbook has no positive-BTR detail rows in Sheet1.")
    return rows


def _casting_match_key(row: dict[str, Any]):
    """Build a stable key shared by campaign and casting-plan Sheet1 rows."""
    return (
        str(_source(row, "SO No.", "SO_No", default="")).strip(),
        str(_source(row, "SO Item", "SOItm", "SO_Item", default="")).strip(),
        str(_source(row, "Material", default="")).strip(),
        round(_number(_source(row, "Cutting Length", "Length", default=0)), 4),
        round(_number(_source(row, "BTR", default=0)), 4),
    )


def _read_casting_sheet_rows(casting_plan_path: Path | None):
    """Read hard-coded Sheet1 inputs/results from the previously generated casting plan."""
    result = defaultdict(deque)
    if not casting_plan_path or not Path(casting_plan_path).is_file():
        return result
    casting_book = load_workbook(casting_plan_path, data_only=True, read_only=True)
    try:
        if "Sheet1" not in casting_book.sheetnames:
            return result
        casting_sheet = casting_book["Sheet1"]
        headers = [_clean_header(cell.value) for cell in casting_sheet[1]]
        index = {header: column for column, header in enumerate(headers, 1) if header}
        for row_number in range(2, casting_sheet.max_row + 1):
            row = {
                header: casting_sheet.cell(row_number, column).value
                for header, column in index.items()
            }
            if _number(row.get("BTR")) > 0:
                result[_casting_match_key(row)].append(row)
    finally:
        casting_book.close()
    return result


def _source(row: dict[str, Any], *headers: str, default=None):
    for header in headers:
        value = row.get(header)
        if value not in (None, ""):
            return value
    return default


def _add_grade_mapping_sheet(workbook, grade_rows):
    if "Grade Mapping" in workbook.sheetnames:
        del workbook["Grade Mapping"]
    sheet = workbook.create_sheet("Grade Mapping")
    for row_index, row in enumerate(grade_rows, 1):
        for col_index, value in enumerate(row[:5], 1):
            sheet.cell(row_index, col_index, value)
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="D97706")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center")
    widths = [30, 25, 25, 12, 38]
    for col, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(col)].width = width
    sheet.freeze_panes = "A2"


def _add_raw_material_sheet(workbook, source_rows):
    if "Raw Material Mapping" in workbook.sheetnames:
        del workbook["Raw Material Mapping"]
    sheet = workbook.create_sheet("Raw Material Mapping")
    sheet.append(["Raw material size", "Sectional weight (kg/m)"])
    for raw_size, weight in BILLET_SECTION_WEIGHTS.items():
        sheet.append([raw_size, weight])
    sheet.append([])
    sheet.append(["Source: Product List in RAW MATERIAL LENGTH TOLERANCE.xlsx"])
    for row in source_rows:
        sheet.append(list(row[:6]))
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="D97706")
        cell.font = Font(bold=True, color="FFFFFF")
    sheet.column_dimensions["A"].width = 34
    sheet.column_dimensions["B"].width = 28
    sheet.freeze_panes = "A2"


def _add_assumptions_sheet(workbook, plan_date: date, camp_number: str, bl_settings):
    if "Plan Assumptions" in workbook.sheetnames:
        del workbook["Plan Assumptions"]
    sheet = workbook.create_sheet("Plan Assumptions")
    
    cold_saw = float(getattr(bl_settings, 'cold_saw_cut', 0.0) or 0.0)
    cb_crop = float(getattr(bl_settings, 'cooling_bed_crop', 0.4) or 0.4)
    div_sheer = float(getattr(bl_settings, 'div_sheer_loss', 0.0) or 0.0)
    base_crop = float(getattr(bl_settings, 'base_crop_loss', 0.0) or 0.0)
    sec_tol = float(getattr(bl_settings, 'section_wt_tolerance_pct', 0.5) or 0.5) / 100.0
    hot_pcs = int(getattr(bl_settings, 'hot_pieces_at_cb', 1) or 1)

    rows = [
        ["Parameter", "Value", "Basis / status"],
        ["Billet length (m)", "Calculated Dynamically", "Formula driven in Final Plan"],
        ["Billet sectional weight", "Mapped by raw material", "BL_165=214; BL_250=490; BL_130=132; BB=480 kg/m"],
        ["Maximum cooling-bed length (m)", 86.0, "Final value must be strictly below this limit"],
        ["RM allowance", 0.07, "Used in cumulative bloom allocation"],
        ["Raw-grade basis", "Billet sectional weight", ">300 kg/m Higher; otherwise Lighter"],
        ["Plan date", plan_date, "Entered when rolling plan was generated"],
        ["Camp number", camp_number or "TBD", "Entered when rolling plan was generated"],
        ["", "", ""],
        ["Dynamic Billet Formula Variables", "", ""],
        ["Cold Saw Cut", cold_saw, "Mapped to $B$11"],
        ["Cooling Bed Crop", cb_crop, "Mapped to $B$12"],
        ["Div Sheer Loss", div_sheer, "Mapped to $B$13"],
        ["Base Crop Loss", base_crop, "Mapped to $B$14"],
        ["Sec WT Tolerance %", sec_tol, "Mapped to $B$15"],
        ["Hot Pieces @ CB", hot_pcs, "Mapped to $B$16"],
    ]
    for row in rows:
        sheet.append(row)
        
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="D97706")
        cell.font = Font(bold=True, color="FFFFFF")
        
    for row in range(2, 9):
        sheet.cell(row, 1).font = Font(bold=True)
        sheet.cell(row, 2).fill = PatternFill("solid", fgColor="FEF3C7")
        
    sheet.cell(10, 1).font = Font(bold=True)
    for row in range(11, 17):
        sheet.cell(row, 1).font = Font(bold=True)
        sheet.cell(row, 2).fill = PatternFill("solid", fgColor="E0F2FE")
        
    sheet["B5"].number_format = "0.0%"
    sheet["B15"].number_format = "0.0%"
    sheet["B7"].number_format = "dd.mm.yyyy"
    
    sheet.column_dimensions["A"].width = 34
    sheet.column_dimensions["B"].width = 22
    sheet.column_dimensions["C"].width = 64
    sheet.freeze_panes = "A2"

def _style_final_sheet(sheet, last_row: int):
    thin_black = Side(style="thin", color="000000")
    border_black = Border(left=thin_black, right=thin_black, top=thin_black, bottom=thin_black)
    for row in sheet.iter_rows(min_row=7, max_row=last_row, min_col=1, max_col=29):
        for cell in row:
            cell.border = border_black
            cell.alignment = Alignment(vertical="center", wrap_text=False)
            if cell.row >= 8:
                cell.font = Font(name="Arial", size=8, color="000000")
    for cell in sheet[7]:
        cell.fill = PatternFill(fill_type=None)
        cell.font = Font(name="Arial", size=8, bold=True, color="000000")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_black
    widths = [23, 10, 10, 14, 14, 12, 14, 12, 14, 12, 12, 5, 14, 14, 14, 12, 10, 15, 34, 18, 18, 14, 10, 20, 14, 10, 10, 10, 10]
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width

def generate_rolling_plan_workbook(
    campaign_path: Path,
    grade_map_path: Path,
    raw_material_map_path: Path,
    output_path: Path,
    *,
    plan_date: date | None = None,
    camp_number: str = "",
    furnace_type: str = "new",
    casting_plan_path: Path | None = None,
) -> RollingPlanResult:
    campaign_path = Path(campaign_path)
    grade_map_path = Path(grade_map_path)
    raw_material_map_path = Path(raw_material_map_path)
    output_path = Path(output_path)
    furnace_type = str(furnace_type or "new").strip().lower()
    if furnace_type not in {"old", "new"}:
        raise ValueError("Furnace selection must be either 'old' or 'new'.")
    
    if not campaign_path.is_file():
        raise ValueError("Generated campaign workbook is not available.")
    if not grade_map_path.is_file():
        raise ValueError("Grade details workbook is missing from data/reference.")
    if not raw_material_map_path.is_file():
        raise ValueError("Raw-material mapping workbook is missing from data/reference.")

    workbook = load_workbook(campaign_path, keep_vba=True, data_only=False)
    if "Sheet1" not in workbook.sheetnames:
        workbook.close()
        raise ValueError("Campaign workbook must contain Sheet1.")
        
    rows = _campaign_rows(workbook["Sheet1"])
    casting_rows = _read_casting_sheet_rows(casting_plan_path)
    planning_material = str(_source(rows[0], "Planning Mat", default="")).strip()
    grade_rows, grade_map = _read_grade_map(grade_map_path)
    raw_material_map, raw_source_rows = _read_raw_material_map(raw_material_map_path)
    plan_date = plan_date or date.today()
    
    # Generate the Order Ref No base: YYMM + 4-digit Camp No (Default to 3915)
    date_prefix = plan_date.strftime("%y%m")
    camp_match = re.search(r'\d{4}', camp_number)
    camp_str = camp_match.group(0) if camp_match else "3915"
    order_ref_base = f"{date_prefix}{camp_str}"
    order_serial = 1  # Initialize the sequential 3-digit counter
    
    from casting_plan.models import BilletLengthSetting, RailMillCalculationMaster
    bl_settings = BilletLengthSetting.objects.first()
    if not bl_settings:
        bl_settings = BilletLengthSetting.objects.create()

    calculation_master_rows = list(
        RailMillCalculationMaster.objects.filter(
            planning_mat__iexact=planning_material
        )
    )
    from casting_plan.btr_service import compute_furnace_cut_pcs, round_to_005

    for row in rows:
        row["_planning_material"] = str(_source(row, "Planning Mat", default="")).strip()
        row["_raw_material"] = _resolve_raw_material(row["_planning_material"], raw_material_map)
        row["_billet_weight"] = BILLET_SECTION_WEIGHTS[row["_raw_material"]]
        initial_mapping, _ = _select_grade_mapping(grade_map, str(_source(row, "Grade", default="")), "")
        row["_std"] = _extract_std(row.get("Remarks"), initial_mapping)
        # Prioritize the exact column from Campaign Sheet1
        # 1. Store the exact original text for DISPLAY in the Excel file
        raw_emb = str(_source(row, "Embossing", "Embossing (from source Emb*)", default="")).strip()
        row["_embossing"] = raw_emb
        
        # 2. Create a NORMALIZED version strictly for the GROUPING logic
        emb_upper = raw_emb.upper()
        if emb_upper == "JINDAL STEEL PG HT":
            row["_embossing_group"] = "JINDAL STEEL HT"
        elif emb_upper == "JINDAL STEEL PG MS":
            row["_embossing_group"] = "JINDAL STEEL MS"
        else:
            row["_embossing_group"] = emb_upper

        # Determine raw grade for sorting
        ext_grade = str(_source(row, "Grade", default="")).strip()
        use_higher = _is_higher_section(row["_billet_weight"])
        raw_grade = (initial_mapping or {}).get("higher" if use_higher else "lighter")
        if not raw_grade:
            raw_grade = f"CHECK ({ext_grade}/{row['_std'] or 'STD'})"
        row["_raw_grade"] = raw_grade

        campaign_material = str(_source(row, "Material", default="")).strip()
        casting_candidates = casting_rows.get(_casting_match_key(row))
        casting_row = casting_candidates.popleft() if casting_candidates else None
        row["_casting_row"] = casting_row

        # Section weight and planned quantity come from casting-plan Sheet1
        # whenever that generated workbook is available.
        sec_wt = _number(
            casting_row.get("Sec Wt") if casting_row else row.get("Sec Wt"),
            0.0,
        )
        row["Sec Wt"] = sec_wt
        cutting_len = _number(row.get("Cutting Length"), 0.0)
        btr_qty = _number(
            casting_row.get("BTR") if casting_row else row.get("BTR"),
            0.0,
        )
        row["BTR"] = btr_qty
        master_record = _select_calculation_master(
            calculation_master_rows,
            row["_planning_material"],
            campaign_material,
            ext_grade,
            sec_wt,
        )
        final_raw_grade = (
            _display_raw_grade(master_record.rm_grade)
            if master_record and master_record.rm_grade
            else _display_raw_grade(raw_grade)
        )
        if not final_raw_grade or final_raw_grade.startswith("CHECK"):
            final_raw_grade = "ASTMA572GR50"

        order_pcs = max(
            1.0,
            round(btr_qty / max(0.1, (sec_wt * cutting_len / 1000.0)), 1),
        )
        is_480_bloom = "480" in str(row["_raw_material"]) and "420" in str(row["_raw_material"])
        rm_sec_wt = (
            _number(master_record.bloom_weight_kg_m)
            if master_record and master_record.bloom_weight_kg_m
            else (287.0 if is_480_bloom else 873.0)
        )
        end_cut_margin = (
            _number(master_record.end_cut_margin_req, 1.0)
            if master_record else 1.0
        )
        if furnace_type == "old":
            furnace_min = _number(master_record.old_furnace_min_len, 8.30) if master_record else 8.30
            furnace_max = _number(master_record.old_furnace_max_len, 9.10) if master_record else 9.10
            casting_cut_header = "Cut Pcs (OLD)"
        else:
            furnace_min = _number(master_record.new_furnace_min_len, 8.30) if master_record else 8.30
            furnace_max = _number(master_record.new_furnace_max_len, 10.70) if master_record else 10.70
            casting_cut_header = "Cut Pcs (NEW)"

        cut_pcs, finish_len, calc_bloom_len, final_bloom_len = compute_furnace_cut_pcs(
            cutting_len,
            sec_wt,
            rm_sec_wt,
            end_cut_margin,
            furnace_min,
            furnace_max,
            hot_saw_loss=0.15,
            order_pcs=order_pcs,
            btr_qty=btr_qty,
        )
        casting_cut_pcs = _number(
            casting_row.get(casting_cut_header) if casting_row else None,
            0.0,
        )
        if casting_cut_pcs > 0:
            cut_pcs = int(casting_cut_pcs)
            finish_len = (cut_pcs * (cutting_len + 0.15)) + end_cut_margin
            calc_bloom_len = (finish_len * sec_wt) / (rm_sec_wt * 0.97)
            calc_bloom_len = max(furnace_min, min(furnace_max, calc_bloom_len))
            final_bloom_len = round_to_005(calc_bloom_len)
        bloom_pcs_cnt = (
            math.ceil(order_pcs / max(1, cut_pcs)) if cut_pcs > 0 else 1
        )
        as_rolled_qty = cutting_len * sec_wt * cut_pcs * bloom_pcs_cnt / 1000.0

        row["_campaign_material"] = campaign_material
        row["_master_record"] = master_record
        row["_final_raw_grade"] = final_raw_grade
        row["_cut_pcs"] = cut_pcs
        row["_final_bloom_len"] = final_bloom_len
        row["_bloom_pcs_cnt"] = bloom_pcs_cnt
        row["_as_rolled_qty"] = as_rolled_qty

        thickness = row["_thickness"]
        thickness_text = (
            f"X{thickness:g}"
            if _is_angle_section(row["_planning_material"])
            and math.isfinite(thickness) and thickness > 0
            else ""
        )
        row["_rolling_section"] = (
            f"{row['_planning_material']}{thickness_text} "
            f"{_rolling_suffix(row.get('Remarks'), row['_std'])}"
        ).strip()
        row["_fourth_sort"] = _fourth_sort_value(
            row["_planning_material"], row["_thickness"], row.get("Sec Wt")
        )
        
        row["_planned_qty_sort"] = _number(row.get("BTR"), 0.0)
        row["_finished_length_sort"] = _number(row.get("Cutting Length"), 0.0)
        row["_raw_grade_rank_sort"] = _raw_grade_rank(row["_final_raw_grade"])
        row["_section_weight_sort"] = _number(row.get("Sec Wt"), 0.0)

    # Apply four independent stable sorts in the exact requested sequence.
    # Every sort works on the result produced by the preceding sort.
    rows.sort(key=lambda row: row["_planned_qty_sort"], reverse=True)
    rows.sort(key=lambda row: row["_finished_length_sort"], reverse=True)
    rows.sort(key=lambda row: row["_raw_grade_rank_sort"])
    rows.sort(key=lambda row: row["_section_weight_sort"], reverse=True)

    _add_grade_mapping_sheet(workbook, grade_rows)
    _add_raw_material_sheet(workbook, raw_source_rows)
    _add_assumptions_sheet(workbook, plan_date, camp_number, bl_settings)
    if "Final Plan" in workbook.sheetnames:
        del workbook["Final Plan"]
    sheet = workbook.create_sheet("Final Plan")
    sheet.views.sheetView[0].showGridLines = True

    # Styling definitions
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    title_font = Font(name="Calibri", size=16, bold=True, color="000000")
    header_font = Font(name="Calibri", size=11, bold=True, color="000000")
    data_font = Font(name="Calibri", size=11, color="000000")
    bold_data_font = Font(name="Calibri", size=11, bold=True, color="000000")
    red_data_font = Font(name="Calibri", size=13, bold=True, color="FF0000")

    fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Row 1-3 Titles
    sheet.merge_cells("A1:W1")
    sheet.merge_cells("A2:W2")
    sheet.merge_cells("A3:V3")
    sheet.merge_cells("W3:X3")
    sheet["A1"] = "JINDAL STEEL"
    sheet["A2"] = "(RAIL MILL)"
    title_material = re.sub(r"\s+", " ", planning_material.replace("_", " ")).replace("X", " x ")
    sheet["A3"] = f"REVISED CAMPAIGN ROLLING PLAN FOR {title_material} (SAP CAMPAIGN NO-{camp_str}) ( REV-01)"
    sheet["X2"] = f"Date: {plan_date.strftime('%d.%m.%Y')}"
    sheet["W3"] = "FORMAT NO.: F-17(8.5.1-21)"
    for coord in ("A1", "A2", "A3"):
        sheet[coord].font = title_font
        sheet[coord].alignment = Alignment(horizontal="center", vertical="center")
    for coord in ("X2", "W3"):
        sheet[coord].font = header_font
        sheet[coord].alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 27.75
    sheet.row_dimensions[2].height = 22.5
    sheet.row_dimensions[3].height = 30

    # Row 4 Column Headers (Cols A to Y / 1 to 25)
    headers_row4 = [
        ("A", "Date"),
        ("B", "Sales Order"),
        ("C", "SO Item"),
        ("D", "Order"),
        ("E", "Planned Order"),
        ("F", "Party name"),
        ("G", "Rolled Grade"),
        ("H", "Section to be Rolled"),
        ("I", "Raw material size"),
        ("J", "R/M Length"),
        ("K", "Grade of raw material"),
        ("L", "No. of Pcs"),
        ("M", "Cumulative no. of Pcs"),
        ("N", "Raw material Qty"),
        ("O", "As Rolled Qty"),
        ("P", "Planned Qty"),
        ("Q", "Cut length at Hot saw (mtr)"),
        ("R", "As Finished length (After Hot saw)"),
        ("S", "Tolerance"),
        ("T", "Export"),
        ("U", "Domestics"),
        ("V", "Internal"),
        ("W", "Stock"),
        ("X", "Explanation"),
        ("Y", "Excess Qty"),
        ("Z", "KG"),
        ("AA", "T/M length"),
        ("AB", "Cut Pcs"),
        ("AC", ""),
        ("AD", "Endcrop"),
    ]

    for col_idx, (c_let, h_name) in enumerate(headers_row4, 1):
        cell = sheet.cell(row=4, column=col_idx, value=h_name)
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = PatternFill(fill_type=None)
        if c_let in ("S", "T", "U", "V", "W"):
            cell.font = Font(name="Calibri", size=10, bold=True, color="000000")
        elif c_let == "X":
            cell.font = Font(name="Calibri", size=12, bold=True, color="000000")

    sheet.row_dimensions[4].height = 54

    output_row = 5
    total_rm_qty = 0.0
    total_blooms = 0
    max_cooling = 0.0
    inferred_grades = 0
    previous_sec_wt = None

    for row in rows:
        current = output_row

        ext_grade = str(_source(row, "Grade", default="A572GR50")).strip()
        cutting_len = _number(row.get("Cutting Length"), 12.0)
        sec_wt = _number(row.get("Sec Wt"), 83.7)
        btr_qty = _number(row.get("BTR"), 0.0)

        planning_mat_str = str(
            _source(row, "Planning Mat", default=planning_material)
        ).strip()
        campaign_material = row["_campaign_material"]
        master_record = row["_master_record"]
        raw_grade = row["_final_raw_grade"]
        cut_pcs = row["_cut_pcs"]
        final_bloom_len = row["_final_bloom_len"]
        bloom_pcs_cnt = row["_bloom_pcs_cnt"]

        # Section to be Rolled must be the Campaign Plan Material value.
        sec_to_be_rolled = campaign_material
        if not sec_to_be_rolled:
            sec_to_be_rolled = (
                f"AG_{ext_grade}_{planning_mat_str}_10-13"
                if not planning_mat_str.startswith("AG_")
                else planning_mat_str
            )

        master_rm_size = (
            master_record.rm_size
            if master_record and master_record.rm_size
            else row.get("_raw_material")
        )
        raw_mat_size = _raw_material_display(master_rm_size, raw_grade)
        rm_qty_factor = 0.98 if (
            "480X420X120" in _normalise_grade(raw_mat_size)
            or "480X420X120" in _normalise_grade(master_rm_size)
        ) else 0.873

        # Every rolling-plan detail row uses the current generation date.
        so_date_val = date.today()

        party_name = str(_source(row, "Sold To Party Name", "Ship to Party", "Sold_To_Party_Name", default=""))

        # Row Data Dict (Cols A to Y / 1 to 25)
        row_values = {
            1: so_date_val,  # Date
            2: str(_source(row, "SO No.", "SO_No", "CO", default="")),  # Sales Order
            3: str(_source(row, "SO Item", "SO_Item", default="")),  # SO Item
            4: "",  # Order (BLANK as requested)
            5: "",  # Planned Order (BLANK as requested)
            6: party_name,  # Party name
            7: ext_grade,  # Rolled Grade
            8: sec_to_be_rolled,  # Section to be Rolled
            9: raw_mat_size,  # Raw material size
            10: final_bloom_len,  # RM Length
            11: raw_grade,  # Grade of raw material
            12: bloom_pcs_cnt,  # No of Pcs
            13: f"=SUM($L$5:L{current})",  # Cumulative no. of Pcs
            14: f"=L{current}*J{current}*{rm_qty_factor:g}",  # Raw material Qty
            # Rolled Qty uses the Casting Plan Sheet1 cut-pieces result.
            15: f"=R{current}*Z{current}*AB{current}*L{current}/1000",
            16: btr_qty,  # Planned Qty: Campaign Plan BTR Qty
            17: f"=R{current}+IF(R{current}>=8.5,0.15,0.10)",  # Hot-saw allowance
            18: cutting_len,  # As Finished length (After Hot saw)
            19: "mm(+0mm,-0mm)",  # Tolerance
            20: btr_qty if str(_source(row, "SO Type", default="")).upper() == "ZEXP" else "",  # Export
            21: btr_qty if str(_source(row, "SO Type", default="")).upper() != "ZEXP" else btr_qty,  # Domestic
            22: "",  # Internal
            23: "",  # Stock
            24: f'="LENGTH -" & TEXT(R{current},"0.00") & " MTRS"',  # Explanation
            25: f"=O{current}-P{current}",  # Excess Qty
            26: sec_wt,  # KG / section weight helper used by As Rolled Qty
            27: f"=J{current}*980*0.97/Z{current}",  # T/M length for both RM sizes
            28: cut_pcs,  # Casting-plan cut pieces used by As Rolled Qty
            29: f"=Q{current}*AB{current}",  # Total cut length
            30: f"=AA{current}-AC{current}",  # Endcrop
        }

        is_first_weight = previous_sec_wt is None
        weight_changed = (
            previous_sec_wt is not None and abs(sec_wt - previous_sec_wt) > 0.001
        )
        boundary_row = is_first_weight or weight_changed
        rolled_grade_color = "0070C0" if (
            "S460" in ext_grade.upper() or "EN10025" in ext_grade.upper()
        ) else "FF0000"

        for col_idx, val in row_values.items():
            cell = sheet.cell(row=current, column=col_idx, value=val)
            cell.border = thin_border
            font_sizes = {7: 13, 8: 16 if boundary_row else 11, 10: 23, 11: 14,
                          12: 24, 13: 24, 14: 16, 15: 16, 16: 16,
                          17: 22, 18: 22, 24: 18, 25: 11}
            always_bold = col_idx in (7, 11, 17, 18, 25)
            font_color = "000000"
            if col_idx == 7:
                font_color = rolled_grade_color
            elif col_idx == 11 and "ASTM" in raw_grade.upper():
                font_color = "FF0000"
            cell.font = Font(
                name="Calibri",
                size=font_sizes.get(col_idx, 11),
                bold=boundary_row or always_bold,
                color=font_color,
            )

            if weight_changed and col_idx in (15, 16):
                cell.fill = fill_yellow

            if col_idx in (10, 17, 18):
                cell.number_format = '0.00'
            elif col_idx == 1:
                cell.number_format = 'dd.mm.yyyy'
            elif col_idx in (4, 12, 13, 14, 15, 16, 20, 21, 22, 23, 25, 28):
                cell.number_format = '0'
            elif col_idx in (26, 27, 29, 30):
                cell.number_format = '0.00'

            if col_idx in (2, 3, 4, 5, 7, 10, 11, 12, 13, 14, 15, 16, 17, 18, 20, 21, 22, 23, 25):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

        sheet.row_dimensions[current].height = 32.25
        previous_sec_wt = sec_wt

        rm_req_est = bloom_pcs_cnt * final_bloom_len * rm_qty_factor
        total_rm_qty += rm_req_est
        total_blooms += bloom_pcs_cnt
        output_row += 1

    # Total Row at Bottom
    tot_row = output_row
    sheet.cell(row=tot_row, column=6, value="TOTAL").font = header_font
    sheet.cell(row=tot_row, column=12, value=f"=SUM(L5:L{tot_row-1})").font = header_font
    sheet.cell(row=tot_row, column=13, value=f"=MAX(M5:M{tot_row-1})").font = header_font
    sheet.cell(row=tot_row, column=14, value=f"=SUM(N5:N{tot_row-1})").font = header_font
    sheet.cell(row=tot_row, column=15, value=f"=SUM(O5:O{tot_row-1})").font = header_font
    sheet.cell(row=tot_row, column=16, value=f"=SUM(P5:P{tot_row-1})").font = header_font
    sheet.cell(row=tot_row, column=25, value=f"=SUM(Y5:Y{tot_row-1})").font = header_font

    for col_idx in range(1, 31):
        c = sheet.cell(row=tot_row, column=col_idx)
        c.border = thin_border
        if col_idx in (12, 13, 14, 15, 16, 25):
            c.number_format = '0.00'
            c.alignment = Alignment(horizontal="right", vertical="center")

    # Set column widths
    col_widths = [8.332, 11.555, 6.555, 15, 9.664, 26.109, 23, 42.441,
                  30.109, 11.664, 21.332, 10.441, 12.441, 11.332, 11.887,
                  13, 13, 15, 19.555, 9.555, 10.441, 7.664, 6.555, 32.887, 13.664,
                  9, 12, 9, 12, 10]
    for idx, width in enumerate(col_widths, 1):
        sheet.column_dimensions[get_column_letter(idx)].width = width

    sheet.freeze_panes = None
    sheet.page_setup.orientation = "landscape"
    sheet.print_area = f"A1:AD{tot_row}"

    # Generate the Casting Plan as a separate file
    casting_plan_filename = output_path.name.replace("ROLLING_PLAN", "CASTING_PLAN").replace(".xlsm", ".xlsx")
    casting_plan_path = output_path.parent / casting_plan_filename
    try:
        from casting_plan.btr_service import create_auto_casting_plan_excel
        create_auto_casting_plan_excel(rows, planning_material, out_path=casting_plan_path)
    except Exception:
        generate_casting_plan_file(rows, planning_material, casting_plan_path)

    # Delete all auxiliary sheets, keeping ONLY 'Final Plan'
    for sname in list(workbook.sheetnames):
        if sname != "Final Plan":
            del workbook[sname]

    # Force auto calculation on load
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    workbook.close()

    total_rolling_qty = sum(_number(row.get("BTR")) for row in rows)
    yield_pct = (total_rolling_qty / total_rm_qty * 100.0) if total_rm_qty > 0 else 0.0

    return RollingPlanResult(
        output_path=output_path,
        row_count=len(rows),
        rolling_qty=total_rolling_qty,
        rm_qty=round(total_rm_qty, 3),
        bloom_count=total_blooms,
        yield_pct=round(yield_pct, 2),
        max_cooling_bed_length=86.0,
        inferred_grade_count=0,
    )
