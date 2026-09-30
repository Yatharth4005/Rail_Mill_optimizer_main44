from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

def generate_casting_plan_file(rows, planning_material, output_path: Path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Casting Plan"
    
    # Aggregate data grouped by the Rolling Plan Group AND Billet Length
    summary = {}
    for row in rows:
        qty = row.get("_final_rolling_qty", 0)
        if qty <= 0:
            continue
            
        rp_group = (
            row["_raw_grade"].upper(),
            row["_planning_material"].upper(),
            row["_embossing_group"],
            row["_std"].upper(),
            row["_fourth_sort"]
        )
        length = row.get("_final_best_length", 0)
        
        key = (rp_group, length)
        
        if key not in summary:
            summary[key] = {
                'sec': row["_raw_material"],
                'grade': row["_raw_grade"],
                'qty': 0,
                'length': length,
                'thicknesses': set()
            }
        summary[key]['qty'] += qty
        
        thickness = row.get("_thickness")
        if thickness and thickness != float('inf'):
            summary[key]['thicknesses'].add(f"{thickness:g}")
            
    # Format Title
    sheet.merge_cells("A1:G1")
    sheet["A1"] = f"Casting plan for {planning_material}"
    sheet["A1"].font = Font(bold=True, size=14)
    sheet["A1"].alignment = Alignment(horizontal="center", vertical="center")
    
    # Format Headers
    headers = ["Section", "Grade", "Qty(MT)", "Length", "Tolerance", "Location", "Remarks"]
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    for col, header in enumerate(headers, 1):
        cell = sheet.cell(2, col, header)
        cell.font = Font(bold=True)
        cell.border = border
        cell.alignment = Alignment(horizontal="center")
        
    # Write Aggregated Data
    row_idx = 3
    for key, data in summary.items():
        sheet.cell(row_idx, 1, data['sec']).border = border
        
        # Construct grade string matching the screenshot (e.g., BL_JSPHTHS(V)_250X250)
        parts = data['sec'].split('_', 1)
        formatted_grade = f"{parts[0]}_{data['grade']}_{parts[1]}" if len(parts) == 2 else data['grade']
        sheet.cell(row_idx, 2, formatted_grade).border = border
        
        sheet.cell(row_idx, 3, data['qty']).border = border
        sheet.cell(row_idx, 4, f"{data['length']:.2f} Mtr").border = border
        sheet.cell(row_idx, 5, "-0 to 40 mm").border = border
        sheet.cell(row_idx, 6, "MCAP").border = border
        
        # Format remarks as comma-separated thicknesses (e.g., "26 , 17 mm")
        thk_list = sorted(list(data['thicknesses']), key=float, reverse=True)
        thk_str = " , ".join(thk_list) + " mm" if thk_list else ""
        sheet.cell(row_idx, 7, thk_str).border = border
        
        row_idx += 1
        
    # Adjust column widths
    widths = [18, 35, 12, 15, 15, 12, 20]
    for col, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(col)].width = width
    sheet.row_dimensions[1].height = 25
    
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    workbook.close()