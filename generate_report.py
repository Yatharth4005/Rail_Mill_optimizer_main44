import json
import os

json_file = "data_inspection_summary.json"
if not os.path.exists(json_file):
    print("❌ Error: data_inspection_summary.json not found. Please run 'py inspect_rail_mill_excel.py' first.")
    exit(1)

with open(json_file, "r", encoding="utf-8") as f:
    sheets_data = json.load(f)

md = []
md.append("# Rail Mill Master Data Analysis — Excel Sheet Breakdown\n")
md.append(f"**Total Master Data Sheets**: `{len(sheets_data)}`\n")

for i, sheet in enumerate(sheets_data, 1):
    name = sheet.get("sheet_name")
    rows = sheet.get("row_count")
    cols = sheet.get("col_count")
    sample = sheet.get("sample_data", [])
    
    md.append(f"## {i}. Sheet: `{name}`")
    md.append(f"- **Dimensions**: `{rows}` rows × `{cols}` columns")
    
    if sample:
        md.append("- **Header / Top Record Fields**:")
        first_row = sample[0]
        field_list = [f"`{k}`: `{v[:40] if isinstance(v, str) else v}`" for k, v in first_row.items() if v != "" and v != "nan" and not str(k).startswith("Unnamed:")]
        if not field_list:
            field_list = [f"`{k}`: `{v[:40] if isinstance(v, str) else v}`" for k, v in list(first_row.items())[:8]]
        for f_str in field_list[:12]:
            md.append(f"  - {f_str}")
    md.append("\n---\n")

report_path = "MASTER_DATA_SHEETS_ANALYSIS.md"
with open(report_path, "w", encoding="utf-8") as f:
    f.write("\n".join(md))

print(f"✅ Generated {report_path} successfully!")
