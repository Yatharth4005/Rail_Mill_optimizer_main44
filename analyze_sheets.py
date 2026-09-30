import openpyxl
import pandas as pd
import json

file_path = "Rail Mill Length calculation.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)

md_out = []
md_out.append("# Analysis of Rail Mill Length calculation.xlsx Master Data Sheets\n")

for sheet in wb.sheetnames:
    df = pd.read_excel(file_path, sheet_name=sheet)
    md_out.append(f"## Sheet: `{sheet}`")
    md_out.append(f"- **Dimensions**: {df.shape[0]} rows × {df.shape[1]} columns\n")
    
    # Try to find header row if there are Unnamed columns
    unnamed_cols = [c for c in df.columns if str(c).startswith("Unnamed:")]
    if len(unnamed_cols) > len(df.columns) / 2:
        # Search first 10 rows for real header
        for idx in range(min(10, len(df))):
            row_vals = [str(x).strip() for x in df.iloc[idx] if pd.notna(x)]
            if len(row_vals) > 3:
                md_out.append(f"- **Header Row Candidate (Index {idx})**: {row_vals[:10]}")
    else:
        md_out.append(f"- **Columns**: `{list(df.columns[:15])}`")
    
    md_out.append("\n---\n")

with open("scratch_sheets_summary.md", "w", encoding="utf-8") as f:
    f.write("\n".join(md_out))

print("Done writing scratch_sheets_summary.md")
