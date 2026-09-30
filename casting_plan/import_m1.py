import os
import django
import pandas as pd

# Setup Django environment
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'rail_casting_plan.settings')
django.setup()

from casting_plan.models import RootRadiusMaster

def run_import():
    # Adjust this path if your Excel file is saved somewhere else
    file_path = "data/reference/M-1 Loss with Section.xlsx"
    
    if not os.path.exists(file_path):
        print(f"❌ Error: File not found at {file_path}")
        print("Please ensure the M-1 Loss file is placed in the data/reference folder.")
        return
        
    print(f"Reading {file_path}...")
    df = pd.read_excel(file_path, sheet_name=0)
    
    count = 0
    for index, row in df.iterrows():
        mat = str(row.get('Planning Mat', '')).strip().upper()
        if not mat.startswith('AL_'):
            continue
            
        thick = row.get('Thick')
        # Handle variations in column names across different sheets
        sec_wt = row.get('Corrected_Sec_wt') if pd.notna(row.get('Corrected_Sec_wt')) else row.get('SecWt_SAP')
        std = str(row.get('STD', '')).strip().upper()
        rr = row.get('RR')
        
        if pd.isna(thick) or pd.isna(rr):
            continue
            
        # Insert or update the record in the database
        obj, created = RootRadiusMaster.objects.update_or_create(
            planning_material=mat,
            thickness=float(thick),
            sec_wt=float(sec_wt) if pd.notna(sec_wt) else None,
            std=std,
            defaults={'root_radius': float(rr)}
        )
        count += 1
        
    print(f"✅ Successfully imported {count} Root Radius records into the database!")

if __name__ == '__main__':
    run_import()