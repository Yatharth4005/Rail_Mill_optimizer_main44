from django.core.management.base import BaseCommand
from pathlib import Path
import pandas as pd, zipfile, shutil
from casting_plan.models import Order,SectionRule,GradeMap,BatchRule,PlanLine,SectionWeightMap
BASE=Path(__file__).resolve().parents[3]
REF=BASE/'data'/'reference'

def clean(v): return '' if pd.isna(v) else str(v).strip()
def num(v):
    try:return float(v)
    except:return 0.0

def raw_family_helper(section):
    sec = (section or '').upper()
    if any(x in sec for x in ['305','356','406','254']):
        return 'SBB', 'BB_355X280X90'
    if any(x in sec for x in ['203','200','180','152']):
        return '250sq', 'BL_250X250'
    return '165SQ', 'BL_165X165'

class Command(BaseCommand):
    help='Import the supplied SPM Excel workbooks and section files into SQLite.'
    def handle(self,*args,**kwargs):
        SectionRule.objects.all().delete(); GradeMap.objects.all().delete(); BatchRule.objects.all().delete(); PlanLine.objects.all().delete()
        SectionWeightMap.objects.all().delete()

        # Seed core predefined mappings
        seed_data = [
            ("UB_305X165", 40.3, "SBB", "BB_355X280X90", 10.98, "0", "MCAP"),
            ("UB_305X165", 46.1, "SBB", "BB_355X280X90", 10.00, "0", "MCAP"),
            ("UB_305X165", 54.0, "SBB", "BB_355X280X90", 11.82, "0", "MCAP"),
            ("UB_356X171", 45.0, "SBB", "BB_355X280X90", 9.98, "0/40", "MCAP"),
            ("UB_356X171", 51.0, "SBB", "BB_355X280X90", 9.66, "0/40", "MCAP"),
            ("UB_356X171", 57.0, "SBB", "BB_355X280X90", 10.73, "0/40", "MCAP"),
            ("UB_356X171", 67.0, "SBB", "BB_355X280X90", 11.00, "0/40", "MCAP"),
            ("UC_203X203", 46.1, "SBB", "BB_355X280X90", 11.40, "0/40", "MCAP"),
            ("UC_203X203", 52.0, "SBB", "BB_355X280X90", 11.40, "0/40", "MCAP"),
            ("UC_203X203", 60.0, "SBB", "BB_355X280X90", 9.80, "0/40", "MCAP"),
            ("UC_203X203", 71.0, "SBB", "BB_355X280X90", 11.40, "0/40", "MCAP"),
            ("UC_203X203", 86.1, "SBB", "BB_355X280X90", 11.65, "0/40", "MCAP"),
            ("NPB_400X180", 57.38, "SBB", "BB_355X280X90", 10.63, "0/40", "MCAP"),
            ("NPB_400X180", 66.31, "SBB", "BB_355X280X90", 10.85, "0/40", "MCAP"),
            ("NPB_400X180", 75.67, "SBB", "BB_355X280X90", 10.30, "0/40", "MCAP"),
            ("UB_406X178", 51.0, "SBB", "BB_355X280X90", 10.15, "0/40", "MCAP"),
            ("UB_406X178", 60.1, "SBB", "BB_355X280X90", 11.15, "0/40", "MCAP"),
            ("UB_406X178", 67.1, "SBB", "BB_355X280X90", 11.05, "0/40", "MCAP"),
            ("UB_406X178", 74.2, "SBB", "BB_355X280X90", 11.90, "0/40", "MCAP"),
        ]
        for sec, wt, fam, raw, length, tol, loc in seed_data:
            SectionWeightMap.objects.update_or_create(
                planning_section=sec,
                section_weight=wt,
                defaults={
                    "rm_family": fam,
                    "raw_section": raw,
                    "casting_length": length,
                    "tolerance": tol,
                    "location": loc
                }
            )
        f=REF/'RAW MATERIAL LENGTH TOLERANCE.xlsx'; df=pd.read_excel(f,sheet_name='Length tollerance',header=1)
        for _,r in df.iterrows():
            sec=clean(r.get('Section'))
            if sec: SectionRule.objects.update_or_create(section=sec,defaults={'min_length':num(r.get('Min. Length')),'max_length':num(r.get('Max. Length')),'two_v':clean(r.get('2V Using or bypass'))})
        g=pd.read_excel(REF/'Grade details of ex grade and raw grade.xlsx')
        for _,r in g.iterrows():
            ext=clean(r.get('Ext Grade'))
            if ext: GradeMap.objects.create(ext_grade=ext,raw_higher=clean(r.get('Raw grade (Higher section)')),raw_lighter=clean(r.get('Raw grade (Lighter section) ')),std=clean(r.get('STD')),remarks=clean(r.get('Remarks')))
        b=pd.read_excel(REF/'LFA BATCH DETAILS.xlsx',header=1)
        for _,r in b.iterrows():
            sec=clean(r.get('SECTION '))
            if sec: BatchRule.objects.create(section=sec,bars_per_batch=num(r.get('NO OF BAR IN BATCH')),section_weight=num(r.get('SECTION WT')),cb_length=num(r.get('C/B LENGTH')),lfa_batch_length=num(r.get(' LFA BATCH LENGTH')),bars_at_cb=num(r.get('NO OF BAR AT C/B IN 1 BILLET')),billet_length=num(r.get('BILLET LENGTH')),billet_sectional_weight=num(r.get('BILLET SECTIONAL WT')))
        z=REF/'all sections.zip'; out=BASE/'data'/'section_workbooks'; out.mkdir(parents=True,exist_ok=True)
        with zipfile.ZipFile(z) as zz: zz.extractall(out)
        count=0
        for p in out.rglob('*'):
            if p.suffix.lower() not in ('.xlsx','.xls'): continue
            try:
                # 1. Discover section weight maps dynamically
                xl = pd.ExcelFile(p)
                sheet_name = next((s for s in xl.sheet_names if "final" in s.lower() and "plan" in s.lower()), None)
                if not sheet_name:
                    sheet_name = next((s for s in xl.sheet_names if "plan" in s.lower()), None)
                if sheet_name:
                    d_plan = xl.parse(sheet_name, header=None)
                    header_row_idx = None
                    for idx, row in d_plan.iterrows():
                        row_strs = [str(x).strip().upper() for x in row]
                        if any("ROLLING SECTION" in x or x == "SECTION" for x in row_strs):
                            header_row_idx = idx
                            break
                    if header_row_idx is not None:
                        headers = [str(x).strip() for x in d_plan.iloc[header_row_idx]]
                        def get_col_idx(names):
                            for name in names:
                                for i, h in enumerate(headers):
                                    if name.upper() in h.upper(): return i
                            return None
                        sec_idx = get_col_idx(["ROLLING SECTION", "SECTION"])
                        wt_idx = get_col_idx(["SEC WT", "SECTION WT"])
                        len_idx = get_col_idx(["BLOOM LENGTH", "BILLET LENGTH", "LENGTH"])
                        if None not in (sec_idx, wt_idx, len_idx):
                            for idx in range(header_row_idx + 1, len(d_plan)):
                                row = d_plan.iloc[idx]
                                sec_val = clean(row[sec_idx])
                                if sec_val and not sec_val.startswith("Total") and not sec_val.startswith("Casting"):
                                    wt_val = clean(row[wt_idx])
                                    len_val = clean(row[len_idx])
                                    try:
                                        wt_float = float(wt_val)
                                        len_float = float(len_val)
                                        if wt_float > 0 and len_float > 0:
                                            fam, raw_sec = raw_family_helper(sec_val)
                                            tol = "0" if "SBB" in fam else "0/40"
                                            SectionWeightMap.objects.get_or_create(
                                                planning_section=sec_val,
                                                section_weight=wt_float,
                                                defaults={
                                                    "rm_family": fam,
                                                    "raw_section": raw_sec,
                                                    "casting_length": len_float,
                                                    "tolerance": tol,
                                                    "location": "MCAP"
                                                }
                                            )
                                    except Exception:
                                        pass

                # 2. Historical plan lines ingestion (existing logic)
                sheets=xl.sheet_names
                for sh in sheets:
                    d=pd.read_excel(p,sheet_name=sh)
                    cols={str(c).strip().upper():c for c in d.columns}
                    sec_col=next((cols[k] for k in cols if 'ROLLING SECTION' in k or k=='SECTION'),None)
                    qty_col=next((cols[k] for k in cols if 'ROLLING QTY' in k or k=='QTY'),None)
                    if sec_col and qty_col:
                        for _,r in d.iterrows():
                            sec=clean(r.get(sec_col)); qty=num(r.get(qty_col))
                            if sec and qty: PlanLine.objects.create(section=sec,grade=clean(r.get(cols.get('EXT GRADE'))),raw_grade=clean(r.get(cols.get('R.M. GRADE'))),rolling_qty=qty,billet_length=num(r.get(cols.get('BILLET LENGTH'))),rm_qty=num(r.get(cols.get('RM'))),rolled_qty=num(r.get(cols.get('ROLLED QTY'))),source_file=p.name); count+=1
            except Exception as e:
                import traceback
                self.stderr.write(self.style.WARNING(f"Error importing {p.name}: {e}"))
                traceback.print_exc()
        self.stdout.write(self.style.SUCCESS(f'Imported {SectionRule.objects.count()} section rules, {GradeMap.objects.count()} grade mappings, {BatchRule.objects.count()} LFA rules, {SectionWeightMap.objects.count()} section weight maps, {PlanLine.objects.count()} historical plan lines. Total counted plan lines: {count}.'))
