import os
import pandas as pd
import numpy as np
from pathlib import Path
from django.core.management.base import BaseCommand
from casting_plan.models import (
    RailMillCalculationMaster,
    RailMillChemistryMaster,
    RailMillSectionGradeStandardMap,
    RailMillSecWiseMaster,
    RailMillSawTolerance,
    RailMillRMSectionWeight,
    RailMillPlanningMaterial
)

def clean_val(val):
    if pd.isna(val) or val is None or str(val).strip().lower() in ['nan', 'none', '', '-']:
        return ''
    return str(val).strip()

def float_val(val):
    try:
        if pd.isna(val) or val is None:
            return None
        v = float(val)
        if np.isnan(v):
            return None
        return v
    except (ValueError, TypeError):
        return None

class Command(BaseCommand):
    help = "Import Rail Mill Length Calculation workbook sheets into database models"

    def add_arguments(self, parser):
        parser.add_argument(
            '--file',
            dest='file_path',
            default='Rail Mill Length calculation.xlsx',
            help='Path to the Rail Mill PPC/Length Calculation workbook.',
        )

    def handle(self, *args, **options):
        file_path = options['file_path']
        if not os.path.exists(file_path):
            self.stderr.write(self.style.ERROR(f"File not found: {file_path}"))
            return

        self.stdout.write(self.style.SUCCESS(f"Reading {file_path}..."))
        xl = pd.ExcelFile(file_path)

        # -----------------------------------------------------------------
        # 1. Sheet: Calculation sheet
        # -----------------------------------------------------------------
        if "Calculation sheet" in xl.sheet_names:
            # Excel row 2 contains the real column headings; row 1 only has
            # grouped furnace captions.  Reading with header=1 preserves all
            # 946 detail rows without treating the first detail row as a header.
            df_calc = xl.parse("Calculation sheet", header=1)
            RailMillCalculationMaster.objects.all().delete()
            calc_objects = []

            for idx, row in df_calc.iterrows():
                p_mat = clean_val(row.iloc[0])
                if not p_mat or p_mat == 'Planning Mat':
                    continue

                obj = RailMillCalculationMaster(
                    planning_mat=p_mat,
                    rolled_section=clean_val(row.iloc[1]),
                    material=clean_val(row.iloc[2]),
                    product=clean_val(row.iloc[3]),
                    sec_wt=float_val(row.iloc[4]),
                    ext_grade=clean_val(row.iloc[5]),
                    grade_group=clean_val(row.iloc[6]),
                    rm_grade=clean_val(row.iloc[7]),
                    rm_size=clean_val(row.iloc[8]),
                    bloom_length_m=float_val(row.iloc[9]),
                    bloom_weight_kg_m=float_val(row.iloc[10]),
                    total_rolled_length_m=float_val(row.iloc[11]),
                    total_finish_length_m=float_val(row.iloc[12]),
                    cut_pcs=clean_val(row.iloc[13]),
                    end_cut_margin=float_val(row.iloc[14]),
                    rm_calculated_length_m=float_val(row.iloc[15]),
                    end_cut_margin_req=float_val(row.iloc[16]),
                    maximum_cooling_bed_length_m=float_val(row.iloc[17]),
                    old_furnace_min_len=float_val(row.iloc[18]),
                    old_furnace_max_len=float_val(row.iloc[19]),
                    new_furnace_min_len=float_val(row.iloc[21]),
                    new_furnace_max_len=float_val(row.iloc[22])
                )
                calc_objects.append(obj)

            RailMillCalculationMaster.objects.bulk_create(calc_objects)
            self.stdout.write(self.style.SUCCESS(f"✅ Imported {len(calc_objects)} Calculation Master records."))

        # -----------------------------------------------------------------
        # 2. Sheet: Chemistry
        # -----------------------------------------------------------------
        if "Chemistry" in xl.sheet_names:
            df_chem = xl.parse("Chemistry")
            RailMillChemistryMaster.objects.all().delete()
            chem_objects = []

            for idx in range(3, len(df_chem)):
                row = df_chem.iloc[idx]
                sl = clean_val(row.iloc[1])
                prod = clean_val(row.iloc[2])
                int_grd = clean_val(row.iloc[3])
                if not sl and not prod and not int_grd:
                    continue

                obj = RailMillChemistryMaster(
                    sl_no=sl,
                    product=prod,
                    internal_grade=int_grd,
                    proposed_internal_grade=clean_val(row.iloc[4]),
                    ext_grade=clean_val(row.iloc[5]),
                    sections=clean_val(row.iloc[6]),
                    range_type=clean_val(row.iloc[7]),
                    c_pct=clean_val(row.iloc[8]),
                    mn_pct=clean_val(row.iloc[9]),
                    s_pct=clean_val(row.iloc[10]),
                    p_pct=clean_val(row.iloc[11]),
                    si_pct=clean_val(row.iloc[12]),
                    al_pct=clean_val(row.iloc[13]),
                    cu_pct=clean_val(row.iloc[14]),
                    ni_pct=clean_val(row.iloc[15]),
                    cr_pct=clean_val(row.iloc[16]),
                    mo_pct=clean_val(row.iloc[17]),
                    ti_pct=clean_val(row.iloc[18]),
                    v_pct=clean_val(row.iloc[19]),
                    sn_pct=clean_val(row.iloc[20]),
                    nb_pct=clean_val(row.iloc[21]),
                    b_pct=clean_val(row.iloc[22]),
                    as_pct=clean_val(row.iloc[23]),
                    ce=clean_val(row.iloc[24]),
                    n2_ppm=clean_val(row.iloc[25]),
                    o2_ppm=clean_val(row.iloc[26]),
                    h2_ppm=clean_val(row.iloc[27]) if len(row) > 27 else ''
                )
                chem_objects.append(obj)

            RailMillChemistryMaster.objects.bulk_create(chem_objects)
            self.stdout.write(self.style.SUCCESS(f"✅ Imported {len(chem_objects)} Chemistry Master records."))

        # -----------------------------------------------------------------
        # 3. Sheet: ANNEXURE- SECTION
        # -----------------------------------------------------------------
        if "ANNEXURE- SECTION" in xl.sheet_names:
            df_annex = xl.parse("ANNEXURE- SECTION")
            RailMillSectionGradeStandardMap.objects.all().delete()
            annex_objects = []

            for idx, row in df_annex.iterrows():
                sec = clean_val(row.iloc[1])
                if not sec or sec == 'Rolled Section':
                    continue

                obj = RailMillSectionGradeStandardMap(
                    sl_no=clean_val(row.iloc[0]),
                    rolled_section=sec,
                    flange_thickness_mm=float_val(row.iloc[2]),
                    product=clean_val(row.iloc[3]),
                    std_e250_e275=clean_val(row.iloc[4]),
                    std_e350_s355=clean_val(row.iloc[5]),
                    std_e410_jis=clean_val(row.iloc[6]),
                    std_s420n=clean_val(row.iloc[7]),
                    std_s460_e450=clean_val(row.iloc[8]),
                    std_s500=clean_val(row.iloc[9]),
                    std_e550=clean_val(row.iloc[10]),
                    std_s355nl=clean_val(row.iloc[11]),
                    fire_resistance=clean_val(row.iloc[12]),
                    is_3443=clean_val(row.iloc[13]),
                    irs_t12=clean_val(row.iloc[14]),
                    en_13674=clean_val(row.iloc[15]) if len(row) > 15 else ''
                )
                annex_objects.append(obj)

            RailMillSectionGradeStandardMap.objects.bulk_create(annex_objects)
            self.stdout.write(self.style.SUCCESS(f"✅ Imported {len(annex_objects)} Section Grade Standard records."))

        # -----------------------------------------------------------------
        # 4. Sheet: SEC WISE
        # -----------------------------------------------------------------
        if "SEC WISE" in xl.sheet_names:
            df_secwise = xl.parse("SEC WISE")
            RailMillSecWiseMaster.objects.all().delete()
            secwise_objects = []

            for idx, row in df_secwise.iterrows():
                sec = clean_val(row.iloc[0])
                if not sec or sec == 'SECTION':
                    continue

                obj = RailMillSecWiseMaster(
                    section=sec,
                    section_weight_desc=clean_val(row.iloc[1]),
                    input_size=clean_val(row.iloc[2]),
                    bloom_length_m=float_val(row.iloc[3]),
                    bloom_weight_kg_m=float_val(row.iloc[4]),
                    total_rolled_length_m=float_val(row.iloc[5]),
                    total_finish_length_m=float_val(row.iloc[6]),
                    cut_pcs=clean_val(row.iloc[7])
                )
                secwise_objects.append(obj)

            RailMillSecWiseMaster.objects.bulk_create(secwise_objects)
            self.stdout.write(self.style.SUCCESS(f"✅ Imported {len(secwise_objects)} Section Wise Master records."))

        # -----------------------------------------------------------------
        # 5. Sheet: Sheet2 (Saw Tolerances & RM Section Weights)
        # -----------------------------------------------------------------
        if "Sheet2" in xl.sheet_names:
            df_s2 = xl.parse("Sheet2")
            RailMillSawTolerance.objects.all().delete()
            RailMillRMSectionWeight.objects.all().delete()

            # Row 0: >8.5 -> 0.15
            # Row 1: <8.5 -> 0.100
            # Row 2: BL 285 x 390 -> 0.873
            # Row 3: BB 480x420x120 -> 0.980
            RailMillSawTolerance.objects.create(condition=">8.5 m", tolerance_m=0.15)
            RailMillSawTolerance.objects.create(condition="<8.5 m", tolerance_m=0.100)
            RailMillRMSectionWeight.objects.create(rm_sec="BL 285 x 390", weight_t_per_m=0.873)
            RailMillRMSectionWeight.objects.create(rm_sec="BB 480x420x120", weight_t_per_m=0.980)

            self.stdout.write(self.style.SUCCESS(f"✅ Imported Saw Tolerances & RM Weight records."))

        # -----------------------------------------------------------------
        # 6. Sheet: Sheet1 (Planning Materials)
        # -----------------------------------------------------------------
        if "Sheet1" in xl.sheet_names:
            df_s1 = xl.parse("Sheet1")
            RailMillPlanningMaterial.objects.all().delete()
            pm_objects = []

            for idx, row in df_s1.iterrows():
                pm = clean_val(row.iloc[0])
                if pm and pm != 'Planning Mat':
                    pm_objects.append(RailMillPlanningMaterial(planning_mat=pm))

            RailMillPlanningMaterial.objects.bulk_create(pm_objects)
            self.stdout.write(self.style.SUCCESS(f"✅ Imported {len(pm_objects)} Planning Materials."))

        self.stdout.write(self.style.SUCCESS("\n🎉 All Rail Mill Length Calculation sheets imported successfully!"))
