from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile

from django.conf import settings
from django.test import SimpleTestCase
from openpyxl import load_workbook

from .btr_services import generate_campaign_workbook, inspect_btr_workbook
from .rolling_plan_services import (
    BILLET_SECTION_WEIGHTS,
    _fourth_sort_value,
    _is_higher_section,
    _read_raw_material_map,
    _remarks_suffix,
    _resolve_raw_material,
    _rolling_suffix,
    generate_rolling_plan_workbook,
)


class BTRCampaignServiceTests(SimpleTestCase):
    def setUp(self):
        self.reference_dir = Path(settings.BASE_DIR) / 'data' / 'reference'
        self.btr = self.reference_dir / 'BTR_file_sample.xlsx'
        self.template = self.reference_dir / 'CAMP PLAN FORMAT.xlsm'

    def test_sample_btr_is_inspected(self):
        inspection = inspect_btr_workbook(self.btr)
        self.assertEqual(inspection.row_count, 49)
        names = {item['name'] for item in inspection.planning_materials}
        self.assertIn('ISMC_150X75', names)

    def test_campaign_preserves_macro_and_calculates_requested_columns(self):
        with TemporaryDirectory() as temporary:
            output = Path(temporary) / 'campaign.xlsm'
            result = generate_campaign_workbook(
                self.btr, self.template, output, 'ISMC_150X75'
            )
            self.assertEqual(result.input_row_count, 7)
            self.assertEqual(result.output_row_count, 7)
            self.assertEqual(result.zero_length_resolved_count, 2)
            self.assertAlmostEqual(result.total_btr_qty, 1300.940, places=3)

            with ZipFile(output) as archive:
                self.assertIn('xl/vbaProject.bin', archive.namelist())

            workbook = load_workbook(output, data_only=False, read_only=True)
            sheet = workbook['Sheet1']
            data_rows = [
                row for row in range(2, sheet.max_row + 1)
                if isinstance(sheet.cell(row, 1).value, int)
            ]
            self.assertEqual(len(data_rows), 7)
            self.assertTrue(str(sheet.cell(data_rows[0], 17).value).startswith('='))
            self.assertTrue(str(sheet.cell(data_rows[0], 18).value).startswith('='))
            labels = [sheet.cell(row, 1).value for row in range(2, sheet.max_row + 1)]
            self.assertFalse(any(str(label).endswith(' GROUP') for label in labels))
            self.assertEqual(
                sum(str(label).startswith('Subtotal / Thickness') for label in labels),
                1,
            )
            self.assertEqual(result.bucket_counts, {'MS': 1, 'PANTHER': 6})
            self.assertNotIn('Bal Bill Qty', [cell.value for cell in workbook['Sheet2'][1]])
            workbook.close()

    def test_rolling_plan_is_generated_from_campaign_with_formulas_and_vba(self):
        with TemporaryDirectory() as temporary:
            campaign_path = Path(temporary) / 'campaign.xlsm'
            campaign = generate_campaign_workbook(
                self.btr, self.template, campaign_path, 'ISMC_150X75'
            )
            rolling_path = Path(temporary) / 'rolling_plan.xlsm'
            result = generate_rolling_plan_workbook(
                campaign_path,
                self.reference_dir / 'Grade details of ex grade and raw grade.xlsx',
                self.reference_dir / 'RAW MATERIAL LENGTH TOLERANCE.xlsx',
                rolling_path,
                camp_number='TEST-1',
            )

            self.assertEqual(result.row_count, campaign.output_row_count)
            self.assertAlmostEqual(result.rolling_qty, campaign.total_btr_qty, places=3)
            self.assertLess(result.max_cooling_bed_length, 87)
            self.assertGreater(result.bloom_count, 0)

            with ZipFile(rolling_path) as archive:
                self.assertIn('xl/vbaProject.bin', archive.namelist())

            workbook = load_workbook(rolling_path, data_only=False, read_only=False)
            self.assertIn('Final Plan', workbook.sheetnames)
            self.assertIn('Grade Mapping', workbook.sheetnames)
            self.assertIn('Raw Material Mapping', workbook.sheetnames)
            self.assertIn('Plan Assumptions', workbook.sheetnames)
            sheet = workbook['Final Plan']
            self.assertEqual(sheet['W2'].value, 'TEST-1')
            first = 9
            self.assertTrue(str(sheet.cell(first, 1).value).startswith('ISMC_150X75 '))
            self.assertEqual(sheet.cell(first, 5).value, 'S275LS')
            self.assertEqual(sheet.cell(first, 9).value, 214)
            for column in (8, 10, 11, 13, 14, 15, 16):
                self.assertTrue(str(sheet.cell(first, column).value).startswith('='))
            self.assertIn("INT(N9/'Plan Assumptions'!$B$4)+1", sheet['O9'].value)
            labels = [sheet.cell(row, 1).value for row in range(9, sheet.max_row + 1)]
            self.assertFalse(any(' / ' in str(label) and str(label).endswith(' TOTAL') for label in labels))
            self.assertEqual(labels.count('ISMC_150X75 BIS TOTAL'), 1)
            self.assertEqual(labels.count('GRAND TOTAL'), 1)
            self.assertEqual(sheet['F16'].value, '=SUM(F9:F15)')
            self.assertEqual(sheet['J16'].value, '=SUM(J9:J15)')
            self.assertEqual(
                sheet['K15'].value,
                '=SUMIF($A$9:A15,"<>*TOTAL",$J$9:J15)',
            )
            workbook.close()

    def test_rolling_plan_rule_helpers(self):
        self.assertEqual(_remarks_suffix('TLT-BSEN'), 'BSEN')
        self.assertEqual(_remarks_suffix('TLT-ASTM-GR60'), 'GR60')
        self.assertEqual(_rolling_suffix('TLT-ASTM-GR60'), 'ASTMGR60')
        self.assertFalse(_is_higher_section(300))
        self.assertFalse(_is_higher_section(BILLET_SECTION_WEIGHTS['BL_165X165']))
        self.assertTrue(_is_higher_section(BILLET_SECTION_WEIGHTS['BL_250X250']))
        self.assertEqual(_fourth_sort_value('AL_110X110', 8, 19.68), 8)
        self.assertEqual(_fourth_sort_value('UB_305X165', 8, 54), 54)

        mappings, _ = _read_raw_material_map(
            self.reference_dir / 'RAW MATERIAL LENGTH TOLERANCE.xlsx'
        )
        self.assertEqual(_resolve_raw_material('ISMC_150X75', mappings), 'BL_165X165')
        self.assertEqual(_resolve_raw_material('UB_305X165', mappings), 'BB_355X280X90')
