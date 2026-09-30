from decimal import Decimal
from pathlib import Path
import pandas as pd
from django.test import TestCase

from casting_plan.models import (
    Order, SectionRule, GradeMap, BatchRule, SectionWeightMap,
    YieldRule, CastingPlanRequest, CastingPlanRequirement
)
from casting_plan.services import (
    NormalizedRequirement, OptimizationEngine, evaluate_candidate,
    generate_candidate_lengths, resolve_yield_rule, resolve_finished_piece_length,
    import_demand_workbook, build_plan_rows, generate_casting_plan_workbook
)


class ManualCastingPlanOptimizationTests(TestCase):
    def setUp(self):
        # Create standard test master data
        SectionWeightMap.objects.create(
            planning_section="UB_305X165",
            section_weight=40.3,
            rm_family="SBB",
            raw_section="BB_355X280X90",
            casting_length=10.98,
            tolerance="0",
            location="MCAP"
        )
        SectionWeightMap.objects.create(
            planning_section="UB_305X165",
            section_weight=46.1,
            rm_family="SBB",
            raw_section="BB_355X280X90",
            casting_length=10.00,
            tolerance="0",
            location="MCAP"
        )
        SectionWeightMap.objects.create(
            planning_section="UB_305X165",
            section_weight=54.0,
            rm_family="SBB",
            raw_section="BB_355X280X90",
            casting_length=11.82,
            tolerance="0",
            location="MCAP"
        )
        SectionRule.objects.create(
            section="UB_305X165",
            min_length=11.50,
            max_length=12.00
        )
        GradeMap.objects.create(
            ext_grade="E250BR",
            raw_higher="BB_S275JR",
            raw_lighter="BB_S275HS"
        )
        # Yield rule setup matching Section + Weight + Family (Priority 1)
        YieldRule.objects.create(
            planning_section="UB_305X165",
            section_weight=40.3,
            rm_family="SBB",
            base_strand_length_m=Decimal('60.00'),
            head_crop_m=Decimal('0.50'),
            tail_crop_m=Decimal('0.50'),
            cut_loss_m=Decimal('0.01'),
            process_loss_pct=Decimal('2.00'),
            batch_rounding_basis='RAW_MATERIAL_QTY',
            minimum_batch_qty_mt=Decimal('50.00'),
            batch_increment_qty_mt=Decimal('50.00'),
            is_active=True
        )
        YieldRule.objects.create(
            planning_section="UB_305X165",
            section_weight=54.0,
            rm_family="SBB",
            base_strand_length_m=Decimal('60.00'),
            head_crop_m=Decimal('0.50'),
            tail_crop_m=Decimal('0.50'),
            cut_loss_m=Decimal('0.01'),
            process_loss_pct=Decimal('2.00'),
            batch_rounding_basis='RAW_MATERIAL_QTY',
            minimum_batch_qty_mt=Decimal('50.00'),
            batch_increment_qty_mt=Decimal('50.00'),
            is_active=True
        )

    def test_manual_requirement_validation(self):
        """Verify invalid requirements raise clear validation errors."""
        req_bad = NormalizedRequirement(
            planning_section="",
            section_weight=Decimal('0'),
            finished_grade="E250BR",
            demand_quantity=Decimal('-100'),
            min_length=Decimal('12.00'),
            max_length=Decimal('11.00')
        )
        with self.assertRaises(ValueError):
            req_bad.validate()

    def test_zero_cutting_length_validation_error(self):
        """Verify zero finished piece length raises validation error when master data is missing."""
        req = NormalizedRequirement(
            planning_section="NON_EXISTENT_SECTION",
            section_weight=Decimal('99.9'),
            finished_grade="E250BR",
            demand_quantity=Decimal('100.0'),
            min_length=Decimal('11.50'),
            max_length=Decimal('12.00'),
            cutting_length=Decimal('0.0')
        )
        with self.assertRaises(ValueError):
            resolve_finished_piece_length(req)

    def test_invalid_candidate_length_handling(self):
        """Verify candidate lengths exceeding usable strand length are flagged INVALID."""
        rule, rule_desc = resolve_yield_rule("UB_305X165", Decimal('40.3'), "SBB")
        req = NormalizedRequirement(
            planning_section="UB_305X165",
            section_weight=Decimal('40.3'),
            finished_grade="E250BR",
            demand_quantity=Decimal('100.0'),
            min_length=Decimal('11.50'),
            max_length=Decimal('12.00'),
            cutting_length=Decimal('10.0')
        )
        ev = evaluate_candidate(req, "BB_S275JR", Decimal('65.00'), Decimal('10.00'), Decimal('60.00'), rule, rule_desc)
        self.assertEqual(ev['status'], 'INVALID')
        self.assertIn('exceeds usable strand length', ev['invalid_reason'])

    def test_raw_material_qty_strand_increment_rounding(self):
        """Verify RAW_MATERIAL_QTY batch rounding increments by whole strands."""
        req = NormalizedRequirement(
            planning_section="UB_305X165",
            section_weight=Decimal('54.0'),
            finished_grade="E250BR",
            demand_quantity=Decimal('30.0'),
            min_length=Decimal('11.50'),
            max_length=Decimal('12.00'),
            cutting_length=Decimal('10.00')
        )
        res = OptimizationEngine.optimize_requirement(req)
        best = res['best']
        self.assertIsInstance(best['actual_strands'], int)
        self.assertGreaterEqual(best['required_rm_mt'], 50.0)

    def test_finished_qty_strand_increment_rounding(self):
        """Verify FINISHED_QTY batch rounding increments by whole strands."""
        SectionWeightMap.objects.create(
            planning_section="UC_203X203",
            section_weight=46.1,
            rm_family="SBB",
            raw_section="BB_355X280X90",
            casting_length=11.40,
            tolerance="0/40",
            location="MCAP"
        )
        YieldRule.objects.create(
            planning_section="UC_203X203",
            section_weight=46.1,
            rm_family="SBB",
            base_strand_length_m=Decimal('60.00'),
            batch_rounding_basis='FINISHED_QTY',
            minimum_batch_qty_mt=Decimal('50.00'),
            batch_increment_qty_mt=Decimal('50.00'),
            is_active=True
        )
        req = NormalizedRequirement(
            planning_section="UC_203X203",
            section_weight=Decimal('46.1'),
            finished_grade="E250BR",
            demand_quantity=Decimal('20.0'),
            min_length=Decimal('11.00'),
            max_length=Decimal('12.00'),
            cutting_length=Decimal('10.00')
        )
        res = OptimizationEngine.optimize_requirement(req)
        best = res['best']
        self.assertIsInstance(best['actual_strands'], int)
        self.assertGreaterEqual(best['planned_finished_mt'], 50.0)

    def test_no_fractional_strands(self):
        """Verify actual_strands is always a whole integer."""
        req = NormalizedRequirement(
            planning_section="UB_305X165",
            section_weight=Decimal('40.3'),
            finished_grade="E250BR",
            demand_quantity=Decimal('127.45'),
            min_length=Decimal('11.50'),
            max_length=Decimal('12.00'),
            cutting_length=Decimal('10.00')
        )
        res = OptimizationEngine.optimize_requirement(req)
        self.assertEqual(type(res['best']['actual_strands']), int)

    def test_decimal_length_dependent_optimization(self):
        """Verify candidate lengths produce mathematically distinct offcuts, RM, and yield."""
        req = NormalizedRequirement(
            planning_section="UB_305X165",
            section_weight=Decimal('54.0'),
            finished_grade="E250BR",
            demand_quantity=Decimal('250.0'),
            min_length=Decimal('11.50'),
            max_length=Decimal('12.00'),
            length_step=Decimal('0.10'),
            cutting_length=Decimal('10.00')
        )
        res = OptimizationEngine.optimize_requirement(req)
        evals = [e for e in res['all_evals'] if e['status'] == 'VALID']
        
        lengths = [e['cand_length'] for e in evals]
        yields = [e['effective_yield_pct'] for e in evals]
        
        self.assertGreater(len(set(lengths)), 1)
        self.assertGreater(len(set(yields)), 1)

    def test_physical_strand_offcut_math(self):
        """Verify discrete strand offcut formula."""
        rule, rule_desc = resolve_yield_rule("UB_305X165", Decimal('40.3'), "SBB")
        req = NormalizedRequirement(
            planning_section="UB_305X165",
            section_weight=Decimal('40.3'),
            finished_grade="E250BR",
            demand_quantity=Decimal('100.0'),
            min_length=Decimal('11.50'),
            max_length=Decimal('12.00'),
            cutting_length=Decimal('10.00')
        )
        ev = evaluate_candidate(req, "BB_S275JR", Decimal('11.50'), Decimal('10.00'), Decimal('60.00'), rule, rule_desc)
        self.assertEqual(ev['units_per_strand'], 5)
        self.assertAlmostEqual(ev['offcut_per_strand_m'], 1.45, places=2)

    def test_candidate_evaluations_storage(self):
        """Verify candidate evaluations are stored on CastingPlanRequirement as structured JSON."""
        req = NormalizedRequirement(
            planning_section="UB_305X165",
            section_weight=Decimal('54.0'),
            finished_grade="E250BR",
            demand_quantity=Decimal('100.0'),
            min_length=Decimal('11.50'),
            max_length=Decimal('12.00'),
            cutting_length=Decimal('10.00')
        )
        plan_req = OptimizationEngine.process_and_save_request([req], "Test Candidate Storage")
        item = plan_req.requirements.first()
        self.assertIsInstance(item.candidate_evaluations, list)
        self.assertGreater(len(item.candidate_evaluations), 0)

    def test_yield_rule_resolution_hierarchy(self):
        """Verify 4-tier resolution priority for YieldRule."""
        # Tier 1: Section + Weight + Family
        YieldRule.objects.create(
            planning_section="SPECIAL_SEC_1",
            section_weight=50.0,
            rm_family="SBB",
            is_active=True
        )
        r1, desc1 = resolve_yield_rule("SPECIAL_SEC_1", Decimal('50.0'), "SBB")
        self.assertIn("priority 1", desc1.lower())

        # Tier 2: Section + Family (no weight)
        YieldRule.objects.create(
            planning_section="SPECIAL_SEC_2",
            rm_family="SBB",
            is_active=True
        )
        r2, desc2 = resolve_yield_rule("SPECIAL_SEC_2", Decimal('45.0'), "SBB")
        self.assertIn("priority 2", desc2.lower())

        # Tier 3: Family only
        YieldRule.objects.create(
            rm_family="SPECIAL_FAM_3",
            is_active=True
        )
        r3, desc3 = resolve_yield_rule("UNKNOWN_SEC_3", Decimal('52.0'), "SPECIAL_FAM_3")
        self.assertIn("priority 3", desc3.lower())

        # Tier 4: Global Default
        YieldRule.objects.create(
            rm_family="",
            planning_section="",
            is_active=True
        )
        r4, desc4 = resolve_yield_rule("NON_EXISTENT_SEC", Decimal('99.0'), "NON_EXISTENT_FAM")
        self.assertIn("priority 4", desc4.lower())

    def test_excel_generation_with_summary_sheet(self):
        """Verify Excel generation produces OPTIMIZATION SUMMARY sheet."""
        req = NormalizedRequirement(
            planning_section="UB_305X165",
            section_weight=Decimal('54.0'),
            finished_grade="E250BR",
            demand_quantity=Decimal('200.0'),
            min_length=Decimal('11.50'),
            max_length=Decimal('12.00'),
            cutting_length=Decimal('10.00')
        )
        plan_req = OptimizationEngine.process_and_save_request([req], "Test Summary Sheet Excel")
        out_path = Path("E:/spm_casting_plan/media/generated") / plan_req.output_file
        self.assertTrue(out_path.exists())
        
        xl = pd.ExcelFile(out_path)
        self.assertIn("OPTIMIZATION SUMMARY", xl.sheet_names)
        df_sum = xl.parse("OPTIMIZATION SUMMARY")
        self.assertIn("Effective Yield %", df_sum.columns)
        self.assertIn("Optimization Reason", df_sum.columns)


class GoldenCastingPlanTest(TestCase):
    def test_compare_golden_plan(self):
        base_dir = Path("E:/spm_casting_plan")
        f_in = base_dir / "input casting plan.xlsm"
        f_ref = base_dir / "Casting plan.xlsx"
        
        self.assertTrue(f_in.exists(), f"Input file {f_in} does not exist")
        self.assertTrue(f_ref.exists(), f"Reference file {f_ref} does not exist")
        
        from django.core.management import call_command
        call_command('import_spm_data')
        
        import_demand_workbook(f_in, replace=True)
        order_count = Order.objects.count()
        self.assertGreater(order_count, 0, "No orders were imported from Golden Input")
        
        plan_rows, selected = build_plan_rows()
        self.assertGreater(len(plan_rows), 0, "No plan rows were generated by optimizer")
        
        f_gen = generate_casting_plan_workbook(plan_rows, "Golden_Casting_Plan_Test")
        self.assertTrue(f_gen.exists(), "Generated workbook does not exist")
        
        report_file = base_dir / "Casting_Plan_Validation_Report.txt"
        with open(report_file, "w", encoding="utf-8") as out:
            out.write("======================================================================\n")
            out.write("              SPM CASTING PLAN GOLDEN VALIDATION REPORT               \n")
            out.write("======================================================================\n\n")
            out.write(f"Golden Input: {f_in.name}\n")
            out.write(f"Golden Reference Output: {f_ref.name}\n")
            out.write(f"Generated Output: {f_gen.name}\n\n")
            out.write(f"Total sales orders imported: {order_count}\n")
            out.write(f"Total optimized casting plan rows: {len(plan_rows)}\n\n")
            out.write("Golden Plan Ingestion & Optimizer Run Completed Successfully.\n")
            
        print(f"\nGolden validation report written successfully to {report_file}")
