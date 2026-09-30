from django.db import models
from django.conf import settings

class Order(models.Model):
    so_no = models.CharField(max_length=80, blank=True)
    so_item = models.CharField(max_length=80, blank=True)
    material = models.CharField(max_length=180, blank=True)
    section = models.CharField(max_length=100, blank=True)
    grade = models.CharField(max_length=150, blank=True)
    quantity = models.FloatField(default=0)  # source/order quantity
    pgi_qty = models.FloatField(default=0)
    balance_qty = models.FloatField(default=0)
    eligible_stock = models.FloatField(default=0)
    net_demand = models.FloatField(default=0)
    cutting_length = models.FloatField(default=0)
    section_weight = models.FloatField(default=0)
    promise_date = models.CharField(max_length=40, blank=True)
    required_delivery_date = models.CharField(max_length=40, blank=True)
    priority = models.CharField(max_length=30, default='Normal')
    status = models.CharField(max_length=40, default='Pending')
    source_file = models.CharField(max_length=200, blank=True)
    raw_data = models.JSONField(default=dict, blank=True)
    is_selected = models.BooleanField(default=True)

    class Meta:
        ordering = ['section', 'id']

class SectionRule(models.Model):
    section = models.CharField(max_length=100, unique=True)
    min_length = models.FloatField(null=True)
    max_length = models.FloatField(null=True)
    raw_sizes = models.CharField(max_length=300, blank=True)
    two_v = models.CharField(max_length=80, blank=True)

class SectionWeightMap(models.Model):
    planning_section = models.CharField(max_length=100)
    section_weight = models.FloatField()
    rm_family = models.CharField(max_length=50)
    raw_section = models.CharField(max_length=100)
    casting_length = models.FloatField()
    tolerance = models.CharField(max_length=40, default="0/40")
    location = models.CharField(max_length=100, default="MCAP")
    class Meta:
        unique_together = ('planning_section', 'section_weight')

class GradeMap(models.Model):
    ext_grade = models.CharField(max_length=150)
    raw_higher = models.CharField(max_length=150, blank=True)
    raw_lighter = models.CharField(max_length=150, blank=True)
    std = models.CharField(max_length=80, blank=True)
    remarks = models.CharField(max_length=200, blank=True)

class BatchRule(models.Model):
    section = models.CharField(max_length=100)
    bars_per_batch = models.FloatField(default=0)
    section_weight = models.FloatField(default=0)
    cb_length = models.FloatField(default=0)
    lfa_batch_length = models.FloatField(default=0)
    bars_at_cb = models.FloatField(default=0)
    billet_length = models.FloatField(default=0)
    billet_sectional_weight = models.FloatField(default=0)

class PlanLine(models.Model):
    section = models.CharField(max_length=100)
    grade = models.CharField(max_length=150, blank=True)
    raw_grade = models.CharField(max_length=150, blank=True)
    rolling_qty = models.FloatField(default=0)
    billet_length = models.FloatField(default=0)
    rm_qty = models.FloatField(default=0)
    rolled_qty = models.FloatField(default=0)
    yield_pct = models.FloatField(default=0)
    source_file = models.CharField(max_length=200, blank=True)

class GeneratedPlan(models.Model):
    created_at = models.DateTimeField(auto_now_add=True)
    name = models.CharField(max_length=160)
    planning_date = models.DateField(null=True, blank=True)
    selected_order_count = models.PositiveIntegerField(default=0)
    demand_qty = models.FloatField(default=0)
    rm_qty = models.FloatField(default=0)
    yield_pct = models.FloatField(default=0)
    output_file = models.CharField(max_length=255, blank=True)


class BTRCampaign(models.Model):
    """Audit record for a campaign workbook generated from an uploaded BTR."""

    source_file = models.CharField(max_length=255)
    planning_material = models.CharField(max_length=120)
    input_row_count = models.PositiveIntegerField(default=0)
    output_row_count = models.PositiveIntegerField(default=0)
    skipped_row_count = models.PositiveIntegerField(default=0)
    zero_length_resolved_count = models.PositiveIntegerField(default=0)
    total_btr_qty = models.FloatField(default=0)
    bucket_counts = models.JSONField(default=dict, blank=True)
    output_file = models.CharField(max_length=255)
    rolling_plan_file = models.CharField(max_length=255, blank=True)
    output_file = models.CharField(max_length=255)
    rolling_plan_file = models.CharField(max_length=255, blank=True)
    casting_plan_file = models.CharField(max_length=255, blank=True)  
    embossing_plan_file = models.CharField(max_length=255, blank=True)
    rolling_plan_created_at = models.DateTimeField(null=True, blank=True)
    rolling_plan_created_at = models.DateTimeField(null=True, blank=True)
    rolling_plan_rm_qty = models.FloatField(default=0)
    rolling_plan_bloom_count = models.PositiveIntegerField(default=0)
    rolling_plan_yield_pct = models.FloatField(default=0)
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='btr_campaigns',
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-created_at']

class Notification(models.Model):
    title = models.CharField(max_length=200)
    message = models.TextField()
    is_read = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)
    class Meta:
        ordering = ['-created_at']

# New Models for Configurable Yield Loss Rules and Manual Casting Plans

class YieldRule(models.Model):
    """
    Plant production & yield loss rule system with 4-level resolution priority:
    1. Planning Section + Section Weight + RM Family
    2. Planning Section + RM Family
    3. RM Family
    4. Global Fallback
    """
    rm_family = models.CharField(max_length=50, blank=True)
    planning_section = models.CharField(max_length=100, blank=True)
    section_weight = models.FloatField(null=True, blank=True)
    
    base_strand_length_m = models.DecimalField(max_digits=10, decimal_places=3, default=60.000)
    head_crop_m = models.DecimalField(max_digits=10, decimal_places=3, default=0.500)
    tail_crop_m = models.DecimalField(max_digits=10, decimal_places=3, default=0.500)
    cut_loss_m = models.DecimalField(max_digits=10, decimal_places=3, default=0.010)
    process_loss_pct = models.DecimalField(max_digits=5, decimal_places=2, default=2.00)
    
    batch_rounding_basis = models.CharField(
        max_length=30,
        default='RAW_MATERIAL_QTY',
        choices=[
            ('RAW_MATERIAL_QTY', 'Raw Material Tonnage Increment'),
            ('FINISHED_QTY', 'Finished Production Tonnage Increment'),
            ('STRAND_COUNT', 'Strand Count Increment')
        ]
    )
    minimum_batch_qty_mt = models.DecimalField(max_digits=10, decimal_places=3, default=50.000)
    batch_increment_qty_mt = models.DecimalField(max_digits=10, decimal_places=3, default=50.000)
    minimum_batch_strands = models.PositiveIntegerField(default=1)
    batch_increment_strands = models.PositiveIntegerField(default=1)
    
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-is_active', 'planning_section', 'rm_family']


class CastingPlanRequest(models.Model):
    """
    Parent request container for manual or uploaded casting plan requests.
    """
    name = models.CharField(max_length=200)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    status = models.CharField(max_length=40, default='Draft')
    optimization_status = models.CharField(max_length=80, default='Pending')
    total_demand = models.FloatField(default=0.0)
    total_planned_qty = models.FloatField(default=0.0)
    avg_estimated_yield = models.FloatField(default=0.0)
    output_file = models.CharField(max_length=255, blank=True)

    class Meta:
        ordering = ['-created_at']


class CastingPlanRequirement(models.Model):
    """
    Child requirement model storing input specifications and complete optimization outcomes.
    """
    request = models.ForeignKey(CastingPlanRequest, on_delete=models.CASCADE, related_name='requirements')
    
    # Input Fields
    planning_section = models.CharField(max_length=100)
    section_weight = models.FloatField()
    finished_grade = models.CharField(max_length=150)
    demand_quantity = models.FloatField()
    requirement_type = models.CharField(max_length=40, default='Firm')
    priority = models.CharField(max_length=30, default='Medium')
    
    min_length = models.FloatField()
    max_length = models.FloatField()
    preferred_length = models.FloatField(null=True, blank=True)
    length_step = models.FloatField(default=0.01)
    cutting_length = models.FloatField(default=0.0)
    tolerance = models.CharField(max_length=40, default='0/40')
    location = models.CharField(max_length=100, default='MCAP')
    remarks = models.CharField(max_length=255, blank=True)
    
    # Resolved / Optimized Output Fields
    resolved_rm_family = models.CharField(max_length=50, blank=True)
    resolved_rm_section = models.CharField(max_length=100, blank=True)
    resolved_rm_grade = models.CharField(max_length=150, blank=True)
    selected_casting_length = models.FloatField(null=True, blank=True)
    finished_piece_length = models.FloatField(null=True, blank=True)
    planned_quantity = models.FloatField(default=0.0)
    required_rm_quantity = models.FloatField(default=0.0)
    estimated_yield_pct = models.FloatField(default=0.0)
    
    crop_loss_mt = models.FloatField(default=0.0)
    offcut_loss_mt = models.FloatField(default=0.0)
    process_loss_mt = models.FloatField(default=0.0)
    excess_quantity = models.FloatField(default=0.0)
    
    rule_used_description = models.CharField(max_length=255, blank=True)
    optimization_reason = models.TextField(blank=True)
    candidate_evaluations = models.JSONField(default=list, blank=True)

    class Meta:
        ordering = ['id']

class BilletLengthSetting(models.Model):
    cold_saw_cut = models.FloatField(default=0.0)
    cooling_bed_crop = models.FloatField(default=0.4)
    div_sheer_loss = models.FloatField(default=0.0)
    base_crop_loss = models.FloatField(default=0.0)
    section_wt_tolerance_pct = models.FloatField(default=0.5)
    hot_pieces_at_cb = models.IntegerField(default=1)

    class Meta:
        verbose_name = "Billet Length Formula Setting"

class RootRadiusMaster(models.Model):
    planning_material = models.CharField(max_length=50, help_text="e.g., AL_130X130")
    thickness = models.FloatField()
    sec_wt = models.FloatField(null=True, blank=True, help_text="Corrected Section Weight")
    std = models.CharField(max_length=20, help_text="e.g., BIS or BSEN")
    root_radius = models.FloatField()

    class Meta:
        verbose_name_plural = "Root Radius Master (M-1)"
        unique_together = ('planning_material', 'thickness', 'sec_wt', 'std')

    def __str__(self):
        return f"{self.planning_material} | {self.thickness} | {self.std} -> RR: {self.root_radius}"


# --- New Master Data Models for Rail Mill Length Calculation Workbook ---

class RailMillCalculationMaster(models.Model):
    planning_mat = models.CharField(max_length=100, blank=True)
    rolled_section = models.CharField(max_length=150, blank=True)
    material = models.CharField(max_length=150, blank=True)
    product = models.CharField(max_length=80, blank=True)
    sec_wt = models.FloatField(null=True, blank=True)
    ext_grade = models.CharField(max_length=100, blank=True)
    grade_group = models.CharField(max_length=80, blank=True)
    rm_grade = models.CharField(max_length=100, blank=True)
    rm_size = models.CharField(max_length=100, blank=True)
    bloom_length_m = models.FloatField(null=True, blank=True)
    bloom_weight_kg_m = models.FloatField(null=True, blank=True)
    total_rolled_length_m = models.FloatField(null=True, blank=True)
    total_finish_length_m = models.FloatField(null=True, blank=True)
    cut_pcs = models.CharField(max_length=100, blank=True)
    end_cut_margin = models.FloatField(null=True, blank=True)
    rm_calculated_length_m = models.FloatField(null=True, blank=True)
    end_cut_margin_req = models.FloatField(null=True, blank=True)
    maximum_cooling_bed_length_m = models.FloatField(null=True, blank=True)
    old_furnace_min_len = models.FloatField(null=True, blank=True)
    old_furnace_max_len = models.FloatField(null=True, blank=True)
    new_furnace_min_len = models.FloatField(null=True, blank=True)
    new_furnace_max_len = models.FloatField(null=True, blank=True)

    class Meta:
        verbose_name_plural = "Rail Mill Calculation Master"


class RailMillChemistryMaster(models.Model):
    sl_no = models.CharField(max_length=40, blank=True)
    product = models.CharField(max_length=100, blank=True)
    internal_grade = models.CharField(max_length=100, blank=True)
    proposed_internal_grade = models.CharField(max_length=100, blank=True)
    ext_grade = models.CharField(max_length=200, blank=True)
    sections = models.CharField(max_length=200, blank=True)
    range_type = models.CharField(max_length=30, blank=True)  # Min / Max
    c_pct = models.CharField(max_length=40, blank=True)
    mn_pct = models.CharField(max_length=40, blank=True)
    s_pct = models.CharField(max_length=40, blank=True)
    p_pct = models.CharField(max_length=40, blank=True)
    si_pct = models.CharField(max_length=40, blank=True)
    al_pct = models.CharField(max_length=40, blank=True)
    cu_pct = models.CharField(max_length=40, blank=True)
    ni_pct = models.CharField(max_length=40, blank=True)
    cr_pct = models.CharField(max_length=40, blank=True)
    mo_pct = models.CharField(max_length=40, blank=True)
    ti_pct = models.CharField(max_length=40, blank=True)
    v_pct = models.CharField(max_length=40, blank=True)
    sn_pct = models.CharField(max_length=40, blank=True)
    nb_pct = models.CharField(max_length=40, blank=True)
    b_pct = models.CharField(max_length=40, blank=True)
    as_pct = models.CharField(max_length=40, blank=True)
    ce = models.CharField(max_length=40, blank=True)
    n2_ppm = models.CharField(max_length=40, blank=True)
    o2_ppm = models.CharField(max_length=40, blank=True)
    h2_ppm = models.CharField(max_length=40, blank=True)

    class Meta:
        verbose_name_plural = "Rail Mill Chemistry Master"


class RailMillSectionGradeStandardMap(models.Model):
    sl_no = models.CharField(max_length=40, blank=True)
    rolled_section = models.CharField(max_length=150, blank=True)
    flange_thickness_mm = models.FloatField(null=True, blank=True)
    product = models.CharField(max_length=80, blank=True)
    std_e250_e275 = models.CharField(max_length=150, blank=True)
    std_e350_s355 = models.CharField(max_length=150, blank=True)
    std_e410_jis = models.CharField(max_length=150, blank=True)
    std_s420n = models.CharField(max_length=150, blank=True)
    std_s460_e450 = models.CharField(max_length=150, blank=True)
    std_s500 = models.CharField(max_length=150, blank=True)
    std_e550 = models.CharField(max_length=150, blank=True)
    std_s355nl = models.CharField(max_length=150, blank=True)
    fire_resistance = models.CharField(max_length=150, blank=True)
    is_3443 = models.CharField(max_length=150, blank=True)
    irs_t12 = models.CharField(max_length=150, blank=True)
    en_13674 = models.CharField(max_length=150, blank=True)

    class Meta:
        verbose_name_plural = "Rail Mill Section Grade Standards"


class RailMillSecWiseMaster(models.Model):
    section = models.CharField(max_length=100, blank=True)
    section_weight_desc = models.CharField(max_length=150, blank=True)
    input_size = models.CharField(max_length=100, blank=True)
    bloom_length_m = models.FloatField(null=True, blank=True)
    bloom_weight_kg_m = models.FloatField(null=True, blank=True)
    total_rolled_length_m = models.FloatField(null=True, blank=True)
    total_finish_length_m = models.FloatField(null=True, blank=True)
    cut_pcs = models.CharField(max_length=150, blank=True)

    class Meta:
        verbose_name_plural = "Rail Mill Section Wise Matrix"


class RailMillSawTolerance(models.Model):
    condition = models.CharField(max_length=50)  # e.g. ">8.5" or "<8.5"
    tolerance_m = models.FloatField()

    class Meta:
        verbose_name_plural = "Rail Mill Saw Length Tolerances"


class RailMillRMSectionWeight(models.Model):
    rm_sec = models.CharField(max_length=100)  # e.g. "BL 285 x 390"
    weight_t_per_m = models.FloatField()

    class Meta:
        verbose_name_plural = "Rail Mill RM Section Weights"


class RailMillPlanningMaterial(models.Model):
    planning_mat = models.CharField(max_length=100, unique=True)

    class Meta:
        verbose_name_plural = "Rail Mill Planning Materials"

