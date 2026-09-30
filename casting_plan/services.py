from collections import defaultdict
from copy import copy
from datetime import date
from decimal import Decimal, ROUND_HALF_UP, ROUND_CEILING
import math
from pathlib import Path
import re
import uuid

from django.conf import settings
from openpyxl import load_workbook

from .models import (
    Order, SectionRule, GradeMap, BatchRule, SectionWeightMap,
    YieldRule, CastingPlanRequest, CastingPlanRequirement
)


def _num(v):
    try:
        if v in (None, '', '-', ' '):
            return 0.0
        return float(v)
    except Exception:
        return 0.0


def _dec(v, default="0.0"):
    try:
        if v in (None, '', '-', ' '):
            return Decimal(default)
        return Decimal(str(v).strip())
    except Exception:
        return Decimal(default)


def _quantize(val, places=3):
    if not isinstance(val, Decimal):
        val = Decimal(str(val))
    fmt = "0." + "0" * places if places > 0 else "0"
    return val.quantize(Decimal(fmt), rounding=ROUND_HALF_UP)


class NormalizedRequirement:
    """
    Unified, high-precision Decimal internal representation for casting plan requirements.
    """
    def __init__(
        self,
        planning_section: str,
        section_weight: Decimal,
        finished_grade: str,
        demand_quantity: Decimal,
        min_length: Decimal,
        max_length: Decimal,
        preferred_length: Decimal = None,
        length_step: Decimal = Decimal('0.01'),
        cutting_length: Decimal = Decimal('0.0'),
        tolerance: str = "0/40",
        location: str = "MCAP",
        remarks: str = "",
        requirement_type: str = "Firm",
        priority: str = "Medium",
        source_id: str = None
    ):
        self.planning_section = str(planning_section or '').strip()
        self.section_weight = _quantize(_dec(section_weight), 2)
        self.finished_grade = str(finished_grade or '').strip()
        self.demand_quantity = _quantize(_dec(demand_quantity), 3)
        self.min_length = _quantize(_dec(min_length), 2)
        self.max_length = _quantize(_dec(max_length), 2)
        self.preferred_length = _quantize(_dec(preferred_length), 2) if preferred_length is not None and preferred_length > 0 else None
        self.length_step = _dec(length_step, '0.01') if length_step > 0 else Decimal('0.01')
        self.cutting_length = _quantize(_dec(cutting_length), 2)
        self.tolerance = str(tolerance or '0/40').strip()
        self.location = str(location or 'MCAP').strip()
        self.remarks = str(remarks or '').strip()
        self.requirement_type = str(requirement_type or 'Firm').strip()
        self.priority = str(priority or 'Medium').strip()
        self.source_id = source_id

    def validate(self):
        errors = []
        if not self.planning_section:
            errors.append("Planning Section is required.")
        if self.section_weight <= Decimal('0'):
            errors.append(f"Section weight must be > 0 (got {self.section_weight}).")
        if not self.finished_grade:
            errors.append("Finished Grade is required.")
        if self.demand_quantity <= Decimal('0'):
            errors.append(f"Demand quantity must be > 0 (got {self.demand_quantity}).")
        if self.min_length <= Decimal('0'):
            errors.append(f"Minimum length must be > 0 (got {self.min_length}).")
        if self.max_length < self.min_length:
            errors.append(f"Minimum length ({self.min_length}m) cannot exceed maximum length ({self.max_length}m).")
        if self.preferred_length is not None:
            if self.preferred_length < self.min_length or self.preferred_length > self.max_length:
                errors.append(f"Preferred length ({self.preferred_length}m) must be within [{self.min_length}m, {self.max_length}m].")
        if errors:
            raise ValueError(" ; ".join(errors))


def choose_raw_grade_candidates(ext, section=None):
    ext_clean = (ext or '').strip()
    rows = list(GradeMap.objects.filter(ext_grade__iexact=ext_clean))
    if not rows:
        return [ext_clean or 'UNKNOWN']
    
    candidates = []
    for r in rows:
        if r.raw_higher and r.raw_higher.strip():
            candidates.append(r.raw_higher.strip())
        if r.raw_lighter and r.raw_lighter.strip():
            candidates.append(r.raw_lighter.strip())
            
    if not candidates:
        candidates.append(ext_clean or 'UNKNOWN')
        
    seen = set()
    unique = []
    for c in candidates:
        if c not in seen:
            seen.add(c)
            unique.append(c)
    return unique


def choose_raw_grade(ext, qty=0, section=None):
    cands = choose_raw_grade_candidates(ext, section)
    return cands[0] if cands else (ext or 'UNKNOWN').strip()


def lookup_casting_rules(planning_section, section_weight):
    sec_clean = str(planning_section).strip().upper()
    wt_val = float(section_weight) if isinstance(section_weight, (Decimal, float)) else float(_num(section_weight))
    
    row = SectionWeightMap.objects.filter(planning_section__iexact=sec_clean, section_weight=wt_val).first()
    if row:
        return row, None
    sec_alt = sec_clean.replace('_', ' ').replace('-', ' ')
    row = SectionWeightMap.objects.filter(planning_section__iexact=sec_alt, section_weight=wt_val).first()
    if row:
        return row, None
    return None, f"Missing SectionWeightMap for Section={planning_section}, Sec Wt={section_weight}"


def raw_family(section, source_rm=''):
    rm = (source_rm or '').upper()
    if 'SBB' in rm:
        return 'SBB', 'BB_355X280X90', '355X280X90', 'SBB'
    if '250' in rm:
        return '250sq', 'BL_250X250', '250X250', '250sq'
    if '165' in rm:
        return '165SQ', 'BL_165X165', '165X165', '165SQ'
        
    sec = (section or '').upper()
    if any(x in sec for x in ['305','356','406','254']):
        return 'SBB', 'BB_355X280X90', '355X280X90', 'SBB'
    if any(x in sec for x in ['203','200','180','152']):
        return '250sq', 'BL_250X250', '250X250', '250sq'
    return '165SQ', 'BL_165X165', '165X165', '165SQ'


def resolve_yield_rule(planning_section, section_weight, rm_family):
    sec = (planning_section or '').strip()
    wt = float(section_weight) if section_weight is not None else None
    fam = (rm_family or '').strip()

    if sec and wt is not None and fam:
        r1 = YieldRule.objects.filter(is_active=True, planning_section__iexact=sec, section_weight=wt, rm_family__iexact=fam).first()
        if r1:
            return r1, f"Rule priority 1: {sec} ({wt} kg/m) [{fam}]"
            
    if sec and fam:
        r2 = YieldRule.objects.filter(is_active=True, planning_section__iexact=sec, rm_family__iexact=fam, section_weight__isnull=True).first()
        if r2:
            return r2, f"Rule priority 2: {sec} [{fam}]"
            
    if fam:
        r3 = YieldRule.objects.filter(is_active=True, rm_family__iexact=fam, planning_section='').first()
        if r3:
            return r3, f"Rule priority 3: Family [{fam}]"
            
    r4 = YieldRule.objects.filter(is_active=True, rm_family='', planning_section='').first()
    if r4:
        return r4, "Rule priority 4: Global Default Rule"
        
    class FallbackRule:
        base_strand_length_m = Decimal('60.00')
        head_crop_m = Decimal('0.50')
        tail_crop_m = Decimal('0.50')
        cut_loss_m = Decimal('0.010')
        process_loss_pct = Decimal('2.00')
        batch_rounding_basis = 'RAW_MATERIAL_QTY'
        minimum_batch_qty_mt = Decimal('50.00')
        batch_increment_qty_mt = Decimal('50.00')
        minimum_batch_strands = 1
        batch_increment_strands = 1
        
    return FallbackRule(), "Rule priority 4: Estimated Yield using Default Fallback Rule"


def resolve_finished_piece_length(req: NormalizedRequirement) -> Decimal:
    if req.cutting_length > Decimal('0'):
        return req.cutting_length
        
    map_row, _ = lookup_casting_rules(req.planning_section, req.section_weight)
    if map_row and map_row.casting_length > 0:
        return _dec(map_row.casting_length)
        
    batch = BatchRule.objects.filter(section__iexact=req.planning_section).first()
    if batch:
        if batch.billet_length > 0:
            return _dec(batch.billet_length)
        if batch.cb_length > 0:
            return _dec(batch.cb_length)
            
    raise ValueError(f"Invalid finished piece length for section {req.planning_section} ({req.section_weight} kg/m). Cutting length is 0 and no master data length is configured.")


def generate_candidate_lengths(min_l: Decimal, max_l: Decimal, step: Decimal, pref_l: Decimal = None) -> list:
    candidates = set()
    curr = min_l
    while curr <= max_l:
        candidates.add(_quantize(curr, 2))
        curr += step
    candidates.add(_quantize(max_l, 2))
    
    if pref_l is not None and min_l <= pref_l <= max_l:
        candidates.add(_quantize(pref_l, 2))
        
    return sorted(list(candidates))


def evaluate_candidate(
    req: NormalizedRequirement,
    rm_grade: str,
    cand_len: Decimal,
    finished_len: Decimal,
    base_strand_len: Decimal,
    rule,
    rule_desc: str
) -> dict:
    """
    High-precision physics evaluation for a single candidate (RM Grade + Candidate Casting Length).
    Applies physical integer strand batch rounding safeguard with linear O(1) convergence.
    """
    # 1. Finished piece weight (MT)
    piece_w_mt = (req.section_weight * finished_len) / Decimal('1000.0')
    if piece_w_mt <= Decimal('0'):
        return {
            'status': 'INVALID',
            'invalid_reason': f'Calculated piece weight is <= 0 for length {finished_len}m',
            'cand_length': float(cand_len), 'rm_grade': rm_grade
        }
        
    # 2. Pieces required
    pieces_required = int(math.ceil(req.demand_quantity / piece_w_mt))
    
    # 3. Usable strand length & units per strand check
    usable_strand = base_strand_len - rule.head_crop_m - rule.tail_crop_m
    unit_slot = cand_len + rule.cut_loss_m
    
    if usable_strand <= Decimal('0') or unit_slot <= Decimal('0'):
        return {
            'status': 'INVALID',
            'invalid_reason': 'Usable strand length or candidate slot length is <= 0',
            'cand_length': float(cand_len), 'rm_grade': rm_grade
        }
        
    units_per_strand = int(math.floor(usable_strand / unit_slot))
    
    if units_per_strand <= 0:
        return {
            'status': 'INVALID',
            'invalid_reason': f'Candidate length {cand_len}m exceeds usable strand length {usable_strand:.2f}m',
            'cand_length': float(cand_len), 'rm_grade': rm_grade
        }
        
    offcut_per_strand = usable_strand - (Decimal(units_per_strand) * unit_slot)
    
    # 4. Minimum required whole strands for demand
    raw_strands = int(math.ceil(Decimal(pieces_required) / Decimal(units_per_strand)))
    actual_strands = raw_strands
    
    # 5. Whole strand batch rounding loop
    basis = getattr(rule, 'batch_rounding_basis', 'RAW_MATERIAL_QTY')
    
    if basis == 'STRAND_COUNT':
        min_strands = int(getattr(rule, 'minimum_batch_strands', 1) or 1)
        inc_strands = int(getattr(rule, 'batch_increment_strands', 1) or 1)
        if actual_strands < min_strands:
            actual_strands = min_strands
        else:
            diff = actual_strands - min_strands
            rem = diff % inc_strands
            if rem != 0:
                actual_strands += (inc_strands - rem)
    else:
        # MT Tonnage Based Rounding (RAW_MATERIAL_QTY or FINISHED_QTY)
        min_mt = getattr(rule, 'minimum_batch_qty_mt', Decimal('50.0')) or Decimal('50.0')
        inc_mt = getattr(rule, 'batch_increment_qty_mt', Decimal('50.0')) or Decimal('50.0')
        
        # Determine target minimum batch tonnage threshold
        if inc_mt > Decimal('0'):
            target_min_mt = max(min_mt, Decimal(math.ceil(req.demand_quantity / inc_mt)) * inc_mt)
        else:
            target_min_mt = min_mt

        # Loop whole strands incrementing by +1 strand until target tonnage condition is satisfied
        while True:
            act_pieces = actual_strands * units_per_strand
            plan_fin_mt = Decimal(act_pieces) * piece_w_mt
            
            h_crop_mt = Decimal(actual_strands) * rule.head_crop_m * (req.section_weight / Decimal('1000.0'))
            t_crop_mt = Decimal(actual_strands) * rule.tail_crop_m * (req.section_weight / Decimal('1000.0'))
            kerf_mt = Decimal(actual_strands * units_per_strand) * rule.cut_loss_m * (req.section_weight / Decimal('1000.0'))
            offcut_mt = Decimal(actual_strands) * offcut_per_strand * (req.section_weight / Decimal('1000.0'))
            
            net_solid_mt = plan_fin_mt + h_crop_mt + t_crop_mt + kerf_mt + offcut_mt
            proc_loss_mt = net_solid_mt * (rule.process_loss_pct / Decimal('100.0'))
            req_rm_mt = net_solid_mt + proc_loss_mt
            
            target_val = req_rm_mt if basis == 'RAW_MATERIAL_QTY' else plan_fin_mt
            
            if target_val >= target_min_mt:
                break
                
            actual_strands += 1

    # 6. Recalculate final metrics strictly from integer actual_strands
    actual_pieces = actual_strands * units_per_strand
    planned_finished_mt = Decimal(actual_pieces) * piece_w_mt
    excess_mt = planned_finished_mt - req.demand_quantity
    
    head_crop_mt = Decimal(actual_strands) * rule.head_crop_m * (req.section_weight / Decimal('1000.0'))
    tail_crop_mt = Decimal(actual_strands) * rule.tail_crop_m * (req.section_weight / Decimal('1000.0'))
    kerf_loss_mt = Decimal(actual_strands * units_per_strand) * rule.cut_loss_m * (req.section_weight / Decimal('1000.0'))
    offcut_loss_mt = Decimal(actual_strands) * offcut_per_strand * (req.section_weight / Decimal('1000.0'))
    
    net_solid_mt = planned_finished_mt + head_crop_mt + tail_crop_mt + kerf_loss_mt + offcut_loss_mt
    process_loss_mt = net_solid_mt * (rule.process_loss_pct / Decimal('100.0'))
    required_rm_mt = net_solid_mt + process_loss_mt
    
    effective_yield_pct = (req.demand_quantity / required_rm_mt) * Decimal('100.0') if required_rm_mt > 0 else Decimal('0.0')
    total_crop_offcut_kerf = head_crop_mt + tail_crop_mt + kerf_loss_mt + offcut_loss_mt
    pref_dist = abs(cand_len - req.preferred_length) if req.preferred_length is not None else Decimal('999.0')
    
    return {
        'status': 'VALID',
        'invalid_reason': '',
        'cand_length': float(cand_len),
        'rm_grade': rm_grade,
        'finished_piece_length': float(finished_len),
        'finished_piece_weight_mt': float(_quantize(piece_w_mt, 4)),
        'pieces_required': pieces_required,
        'units_per_strand': units_per_strand,
        'usable_strand_m': float(_quantize(usable_strand, 2)),
        'offcut_per_strand_m': float(_quantize(offcut_per_strand, 3)),
        'actual_strands': actual_strands,
        'actual_pieces': actual_pieces,
        'planned_finished_mt': float(_quantize(planned_finished_mt, 3)),
        'required_rm_mt': float(_quantize(required_rm_mt, 3)),
        'effective_yield_pct': float(_quantize(effective_yield_pct, 2)),
        'head_crop_mt': float(_quantize(head_crop_mt, 3)),
        'tail_crop_mt': float(_quantize(tail_crop_mt, 3)),
        'kerf_loss_mt': float(_quantize(kerf_loss_mt, 3)),
        'crop_loss_mt': float(_quantize(head_crop_mt + tail_crop_mt + kerf_loss_mt, 3)),
        'offcut_loss_mt': float(_quantize(offcut_loss_mt, 3)),
        'process_loss_mt': float(_quantize(process_loss_mt, 3)),
        'excess_mt': float(_quantize(excess_mt, 3)),
        'pref_dist': float(pref_dist),
        'rule_desc': rule_desc,
        'score_tuple': (
            -float(effective_yield_pct),
            float(required_rm_mt),
            float(total_crop_offcut_kerf),
            float(excess_mt),
            float(pref_dist),
            float(cand_len)
        )
    }


class OptimizationEngine:
    """
    Main core optimization pipeline shared by both Manual Input and Excel Upload.
    """
    @staticmethod
    def optimize_requirement(req: NormalizedRequirement) -> dict:
        req.validate()
        
        # 1. Resolve Finished Piece Length
        finished_piece_len = resolve_finished_piece_length(req)
        
        # 2. Section & Weight Map Lookup
        map_row, map_err = lookup_casting_rules(req.planning_section, req.section_weight)
        if map_row:
            sheet_name = map_row.rm_family
            raw_section = map_row.raw_section
            tolerance = req.tolerance or map_row.tolerance
            location = req.location or map_row.location
        else:
            sheet_name, raw_section, raw_size, family = raw_family(req.planning_section)
            tolerance = req.tolerance or '0/40'
            location = req.location or 'MCAP'
            
        # 3. Resolve Yield Rule
        rule, rule_desc = resolve_yield_rule(req.planning_section, req.section_weight, sheet_name)
        
        # 4. Determine Strand Base Length
        base_strand_len = _dec(getattr(rule, 'base_strand_length_m', '60.0'))
        batch = BatchRule.objects.filter(section__iexact=req.planning_section).first()
        if batch and batch.billet_length > 0:
            base_strand_len = Decimal(str(batch.billet_length)) * Decimal('4.0')
            
        # 5. Resolve RM Grade Candidates
        rm_grades = choose_raw_grade_candidates(req.finished_grade, req.planning_section)
        
        # 6. Generate Candidate Casting Lengths
        cand_lengths = generate_candidate_lengths(req.min_length, req.max_length, req.length_step, req.preferred_length)
        
        # 7. Evaluate All Candidates
        all_evals = []
        valid_evals = []
        
        for rm_g in rm_grades:
            for cand_l in cand_lengths:
                ev = evaluate_candidate(req, rm_g, cand_l, finished_piece_len, base_strand_len, rule, rule_desc)
                all_evals.append(ev)
                if ev['status'] == 'VALID':
                    valid_evals.append(ev)
                    
        if not valid_evals:
            raise ValueError(f"No feasible candidate casting length could be found for {req.planning_section} between {req.min_length}m and {req.max_length}m. All candidate lengths were invalid relative to usable strand bounds.")
            
        # 8. Deterministic Candidate Scoring & Selection
        valid_evals.sort(key=lambda x: x['score_tuple'])
        best = valid_evals[0]
        
        raw_size = raw_section.split('_')[-1]
        output_grade = f"{raw_section.split('_')[0]}_{best['rm_grade']}_{raw_size}"
        sec_wt_str = f"{float(req.section_weight):.1f}".rstrip('0').rstrip('.')
        remarks = req.remarks or f"for {sec_wt_str} kg"
        
        opt_reason = (
            f"Selected candidate length {best['cand_length']:.2f}m with grade {best['rm_grade']} "
            f"achieving {best['effective_yield_pct']:.2f}% effective yield. Required RM: {best['required_rm_mt']:.2f} MT, "
            f"Planned Production: {best['planned_finished_mt']:.2f} MT (Excess: {best['excess_mt']:.2f} MT) over {best['actual_strands']} strands."
        )
        
        return {
            'req': req,
            'sheet': sheet_name,
            'raw_section': raw_section,
            'output_grade': output_grade,
            'raw_grade': best['rm_grade'],
            'best': best,
            'all_evals': all_evals,
            'rule_desc': rule_desc,
            'opt_reason': opt_reason,
            'tolerance': tolerance,
            'location': location,
            'remarks': remarks
        }

    @staticmethod
    def process_and_save_request(req_objects: list, plan_name: str = None) -> CastingPlanRequest:
        if not req_objects:
            raise ValueError("No requirements supplied for optimization.")
            
        total_demand = sum(float(r.demand_quantity) for r in req_objects)
        plan_req = CastingPlanRequest.objects.create(
            name=plan_name or f"Casting Plan {date.today().isoformat()}",
            status="Processing",
            total_demand=round(total_demand, 3)
        )
        
        results = []
        for r in req_objects:
            res = OptimizationEngine.optimize_requirement(r)
            results.append(res)
            best = res['best']
            
            CastingPlanRequirement.objects.create(
                request=plan_req,
                planning_section=r.planning_section,
                section_weight=float(r.section_weight),
                finished_grade=r.finished_grade,
                demand_quantity=float(r.demand_quantity),
                requirement_type=r.requirement_type,
                priority=r.priority,
                min_length=float(r.min_length),
                max_length=float(r.max_length),
                preferred_length=float(r.preferred_length) if r.preferred_length is not None else None,
                length_step=float(r.length_step),
                cutting_length=float(r.cutting_length),
                tolerance=res['tolerance'],
                location=res['location'],
                remarks=res['remarks'],
                resolved_rm_family=res['sheet'],
                resolved_rm_section=res['raw_section'],
                resolved_rm_grade=res['output_grade'],
                selected_casting_length=best['cand_length'],
                finished_piece_length=best['finished_piece_length'],
                planned_quantity=best['planned_finished_mt'],
                required_rm_quantity=best['required_rm_mt'],
                estimated_yield_pct=best['effective_yield_pct'],
                crop_loss_mt=best['crop_loss_mt'],
                offcut_loss_mt=best['offcut_loss_mt'],
                process_loss_mt=best['process_loss_mt'],
                excess_quantity=best['excess_mt'],
                rule_used_description=res['rule_desc'],
                optimization_reason=res['opt_reason'],
                candidate_evaluations=res['all_evals']
            )
            
        total_planned = sum(res['best']['planned_finished_mt'] for res in results)
        total_rm = sum(res['best']['required_rm_mt'] for res in results)
        avg_yield = (total_demand / total_rm * 100.0) if total_rm > 0 else 0.0
        
        plan_rows = []
        for res in results:
            b = res['best']
            r = res['req']
            plan_rows.append({
                'sheet': res['sheet'],
                'section': res['raw_section'],
                'grade': res['output_grade'],
                'qty': b['planned_finished_mt'],
                'demand_qty': float(r.demand_quantity),
                'length': b['cand_length'],
                'finished_piece_length': b['finished_piece_length'],
                'tolerance': res['tolerance'],
                'location': res['location'],
                'remarks': res['remarks'],
                'source_section': r.planning_section,
                'source_grade': r.finished_grade,
                'raw_grade': b['rm_grade'],
                'rm_qty': b['required_rm_mt'],
                'crop_loss_mt': b['crop_loss_mt'],
                'offcut_loss_mt': b['offcut_loss_mt'],
                'process_loss_mt': b['process_loss_mt'],
                'excess_mt': b['excess_mt'],
                'yield_pct': b['effective_yield_pct'],
                'rule_desc': res['rule_desc'],
                'opt_reason': res['opt_reason'],
                'sec_wt': float(r.section_weight),
                'req_type': r.requirement_type
            })
            
        wb_path = generate_casting_plan_workbook(plan_rows, plan_req.name)
        
        plan_req.status = "Completed"
        plan_req.optimization_status = "Optimized"
        plan_req.total_planned_qty = round(total_planned, 3)
        plan_req.avg_estimated_yield = round(avg_yield, 2)
        plan_req.output_file = wb_path.name
        plan_req.save()
        
        return plan_req


def optimize(section=None, grade=None, target_qty=None):
    """
    Backwards-compatible API wrapper returning legacy result dict.
    """
    qs = Order.objects.filter(is_selected=True)
    if section:
        qs = qs.filter(section__iexact=section)
    if grade:
        qs = qs.filter(grade__iexact=grade)
    rows = list(qs)
    if not rows:
        return {'status':'No matching orders','section':section or 'All','grade':grade or 'All','required':0,'raw_grade':'-','billet_length':0,'billets':0,'rm_qty':0,'finished_qty':0,'yield_pct':0,'saving':0,'reason':'No selected demand found.'}
        
    req_qty = float(target_qty or sum(max(0, _num(r.net_demand or r.balance_qty or r.quantity)) for r in rows))
    sec = section or rows[0].section
    grd = grade or rows[0].grade
    sec_wt = rows[0].section_weight if rows[0].section_weight > 0 else 40.3
    cut_len = rows[0].cutting_length if rows[0].cutting_length > 0 else 10.0
    
    n_req = NormalizedRequirement(
        planning_section=sec,
        section_weight=Decimal(str(sec_wt)),
        finished_grade=grd,
        demand_quantity=Decimal(str(req_qty)),
        min_length=Decimal('10.0'),
        max_length=Decimal('12.0'),
        cutting_length=Decimal(str(cut_len))
    )
    
    res = OptimizationEngine.optimize_requirement(n_req)
    b = res['best']
    
    return {
        'status': 'Optimized', 'section': sec, 'grade': grd, 'raw_grade': res['raw_grade'],
        'required': round(req_qty, 3), 'billet_length': round(b['cand_length'], 3),
        'billets': b['actual_strands'], 'rm_qty': round(b['required_rm_mt'], 3),
        'finished_qty': round(b['planned_finished_mt'], 3), 'yield_pct': round(b['effective_yield_pct'], 2),
        'saving': round(max(0, (req_qty / 0.93) - b['required_rm_mt']), 3),
        'min_length': 10.0, 'max_length': 12.0, 'reason': res['opt_reason']
    }


def _clean_header(v):
    return re.sub(r'\s+', ' ', str(v or '').strip()).lower()


def _pick(d, *names):
    for n in names:
        key = _clean_header(n)
        if key in d:
            return d[key]
    return None


def import_demand_workbook(path, replace=True):
    path = Path(path)
    if path.suffix.lower() == '.xls':
        import pandas as pd
        xl = pd.ExcelFile(path)
        sheet_name = next((s for s in xl.sheet_names if 'PLAN' in s.upper()), None)
        if not sheet_name:
            raise ValueError(f'No PLAN sheet found in {path.name}.')
        df = xl.parse(sheet_name, header=None)
        all_rows = []
        for idx, r in df.iterrows():
            row_vals = [None if pd.isna(x) else x for x in r]
            all_rows.append(row_vals)
        if not all_rows:
            raise ValueError(f'Empty plan sheet in {path.name}.')
        headers = all_rows[0]
        rows = all_rows[1:]
    else:
        wb = load_workbook(path, read_only=True, data_only=True, keep_vba=path.suffix.lower()=='.xlsm')
        sheet_name = next((s for s in wb.sheetnames if 'PLAN' in s.upper()), None)
        if not sheet_name:
            raise ValueError(f'No PLAN sheet found in {path.name}.')
        ws = wb[sheet_name]
        rows_iter = ws.iter_rows(values_only=True)
        headers = next(rows_iter)
        rows = list(rows_iter)

    header_keys = [_clean_header(x) for x in headers]
    if replace:
        Order.objects.all().delete()
    created = 0
    for values in rows:
        if not any(v not in (None, '') for v in values):
            continue
        raw = {header_keys[i]: values[i] for i in range(min(len(header_keys), len(values)))}
        section = str(_pick(raw, 'Planning Mat', 'Section', 'ROLLING SECTION') or '').strip()
        grade = str(_pick(raw, 'Grade', 'Ext Grade', 'EXT. GRADE', 'R.M. GRADE') or '').strip()
        so_no = str(_pick(raw, 'SO No.', 'SO No') or '').strip()
        
        order_qty = _num(_pick(raw, 'Order Qty', 'ROLLING QTY', 'ROLLING    QTY', 'ORDER QTY', 'ORDER  QTY', 'ORDER    QTY'))
        pgi_qty = _num(_pick(raw, 'PGI Qty'))
        bal = _num(_pick(raw, 'Bal2PGI', 'ROLLING QTY', 'ROLLING    QTY', 'ORDER QTY', 'ORDER  QTY', 'ORDER    QTY'))
        
        if not section and not grade:
            continue
            
        if not so_no:
            if order_qty > 0:
                so_no = f"MOCK_{section}_{created + 1}"
            else:
                continue
                
        if bal <= 0:
            bal = max(0, order_qty - pgi_qty)
        if bal <= 0:
            bal = order_qty

        stock = sum(_num(_pick(raw, n)) for n in ['SO Stk(UR)', 'SO Stk(QI)'])
        net = max(0, bal - stock)
        if net == 0 and bal > 0:
            net = bal
            
        Order.objects.create(
            so_no=so_no,
            so_item=str(_pick(raw, 'SO Item') or '10').strip(),
            material=str(_pick(raw, 'Material') or section).strip(),
            section=section,
            grade=grade,
            quantity=order_qty,
            pgi_qty=pgi_qty,
            balance_qty=bal,
            eligible_stock=stock,
            net_demand=net,
            cutting_length=_num(_pick(raw, 'Cutting Length', 'BILLET LENGTH (MTR)', 'BLOOM LENGTH (MTR)', 'LENGTH')),
            section_weight=_num(_pick(raw, 'Sec Wt', 'SEC WT', 'SEC. WT.')),
            promise_date=str(_pick(raw, 'Promise Date') or '').strip(),
            required_delivery_date=str(_pick(raw, 'Req. Deli. Date') or '').strip(),
            priority='Priority' if _num(_pick(raw, 'Age')) >= 30 else 'Normal',
            status=str(_pick(raw, 'RM Rel Sta') or 'Pending').strip(),
            source_file=path.name,
            raw_data={k: ('' if v is None else str(v)) for k,v in raw.items()},
            is_selected=True,
        )
        created += 1
    return {'sheet': sheet_name, 'created': created, 'headers': list(headers)}


def build_plan_rows(order_ids=None):
    qs = Order.objects.filter(is_selected=True)
    if order_ids is not None:
        qs = qs.filter(id__in=order_ids)
    selected = list(qs)
    
    groups = defaultdict(list)
    for o in selected:
        groups[(o.section, o.section_weight or 40.3, o.grade)].append(o)
        
    normalized_reqs = []
    for (section, sec_wt, grade), orders in groups.items():
        total_net = sum(max(0, o.net_demand or o.balance_qty or o.quantity) for o in orders)
        if total_net <= 0:
            continue
        cut_len = max([o.cutting_length for o in orders if o.cutting_length > 0] or [10.0])
        
        sec_rule = SectionRule.objects.filter(section__iexact=section).first()
        min_l = sec_rule.min_length if sec_rule and sec_rule.min_length else 10.0
        max_l = sec_rule.max_length if sec_rule and sec_rule.max_length else 12.0
        
        n_req = NormalizedRequirement(
            planning_section=section,
            section_weight=Decimal(str(sec_wt)),
            finished_grade=grade,
            demand_quantity=Decimal(str(total_net)),
            min_length=Decimal(str(min_l)),
            max_length=Decimal(str(max_l)),
            cutting_length=Decimal(str(cut_len)),
            requirement_type='Firm'
        )
        normalized_reqs.append(n_req)
        
    if not normalized_reqs:
        return [], selected
        
    plan_rows = []
    for n_req in normalized_reqs:
        res = OptimizationEngine.optimize_requirement(n_req)
        b = res['best']
        plan_rows.append({
            'sheet': res['sheet'],
            'section': res['raw_section'],
            'grade': res['output_grade'],
            'qty': b['planned_finished_mt'],
            'demand_qty': float(n_req.demand_quantity),
            'length': b['cand_length'],
            'finished_piece_length': b['finished_piece_length'],
            'tolerance': res['tolerance'],
            'location': res['location'],
            'remarks': res['remarks'],
            'source_section': n_req.planning_section,
            'source_grade': n_req.finished_grade,
            'raw_grade': b['rm_grade'],
            'rm_qty': b['required_rm_mt'],
            'crop_loss_mt': b['crop_loss_mt'],
            'offcut_loss_mt': b['offcut_loss_mt'],
            'process_loss_mt': b['process_loss_mt'],
            'excess_mt': b['excess_mt'],
            'yield_pct': b['effective_yield_pct'],
            'rule_desc': res['rule_desc'],
            'opt_reason': res['opt_reason'],
            'sec_wt': float(n_req.section_weight),
            'req_type': n_req.requirement_type
        })
        
    req_type_order = {'Firm': 0, 'Additional': 1, 'Export': 2}
    
    # Advanced Sorting Hierarchy for Casting Plan
    plan_rows.sort(key=lambda r: (
        r['sheet'],                               # 1. RM Family Sheet
        req_type_order.get(r['req_type'], 9),     # 2. Block type order (Firm, Additional, Export)
        r['source_section'],                      # 3. Section Planned Material
        r['raw_grade'],                           # 4. Grade-wise (Same raw material chemistry)
        -r['sec_wt'],                             # 5. Thickness (Section Wt) in Decreasing order
        r['finished_piece_length'],               # 6. Cutting plan length (Ascending)
        r['length']                               # 7. Candidate casting length (Ascending fallback)
    ))
    
    return plan_rows, selected


def _copy_row_style(ws, source_row, target_row, max_col=7):
    for col in range(1, max_col+1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy(src.alignment)
        if src.fill:
            dst.fill = copy(src.fill)
        if src.font:
            dst.font = copy(src.font)
        if src.border:
            dst.border = copy(src.border)


def generate_casting_plan_workbook(plan_rows, plan_name=None):
    template = Path(settings.BASE_DIR) / 'data' / 'reference' / 'Casting plan.xlsx'
    wb = load_workbook(template)
    grouped = defaultdict(list)
    for r in plan_rows:
        grouped[r['sheet']].append(r)

    for sheet_name in ['SBB','250sq','165SQ']:
        ws = wb[sheet_name]
        for merged_range in list(ws.merged_cells.ranges):
            try:
                ws.unmerge_cells(str(merged_range))
            except Exception:
                try:
                    ws.merged_cells.remove(merged_range)
                except Exception:
                    pass

        for row in range(1, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row, col)
                if type(cell).__name__ != 'MergedCell':
                    try:
                        cell.value = None
                    except AttributeError:
                        pass
        current = 1
        rows = grouped.get(sheet_name, [])
        by_source = defaultdict(list)
        for r in rows:
            by_source[(r['source_section'], r.get('req_type','Firm'))].append(r)
            
        req_type_order = {'Firm': 0, 'Additional': 1, 'Export': 2}
        sorted_groups = sorted(by_source.items(), key=lambda x: (x[0][0], req_type_order.get(x[0][1], 9)))
        
        for (source_section, req_type), items in sorted_groups:
            _copy_row_style(ws, 1, current, 7)
            if req_type == 'Firm':
                title = f'Casting plan for {source_section}'
            elif req_type == 'Additional':
                title = f'Additional Casting plan for {source_section}'
            elif req_type == 'Export':
                title = f'Casting plan {source_section}(Export)'
            else:
                title = f'{req_type} Casting plan for {source_section}'
                
            ws.cell(current, 1).value = title
            try:
                ws.merge_cells(start_row=current, start_column=1, end_row=current, end_column=7)
            except Exception:
                pass
            current += 1
            _copy_row_style(ws, 2, current, 7)
            headers = ['Section','Grade','Qty(MT)','Length','Tolerance','Location','Remarks']
            for i,h in enumerate(headers,1): ws.cell(current,i).value=h
            current += 1
            total = 0
            for item in items:
                _copy_row_style(ws, 3, current, 7)
                vals = [item['section'], item['grade'], item['qty'], f"{item['length']:.2f} Mtr", item['tolerance'], item['location'], item['remarks']]
                for i,v in enumerate(vals,1): ws.cell(current,i).value=v
                total += item['qty']; current += 1
            _copy_row_style(ws, 9 if sheet_name=='SBB' else 5, current, 7)
            ws.cell(current,1).value='Total Qty(MT)'; ws.cell(current,3).value=round(total,3)
            try:
                ws.merge_cells(start_row=current, start_column=1, end_row=current, end_column=2)
            except Exception:
                pass
            current += 2
            
        if current <= ws.max_row:
            ws.delete_rows(current, ws.max_row - current + 1)
            
        if not rows:
            ws.cell(1,1).value='No generated casting plan lines for this RM family'

    if "OPTIMIZATION SUMMARY" in wb.sheetnames:
        del wb["OPTIMIZATION SUMMARY"]
    ws_sum = wb.create_sheet(title="OPTIMIZATION SUMMARY")
    
    summary_headers = [
        "Planning Section", "Section Weight (kg/m)", "Finished Grade", "Demand Qty (MT)",
        "RM Family", "RM Section", "RM Grade", "Selected Casting Length (m)",
        "Finished Piece Length (m)", "Planned Qty (MT)", "Required RM Qty (MT)",
        "Crop Loss (MT)", "Offcut Loss (MT)", "Process Loss (MT)", "Excess Qty (MT)",
        "Effective Yield %", "Yield Rule Used", "Optimization Reason"
    ]
    
    ws_sum.append(summary_headers)
        
    for r in plan_rows:
        ws_sum.append([
            r.get('source_section', ''),
            r.get('sec_wt', 0.0),
            r.get('source_grade', ''),
            r.get('demand_qty', 0.0),
            r.get('sheet', ''),
            r.get('section', ''),
            r.get('grade', ''),
            r.get('length', 0.0),
            r.get('finished_piece_length', 0.0),
            r.get('qty', 0.0),
            r.get('rm_qty', 0.0),
            r.get('crop_loss_mt', 0.0),
            r.get('offcut_loss_mt', 0.0),
            r.get('process_loss_mt', 0.0),
            r.get('excess_mt', 0.0),
            r.get('yield_pct', 0.0),
            r.get('rule_desc', ''),
            r.get('opt_reason', '')
        ])

    out_dir = Path(settings.BASE_DIR) / 'media' / 'generated'
    out_dir.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r'[^A-Za-z0-9_-]+','_', plan_name or f'Casting_Plan_{date.today().isoformat()}')
    filename = f'{safe}_{uuid.uuid4().hex[:8]}.xlsx'
    out_path = out_dir / filename
    wb.save(out_path)
    return out_path
