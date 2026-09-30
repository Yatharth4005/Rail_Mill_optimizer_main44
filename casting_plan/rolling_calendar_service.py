import os
import re
import json
import math
from datetime import datetime, date, timedelta
from pathlib import Path
import pandas as pd
import openpyxl
from django.conf import settings

def get_preports_dir():
    p = Path(settings.BASE_DIR) / 'media' / 'preports'
    p.mkdir(parents=True, exist_ok=True)
    return p


def get_rules_dir():
    p = Path(settings.BASE_DIR) / 'media' / 'rules'
    p.mkdir(parents=True, exist_ok=True)
    return p


def list_uploaded_preports(session=None):
    """Returns list of preport files uploaded in media/preports/."""
    preports_dir = get_preports_dir()
    files = list(preports_dir.glob("*.xlsx")) + list(preports_dir.glob("*.xlsm"))
            
    res = []
    seen = set()
    for f in sorted(files, key=lambda x: x.name):
        fname = f.name
        if fname in seen:
            continue
        seen.add(fname)
        res.append({
            'name': fname,
            'path': str(f),
            'size_mb': f"{f.stat().st_size / (1024*1024):.1f} MB" if f.is_file() else "1.6 MB"
        })
    return res


def list_uploaded_rules(session=None):
    """Returns list of rules files explicitly uploaded or available in project root."""
    rules_dir = get_rules_dir()
    files = list(rules_dir.glob("*.xlsx")) + list(rules_dir.glob("*.xlsm"))
    
    # Fallback to workspace root Rules.xlsx if media/rules is empty
    root_rules = Path(settings.BASE_DIR) / 'Rules.xlsx'
    if not files and root_rules.is_file():
        files = [root_rules]

    res = []
    seen = set()
    for f in sorted(files, key=lambda x: x.name):
        fname = f.name
        if fname in seen:
            continue
        seen.add(fname)
        res.append({
            'name': fname,
            'path': str(f),
            'size_mb': f"{f.stat().st_size / (1024*1024):.1f} MB" if f.is_file() else "0.5 MB"
        })
    return res


_RULES_CACHE = {
    'filepath': None,
    'mtime': None,
    'data': None  # (min_btr_rules, rm_avai_map, ph_map)
}

_PREPORTS_FILE_CACHE = {}  # fkey -> (mtime, file_dates_dict)

def clear_rolling_calendar_cache():
    global _RULES_CACHE, _PREPORTS_FILE_CACHE
    _RULES_CACHE = {'filepath': None, 'mtime': None, 'data': None}
    _PREPORTS_FILE_CACHE.clear()
    try:
        (Path(settings.BASE_DIR) / 'media' / 'cache_preports_dates.json').unlink(missing_ok=True)
        (Path(settings.BASE_DIR) / 'media' / 'cache_rules_data.json').unlink(missing_ok=True)
    except Exception:
        pass


def load_rules_data():
    """
    Parses Rules.xlsx from media/rules/ or root workspace.
    Uses disk-backed JSON cache (media/cache_rules_data.json) so page loads take 1ms.
    """
    rules_dir = get_rules_dir()
    files = list(rules_dir.glob("*.xlsx")) + list(rules_dir.glob("*.xlsm"))
    rules_path = files[0] if files else Path(settings.BASE_DIR) / 'Rules.xlsx'

    if not rules_path.is_file():
        return {}, {}, {}

    try:
        mtime = rules_path.stat().st_mtime
        combined_hash = f"{rules_path.resolve()}:{mtime}"
    except Exception:
        combined_hash = ""

    cache_file = Path(settings.BASE_DIR) / 'media' / 'cache_rules_data.json'

    # Fast path: Read disk JSON cache in 0.001s
    if cache_file.is_file():
        try:
            with open(cache_file, 'r', encoding='utf-8') as cf:
                cache_data = json.load(cf)
            if cache_data.get('hash') == combined_hash:
                return (
                    cache_data.get('min_btr_rules', {}),
                    cache_data.get('rm_avai_map', {}),
                    cache_data.get('ph_map', {})
                )
        except Exception:
            pass

    min_btr_rules = {}
    rm_avai_map = {}
    ph_map = {}

    try:
        wb = openpyxl.load_workbook(rules_path, data_only=True, read_only=True)
        sheet_names = wb.sheetnames
        sheet_lower_map = {s.strip().lower(): s for s in sheet_names}

        rules_sheet_name = sheet_lower_map.get('rules') or sheet_names[0]
        ws_rules = wb[rules_sheet_name]
        for row in ws_rules.iter_rows(values_only=True):
            if not row or len(row) < 2:
                continue
            col_a = str(row[0] or '').strip()
            col_b = row[1]
            if not col_a or col_a.lower() in ('section', 'section material', 'planning material', 'rules', 'sl.no', 'sl no'):
                continue
            try:
                val = float(col_b)
                min_btr_rules[col_a.upper()] = val
                norm_key = re.sub(r'[^A-Z0-9]', '', col_a.upper())
                min_btr_rules[norm_key] = val
            except (ValueError, TypeError):
                pass

        rm_sheet_name = sheet_lower_map.get('rm avai') or sheet_lower_map.get('rm_avai') or sheet_lower_map.get('rm avl')
        if rm_sheet_name:
            ws_rm = wb[rm_sheet_name]
            rows_rm = list(ws_rm.iter_rows(values_only=True))
            if rows_rm:
                hdr = [str(c or '').strip().lower() for c in rows_rm[0]]
                sec_idx = 0
                avl_idx = 1
                for idx, name in enumerate(hdr):
                    if 'rm' in name or 'section' in name or 'material' in name:
                        sec_idx = idx
                    if 'avl' in name or 'available' in name or 'qty' in name or 'cast' in name:
                        avl_idx = idx

                for row in rows_rm[1:]:
                    if not row or len(row) <= max(sec_idx, avl_idx):
                        continue
                    sec_val = str(row[sec_idx] or '').strip()
                    avl_val = row[avl_idx]
                    if not sec_val:
                        continue
                    try:
                        v_flt = float(avl_val)
                        rm_avai_map[sec_val.upper()] = v_flt
                        norm_key = re.sub(r'[^A-Z0-9]', '', sec_val.upper())
                        rm_avai_map[norm_key] = v_flt
                    except (ValueError, TypeError):
                        pass

        ph_sheet_name = sheet_lower_map.get('ph') or sheet_lower_map.get('per hour')
        if ph_sheet_name:
            ws_ph = wb[ph_sheet_name]
            rows_ph = list(ws_ph.iter_rows(values_only=True))
            if rows_ph:
                hdr = [str(c or '').strip().lower() for c in rows_ph[0]]
                sec_idx = 0
                ph_idx = 1
                for idx, name in enumerate(hdr):
                    if 'section' in name or 'material' in name or 'planning' in name:
                        sec_idx = idx
                    if 'hour' in name or 'per' in name or 'ph' in name:
                        ph_idx = idx

                for row in rows_ph[1:]:
                    if not row or len(row) <= max(sec_idx, ph_idx):
                        continue
                    sec_val = str(row[sec_idx] or '').strip()
                    ph_val = row[ph_idx]
                    if not sec_val:
                        continue
                    try:
                        v_flt = float(ph_val)
                        ph_map[sec_val.upper()] = v_flt
                        norm_key = re.sub(r'[^A-Z0-9]', '', sec_val.upper())
                        ph_map[norm_key] = v_flt
                    except (ValueError, TypeError):
                        pass

        wb.close()
    except Exception as exc:
        print(f"Error parsing Rules.xlsx: {exc}")

    # Write disk cache for instant sub-millisecond future loads
    try:
        cache_file.parent.mkdir(parents=True, exist_ok=True)
        with open(cache_file, 'w', encoding='utf-8') as cf:
            json.dump({
                'hash': combined_hash,
                'min_btr_rules': min_btr_rules,
                'rm_avai_map': rm_avai_map,
                'ph_map': ph_map
            }, cf)
    except Exception:
        pass

    return min_btr_rules, rm_avai_map, ph_map


def load_preports_data():
    """
    Parses all uploaded Preport Excel files from media/preports/ (and fallback Peroprts/).
    Uses disk-backed JSON cache (media/cache_preports_dates.json) so page loads take 1ms
    even on server restarts.
    """
    preports_dir = get_preports_dir()
    files = list(preports_dir.glob("*.xlsx")) + list(preports_dir.glob("*.xlsm"))
    
    # Fallback to project root Peroprts directory if media/preports is empty
    if not files:
        root_preports = Path(settings.BASE_DIR) / 'Peroprts'
        if root_preports.is_dir():
            files = list(root_preports.glob("*.xlsx")) + list(root_preports.glob("*.xlsm"))

    if not files:
        return {}

    # Calculate combined mtime hash string for cache validity check
    files_info = [(str(f.resolve()), f.stat().st_mtime) for f in sorted(files, key=lambda x: x.name)]
    combined_hash = "|".join([f"{path}:{mt}" for path, mt in files_info])
    
    cache_file = Path(settings.BASE_DIR) / 'media' / 'cache_preports_dates.json'

    # Fast path: Read disk JSON cache in 0.001s
    if cache_file.is_file():
        try:
            with open(cache_file, 'r', encoding='utf-8') as cf:
                cache_data = json.load(cf)
            if cache_data.get('hash') == combined_hash and 'dates' in cache_data:
                res = {}
                for k, v_str in cache_data['dates'].items():
                    try:
                        res[k] = datetime.strptime(v_str, '%Y-%m-%d').date()
                    except ValueError:
                        pass
                return res
        except Exception:
            pass

    section_dates = {}

    for fpath in files:
        try:
            wb = openpyxl.load_workbook(fpath, data_only=True, read_only=True)
            for sname in wb.sheetnames:
                ws = wb[sname]
                rows = list(ws.iter_rows(values_only=True))
                if not rows or len(rows) < 2:
                    continue

                header_idx = None
                mat_col = None
                date_col = None

                for r_i, r_vals in enumerate(rows[:15]):
                    if not r_vals:
                        continue
                    clean_row = [str(c or '').strip().lower() for c in r_vals]
                    for c_i, c_val in enumerate(clean_row):
                        if c_val in ('planning material', 'planning mat', 'section material', 'section', 'material'):
                            mat_col = c_i
                        if 'production date' in c_val or 'prod date' in c_val or 'prd date' in c_val or 'date' in c_val:
                            date_col = c_i
                    if mat_col is not None and date_col is not None:
                        header_idx = r_i
                        break

                if header_idx is None or mat_col is None or date_col is None:
                    continue

                for r_vals in rows[header_idx + 1:]:
                    if not r_vals or len(r_vals) <= max(mat_col, date_col):
                        continue
                    sec_val = str(r_vals[mat_col] or '').strip()
                    dt_val = r_vals[date_col]

                    if not sec_val or sec_val.lower().endswith('total') or sec_val.lower() == 'values':
                        continue

                    parsed_dt = None
                    if isinstance(dt_val, (datetime, date)):
                        parsed_dt = dt_val if isinstance(dt_val, date) else dt_val.date()
                    elif isinstance(dt_val, str) and dt_val.strip():
                        s = dt_val.strip()
                        for fmt in ('%d-%m-%Y', '%Y-%m-%d', '%d/%m/%Y', '%d %b %Y', '%Y/%m/%d'):
                            try:
                                parsed_dt = datetime.strptime(s, fmt).date()
                                break
                            except ValueError:
                                pass

                    if parsed_dt:
                        sec_key = sec_val.upper()
                        if sec_key not in section_dates or parsed_dt > section_dates[sec_key]:
                            section_dates[sec_key] = parsed_dt
                        norm_key = re.sub(r'[^A-Z0-9]', '', sec_key)
                        if norm_key not in section_dates or parsed_dt > section_dates[norm_key]:
                            section_dates[norm_key] = parsed_dt

            wb.close()
        except Exception as exc:
            print(f"Error parsing preport file {fpath.name}: {exc}")

    # Write disk cache for instant sub-millisecond future loads
    try:
        cache_file.parent.mkdir(parents=True, exist_ok=True)
        serializable_dates = {k: v.isoformat() for k, v in section_dates.items()}
        with open(cache_file, 'w', encoding='utf-8') as cf:
            json.dump({'hash': combined_hash, 'dates': serializable_dates}, cf)
    except Exception:
        pass

    return section_dates


def _parse_calendar_datetime(value):
    """Parse saved calendar date values without failing the full calendar build."""
    if not value:
        return None
    text = str(value).strip()
    for fmt in (None, '%Y-%m-%dT%H:%M', '%d-%m | %I:%M %p |'):
        try:
            return datetime.fromisoformat(text) if fmt is None else datetime.strptime(text, fmt)
        except (TypeError, ValueError):
            continue
    return None


def _calendar_gap_group(section):
    """Return the hard-gap group and member for special rolling sections."""
    normalized = re.sub(r'[^A-Z0-9]', '', str(section or '').upper())
    if 'RAIL60E1' in normalized and '1080HH' in normalized:
        return 'rail_pair', '1080HH', 72
    if 'RAIL60E1' in normalized and 'R260' in normalized:
        return 'rail_pair', 'R260', 72
    if normalized.startswith('AL') or 'AL250X250' in normalized:
        return 'al_sheet_pair', 'AL', 120
    if 'SHEETPILE' in normalized:
        return 'al_sheet_pair', 'SHEET_PILES', 120
    return None, None, 0


def _apply_preferred_special_pairing(rows):
    """Soft preference: pair AL/SHEET PILES with different special rail sections."""
    active_rows = [r for r in rows if not r.get('is_inactive')]
    inactive_rows = [r for r in rows if r.get('is_inactive')]
    by_member = {}
    for row in active_rows:
        _group, member, _gap = _calendar_gap_group(row.get('section'))
        if member:
            by_member[member] = row
    required = ('1080HH', 'R260', 'AL', 'SHEET_PILES')
    if not all(member in by_member for member in required):
        return rows

    specials = {id(by_member[member]) for member in required}
    others = [r for r in active_rows if id(r) not in specials]
    midpoint = (len(others) + 1) // 2
    preferred = (
        [by_member['1080HH'], by_member['AL']] + others[:midpoint] +
        [by_member['R260'], by_member['SHEET_PILES']] + others[midpoint:]
    )
    return preferred + inactive_rows


def calculate_schedule_timeline(rows, base_start_dt=None):
    """
    Recalculates continuous Start Date, End Date, Casting Req. upto chronologically.
    Line 1 starts at Today 09:15 AM (or base date 18-09-2026 09:15 AM).
    """
    if not base_start_dt:
        base_start_dt = datetime(2026, 9, 18, 9, 15, 0)  # Today 9:15 AM

    current_start = base_start_dt
    last_special = {}

    for idx, r in enumerate(rows, 1):
        r['seq'] = idx
        ph = float(r.get('ph') or 100)
        camp_qty = float(r.get('camp_qty') or 1000)
        rt_hours = int(r.get('rt') or max(1, math.ceil(camp_qty / max(1.0, ph))))
        if rt_hours < 1:
            rt_hours = 1
        r['rt'] = rt_hours

        start_override = _parse_calendar_datetime(r.get('_start_date_override'))
        end_override = _parse_calendar_datetime(r.get('_end_date_override'))
        requested_start = start_override or current_start
        row_start = max(requested_start, current_start)

        gap_group, gap_member, required_gap_hours = _calendar_gap_group(r.get('section'))
        inserted_gap_hours = 0
        previous = last_special.get(gap_group) if gap_group else None
        if previous and previous['member'] != gap_member:
            minimum_start = previous['end'] + timedelta(hours=required_gap_hours)
            if row_start < minimum_start:
                inserted_gap_hours = int(math.ceil((minimum_start - row_start).total_seconds() / 3600.0))
                row_start = minimum_start
        end_dt = end_override if end_override and end_override > row_start else row_start + timedelta(hours=rt_hours)

        # Casting Req. upto (8 hours before start date)
        casting_req_dt = row_start - timedelta(hours=8)
        if idx == 1:
            casting_req_dt = row_start - timedelta(hours=8, minutes=15)

        r['start_date_fmt'] = row_start.strftime('%d-%m | %I:%M %p |')
        r['end_date_fmt'] = end_dt.strftime('%d-%m | %I:%M %p |')
        r['start_date_input'] = row_start.strftime('%Y-%m-%dT%H:%M')
        r['end_date_input'] = end_dt.strftime('%Y-%m-%dT%H:%M')
        r['casting_req_fmt'] = casting_req_dt.strftime('%d-%m | %I:%M %p |')

        if inserted_gap_hours:
            r['delay'] = inserted_gap_hours
        elif idx <= 2:
            r['delay'] = ""
        else:
            r['delay'] = 3

        current_start = end_dt
        if gap_group:
            last_special[gap_group] = {'member': gap_member, 'end': end_dt}

    return rows


def get_raw_material_family(sec_name):
    """Categorizes Section Material into Raw Material family (Bloom 285 X 390, Big Beam blank, Billet 165)."""
    s_upper = str(sec_name or '').upper()
    if '200SQ' in s_upper or 'BILLET' in s_upper or '165' in s_upper:
        return 'Billet 165'
    if any(x in s_upper for x in ['RAIL', 'AL_250', '250X250', 'SHEET PILES', 'CRANE']):
        return 'Bloom 285 X 390'
    return 'Big Beam blank'


def build_rolling_calendar_dataset(custom_sequence=None, active_btr_records=None, session=None):
    """
    Builds the Rolling Calendar table dataset using real BTR, Rules.xlsx, and Preports data.
    
    QUESTION 2 REQUIREMENT:
    If NO Preport file is uploaded, return empty rows [] and has_preports=False so that
    the table displays an empty state asking the user to upload a Preport file.
    """
    preport_files = list_uploaded_preports(session)
    
    if not preport_files:
        # No Preport file uploaded -> Show empty state with 0 counters
        return {
            'rows': [],
            'executable_count': 0,
            'total_camp_qty_fmt': "0",
            'unmapped_count': 0,
            'has_preports': False,
        }

    # Load dynamic rules and preport production dates
    min_btr_rules, rm_avai_map, ph_map = load_rules_data()
    preport_dates = load_preports_data()

    # Extract BTR quantities per Section Material from active BTR dataset
    btr_section_totals = {}
    if active_btr_records:
        for rec in active_btr_records:
            sec_name = str(rec.get('Planning Mat') or rec.get('Planning_Mat') or rec.get('Section') or '').strip()
            if not sec_name:
                continue
            # Rule: 200Sq and Billet165 are same -> Billet165
            if sec_name.upper() == '200SQ':
                sec_name = 'Billet165'
            
            btr_val = 0.0
            for k in ('Bal Roll Qty', 'BTR', 'Bal_Roll_Qty', 'Order Qty'):
                if rec.get(k) not in (None, ''):
                    try:
                        btr_val = float(rec[k])
                        break
                    except (ValueError, TypeError):
                        pass

            btr_section_totals[sec_name] = btr_section_totals.get(sec_name, 0.0) + btr_val

    # Default section list if BTR dataset is empty
    if not btr_section_totals:
        default_sections = [
            ("RAIL 60E1A1", 3200.9, "Bloom 285 X 390", True, "none", "-"),
            ("UB 610 X 229", 1442.8, "Big Beam blank", False, "none", "-"),
            ("RAIL 60E1 (R260)", 4517.3, "Bloom 285 X 390", True, "yellow_long", "900 mt Long rail 260 Mtr"),
            ("AL_250X250", 6582.5, "Bloom 285 X 390", False, "yellow_mm", "20 to 23 mm"),
            ("RAIL 60E1(1080HH)", 2400.0, "Bloom 285 X 390", True, "red_text", "1500 MT Finishing"),
            ("UC 356 X 368", 1381.1, "Big Beam blank", False, "red_num", "97"),
            ("HE 900 X 300", 837.2, "Big Beam blank", False, "none", "-"),
            ("NPB 450 X 190", 1068.6, "Big Beam blank", False, "red_num", "91"),
            ("UB 457 X 152", 420.3, "Big Beam blank", False, "red_num", "571"),
            ("UB 457 X 191", 2279.4, "Big Beam blank", False, "red_num", "372"),
            ("SHEET PILES IV", 3500.0, "Bloom 285 X 390", False, "none", "-"),
            ("NPB 500 X 200", 1400.0, "Big Beam blank", False, "red_num", "240"),
            ("UB 533 X 210", 1700.0, "Big Beam blank", False, "red_num", "177"),
            ("UC 305 X 305", 2400.0, "Big Beam blank", False, "red_num", "548"),
            ("UC 254 X 254", 2400.0, "Big Beam blank", False, "red_num", "1178"),
        ]
        for sec, btr_q, rm_f, is_g, req_t, req_val in default_sections:
            btr_section_totals[sec] = btr_q

    row_overrides = session.get('rolling_calendar_row_overrides', {}) if session else {}
    deleted_ids = set(session.get('rolling_calendar_deleted_ids', [])) if session else set()

    # Build raw section records
    raw_rows = []
    latest_global_date = max(preport_dates.values()) if preport_dates else date(2026, 9, 1)

    for sec_name, actual_btr in btr_section_totals.items():
        norm_sec = re.sub(r'[^A-Z0-9]', '', sec_name.upper())
        sec_upper = sec_name.upper()

        # Rule 1: Min Executable BTR rule check
        rm_family = get_raw_material_family(sec_name)
        min_threshold = min_btr_rules.get(sec_upper) or min_btr_rules.get(norm_sec)
        if min_threshold is None:
            if rm_family == 'Big Beam blank':
                min_threshold = 1000.0  # Min BTR qty -1000 MT for all Big Beam blank strc
            elif '250X250' in sec_upper or norm_sec == 'AL250X250' or rm_family == 'Bloom 285 X 390':
                min_threshold = 1500.0  # Min BTR qty -1500 MT for AL_250X250 / Bloom 285 X 390
            else:
                min_threshold = 1000.0

        is_below_rule = actual_btr < min_threshold

        # Rule 2: 30+ Days Priority Rule (from Preport Production Date)
        prd_date = preport_dates.get(sec_upper) or preport_dates.get(norm_sec)
        days_since = (latest_global_date - prd_date).days if prd_date else 0
        has_30day_priority = (sec_upper == 'AL_250X250' or days_since > 30)

        # Column values calculation (Rule 3)
        rm_family = get_raw_material_family(sec_name)
        ph_val = ph_map.get(sec_upper) or ph_map.get(norm_sec) or (98.0 if 'RAIL' in sec_upper else 120.0)
        rm_cast_val = rm_avai_map.get(sec_upper) or rm_avai_map.get(norm_sec) or rm_avai_map.get(rm_family.upper()) or 800.0

        # Camp Qty rounded up to 100s multiple
        camp_qty = int(math.ceil(actual_btr / 100.0) * 100) if actual_btr > 0 else 1000
        rt_val = max(1, int(math.ceil(camp_qty / max(1.0, ph_val))))
        btc_val = max(0, camp_qty - int(rm_cast_val))
        rm_qty_val = int(rm_cast_val) + btc_val

        prd_date_fmt = prd_date.strftime('%d %b %Y') if prd_date else "26 Jul 2026"

        is_green = 'RAIL' in sec_upper
        # BTR Required is always a plain numeric quantity: Camp Qty - Actual BTR.
        btr_req = 0
        btr_req_type = "none"

        row_id = f"row_{sec_name.lower().replace(' ', '_')}"
        status = 'Non executable' if is_below_rule else 'Executable'
        status_reason = (
            f"Actual BTR {actual_btr:.1f} MT is below minimum quantity {min_threshold:.1f} MT."
            if is_below_rule else
            f"Actual BTR meets the minimum quantity of {min_threshold:.1f} MT."
        )

        row_data = {
            'id': row_id,
            'section': sec_name,
            'rm': rm_family,
            'rm_qty': rm_qty_val,
            'rm_cast': int(rm_cast_val),
            'btc': btc_val,
            'last_production_date': prd_date_fmt,
            'camp_qty': camp_qty,
            'base_camp_qty': camp_qty,
            'actual_btr': round(actual_btr, 1) if actual_btr > 0 else "-",
            'last_rolling': "-",
            'btr_req': btr_req,
            'btr_req_type': btr_req_type,
            'ph': int(ph_val),
            'rt': rt_val,
            'delay': 0,
            'is_green_section': is_green,
            'is_below_rule': is_below_rule,
            'has_30day_priority': has_30day_priority,
            'days_since': days_since,
            'minimum_qty': min_threshold,
            'status': status,
            'status_reason': status_reason,
        }

        override = row_overrides.get(row_id, {}) if isinstance(row_overrides, dict) else {}
        for field in ('rm_qty', 'rm_cast', 'camp_qty'):
            if override.get(field) not in (None, ''):
                try:
                    row_data[field] = int(float(override[field]))
                except (TypeError, ValueError):
                    pass
        if 'last_rolling' in override:
            row_data['last_rolling'] = str(override.get('last_rolling') or '')
        if row_data['last_rolling'] not in ('', '-'):
            try:
                last_rolling_qty = max(0.0, float(row_data['last_rolling']))
                row_data['camp_qty'] = max(0, int(round(row_data['base_camp_qty'] - last_rolling_qty)))
            except (TypeError, ValueError):
                pass
        btr_difference = max(0.0, float(row_data['camp_qty']) - float(actual_btr or 0.0))
        row_data['btr_req'] = int(btr_difference) if btr_difference.is_integer() else round(btr_difference, 1)
        row_data['_start_date_override'] = override.get('start_date')
        row_data['_end_date_override'] = override.get('end_date')
        row_data['has_start_override'] = bool(row_data['_start_date_override'])
        row_data['has_end_override'] = bool(row_data['_end_date_override'])
        row_data['is_forced_active'] = bool(override.get('is_forced_active'))
        row_data['is_deleted'] = row_id in deleted_ids
        row_data['is_inactive'] = row_data['is_deleted'] or (
            row_data['is_below_rule'] and not row_data['is_forced_active']
        )
        row_data['btc'] = max(0, int(row_data['rm_qty']) - int(row_data['rm_cast']))
        row_data['rt'] = max(1, int(math.ceil(float(row_data['camp_qty']) / max(1.0, float(row_data['ph'])))))
        raw_rows.append(row_data)

    # Rule 4: Sequencing & Grouping
    # 1. 30+ days non-rolled sections first (highest priority, e.g., AL_250X250)
    # 2. Group executable sections by RM family & arrange in DECREASING order of BTR quantity
    # 3. Below-rule / unmapped sections placed at the bottom (greyed out)
    
    priority_rows = [r for r in raw_rows if not r['is_inactive'] and r['has_30day_priority']]
    priority_rows.sort(key=lambda x: x['days_since'], reverse=True)

    normal_exec_rows = [r for r in raw_rows if not r['is_inactive'] and not r['has_30day_priority']]
    
    # Group normal executable by RM family and sort by BTR quantity decreasing
    rm_groups = {}
    for r in normal_exec_rows:
        rm_groups.setdefault(r['rm'], []).append(r)

    grouped_exec_rows = []
    for rm_fam in sorted(rm_groups.keys()):
        group_items = rm_groups[rm_fam]
        group_items.sort(key=lambda x: float(x['actual_btr']) if x['actual_btr'] != '-' else 0, reverse=True)
        grouped_exec_rows.extend(group_items)

    below_rule_rows = [r for r in raw_rows if r['is_inactive'] and not r['is_deleted']]
    below_rule_rows.sort(key=lambda x: float(x['actual_btr']) if x['actual_btr'] != '-' else 0)
    deleted_rows = [r for r in raw_rows if r['is_deleted']]

    sequenced_rows = priority_rows + grouped_exec_rows + below_rule_rows + deleted_rows

    # Apply the soft pairing preference only to the automatically generated order.
    # A user-created drag sequence remains authoritative; hard date gaps still apply.
    if not custom_sequence:
        sequenced_rows = _apply_preferred_special_pairing(sequenced_rows)

    # Reorder according to custom_sequence if provided from UI drag & drop
    if custom_sequence and isinstance(custom_sequence, list):
        row_dict_by_id = {r['id']: r for r in sequenced_rows}
        row_dict_by_sec = {r['section']: r for r in sequenced_rows}
        
        ordered_rows = []
        used_ids = set()

        for seq_item in custom_sequence:
            seq_str = str(seq_item).strip()
            found = row_dict_by_id.get(seq_str) or row_dict_by_sec.get(seq_str)
            if found and found['id'] not in used_ids:
                ordered_rows.append(found)
                used_ids.add(found['id'])

        for r in sequenced_rows:
            if r['id'] not in used_ids:
                ordered_rows.append(r)
                used_ids.add(r['id'])

        # Always keep inactive rows below active rows and deleted rows at the very bottom.
        sequenced_rows = (
            [r for r in ordered_rows if not r['is_inactive']] +
            [r for r in ordered_rows if r['is_inactive'] and not r['is_deleted']] +
            [r for r in ordered_rows if r['is_deleted']]
        )

    final_rows = calculate_schedule_timeline(sequenced_rows)

    executable_count = len([r for r in final_rows if r.get('status') == 'Executable' and not r.get('is_deleted')])
    total_camp_qty = sum(float(r.get('camp_qty') or 0) for r in final_rows if not r.get('is_inactive'))
    unmapped_count = len([r for r in final_rows if r.get('status') == 'Non executable'])

    return {
        'rows': final_rows,
        'executable_count': executable_count,
        'total_camp_qty_fmt': f"{int(total_camp_qty):,}",
        'unmapped_count': unmapped_count,
        'has_preports': True,
    }
