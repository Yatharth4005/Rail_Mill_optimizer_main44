from datetime import date
from decimal import Decimal
from pathlib import Path
from uuid import uuid4

from django.contrib import messages
from django.contrib.auth import authenticate, login as auth_login, logout as auth_logout
from django.contrib.auth.decorators import login_required
from django.conf import settings
from django.http import JsonResponse, FileResponse, Http404
from django.urls import reverse
from django.shortcuts import render, redirect, get_object_or_404
from django.utils import timezone
from django.utils.text import get_valid_filename
from django.views.decorators.http import require_POST, require_GET

from .models import (
    Order, SectionRule, GradeMap, BatchRule, PlanLine, Notification,
    GeneratedPlan, SectionWeightMap, YieldRule, CastingPlanRequest, CastingPlanRequirement,
    BTRCampaign, BilletLengthSetting
)
from .services import (
    optimize, import_demand_workbook, build_plan_rows,
    generate_casting_plan_workbook, NormalizedRequirement, OptimizationEngine,
    raw_family
)
from .btr_services import inspect_btr_workbook, generate_campaign_workbook
from .btr_service import get_sap_assignment_info, update_sap_assignment_info, get_all_sap_assignments, get_grade_group_for_record
from .rolling_plan_services import generate_rolling_plan_workbook


@login_required
def dashboard(request):
    ctx = {
        'orders': Order.objects.count(),
        'sections': SectionRule.objects.count(),
        'grades': GradeMap.objects.count(),
        'batches': BatchRule.objects.count(),
        'plans': (
            GeneratedPlan.objects.count()
            + CastingPlanRequest.objects.filter(status='Completed').count()
            + BTRCampaign.objects.count()
        ),
        'selected_demand': sum(o.net_demand for o in Order.objects.filter(is_selected=True)),
    }
    return render(request, 'casting_plan/dashboard.html', ctx)


@login_required
def optimizer(request):
    total_orders = Order.objects.count()
    eligible_orders = Order.objects.filter(is_selected=True).count()
    selected_orders = Order.objects.filter(is_selected=True).count()
    pending_qty = sum(o.net_demand for o in Order.objects.filter(is_selected=True))
    
    first_order = Order.objects.first()
    source_file = first_order.source_file if first_order else None
    
    sections = list(SectionRule.objects.values_list('section', flat=True).order_by('section'))
    if not sections:
        sections = list(Order.objects.values_list('section', flat=True).distinct().order_by('section'))
        
    grades = list(GradeMap.objects.values_list('ext_grade', flat=True).distinct().order_by('ext_grade'))
    if not grades:
        grades = list(Order.objects.values_list('grade', flat=True).distinct().order_by('grade'))

    return render(request, 'casting_plan/optimizer.html', {
        'total_orders': total_orders,
        'eligible_orders': eligible_orders,
        'selected_orders': selected_orders,
        'pending_qty': pending_qty,
        'source_file': source_file,
        'sections': sections,
        'grades': grades,
    })


@login_required
def plan(request):
    plans = GeneratedPlan.objects.all()[:30]
    requests = CastingPlanRequest.objects.all()[:30]
    return render(request, 'casting_plan/plan.html', {'plans': plans, 'requests': requests})


@login_required
def create_casting_plan_view(request):
    """
    Step 4: Casting Plan Ready View (Auto-generated from Camp Plan output).
    Displays current planning material, active Camp Plan summary, auto-generation card, and history/downloads table.
    """
    from .btr_service import load_btr_dataset, generate_camp_plan_records, get_casting_plan_history
    
    upload_id = request.session.get('btr_upload_id')
    metadata, records = load_btr_dataset(upload_id) if upload_id else (None, None)
    
    selected_mat = request.GET.get('mat') or request.session.get('btr_selected_material') or 'UC_356_X_406'
    request.session['btr_selected_material'] = selected_mat
    
    camp_records = []
    if records and selected_mat:
        res = generate_camp_plan_records(records, selected_mat, upload_id=upload_id)
        camp_records = res.get('valid_records', [])

    history_items = get_casting_plan_history()
    today_str = date.today().isoformat()

    return render(request, 'casting_plan/create_casting_plan.html', {
        'planning_material': selected_mat,
        'selected_mat': selected_mat,
        'camp_records_count': len(camp_records),
        'camp_records': camp_records[:50],
        'today': today_str,
        'history_items': history_items,
        'metadata': metadata,
    })


@login_required
@require_POST
def btr_casting_plan_generate_view(request):
    """
    Auto-generates the Casting Plan Excel workbook from the active Camp Plan dataset.
    """
    from .btr_service import load_btr_dataset, generate_camp_plan_records, create_auto_casting_plan_excel
    
    planning_mat = request.POST.get('planning_material') or request.session.get('btr_selected_material') or 'UC_356_X_406'
    plan_date = request.POST.get('plan_date') or date.today().isoformat()
    plan_name = request.POST.get('plan_name') or f"Casting Plan {planning_mat}"

    upload_id = request.session.get('btr_upload_id')
    metadata, records = load_btr_dataset(upload_id) if upload_id else (None, None)
    
    camp_records = []
    if records and planning_mat:
        res = generate_camp_plan_records(records, planning_mat, upload_id=upload_id)
        camp_records = res.get('valid_records', [])

    if not camp_records and records:
        camp_records = records

    filename, out_path, bloom_count = create_auto_casting_plan_excel(camp_records, planning_mat, plan_date, plan_name)
    request.session['latest_casting_plan_path'] = str(out_path)
    request.session['latest_casting_plan_mat'] = str(planning_mat)

    Notification.objects.create(
        title="Casting Plan Generated",
        message=f"Casting Plan workbook '{filename}' generated for {planning_mat} ({bloom_count} blooms)."
    )

    messages.success(
        request,
        f"✓ Casting Plan workbook for {planning_mat} successfully generated! Filename: {filename} ({bloom_count} blooms calculated)."
    )
    return redirect('create_casting_plan')


@login_required
def btr_casting_plan_preview_api_view(request, filename):
    """
    JSON endpoint for previewing multi-sheet Casting Plan Excel workbooks in modal viewer.
    """
    import openpyxl
    from .btr_service import sanitize_for_json
    safe = get_valid_filename(filename)
    path = Path(settings.BASE_DIR) / 'media' / 'casting_plans' / safe
    if not path.is_file():
        path = Path(settings.BASE_DIR) / 'media' / 'rolling_plans' / safe
    if not path.is_file():
        path = Path(settings.BASE_DIR) / 'media' / 'generated' / safe
    if not path.is_file():
        return JsonResponse({'status': 'error', 'message': 'Casting plan file not found'}, status=404)

    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        sheets = {}
        sheet_names = wb.sheetnames
        for sheet_name in sheet_names:
            ws = wb[sheet_name]
            header_row = 3 if sheet_name != "Sheet1" else 2
            start_row = header_row + 1
            headers = [ws.cell(header_row, c).value for c in range(1, min(ws.max_column + 1, 20))] if ws.max_row >= header_row else []
            data_rows = []
            for r in range(start_row, min(ws.max_row + 1, 100)):
                row_vals = [ws.cell(r, c).value for c in range(1, len(headers) + 1)]
                if any(v not in (None, '') for v in row_vals):
                    data_rows.append(row_vals)
            sheets[sheet_name] = {
                'headers': headers,
                'rows': data_rows
            }
        wb.close()
        return JsonResponse({
            'status': 'success',
            'filename': safe,
            'sheet_names': sheet_names,
            'sheets': sanitize_for_json(sheets)
        })
    except Exception as exc:
        return JsonResponse({'status': 'error', 'message': str(exc)}, status=500)


@login_required
def btr_casting_plan_download_file_view(request, filename):
    """
    Download a generated Casting Plan Excel file.
    """
    safe = get_valid_filename(filename)
    path = Path(settings.BASE_DIR) / 'media' / 'casting_plans' / safe
    if not path.is_file():
        path = Path(settings.BASE_DIR) / 'media' / 'rolling_plans' / safe
    if not path.is_file():
        path = Path(settings.BASE_DIR) / 'media' / 'generated' / safe
    if not path.is_file():
        raise Http404("Casting plan file not found.")

    return FileResponse(
        open(path, 'rb'),
        as_attachment=True,
        filename=safe,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@login_required
def btr_casting_plan_delete_view(request, filename):
    """
    Deletes a generated Casting Plan file from disk.
    """
    safe = get_valid_filename(filename)
    path = Path(settings.BASE_DIR) / 'media' / 'casting_plans' / safe
    if path.is_file():
        path.unlink(missing_ok=True)
        messages.success(request, f"Deleted casting plan '{safe}'.")
    else:
        messages.error(request, f"Casting plan file '{safe}' not found.")
    return redirect('create_casting_plan')



@login_required
def plan_request_detail_view(request, request_id):
    """
    Renders results dashboard for an optimized CastingPlanRequest.
    """
    plan_req = get_object_or_404(CastingPlanRequest, id=request_id)
    reqs = plan_req.requirements.all()
    return render(request, 'casting_plan/plan_request_detail.html', {
        'plan_req': plan_req,
        'requirements': reqs
    })


@login_required
@require_GET
def api_section_weights(request):
    """
    AJAX endpoint returning valid section weights and resolved RM family preview for a planning section.
    """
    section = request.GET.get('section', '').strip()
    if not section:
        return JsonResponse({'status': 'error', 'message': 'Section is required'}, status=400)
        
    weights = list(SectionWeightMap.objects.filter(
        planning_section__iexact=section
    ).values_list('section_weight', flat=True).distinct().order_by('section_weight'))
    
    if not weights:
        # Fallback search without underscores/hyphens
        alt_sec = section.replace('_', ' ').replace('-', ' ')
        weights = list(SectionWeightMap.objects.filter(
            planning_section__iexact=alt_sec
        ).values_list('section_weight', flat=True).distinct().order_by('section_weight'))
        
    if not weights:
        weights = [40.3, 46.1, 54.0]
        
    sheet_name, raw_sec, raw_size, family = raw_family(section)
    
    return JsonResponse({
        'status': 'success',
        'section': section,
        'weights': [float(w) for w in weights],
        'default_family': sheet_name,
        'default_raw_section': raw_sec
    })


@login_required
@require_POST
def upload_demand(request):
    f = request.FILES.get('demand_file')
    if not f:
        messages.error(request, 'Please choose an Excel demand file.')
        return redirect('data')
    if not f.name.lower().endswith(('.xlsx','.xlsm')):
        messages.error(request, 'Please upload an .xlsx or .xlsm file.')
        return redirect('data')
        
    target_dir = Path(__import__('django.conf').conf.settings.BASE_DIR) / 'media' / 'uploads'
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / f.name
    with open(target, 'wb+') as out:
        for chunk in f.chunks(): out.write(chunk)
    try:
        result = import_demand_workbook(target, replace=request.POST.get('replace') == '1')
        messages.success(request, f"Imported {result['created']} demand records from {result['sheet']}.")
    except Exception as exc:
        messages.error(request, f'Could not import demand file: {exc}')
    return redirect('data')


def _session_btr_path(request):
    """Return a validated upload path kept in this user's session."""
    value = request.session.get('btr_upload_path')
    if not value:
        return None
    upload_root = (Path(settings.BASE_DIR) / 'media' / 'btr_uploads').resolve()
    candidate = Path(value).resolve()
    if upload_root not in candidate.parents or not candidate.is_file():
        request.session.pop('btr_upload_path', None)
        request.session.pop('btr_upload_info', None)
        return None
    return candidate

@login_required
def upload_btr_view(request):
    """Upload/inspect a BTR and display planning-material generation controls."""
    from .btr_service import process_btr_file_path, load_btr_dataset, list_all_btr_uploads, sanitize_for_json
    from django.utils.text import get_valid_filename
    from uuid import uuid4
    from pathlib import Path
    
    # Initialize variables at the very beginning to prevent UnboundLocalError
    metadata = None
    records = None
    
    if request.method == 'POST':
        uploaded = request.FILES.get('btr_file')
        if not uploaded:
            messages.error(request, 'Please choose a BTR Excel file.')
            return redirect('btr_input')
        if not uploaded.name.lower().endswith(('.xlsx', '.xlsm')):
            messages.error(request, 'Only .xlsx or .xlsm BTR files are supported.')
            return redirect('btr_input')
        if uploaded.size > 25 * 1024 * 1024:
            messages.error(request, 'BTR file is larger than the 25 MB upload limit.')
            return redirect('btr_input')

        upload_dir = Path(settings.BASE_DIR) / 'media' / 'btr_uploads'
        upload_dir.mkdir(parents=True, exist_ok=True)
        safe_name = get_valid_filename(Path(uploaded.name).name)

        # Check if file with same filename already exists in active dataset BTR Files history
        btr_uploads_history = list_all_btr_uploads()
        existing_upload = None
        for h in btr_uploads_history:
            if h.get('filename') and h.get('filename').strip().lower() == safe_name.lower():
                existing_upload = h
                break

        if existing_upload:
            existing_id = existing_upload.get('upload_id')
            metadata, records = load_btr_dataset(existing_id)
            if metadata:
                raw_unique = metadata.get('unique_planning_mats', [])
                mats_info = []
                for m in raw_unique:
                    if isinstance(m, dict):
                        n = str(m.get('name') or m.get('mat_name') or '').strip()
                        c = m.get('count') if m.get('count') is not None else m.get('row_count', 0)
                        if n:
                            mats_info.append({'name': n, 'mat_name': n, 'count': c, 'row_count': c})
                    elif isinstance(m, str) and m.strip():
                        n = m.strip()
                        mats_info.append({'name': n, 'mat_name': n, 'count': 0, 'row_count': 0})

                request.session['btr_upload_id'] = existing_id
                request.session['btr_upload_info'] = sanitize_for_json({
                    'upload_id': existing_id,
                    'original_name': metadata.get('filename', safe_name),
                    'sheet_name': metadata.get('sheet_name', 'CFD COS'),
                    'row_count': metadata.get('total_records', len(records or [])),
                    'planning_materials': mats_info,
                    'preview_records': (records or [])[:20],
                })
                messages.info(
                    request,
                    f"( Uploaded from active dataset BTR Files ) : '{safe_name}' already exists in active dataset and was activated."
                )
                return redirect('btr_input')

        target = upload_dir / f'{uuid4().hex}_{safe_name}'
        upload_id = target.stem
        try:
            with target.open('wb') as output:
                for chunk in uploaded.chunks():
                    output.write(chunk)
            metadata, records = process_btr_file_path(target, upload_id=upload_id, original_filename=safe_name)
        except Exception as exc:
            if target.exists():
                target.unlink()
            messages.error(request, f'{exc}')
            return redirect('btr_input')

        raw_unique = metadata.get('unique_planning_mats', [])
        mats_info = []
        for m in raw_unique:
            if isinstance(m, dict):
                n = str(m.get('name') or m.get('mat_name') or '').strip()
                c = m.get('count') if m.get('count') is not None else m.get('row_count', 0)
                if n:
                    mats_info.append({'name': n, 'mat_name': n, 'count': c, 'row_count': c})
            elif isinstance(m, str) and m.strip():
                n = m.strip()
                mats_info.append({'name': n, 'mat_name': n, 'count': 0, 'row_count': 0})

        request.session['btr_upload_id'] = upload_id
        request.session['btr_upload_path'] = str(target)
        request.session['btr_upload_info'] = sanitize_for_json({
            'upload_id': upload_id,
            'original_name': metadata.get('filename', safe_name),
            'sheet_name': metadata.get('sheet_name', 'CFD COS'),
            'row_count': metadata.get('total_records', len(records or [])),
            'planning_materials': mats_info,
            'preview_records': (records or [])[:20],
        })
        messages.success(
            request,
            f"BTR report '{safe_name}' uploaded and processed successfully: {metadata.get('total_records', len(records or []))} records found."
        )
        return redirect('btr_input')

    import json
    from .btr_service import get_btr_material_summary
    
    upload_id = request.session.get('btr_upload_id')
    btr_uploads_history = list_all_btr_uploads()
    
    if upload_id:
        metadata, records = load_btr_dataset(upload_id)
        
    first_mat = None
    if metadata and metadata.get('unique_planning_mats'):
        u0 = metadata['unique_planning_mats'][0]
        first_mat = u0.get('name') or u0.get('mat_name') if isinstance(u0, dict) else str(u0)
        
    selected_mat = request.session.get('btr_selected_material') or first_mat or 'AL_100X100'

    active_mat_summary = get_btr_material_summary(metadata, records, selected_mat) if metadata else {}
    btr_summary_map = (metadata or {}).get('btr_summary') or {}
    
    if metadata and not btr_summary_map and records:
        from .btr_service import compute_all_btr_material_summaries
        planning_col = (metadata or {}).get('planning_mat_col') or 'Planning Mat'
        btr_summary_map = compute_all_btr_material_summaries(records, planning_col)

    preview_records = (records or [])[:20] if records else []
    return render(request, 'casting_plan/upload_btr.html', {
        'metadata': metadata,
        'upload_info': request.session.get('btr_upload_info') if metadata else None,
        'preview_records': preview_records,
        'btr_uploads': btr_uploads_history,
        'active_upload_id': upload_id if metadata else None,
        'today': date.today().strftime('%d %b %Y, %I:%M %p'),
        'active_mat_summary': active_mat_summary,
        'btr_summary_json': json.dumps(btr_summary_map),
    })


@login_required
def btr_rolling_calendar_view(request):
    """Step 2: Upload Rolling Calendar — Renders draggable Rolling Calendar table."""
    from .btr_service import load_btr_dataset
    from .rolling_calendar_service import (
        build_rolling_calendar_dataset, list_uploaded_preports, list_uploaded_rules
    )

    upload_id = request.session.get('btr_upload_id')
    metadata, records = load_btr_dataset(upload_id) if upload_id else (None, None)

    custom_seq = request.session.get('rolling_calendar_custom_sequence')
    calendar_data = build_rolling_calendar_dataset(
        custom_sequence=custom_seq,
        active_btr_records=records,
        session=request.session
    )

    preport_files = list_uploaded_preports(request.session)
    rules_files = list_uploaded_rules(request.session)

    return render(request, 'casting_plan/rolling_calendar.html', {
        'metadata': metadata,
        'upload_info': request.session.get('btr_upload_info') if metadata else None,
        'calendar_rows': calendar_data['rows'],
        'executable_count': calendar_data['executable_count'],
        'total_camp_qty_fmt': calendar_data['total_camp_qty_fmt'],
        'unmapped_count': calendar_data['unmapped_count'],
        'has_preports': calendar_data.get('has_preports', False),
        'preport_files': preport_files,
        'rules_files': rules_files,
        'has_custom_sequence': bool(custom_seq),
        'is_fixed': request.session.get('rolling_calendar_is_fixed', False),
        'active_step': 2,
    })


@login_required
@require_POST
def btr_upload_preport_view(request):
    """Upload Preport Excel file(s)."""
    from .rolling_calendar_service import get_preports_dir, clear_rolling_calendar_cache
    from django.utils.text import get_valid_filename

    uploaded_files = request.FILES.getlist('preport_file') or [request.FILES.get('preport_file')]
    uploaded_files = [f for f in uploaded_files if f]

    if not uploaded_files:
        messages.error(request, 'Please choose a Preport Excel file.')
        return redirect('btr_rolling_calendar')

    pdir = get_preports_dir()
    saved_names = []

    for uploaded in uploaded_files:
        if not uploaded.name.lower().endswith(('.xlsx', '.xlsm')):
            continue
        safe_name = get_valid_filename(uploaded.name)
        target = pdir / safe_name
        with target.open('wb') as dest:
            for chunk in uploaded.chunks():
                dest.write(chunk)
        saved_names.append(safe_name)

    if saved_names:
        clear_rolling_calendar_cache()
        messages.success(request, f"Preport file(s) '{', '.join(saved_names)}' uploaded successfully.")
    else:
        messages.error(request, 'Only .xlsx or .xlsm Preport files are supported.')

    return redirect('btr_rolling_calendar')


@login_required
@require_POST
def btr_upload_rules_view(request):
    """Upload Rules Excel file."""
    from .rolling_calendar_service import get_rules_dir, clear_rolling_calendar_cache
    from django.utils.text import get_valid_filename

    uploaded = request.FILES.get('rules_file')
    if not uploaded:
        messages.error(request, 'Please choose a Rules Excel file.')
        return redirect('btr_rolling_calendar')

    if not uploaded.name.lower().endswith(('.xlsx', '.xlsm')):
        messages.error(request, 'Only .xlsx or .xlsm Rules files are supported.')
        return redirect('btr_rolling_calendar')

    rdir = get_rules_dir()
    safe_name = get_valid_filename(uploaded.name)
    target = rdir / safe_name

    with target.open('wb') as dest:
        for chunk in uploaded.chunks():
            dest.write(chunk)

    clear_rolling_calendar_cache()
    messages.success(request, f"Rules report file '{safe_name}' uploaded successfully.")
    return redirect('btr_rolling_calendar')


@login_required
@require_POST
def btr_save_calendar_sequence_view(request):
    """Save drag order, row edits, activation state, and soft-deleted rows."""
    import json
    try:
        data = json.loads(request.body)
        sequence = data.get('sequence')
        overrides = data.get('overrides', {})
        deleted_ids = data.get('deleted_ids', [])
        if isinstance(sequence, list) and isinstance(overrides, dict) and isinstance(deleted_ids, list):
            request.session['rolling_calendar_custom_sequence'] = sequence
            request.session['rolling_calendar_row_overrides'] = overrides
            request.session['rolling_calendar_deleted_ids'] = deleted_ids
            request.session.modified = True
            return JsonResponse({'status': 'success', 'message': 'Calendar changes saved.'})
        return JsonResponse({'status': 'error', 'message': 'Invalid sequence data.'}, status=400)
    except Exception as exc:
        return JsonResponse({'status': 'error', 'message': str(exc)}, status=500)


@login_required
def btr_reset_calendar_sequence_view(request):
    """Resets calendar row sequence back to default."""
    request.session.pop('rolling_calendar_custom_sequence', None)
    request.session.pop('rolling_calendar_is_fixed', None)
    request.session.pop('rolling_calendar_row_overrides', None)
    request.session.pop('rolling_calendar_deleted_ids', None)
    if request.headers.get('x-requested-with') == 'XMLHttpRequest' or request.GET.get('ajax') == '1':
        return JsonResponse({'status': 'success', 'message': 'Sequence reset to default.'})
    messages.info(request, "Rolling Calendar sequencing reset to default order.")
    return redirect('btr_rolling_calendar')


@login_required
def btr_save_and_fix_calendar_view(request):
    """
    Fixes and locks the current Rolling Calendar sequence and dataset for the active reports,
    then moves forward to Step 3: Planning Material.
    """
    from .btr_service import load_btr_dataset
    from .rolling_calendar_service import build_rolling_calendar_dataset

    upload_id = request.session.get('btr_upload_id')
    metadata, records = load_btr_dataset(upload_id) if upload_id else (None, None)
    custom_seq = request.session.get('rolling_calendar_custom_sequence')
    calendar_data = build_rolling_calendar_dataset(
        custom_sequence=custom_seq,
        active_btr_records=records,
        session=request.session
    )

    if not calendar_data.get('has_preports'):
        messages.error(request, "Cannot fix calendar: Please upload at least one Preport file first.")
        return redirect('btr_rolling_calendar')

    request.session['rolling_calendar_is_fixed'] = True
    messages.success(
        request,
        "✓ Rolling Calendar sequence and dataset successfully saved and fixed for current Preport, Rules, and BTR reports!"
    )
    return redirect('btr_planning_material')


@login_required
def btr_remove_calendar_file_view(request):
    """Remove uploaded preport or rules file from media directory."""
    file_type = request.GET.get('type') or request.POST.get('type')
    filename = request.GET.get('filename') or request.POST.get('filename')

    if filename and file_type:
        from django.utils.text import get_valid_filename
        from .rolling_calendar_service import get_preports_dir, get_rules_dir, clear_rolling_calendar_cache
        safe = get_valid_filename(filename)
        target_dir = get_preports_dir() if file_type == 'preport' else get_rules_dir()
        target_file = target_dir / safe
        if target_file.is_file():
            try:
                target_file.unlink()
                clear_rolling_calendar_cache()
                messages.success(request, f"Removed file '{safe}'.")
            except Exception:
                pass
    return redirect('btr_rolling_calendar')



@login_required
def select_btr_upload_view(request, upload_id):
    """Set selected BTR Upload dataset as active in user session."""
    from .btr_service import load_btr_dataset, sanitize_for_json
    metadata, records = load_btr_dataset(upload_id)
    if metadata:
        raw_unique = metadata.get('unique_planning_mats', [])
        mats_info = []
        for m in raw_unique:
            if isinstance(m, dict):
                n = str(m.get('name') or m.get('mat_name') or '').strip()
                c = m.get('count') if m.get('count') is not None else m.get('row_count', 0)
                if n:
                    mats_info.append({'name': n, 'mat_name': n, 'count': c, 'row_count': c})
            elif isinstance(m, str) and m.strip():
                n = m.strip()
                mats_info.append({'name': n, 'mat_name': n, 'count': 0, 'row_count': 0})

        request.session['btr_upload_id'] = upload_id
        request.session['btr_upload_info'] = sanitize_for_json({
            'upload_id': upload_id,
            'original_name': metadata.get('filename', 'BTR Report.xlsx'),
            'sheet_name': metadata.get('sheet_name', 'CFD COS'),
            'row_count': metadata.get('total_records', len(records or [])),
            'planning_materials': mats_info,
            'preview_records': (records or [])[:20],
        })
        messages.success(request, f"BTR report '{metadata.get('filename')}' set as active dataset.")
    else:
        messages.error(request, "Selected BTR report dataset not found.")
    return redirect('btr_input')


@login_required
def clear_btr_upload_view(request):
    """Deactivate / clear active BTR upload dataset from user session."""
    filename = request.session.get('btr_upload_info', {}).get('original_name', 'BTR report')
    request.session.pop('btr_upload_id', None)
    request.session.pop('btr_upload_path', None)
    request.session.pop('btr_upload_info', None)
    request.session.pop('btr_selected_material', None)
    messages.info(request, f"Active BTR dataset '{filename}' cleared. No active dataset is currently selected.")
    return redirect('btr_input')


@login_required
@require_POST
def delete_btr_upload_view(request, upload_id):
    """Delete a BTR report from upload history."""
    from .btr_service import load_btr_dataset, delete_btr_upload_dataset
    metadata, _ = load_btr_dataset(upload_id)
    filename = metadata.get('filename', 'BTR report') if metadata else 'BTR report'
    delete_btr_upload_dataset(upload_id)
    if request.session.get('btr_upload_id') == upload_id:
        request.session.pop('btr_upload_id', None)
        request.session.pop('btr_upload_info', None)
        request.session.pop('btr_selected_material', None)
    messages.success(request, f"BTR upload history for {filename} was deleted.")
    return redirect('btr_input')


@login_required
def btr_upload_detail_api_view(request, upload_id):
    """API Endpoint: Returns BTR dataset metadata & top 20 preview records for modal preview."""
    import json
    from .btr_service import load_btr_dataset, sanitize_for_json
    metadata, records = load_btr_dataset(upload_id)
    if not metadata:
        return JsonResponse({'error': 'BTR report dataset not found.'}, status=404)
    preview = sanitize_for_json(records[:20]) if records else []
    payload = {
        'status': 'success',
        'metadata': sanitize_for_json(metadata),
        'preview_records': preview,
        'records_count': len(records) if records else 0
    }
    return JsonResponse(payload)


@login_required
def btr_planning_material_view(request):
    """Step 2: Planning Material Selection"""
    from .btr_service import load_btr_dataset
    upload_id = request.session.get('btr_upload_id')
    if not upload_id:
        messages.error(request, 'Please upload a BTR report first.')
        return redirect('btr_input')
    metadata, records = load_btr_dataset(upload_id)
    if not metadata:
        messages.error(request, 'BTR dataset not found. Please upload again.')
        return redirect('btr_input')
    
    if request.GET.get('clear') == '1':
        request.session.pop('btr_selected_material', None)
        return redirect('btr_planning_material')

    if request.method == 'POST':
        selected_mat = request.POST.get('planning_mat') or request.POST.get('planning_material')
        if selected_mat:
            request.session['btr_selected_material'] = selected_mat
            request.session.pop('custom_camp_plan_path', None)
            request.session.pop('custom_camp_plan_name', None)
            request.session.pop('custom_camp_plan_mat', None)
            return redirect('btr_planning_material')

    selected_mat = request.session.get('btr_selected_material', '')
    
    # Extract pre-computed planning materials from metadata in 0ms
    mat_counts = {}
    for m_item in (metadata or {}).get('unique_planning_mats', []):
        if isinstance(m_item, dict):
            m_name = m_item.get('name') or m_item.get('mat_name')
            cnt = m_item.get('count', 0)
        else:
            m_name = str(m_item)
            cnt = 0
        if m_name:
            mat_counts[m_name] = cnt

    # Fetch active Rolling Calendar sections to constrain Planning Material selection
    from .rolling_calendar_service import build_rolling_calendar_dataset
    custom_seq = request.session.get('rolling_calendar_custom_sequence')
    calendar_dataset = build_rolling_calendar_dataset(custom_sequence=custom_seq, active_btr_records=records, session=request.session)
    calendar_rows = calendar_dataset.get('rows', [])

    # Calendar edits are display/scheduling inputs only. Never use edited Camp Qty,
    # RM values, or Last Rolling Qty as BTR/campaign source data for downstream plans.
    # The original BTR metadata remains the single source of truth for material counts.
    if not mat_counts and records:
        planning_col = (metadata or {}).get('planning_mat_col') or 'Planning Mat'
        for source_row in records:
            source_name = source_row.get(planning_col) or source_row.get('Planning Mat') or source_row.get('Planning_Mat')
            source_name = str(source_name or '').strip()
            if source_name:
                mat_counts[source_name] = mat_counts.get(source_name, 0) + 1

    source_names = list(mat_counts.keys())
    source_by_normalized = {
        ''.join(ch for ch in str(name).upper() if ch.isalnum()): name
        for name in source_names
    }
    ordered_source_names = []
    seen_source_names = set()

    # Keep the user's calendar sequence only as a presentation order for valid,
    # non-deleted source materials. No edited calendar quantity is copied forward.
    for crow in calendar_rows:
        if crow.get('is_deleted'):
            continue
        section_name = str(crow.get('section') or '').strip()
        normalized = ''.join(ch for ch in section_name.upper() if ch.isalnum())
        source_name = source_by_normalized.get(normalized)
        if source_name and source_name not in seen_source_names:
            ordered_source_names.append(source_name)
            seen_source_names.add(source_name)

    for source_name in source_names:
        if source_name not in seen_source_names:
            ordered_source_names.append(source_name)
            seen_source_names.add(source_name)

    unique_mats = [
        {
            'mat_name': name,
            'name': name,
            'mat_count': mat_counts.get(name, 0),
            'count': mat_counts.get(name, 0),
            'row_count': mat_counts.get(name, 0),
        }
        for name in ordered_source_names
    ]

    valid_materials = {item['mat_name'] for item in unique_mats}
    if selected_mat and selected_mat not in valid_materials:
        selected_mat = ''
        request.session.pop('btr_selected_material', None)

    if not selected_mat and unique_mats:
        selected_mat = unique_mats[0]['mat_name']
        request.session['btr_selected_material'] = selected_mat

    # Filter records by active selection
    planning_col = (metadata or {}).get('planning_mat_col') or 'Planning Mat'
    filtered_records = []
    if selected_mat and records:
        sel_upper = selected_mat.strip().upper()
        for r in records:
            mat_val = r.get(planning_col) or r.get('Planning Mat') or r.get('Planning_Mat')
            if not mat_val:
                for k, v in r.items():
                    if 'planning' in str(k).lower() and 'mat' in str(k).lower():
                        mat_val = v
                        break
            if mat_val and str(mat_val).strip().upper() == sel_upper:
                filtered_records.append(r)

    import json
    from .btr_service import get_btr_material_summary, compute_all_btr_material_summaries
    active_mat_summary = get_btr_material_summary(metadata, records, selected_mat)
    
    # Pre-build summary map for all unique mats for instant JS tab updating
    btr_summary_map = (metadata or {}).get('btr_summary') or {}
    if not btr_summary_map and records:
        btr_summary_map = compute_all_btr_material_summaries(records, planning_col)

    return render(request, 'casting_plan/btr_planning_material.html', {
        'metadata': metadata,
        'unique_mats': unique_mats,
        'selected_mat': selected_mat,
        'filtered_count': len(filtered_records),
        'filtered_preview': filtered_records[:10],
        'active_mat_summary': active_mat_summary,
        'btr_summary_json': json.dumps(btr_summary_map),
    })


@login_required
def btr_camp_plan_view(request):
    """Step 3: Campaign Plan View"""
    from .btr_service import load_btr_dataset, generate_camp_plan_records, process_uploaded_camp_plan

    if request.GET.get('clear_custom') == '1':
        request.session.pop('custom_camp_plan_path', None)
        request.session.pop('custom_camp_plan_name', None)
        request.session.pop('custom_camp_plan_mat', None)
        messages.info(request, "Reset to auto-generated Campaign Plan dataset.")
        return redirect('btr_camp_plan')

    upload_id = request.session.get('btr_upload_id')
    if not upload_id:
        messages.error(request, 'Please upload a BTR report first.')
        return redirect('btr_input')
    metadata, records = load_btr_dataset(upload_id)
    url_mat = request.GET.get('mat') or request.GET.get('material')
    selected_mat = url_mat or request.session.get('btr_selected_material')
    
    custom_camp_mat = request.session.get('custom_camp_plan_mat')
    if (url_mat or selected_mat) and custom_camp_mat and str(url_mat or selected_mat).strip().upper() != str(custom_camp_mat).strip().upper():
        request.session.pop('custom_camp_plan_path', None)
        request.session.pop('custom_camp_plan_name', None)
        request.session.pop('custom_camp_plan_mat', None)

    if selected_mat:
        request.session['btr_selected_material'] = selected_mat
    if not selected_mat and metadata and metadata.get('unique_planning_mats'):
        selected_mat = metadata['unique_planning_mats'][0]
        request.session['btr_selected_material'] = selected_mat
    if not selected_mat:
        selected_mat = 'AL_110X110'

    custom_camp_path = request.session.get('custom_camp_plan_path')
    is_custom_upload = False
    custom_filename = ""

    if custom_camp_path and Path(custom_camp_path).is_file():
        try:
            valid_records, summary, detected_mat = process_uploaded_camp_plan(custom_camp_path)
            selected_mat = detected_mat or selected_mat
            request.session['btr_selected_material'] = selected_mat
            is_custom_upload = True
            custom_filename = request.session.get('custom_camp_plan_name', 'Uploaded File')

            valid_count = len(valid_records)
            summary.update({
                'total_input': valid_count,
                'input_count': valid_count,
                'valid_count': valid_count,
                'removed_count': 0,
            })

            result = {
                'valid_records': valid_records,
                'removed_records': [],
                'summary': summary,
                'subtotals': {}
            }
        except Exception:
            request.session.pop('custom_camp_plan_path', None)
            result = generate_camp_plan_records(records or [], selected_mat, upload_id=upload_id) if selected_mat else {'valid_records': [], 'removed_records': [], 'summary': {}}
    else:
        result = generate_camp_plan_records(records or [], selected_mat, upload_id=upload_id) if selected_mat else {'valid_records': [], 'removed_records': [], 'summary': {}}

    all_valid_records = result.get('valid_records', [])
    sap_info = get_sap_assignment_info(selected_mat)
    is_sent_to_sap = sap_info.get('is_sent', False) or request.session.get(f'sap_sent_{selected_mat}', False)
    is_assigned_by_sap = sap_info.get('is_assigned', False) or request.session.get(f'sap_assigned_by_sap_{selected_mat}', False)
    sap_camp_no = sap_info.get('camp_plan_no') or request.session.get(f'sap_camp_no_{selected_mat}', '6427')

    return render(request, 'casting_plan/btr_camp_plan.html', {
        'metadata': metadata,
        'selected_mat': selected_mat,
        'planning_material': selected_mat,
        'valid_records': all_valid_records[:100],
        'camp_records': all_valid_records[:100],
        'total_valid_count': len(all_valid_records),
        'removed_records': result.get('removed_records', []),
        'summary': result.get('summary', {}),
        'subtotals': result.get('subtotals', {}),
        'is_custom_upload': is_custom_upload,
        'custom_filename': custom_filename,
        'is_sent_to_sap': is_sent_to_sap,
        'is_assigned_by_sap': is_assigned_by_sap,
        'sap_camp_no': sap_camp_no,
    })


@login_required
@require_POST
def btr_send_to_sap_view(request):
    """Simulate sending Campaign Plan to SAP system and generating assigned Camp Plan No."""
    selected_mat = request.POST.get('planning_material') or request.session.get('btr_selected_material') or 'AL_250X250'
    camp_no = request.POST.get('camp_plan_no') or '6427'
    
    update_sap_assignment_info(selected_mat, is_sent=True, camp_plan_no=camp_no)
    request.session[f'sap_sent_{selected_mat}'] = True
    request.session[f'sap_camp_no_{selected_mat}'] = camp_no
    request.session['btr_selected_material'] = selected_mat
    
    messages.success(
        request,
        f"Campaign Plan for {selected_mat} successfully sent to SAP! Switch to SAP Portal or download SAP Excel."
    )
    return redirect('btr_camp_plan')


def sap_login_view(request):
    """Dedicated SAP ERP Portal Login View"""
    error = None
    if request.method == 'POST':
        u = request.POST.get('username', '').strip()
        p = request.POST.get('password', '').strip()
        if (u.lower() in ('sap', 'sap_user', 'admin') and p in ('sap123', 'admin123')) or authenticate(request, username=u, password=p) is not None:
            request.session['is_sap_user'] = True
            request.session['sap_username'] = u
            messages.success(request, f"Welcome to SAP ERP Portal, {u}!")
            return redirect('sap_portal')
        error = 'Invalid SAP username or password. Default: sap / sap123'
    return render(request, 'casting_plan/sap_login.html', {'error': error})


@login_required
def sap_portal_view(request):
    """SAP ERP Portal Dashboard listing received Campaign Plans from Rail Mill"""
    request.session['is_sap_user'] = True
    from .btr_service import load_btr_dataset
    upload_id = request.session.get('btr_upload_id')
    metadata, records = load_btr_dataset(upload_id) if upload_id else (None, None)
    
    unique_mats = (metadata or {}).get('unique_planning_mats') or []
    mat_names = []
    for m in unique_mats:
        n = m.get('name') if isinstance(m, dict) else str(m)
        if n and n not in mat_names:
            mat_names.append(n)

    # Collect all materials stored in persistent SAP assignments
    all_saved_assignments = get_all_sap_assignments()
    for saved_mat_key, saved_info in all_saved_assignments.items():
        if saved_info.get('is_sent') or saved_info.get('is_assigned'):
            if saved_mat_key not in mat_names:
                mat_names.append(saved_mat_key)
            
    active_mat = request.session.get('btr_selected_material') or (mat_names[0] if mat_names else 'AL_250X250')
    if active_mat and active_mat not in mat_names:
        mat_names.insert(0, active_mat)

    sap_plans = []
    pending_count = 0
    assigned_count = 0

    for m in mat_names:
        info = get_sap_assignment_info(m)
        is_sent = info.get('is_sent', False) or request.session.get(f'sap_sent_{m}', False)
        is_assigned = info.get('is_assigned', False) or request.session.get(f'sap_assigned_by_sap_{m}', False)
        camp_no = info.get('camp_plan_no') or request.session.get(f'sap_camp_no_{m}', '6427')
        
        if is_sent or is_assigned or m == active_mat:
            if is_assigned:
                assigned_count += 1
            else:
                pending_count += 1

            sap_plans.append({
                'planning_material': m,
                'received_at': info.get('updated_at') and timezone.localtime().strftime('%d %b %Y, %I:%M %p') or timezone.localtime().strftime('%d %b %Y, %I:%M %p'),
                'source_file': metadata.get('filename', 'BTR Report.xlsx') if metadata else 'BTR Report',
                'valid_count': len(records or []) if records else 200,
                'total_btr_qty': round(sum(float(r.get('BTR') or r.get('Bal Roll Qty') or 0) for r in (records or [])) or 6580.5, 2),
                'is_assigned': is_assigned,
                'camp_plan_no': camp_no,
            })

    return render(request, 'casting_plan/sap_portal.html', {
        'sap_plans': sap_plans,
        'pending_count': pending_count,
        'assigned_count': assigned_count,
    })


@login_required
@require_POST
def sap_assign_camp_plan_view(request):
    """Assign SAP Camp Plan Number and transmit back to Rail Mill"""
    planning_mat = request.POST.get('planning_material') or 'AL_250X250'
    camp_no = request.POST.get('camp_plan_no') or '6427'
    
    update_sap_assignment_info(planning_mat, is_sent=True, is_assigned=True, camp_plan_no=camp_no)
    request.session[f'sap_sent_{planning_mat}'] = True
    request.session[f'sap_camp_no_{planning_mat}'] = camp_no
    request.session[f'sap_assigned_by_sap_{planning_mat}'] = True
    request.session['btr_selected_material'] = planning_mat
    
    Notification.objects.create(
        title="SAP Camp Plan Number Assigned",
        message=f"SAP ERP System assigned CAMP PLAN NO. {camp_no} to {planning_mat} and transmitted back to Rail Mill."
    )
    
    messages.success(
        request,
        f"✓ Camp Plan No. {camp_no} assigned to {planning_mat} and successfully transmitted back to Rail Mill!"
    )
    return redirect('sap_portal')


@login_required
@require_POST
def btr_camp_plan_upload_view(request):
    """Step 3: Upload custom edited/filtered Campaign Plan Excel file"""
    from .btr_service import process_uploaded_camp_plan
    uploaded_file = request.FILES.get('camp_plan_file') or request.FILES.get('campaign_file')
    if not uploaded_file:
        messages.error(request, 'Please select a Campaign Plan Excel file to upload.')
        return redirect('btr_camp_plan')
    if not uploaded_file.name.lower().endswith(('.xlsx', '.xlsm')):
        messages.error(request, 'Only .xlsx or .xlsm Campaign Plan files are supported.')
        return redirect('btr_camp_plan')

    upload_dir = Path(settings.BASE_DIR) / 'media' / 'btr_uploads' / 'custom_camp_plans'
    upload_dir.mkdir(parents=True, exist_ok=True)
    safe_name = get_valid_filename(Path(uploaded_file.name).name)
    target_path = upload_dir / f"{uuid4().hex[:8]}_{safe_name}"

    try:
        with open(target_path, 'wb+') as dest:
            for chunk in uploaded_file.chunks():
                dest.write(chunk)

        valid_records, summary, detected_mat = process_uploaded_camp_plan(target_path)
    except Exception as exc:
        if target_path.exists():
            target_path.unlink(missing_ok=True)
        messages.error(request, f"Could not process custom Campaign Plan file: {exc}")
        return redirect('btr_camp_plan')

    current_mat = detected_mat or request.session.get('btr_selected_material') or 'AL_110X110'
    request.session['custom_camp_plan_path'] = str(target_path)
    request.session['custom_camp_plan_name'] = safe_name
    request.session['custom_camp_plan_mat'] = current_mat
    request.session['btr_selected_material'] = current_mat

    messages.success(
        request,
        f"Campaign generated for {current_mat}: {summary['valid_count']} rows, {summary['total_btr_qty']} MT. Download it from Campaign History."
    )
    return redirect('btr_camp_plan')


@login_required
def btr_camp_plan_export_view(request):
    """Step 3: Export Campaign Plan Excel"""
    from .btr_service import load_btr_dataset, generate_camp_plan_records, create_camp_plan_excel, process_uploaded_camp_plan
    upload_id = request.session.get('btr_upload_id')
    selected_mat = request.GET.get('mat') or request.GET.get('material') or request.session.get('btr_selected_material')
    if not upload_id or not selected_mat:
        messages.error(request, 'No active BTR dataset or material selected for export.')
        return redirect('btr_input')
    metadata, records = load_btr_dataset(upload_id)

    custom_camp_path = request.session.get('custom_camp_plan_path')
    if custom_camp_path and Path(custom_camp_path).is_file():
        try:
            valid_records, _, detected_mat = process_uploaded_camp_plan(custom_camp_path)
            selected_mat = detected_mat or selected_mat
        except Exception:
            result = generate_camp_plan_records(records or [], selected_mat, upload_id=upload_id)
            valid_records = result.get('valid_records', [])
    else:
        result = generate_camp_plan_records(records or [], selected_mat, upload_id=upload_id)
        valid_records = result.get('valid_records', [])

    sap_info = get_sap_assignment_info(selected_mat)
    is_received_from_sap = sap_info.get('is_assigned', False) or request.session.get(f'sap_assigned_by_sap_{selected_mat}', False)
    is_sap_download = (request.GET.get('sap') == '1')
    include_sap_col = is_received_from_sap or is_sap_download
    sap_camp_no = sap_info.get('camp_plan_no') or request.session.get(f'sap_camp_no_{selected_mat}', '6427')

    source_headers = metadata.get('headers') if metadata else None
    output_filename, excel_path = create_camp_plan_excel(
        valid_records,
        selected_mat,
        source_headers=source_headers,
        source_records=records or [],
        include_sap_col=include_sap_col,
        camp_plan_no=sap_camp_no,
        is_received_from_sap=is_received_from_sap
    )
    return FileResponse(
        open(excel_path, 'rb'),
        as_attachment=True,
        filename=output_filename,
        content_type='application/vnd.ms-excel.sheet.macroEnabled.12'
    )


@login_required
def btr_embossing_plan_view(request):
    """Step 5: Campaign Plan Summary Review View"""
    from collections import OrderedDict
    from .btr_service import load_btr_dataset, generate_camp_plan_records, process_uploaded_camp_plan
    
    upload_id = request.session.get('btr_upload_id')
    metadata, records = load_btr_dataset(upload_id) if upload_id else (None, None)
    selected_mat = request.GET.get('mat') or request.GET.get('material') or request.session.get('btr_selected_material') or 'AL_250X250'
    
    custom_camp_path = request.session.get('custom_camp_plan_path')
    if custom_camp_path and Path(custom_camp_path).is_file():
        try:
            valid_records, _, detected_mat = process_uploaded_camp_plan(custom_camp_path)
            selected_mat = detected_mat or selected_mat
        except Exception:
            result = generate_camp_plan_records(records or [], selected_mat, upload_id=upload_id) if records else {'valid_records': []}
            valid_records = result.get('valid_records', [])
    else:
        result = generate_camp_plan_records(records or [], selected_mat, upload_id=upload_id) if records else {'valid_records': []}
        valid_records = result.get('valid_records', [])

    calc_master_map = {}
    try:
        from .models import RailMillCalculationMaster
        for cm in RailMillCalculationMaster.objects.all():
            if cm.grade_group:
                g_val = cm.grade_group.strip()
                if cm.ext_grade:
                    calc_master_map[cm.ext_grade.strip().upper()] = g_val
                if cm.material:
                    calc_master_map[cm.material.strip().upper()] = g_val
                if cm.planning_mat and cm.ext_grade:
                    calc_master_map[(cm.planning_mat.strip().upper(), cm.ext_grade.strip().upper())] = g_val
    except Exception:
        pass

    from .btr_service import get_grade_group_for_record

    # Group valid_records by (Planning Mat, Sec Wt, Grade group) matching Sheet2 pivot table
    grouped = OrderedDict()
    for r in (valid_records or []):
        p_mat = str(r.get('Planning Mat') or r.get('Planning_Mat') or selected_mat).strip()
        sec_wt_raw = r.get('Sec_Wt') if r.get('Sec_Wt') is not None else (r.get('Sec Wt') if r.get('Sec Wt') is not None else (r.get('Sec_wt') if r.get('Sec_wt') is not None else 0.0))
        try:
            sec_wt_val = float(sec_wt_raw)
            sec_wt_disp = f"{sec_wt_val:g}"
        except (ValueError, TypeError):
            sec_wt_val = 0.0
            sec_wt_disp = str(sec_wt_raw)
        
        grade_grp = get_grade_group_for_record(r, p_mat, calc_master_map)
        tonnage = float(r.get('BTR') or r.get('BTR Qty') or r.get('BTR_Qty') or 0.0)

        key = (p_mat, sec_wt_val, sec_wt_disp, grade_grp)
        grouped[key] = grouped.get(key, 0.0) + tonnage

    sequence_rows = []
    total_tonnage = 0.0
    sec_wts = []
    
    for idx, ((p_mat, sec_wt_val, sec_wt_disp, grade_grp), total_btr) in enumerate(grouped.items(), 1):
        total_tonnage += total_btr
        if sec_wt_val > 0:
            sec_wts.append(sec_wt_val)
        
        formatted_total = int(round(total_btr)) if abs(total_btr - round(total_btr)) < 0.001 else round(total_btr, 2)
        
        sequence_rows.append({
            'seq_no': idx,
            'planning_mat': p_mat,
            'sec_wt': sec_wt_disp,
            'grade_group': grade_grp,
            'total': formatted_total,
            'tonnage': formatted_total,
            'thickness': sec_wt_disp,
            'std': grade_grp,
        })

    thick_range = f"{min(sec_wts):g} – {max(sec_wts):g} mm" if sec_wts else "N/A"
    is_submitted = request.session.get(f'embossing_sent_{selected_mat}', False)
    disp_total_tonnage = int(round(total_tonnage)) if abs(total_tonnage - round(total_tonnage)) < 0.001 else f"{total_tonnage:,.2f}"

    return render(request, 'casting_plan/btr_embossing_plan.html', {
        'planning_material': selected_mat,
        'total_records': len(sequence_rows),
        'thickness_range': thick_range,
        'total_tonnage': disp_total_tonnage,
        'sequence_rows': sequence_rows,
        'is_submitted': is_submitted,
    })

@login_required
@require_POST
def btr_embossing_plan_send_view(request):
    """Redirects to Roll Shop Sequence page after Campaign Summary review"""
    selected_mat = request.POST.get('planning_material') or 'AL_250X250'
    request.session[f'embossing_sent_{selected_mat}'] = True
    messages.success(request, f"Campaign Summary for {selected_mat} verified. Proceeding to Roll Shop Sequence.")
    return redirect(f"{reverse('btr_roll_shop_sequence')}?mat={selected_mat}")

@login_required
def btr_roll_shop_sequence_view(request):
    """Step 6: Roll Shop Sequence Arrangement"""
    import json
    from collections import OrderedDict
    from .btr_service import load_btr_dataset, generate_camp_plan_records, process_uploaded_camp_plan
    
    upload_id = request.session.get('btr_upload_id')
    metadata, records = load_btr_dataset(upload_id) if upload_id else (None, None)
    selected_mat = request.GET.get('mat') or request.GET.get('material') or request.session.get('btr_selected_material') or 'AL_250X250'
    
    if request.GET.get('reset') == '1':
        request.session.pop(f'roll_shop_sequence_{selected_mat}', None)
        return redirect(f"{reverse('btr_roll_shop_sequence')}?mat={selected_mat}")

    custom_camp_path = request.session.get('custom_camp_plan_path')
    if custom_camp_path and Path(custom_camp_path).is_file():
        try:
            valid_records, _, detected_mat = process_uploaded_camp_plan(custom_camp_path)
            selected_mat = detected_mat or selected_mat
        except Exception:
            result = generate_camp_plan_records(records or [], selected_mat, upload_id=upload_id) if records else {'valid_records': []}
            valid_records = result.get('valid_records', [])
    else:
        result = generate_camp_plan_records(records or [], selected_mat, upload_id=upload_id) if records else {'valid_records': []}
        valid_records = result.get('valid_records', [])

    calc_master_map = {}
    try:
        from .models import RailMillCalculationMaster
        for cm in RailMillCalculationMaster.objects.all():
            if cm.grade_group:
                g_val = cm.grade_group.strip()
                if cm.ext_grade:
                    calc_master_map[cm.ext_grade.strip().upper()] = g_val
                if cm.material:
                    calc_master_map[cm.material.strip().upper()] = g_val
                if cm.planning_mat and cm.ext_grade:
                    calc_master_map[(cm.planning_mat.strip().upper(), cm.ext_grade.strip().upper())] = g_val
    except Exception:
        pass

    # Group valid_records by (Planning Mat, Sec Wt, Grade group) matching Sheet2 pivot table
    grouped = OrderedDict()
    for r in (valid_records or []):
        p_mat = str(r.get('Planning Mat') or r.get('Planning_Mat') or selected_mat).strip()
        sec_wt_raw = r.get('Sec_Wt') if r.get('Sec_Wt') is not None else (r.get('Sec Wt') if r.get('Sec Wt') is not None else (r.get('Sec_wt') if r.get('Sec_wt') is not None else 0.0))
        try:
            sec_wt_val = float(sec_wt_raw)
            sec_wt_disp = f"{sec_wt_val:g}"
        except (ValueError, TypeError):
            sec_wt_val = 0.0
            sec_wt_disp = str(sec_wt_raw)
        
        grade_grp = get_grade_group_for_record(r, p_mat, calc_master_map)
        tonnage = float(r.get('BTR') or r.get('BTR Qty') or r.get('BTR_Qty') or 0.0)

        key = (p_mat, sec_wt_val, sec_wt_disp, grade_grp)
        grouped[key] = grouped.get(key, 0.0) + tonnage

    sequence_rows = []
    total_tonnage = 0.0
    sec_wts = []
    
    for idx, ((p_mat, sec_wt_val, sec_wt_disp, grade_grp), total_btr) in enumerate(grouped.items(), 1):
        total_tonnage += total_btr
        if sec_wt_val > 0:
            sec_wts.append(sec_wt_val)
        
        formatted_total = int(round(total_btr)) if abs(total_btr - round(total_btr)) < 0.001 else round(total_btr, 2)
        
        sequence_rows.append({
            'seq_no': idx,
            'planning_mat': p_mat,
            'sec_wt': sec_wt_disp,
            'grade_group': grade_grp,
            'total': formatted_total,
            'tonnage': formatted_total,
            'thickness': sec_wt_disp,
            'std': grade_grp,
        })

    # Check if custom sequence was saved in session
    saved_seq = request.session.get(f'roll_shop_sequence_{selected_mat}')
    if saved_seq and isinstance(saved_seq, list) and len(saved_seq) > 0:
        sequence_rows = saved_seq

    thick_range = f"{min(sec_wts):g} – {max(sec_wts):g} mm" if sec_wts else "N/A"
    disp_total_tonnage = int(round(total_tonnage)) if abs(total_tonnage - round(total_tonnage)) < 0.001 else f"{total_tonnage:,.2f}"

    return render(request, 'casting_plan/btr_roll_shop_sequence.html', {
        'planning_material': selected_mat,
        'total_records': len(sequence_rows),
        'thickness_range': thick_range,
        'total_tonnage': disp_total_tonnage,
        'sequence_rows': sequence_rows,
    })

@login_required
@require_POST
def btr_roll_shop_submit_view(request):
    """Submits Roll Shop sequence and redirects to Step 7: Rolling Plan"""
    import json
    selected_mat = request.POST.get('planning_material') or 'AL_250X250'
    seq_data_raw = request.POST.get('sequence_data')
    if seq_data_raw:
        try:
            seq_list = json.loads(seq_data_raw)
            if isinstance(seq_list, list):
                request.session[f'roll_shop_sequence_{selected_mat}'] = seq_list
        except Exception:
            pass
    messages.success(request, f"Sequence for {selected_mat} finalized by Roll Shop. Proceeding to Rolling Plan.")
    return redirect(f"{reverse('btr_rolling_plan')}?mat={selected_mat}")    

@login_required
def btr_rolling_plan_view(request):
    """Step 4: Rolling Plan Generation & History View"""
    import os
    from datetime import datetime, date
    from .btr_service import load_btr_dataset, get_rolling_plan_file_metrics
    from .models import BTRCampaign
    upload_id = request.session.get('btr_upload_id')
    metadata, records = load_btr_dataset(upload_id) if upload_id else (None, None)
    selected_mat = request.GET.get('material') or request.GET.get('mat') or request.session.get('btr_selected_material')
    if selected_mat:
        request.session['btr_selected_material'] = selected_mat
    if not selected_mat and metadata and metadata.get('unique_planning_mats'):
        selected_mat = metadata['unique_planning_mats'][0]
        request.session['btr_selected_material'] = selected_mat
    if not selected_mat:
        selected_mat = 'AL_110X110'
    
    rolling_dir = Path(settings.BASE_DIR) / 'media' / 'rolling_plans'
    history_items = []
    seen_files = set()
    
    # Query DB campaign records associated with source files
    campaigns = BTRCampaign.objects.all().order_by('-created_at')
    for c in campaigns:
        fname = c.rolling_plan_file or c.output_file
        if fname and fname not in seen_files:
            fpath = rolling_dir / fname
            if fpath.is_file():
                seen_files.add(fname)
                bloom_cnt = getattr(c, 'rolling_plan_bloom_count', None)
                rm_q = getattr(c, 'rolling_plan_rm_qty', None)
                y_pct = getattr(c, 'rolling_plan_yield_pct', None)
                
                if not bloom_cnt or not rm_q or not y_pct or (c.total_btr_qty and rm_q == c.total_btr_qty):
                    m = get_rolling_plan_file_metrics(fpath)
                    if m:
                        bloom_cnt = m.get('bloom_count') or bloom_cnt
                        rm_q = m.get('rm_qty') or rm_q
                        y_pct = m.get('yield_pct') or y_pct
                        c.rolling_plan_bloom_count = bloom_cnt or 0
                        c.rolling_plan_rm_qty = rm_q or 0.0
                        c.rolling_plan_yield_pct = y_pct or 0.0
                        c.save(update_fields=['rolling_plan_bloom_count', 'rolling_plan_rm_qty', 'rolling_plan_yield_pct'])

                history_items.append({
                    'planning_mat': c.planning_material or 'BTR Plan',
                    'created_at_fmt': timezone.localtime(c.created_at).strftime('%d %b %Y, %I:%M %p') if c.created_at else '-',
                    'source_btr_file': c.source_file or 'BTR Report',
                    'bloom_count': bloom_cnt if bloom_cnt is not None and bloom_cnt != '' else '-',
                    'rm_qty': f"{rm_q:.2f}" if isinstance(rm_q, (int, float)) else (rm_q or '-'),
                    'yield_pct': f"{y_pct:.2f}%" if isinstance(y_pct, (int, float)) else (y_pct or '-'),
                    'is_generated': True,
                    'filename': fname,
                    'upload_id': c.id,
                    'plan_date': timezone.localtime(c.created_at).strftime('%Y-%m-%d') if c.created_at else date.today().isoformat(),
                })

    return render(request, 'casting_plan/btr_rolling_plan.html', {
        'metadata': metadata,
        'selected_mat': selected_mat,
        'planning_material': selected_mat,
        'history_items': history_items,
        'today': date.today().isoformat(),
        'selected_furnace': request.session.get('rolling_plan_furnace', 'new'),
    })


@login_required
@require_POST
def btr_rolling_plan_generate_view(request):
    """Step 4: Generate Rolling Plan Workbook"""
    from datetime import date
    from uuid import uuid4
    from django.utils import timezone
    from django.utils.text import get_valid_filename
    from .btr_service import load_btr_dataset, generate_camp_plan_records, create_rolling_campaign_workbook
    from .rolling_plan_services import generate_rolling_plan_workbook
    from .models import BTRCampaign

    upload_id = request.session.get('btr_upload_id')
    selected_mat = request.POST.get('planning_material') or request.session.get('btr_selected_material')
    if not upload_id or not selected_mat:
        messages.error(request, 'Please upload a BTR report and select a planning material first.')
        return redirect('btr_input')
        
    metadata, records = load_btr_dataset(upload_id)

    # Resolve the selected material against the original BTR records. Calendar
    # display labels and edited quantities must never become rolling-plan inputs.
    planning_col = (metadata or {}).get('planning_mat_col') or 'Planning Mat'
    source_materials = {}
    for source_row in records or []:
        source_name = source_row.get(planning_col) or source_row.get('Planning Mat') or source_row.get('Planning_Mat')
        source_name = str(source_name or '').strip()
        if source_name:
            source_materials[source_name.upper()] = source_name
    canonical_selected_mat = source_materials.get(str(selected_mat).strip().upper())
    if not canonical_selected_mat:
        messages.error(request, 'Selected planning material is not present in the active BTR report. Please select it again.')
        request.session.pop('btr_selected_material', None)
        return redirect('btr_planning_material')
    selected_mat = canonical_selected_mat
    request.session['btr_selected_material'] = selected_mat

    # A custom campaign workbook may only be reused for the material it was
    # uploaded for. Reusing another material's workbook produces shifted/wrong
    # rolling-plan columns even though the current material selection looks valid.
    custom_camp_path = request.session.get('custom_camp_plan_path')
    custom_camp_mat = request.session.get('custom_camp_plan_mat')
    custom_matches_selection = (
        custom_camp_mat and
        str(custom_camp_mat).strip().upper() == str(selected_mat).strip().upper()
    )
    if custom_camp_path and custom_matches_selection and Path(custom_camp_path).is_file():
        camp_path = Path(custom_camp_path)
    else:
        custom_camp_path = None
        result = generate_camp_plan_records(records or [], selected_mat)
        camp_path = create_rolling_campaign_workbook(result.get('valid_records', []), selected_mat)
    
    plan_date_val = request.POST.get('plan_date')
    planning_date = date.fromisoformat(plan_date_val) if plan_date_val else date.today()
    camp_num = request.POST.get('camp_number', '3915')
    furnace_type = str(request.POST.get('furnace_type') or 'new').strip().lower()
    if furnace_type not in {'old', 'new'}:
        messages.error(request, 'Please select either Old Furnace or New Furnace.')
        return redirect('btr_rolling_plan')
    request.session['rolling_plan_furnace'] = furnace_type

    casting_plan_path = None
    latest_casting_mat = str(request.session.get('latest_casting_plan_mat') or '').strip()
    latest_casting_path = request.session.get('latest_casting_plan_path')
    if (
        latest_casting_path
        and latest_casting_mat.upper() == str(selected_mat).strip().upper()
        and Path(latest_casting_path).is_file()
    ):
        casting_plan_path = Path(latest_casting_path)
    if casting_plan_path is None:
        casting_dir = Path(settings.BASE_DIR) / 'media' / 'casting_plans'
        casting_material = get_valid_filename(selected_mat.replace('/', '_')) or 'planning_material'
        casting_candidates = sorted(
            casting_dir.glob(f'CASTING_PLAN_{casting_material}_*.xlsx'),
            key=lambda item: item.stat().st_mtime,
            reverse=True,
        ) if casting_dir.is_dir() else []
        if casting_candidates:
            casting_plan_path = casting_candidates[0]
    
    grade_map_path = Path(settings.BASE_DIR) / 'data' / 'reference' / 'Grade details of ex grade and raw grade.xlsx'
    raw_material_map_path = Path(settings.BASE_DIR) / 'data' / 'reference' / 'RAW MATERIAL LENGTH TOLERANCE.xlsx'
    
    safe_material = get_valid_filename(selected_mat.replace('/', '_')) or 'planning_material'
    timestamp = timezone.localtime().strftime('%Y%m%d_%H%M%S')
    output_name = f'ROLLING_PLAN_{safe_material}_{timestamp}_{uuid4().hex[:6]}.xlsm'
    output_path = Path(settings.BASE_DIR) / 'media' / 'rolling_plans' / output_name
    
    res = generate_rolling_plan_workbook(
        camp_path, grade_map_path, raw_material_map_path, output_path,
        plan_date=planning_date, camp_number=camp_num,
        furnace_type=furnace_type, casting_plan_path=casting_plan_path,
    )

    # Clean up temporary auto-generated camp input workbook if created (do NOT delete custom_camp_path)
    if camp_path and not custom_camp_path and Path(camp_path).is_file():
        try:
            Path(camp_path).unlink()
        except Exception:
            pass
    
    BTRCampaign.objects.create(
        source_file=metadata.get('filename', 'BTR Report.xlsx') if metadata else 'BTR Report',
        planning_material=selected_mat,
        input_row_count=len(records) if records else 0,
        output_row_count=res.row_count,
        skipped_row_count=0,
        zero_length_resolved_count=0,
        total_btr_qty=res.rolling_qty,
        output_file=output_name,
        rolling_plan_file=output_name,
        casting_plan_file=output_name.replace("ROLLING_PLAN", "CASTING_PLAN").replace(".xlsm", ".xlsx"),
        embossing_plan_file=output_name.replace("ROLLING_PLAN", "EMBOSSING_PLAN").replace(".xlsm", ".xlsx") if selected_mat.upper().startswith("AL") else "",
        rolling_plan_bloom_count=res.bloom_count,
        rolling_plan_rm_qty=res.rm_qty,
        rolling_plan_yield_pct=res.yield_pct,
        created_by=request.user if request.user.is_authenticated else None,
    )
    furnace_label = 'Old Furnace' if furnace_type == 'old' else 'New Furnace'
    messages.success(request, f"Rolling plan generated successfully for {furnace_label}: {res.row_count} rows, {res.bloom_count} blooms.")
    return redirect('btr_rolling_plan')


@login_required
def btr_rolling_plan_download_view(request):
    """Download latest rolling plan"""
    filename = request.GET.get('filename')
    return btr_rolling_plan_download_file_view(request, filename)


@login_required
def btr_rolling_plan_download_file_view(request, filename):
    if not filename:
        raise Http404("File name required.")
    safe = get_valid_filename(filename)
    path = Path(settings.BASE_DIR) / 'media' / 'rolling_plans' / safe
    if not path.is_file():
        raise Http404("Requested file not found.")
    content_type = 'application/vnd.ms-excel.sheet.macroEnabled.12' if safe.endswith('.xlsm') else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return FileResponse(open(path, 'rb'), as_attachment=True, filename=safe, content_type=content_type)


@login_required
def btr_rolling_plan_delete_view(request, filename):
    from django.db.models import Q
    safe = get_valid_filename(filename) if filename else ''
    
    # 1. If filename parameter is a numeric ID, delete by primary key
    if str(filename).isdigit():
        BTRCampaign.objects.filter(id=int(filename)).delete()

    # 2. Delete database records matching filename in rolling_plan_file, output_file, or casting_plan_file
    if safe:
        matching_campaigns = BTRCampaign.objects.filter(
            Q(rolling_plan_file__icontains=safe) |
            Q(output_file__icontains=safe) |
            Q(casting_plan_file__icontains=safe)
        )
        for comp in matching_campaigns:
            for f_attr in [comp.rolling_plan_file, comp.output_file, comp.casting_plan_file]:
                if f_attr:
                    fpath = Path(settings.BASE_DIR) / 'media' / 'rolling_plans' / get_valid_filename(f_attr)
                    if fpath.is_file():
                        try:
                            fpath.unlink()
                        except Exception:
                            pass
        matching_campaigns.delete()

        # 3. Delete physical file from disk
        path = Path(settings.BASE_DIR) / 'media' / 'rolling_plans' / safe
        if path.is_file():
            try:
                path.unlink()
            except Exception:
                pass

    messages.success(request, "Rolling plan deleted successfully.")
    return redirect('btr_rolling_plan')


@login_required
def btr_rolling_plan_preview_api_view(request, filename):
    import openpyxl
    from .btr_service import sanitize_for_json
    safe = get_valid_filename(filename)
    path = Path(settings.BASE_DIR) / 'media' / 'rolling_plans' / safe
    if not path.is_file():
        return JsonResponse({'status': 'error', 'message': 'File not found'}, status=404)
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        sheets = {}
        sheet_names = wb.sheetnames
        for sheet_name in sheet_names:
            ws = wb[sheet_name]
            header_row = 8 if sheet_name == "Final Plan" else 1
            start_row = 9 if sheet_name == "Final Plan" else 2
            headers = [ws.cell(header_row, c).value for c in range(1, ws.max_column + 1)] if ws.max_row >= header_row else []
            data_rows = []
            for r in range(start_row, min(ws.max_row + 1, 100)):
                row_vals = [ws.cell(r, c).value for c in range(1, len(headers) + 1)]
                data_rows.append(row_vals)
            sheets[sheet_name] = {
                'headers': headers,
                'rows': data_rows
            }
        wb.close()
        return JsonResponse({
            'status': 'success',
            'filename': safe,
            'sheet_names': sheet_names,
            'sheets': sanitize_for_json(sheets)
        })
    except Exception as exc:
        return JsonResponse({'status': 'error', 'message': str(exc)}, status=500)


@login_required
@require_POST
def generate_btr_campaign(request):
    source_path = _session_btr_path(request)
    upload_info = request.session.get('btr_upload_info') or {}
    planning_material = request.POST.get('planning_material', '').strip()
    if not source_path:
        messages.error(request, 'Please upload a BTR file before generating a campaign.')
        return redirect('upload_btr')
    allowed_materials = {
        str(item.get('name', '')).strip()
        for item in upload_info.get('planning_materials', [])
    }
    if not planning_material or planning_material not in allowed_materials:
        messages.error(request, 'Please select a valid planning material from the uploaded BTR.')
        return redirect('upload_btr')

    reference_template = Path(settings.BASE_DIR) / 'data' / 'reference' / 'CAMP PLAN FORMAT.xlsm'
    generated_dir = Path(settings.BASE_DIR) / 'media' / 'btr_campaigns'
    safe_material = get_valid_filename(planning_material.replace('/', '_')) or 'planning_material'
    timestamp = timezone.localtime().strftime('%Y%m%d_%H%M%S')
    output_name = f'CAMP_{safe_material}_{timestamp}_{uuid4().hex[:6]}.xlsm'
    output_path = generated_dir / output_name
    try:
        result = generate_campaign_workbook(
            source_path,
            reference_template,
            output_path,
            planning_material,
        )
    except Exception as exc:
        if output_path.exists():
            output_path.unlink()
        messages.error(request, f'Campaign generation failed: {exc}')
        return redirect('upload_btr')

    campaign = BTRCampaign.objects.create(
        source_file=upload_info.get('original_name', source_path.name),
        planning_material=planning_material,
        input_row_count=result.input_row_count,
        output_row_count=result.output_row_count,
        skipped_row_count=result.skipped_row_count,
        zero_length_resolved_count=result.zero_length_resolved_count,
        total_btr_qty=result.total_btr_qty,
        bucket_counts=result.bucket_counts,
        output_file=result.output_path.name,
        created_by=request.user,
    )
    Notification.objects.create(
        title='BTR campaign generated',
        message=(
            f'{planning_material}: {result.output_row_count} sorted rows, '
            f'{result.total_btr_qty:.3f} MT BTR.'
        ),
    )
    messages.success(
        request,
        f'Campaign generated for {planning_material}: {result.output_row_count} rows, '
        f'{result.total_btr_qty:.3f} MT. Download it from Campaign History.'
    )
    request.session['btr_selected_material'] = planning_material
    return redirect('btr_camp_plan')


@login_required
def download_btr_campaign(request, campaign_id):
    campaign = get_object_or_404(BTRCampaign, id=campaign_id)
    path = Path(settings.BASE_DIR) / 'media' / 'btr_campaigns' / campaign.output_file
    if not path.is_file():
        raise Http404('Generated campaign file is not available.')
    return FileResponse(
        path.open('rb'),
        as_attachment=True,
        filename=campaign.output_file,
        content_type='application/vnd.ms-excel.sheet.macroEnabled.12',
    )


@login_required
@require_POST
def generate_rolling_plan(request, campaign_id):
    campaign = get_object_or_404(BTRCampaign, id=campaign_id)
    campaign_path = Path(settings.BASE_DIR) / 'media' / 'btr_campaigns' / campaign.output_file
    if not campaign_path.is_file():
        messages.error(request, 'Generate or restore the campaign workbook before creating its rolling plan.')
        return redirect('upload_btr')

    plan_date_value = request.POST.get('plan_date', '').strip()
    try:
        planning_date = date.fromisoformat(plan_date_value) if plan_date_value else date.today()
    except ValueError:
        messages.error(request, 'Please enter a valid rolling-plan date.')
        return redirect('upload_btr')
    camp_number = request.POST.get('camp_number', '').strip()[:40]

    grade_map_path = Path(settings.BASE_DIR) / 'data' / 'reference' / 'Grade details of ex grade and raw grade.xlsx'
    raw_material_map_path = Path(settings.BASE_DIR) / 'data' / 'reference' / 'RAW MATERIAL LENGTH TOLERANCE.xlsx'
    generated_dir = Path(settings.BASE_DIR) / 'media' / 'rolling_plans'
    safe_material = get_valid_filename(campaign.planning_material.replace('/', '_')) or 'planning_material'
    timestamp = timezone.localtime().strftime('%Y%m%d_%H%M%S')
    output_name = f'ROLLING_PLAN_{safe_material}_{timestamp}_{uuid4().hex[:6]}.xlsm'
    output_path = generated_dir / output_name
    try:
        result = generate_rolling_plan_workbook(
            campaign_path,
            grade_map_path,
            raw_material_map_path,
            output_path,
            plan_date=planning_date,
            camp_number=camp_number,
        )
    except Exception as exc:
        if output_path.exists():
            output_path.unlink()
        messages.error(request, f'Rolling-plan generation failed: {exc}')
        return redirect('upload_btr')

    old_name = campaign.rolling_plan_file
    old_casting = campaign.casting_plan_file
    
    campaign.rolling_plan_file = result.output_path.name
    # Dynamically generate the matching Casting Plan filename
    campaign.casting_plan_file = result.output_path.name.replace("ROLLING_PLAN", "CASTING_PLAN").replace(".xlsm", ".xlsx")

    if selected_mat.upper().startswith("AL"):
        campaign.embossing_plan_file = result.output_path.name.replace("ROLLING_PLAN", "EMBOSSING_PLAN").replace(".xlsm", ".xlsx")
    else:
        campaign.embossing_plan_file = ""
    
    campaign.rolling_plan_created_at = timezone.now()
    campaign.rolling_plan_rm_qty = result.rm_qty
    campaign.rolling_plan_bloom_count = result.bloom_count
    campaign.rolling_plan_yield_pct = result.yield_pct
    
    campaign.save(update_fields=[
        'rolling_plan_file', 'casting_plan_file', 'rolling_plan_created_at', 
        'rolling_plan_rm_qty', 'rolling_plan_bloom_count', 'rolling_plan_yield_pct',
    ])
    
    # Cleanup old files if regenerated
    if old_name and old_name != result.output_path.name:
        for old_file in [old_name, old_casting]:
            if old_file:
                old_path = Path(settings.BASE_DIR) / 'media' / 'rolling_plans' / Path(old_file).name
                if old_path.is_file():
                    old_path.unlink()

    Notification.objects.create(
        title='Rolling plan generated',
        message=(
            f'{campaign.planning_material}: {result.row_count} rows, '
            f'{result.bloom_count} blooms, {result.yield_pct:.2f}% yield.'
        ),
    )
    messages.success(
        request,
        f'Rolling plan generated: {result.row_count} rows, {result.bloom_count} blooms, '
        f'{result.max_cooling_bed_length:.2f} m maximum cooling-bed length.'
    )
    return redirect('upload_btr')


@login_required
def download_rolling_plan(request, campaign_id):
    campaign = get_object_or_404(BTRCampaign, id=campaign_id)
    if not campaign.rolling_plan_file:
        raise Http404('A rolling plan has not been generated for this campaign.')
    path = Path(settings.BASE_DIR) / 'media' / 'rolling_plans' / Path(campaign.rolling_plan_file).name
    if not path.is_file():
        raise Http404('Generated rolling-plan file is not available.')
    return FileResponse(
        path.open('rb'),
        as_attachment=True,
        filename=campaign.rolling_plan_file,
        content_type='application/vnd.ms-excel.sheet.macroEnabled.12',
    )

@login_required
def download_casting_plan(request, campaign_id):
    campaign = get_object_or_404(BTRCampaign, id=campaign_id)
    if not campaign.casting_plan_file:
        raise Http404('A casting plan has not been generated for this campaign.')
    
    # The casting plan is saved in the same directory as the rolling plan
    path = Path(settings.BASE_DIR) / 'media' / 'rolling_plans' / Path(campaign.casting_plan_file).name
    if not path.is_file():
        raise Http404('Generated casting-plan file is not available.')
        
    return FileResponse(
        path.open('rb'),
        as_attachment=True,
        filename=campaign.casting_plan_file,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )

@login_required
def download_embossing_plan(request, campaign_id):
    from .models import BTRCampaign
    from .btr_service import load_btr_dataset, generate_camp_plan_records
    from .embossing_plan import generate_embossing_plan_file

    campaign = BTRCampaign.objects.filter(id=campaign_id).first()
    
    # Determine planning material and filename
    if campaign:
        planning_mat = campaign.planning_material or 'AL_100X100'
        emb_filename = campaign.embossing_plan_file or f"EMBOSSING_PLAN_{get_valid_filename(planning_mat.replace('/', '_'))}.xlsx"
    else:
        planning_mat = request.session.get('btr_selected_material') or 'AL_100X100'
        emb_filename = f"EMBOSSING_PLAN_{get_valid_filename(planning_mat.replace('/', '_'))}.xlsx"

    rolling_dir = Path(settings.BASE_DIR) / 'media' / 'rolling_plans'
    rolling_dir.mkdir(parents=True, exist_ok=True)
    
    path = rolling_dir / Path(emb_filename).name

    # If embossing plan file doesn't exist on disk, generate it dynamically on the fly!
    if not path.is_file():
        upload_id = request.session.get('btr_upload_id')
        records = None
        if upload_id:
            _, records = load_btr_dataset(upload_id)
        
        camp_records = []
        if records and planning_mat:
            res = generate_camp_plan_records(records, planning_mat)
            camp_records = res.get('valid_records', [])

        generate_embossing_plan_file(camp_records, planning_mat, path)

    if not path.is_file():
        raise Http404(f'Embossing plan file for {planning_mat} could not be generated.')

    if campaign and not campaign.embossing_plan_file:
        campaign.embossing_plan_file = path.name
        campaign.save(update_fields=['embossing_plan_file'])

    return FileResponse(
        path.open('rb'),
        as_attachment=True,
        filename=path.name,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@login_required
@require_POST
def delete_btr_campaign(request, campaign_id):
    campaign = get_object_or_404(BTRCampaign, id=campaign_id)
    path = Path(settings.BASE_DIR) / 'media' / 'btr_campaigns' / campaign.output_file
    if path.is_file():
        path.unlink()
    if campaign.rolling_plan_file:
        rolling_path = Path(settings.BASE_DIR) / 'media' / 'rolling_plans' / Path(campaign.rolling_plan_file).name
        if rolling_path.is_file():
            rolling_path.unlink()
    material = campaign.planning_material
    campaign.delete()
    messages.success(request, f'Campaign history for {material} was deleted.')
    return redirect('upload_btr')


@login_required
@require_POST
def update_demand_selection(request):
    selected_ids = {int(x) for x in request.POST.getlist('selected_order') if str(x).isdigit()}
    visible_ids = {int(x) for x in request.POST.getlist('visible_order') if str(x).isdigit()}
    if visible_ids:
        Order.objects.filter(id__in=visible_ids).update(is_selected=False)
        if selected_ids:
            Order.objects.filter(id__in=selected_ids).update(is_selected=True)
    messages.success(request, f'{len(selected_ids)} demand rows selected for planning.')
    return redirect('data')


@login_required
@require_POST
def generate_plan(request):
    ids = [int(x) for x in request.POST.getlist('order_ids') if str(x).isdigit()]
    if not ids:
        qs = Order.objects.filter(is_selected=True)
        spm_section = request.POST.get('spm_section', '').strip()
        if spm_section:
            qs = qs.filter(section__iexact=spm_section)
        
        specific_so = request.POST.get('specific_so', '').strip()
        if specific_so:
            so_nos = [s.strip() for s in specific_so.split(',') if s.strip()]
            if so_nos:
                qs = qs.filter(so_no__in=so_nos)
                
        ids = list(qs.values_list('id', flat=True))
        
        if spm_section and not ids:
            base_dir = Path(__import__('django.conf').conf.settings.BASE_DIR)
            sec_dir = base_dir / 'data' / 'section_workbooks' / 'All Section'
            matches = list(sec_dir.glob(f"{spm_section}.*"))
            if not matches:
                matches = [p for p in sec_dir.glob("*") if p.stem.replace('_','').replace('-','').lower() == spm_section.replace('_','').replace('-','').lower()]
            if matches:
                try:
                    import_demand_workbook(matches[0], replace=False)
                    qs = Order.objects.filter(is_selected=True, section__iexact=spm_section)
                    if specific_so:
                        so_nos = [s.strip() for s in specific_so.split(',') if s.strip()]
                        if so_nos:
                            qs = qs.filter(so_no__in=so_nos)
                    ids = list(qs.values_list('id', flat=True))
                except Exception:
                    pass
        
    rows, selected = build_plan_rows(ids)
    if not rows:
        messages.error(request, 'No positive pending demand was available for the selected parameters/orders.')
        return redirect('optimizer')
        
    name = request.POST.get('plan_name') or f'Casting Plan {date.today().isoformat()}'
    output = generate_casting_plan_workbook(rows, name)
    demand = sum(r.get('demand_qty', 0) for r in rows)
    rm = sum(r['rm_qty'] for r in rows)
    yld = demand / rm * 100 if rm else 0
    GeneratedPlan.objects.create(
        name=name, planning_date=date.today(), selected_order_count=len(selected),
        demand_qty=round(demand,3), rm_qty=round(rm,3), yield_pct=round(yld,2), output_file=output.name
    )
    Notification.objects.create(title='Casting plan generated', message=f'{name}: {len(rows)} optimized casting lines created.')
    messages.success(request, f'Casting plan generated with {len(rows)} output lines.')
    return redirect('plan')


@login_required
def download_plan(request, plan_id):
    try:
        plan_obj = GeneratedPlan.objects.get(id=plan_id)
    except GeneratedPlan.DoesNotExist:
        try:
            plan_obj = CastingPlanRequest.objects.get(id=plan_id)
        except CastingPlanRequest.DoesNotExist:
            raise Http404('Plan not found')
            
    path = Path(__import__('django.conf').conf.settings.BASE_DIR) / 'media' / 'generated' / plan_obj.output_file
    if not path.exists():
        raise Http404('Generated file is not available')
    return FileResponse(open(path,'rb'), as_attachment=True, filename=path.name)


@login_required
@require_POST
def delete_plan(request, plan_id):
    try:
        plan_obj = GeneratedPlan.objects.get(id=plan_id)
        path = Path(__import__('django.conf').conf.settings.BASE_DIR) / 'media' / 'generated' / plan_obj.output_file
        if path.exists():
            try:
                path.unlink()
            except Exception:
                pass
        plan_obj.delete()
        messages.success(request, f"Casting plan '{plan_obj.name}' was successfully deleted.")
    except GeneratedPlan.DoesNotExist:
        try:
            req_obj = CastingPlanRequest.objects.get(id=plan_id)
            path = Path(__import__('django.conf').conf.settings.BASE_DIR) / 'media' / 'generated' / req_obj.output_file
            if path.exists():
                try:
                    path.unlink()
                except Exception:
                    pass
            req_obj.delete()
            messages.success(request, f"Casting plan '{req_obj.name}' deleted.")
        except CastingPlanRequest.DoesNotExist:
            messages.error(request, 'Casting plan not found.')
    return redirect('plan')


@login_required
def data_view(request):
    orders = Order.objects.all()[:500]
    summary = {
        'total': Order.objects.count(),
        'selected': Order.objects.filter(is_selected=True).count(),
        'net_demand': sum(o.net_demand for o in Order.objects.filter(is_selected=True)),
    }
    return render(request, 'casting_plan/data.html', {
        'orders': orders, 'summary': summary,
        'rules': SectionRule.objects.all()[:100], 'batches': BatchRule.objects.all()[:100]
    })


def health(request):
    return JsonResponse({'status': 'ok', 'service': 'SPM Casting Plan Optimizer'})


def login_view(request):
    from django.contrib.auth.models import User
    if request.user.is_authenticated:
        if request.session.get('is_sap_user') or request.user.username.lower() == 'sap':
            return redirect('sap_portal')
        return redirect('dashboard')
    error = None
    if request.method == 'POST':
        u = request.POST.get('username', '').strip()
        p = request.POST.get('password', '').strip()
        
        # Check SAP credentials fallback or authentic Django user
        if u.lower() == 'sap' and (p == 'sap123' or authenticate(request, username=u, password=p) is not None):
            sap_user = User.objects.filter(username__iexact='sap').first()
            if not sap_user:
                sap_user = User.objects.create_user(username='sap', password='sap123')
            elif p == 'sap123' and not sap_user.check_password('sap123'):
                sap_user.set_password('sap123')
                sap_user.save()
            auth_login(request, sap_user)
            request.session['is_sap_user'] = True
            request.session['sap_username'] = 'sap'
            messages.success(request, "Welcome to SAP ERP Portal, sap!")
            return redirect('sap_portal')
            
        user = authenticate(request, username=u, password=p)
        if user is not None:
            auth_login(request, user)
            if user.username.lower() == 'sap':
                request.session['is_sap_user'] = True
                request.session['sap_username'] = user.username
                return redirect('sap_portal')
            else:
                request.session['is_sap_user'] = False
                return redirect('/')
        error = 'Invalid username or password'
    return render(request, 'casting_plan/login.html', {'error': error})


def sap_login_view(request):
    """Redirect to unified login view."""
    return redirect('login')


def logout_view(request):
    request.session.pop('is_sap_user', None)
    request.session.pop('sap_username', None)
    auth_logout(request)
    return redirect('login')


@login_required
@require_POST
def mark_all_notifications_read(request):
    Notification.objects.filter(is_read=False).update(is_read=True)
    return JsonResponse({'status':'success'})


@login_required
def master_data_view(request):
    selected_tab = request.GET.get('tab', 'rm_calc')
    sec_wts = SectionWeightMap.objects.all().order_by('planning_section')
    rules = SectionRule.objects.all().order_by('section')
    grades = GradeMap.objects.all().order_by('ext_grade')
    batches = BatchRule.objects.all().order_by('section')
    yield_rules = YieldRule.objects.all().order_by('-is_active', 'planning_section')

    # New Rail Mill Master Data Queries
    from .models import (
        RailMillCalculationMaster,
        RailMillChemistryMaster,
        RailMillSectionGradeStandardMap,
        RailMillSecWiseMaster,
        RailMillSawTolerance,
        RailMillRMSectionWeight,
        RailMillPlanningMaterial
    )

    # Show every Calculation-sheet record. The previous [:500] slice hid
    # records 501 onward even though they were present in the database.
    rm_calc = RailMillCalculationMaster.objects.all().order_by('id')
    rm_chem = RailMillChemistryMaster.objects.all()[:200]
    rm_standards = RailMillSectionGradeStandardMap.objects.all()[:200]
    rm_secwise = RailMillSecWiseMaster.objects.all()[:200]
    rm_saw = RailMillSawTolerance.objects.all()
    rm_weights = RailMillRMSectionWeight.objects.all()
    rm_materials = RailMillPlanningMaterial.objects.all()

    return render(request, 'casting_plan/master_data.html', {
        'selected_tab': selected_tab,
        'sec_wts': sec_wts,
        'rules': rules,
        'grades': grades,
        'batches': batches,
        'yield_rules': yield_rules,
        'rm_calc': rm_calc,
        'rm_chem': rm_chem,
        'rm_standards': rm_standards,
        'rm_secwise': rm_secwise,
        'rm_saw': rm_saw,
        'rm_weights': rm_weights,
        'rm_materials': rm_materials,
        'counts': {
            'rm_calc': RailMillCalculationMaster.objects.count(),
            'rm_chem': RailMillChemistryMaster.objects.count(),
            'rm_standards': RailMillSectionGradeStandardMap.objects.count(),
            'rm_secwise': RailMillSecWiseMaster.objects.count(),
            'rm_materials': RailMillPlanningMaterial.objects.count(),
        }
    })



@login_required
def validation_report_view(request):
    base_dir = Path(__import__('django.conf').conf.settings.BASE_DIR)
    report_file = base_dir / "Casting_Plan_Validation_Report.txt"
    content = ""
    if report_file.exists():
        with open(report_file, "r", encoding="utf-8") as f:
            content = f.read()
    else:
        content = "Validation report not generated yet. Please run the automated test suite to produce the validation report: python manage.py test casting_plan"
    return render(request, 'casting_plan/validation_report.html', {'content': content})


@login_required
def settings_view(request):
    settings_obj, _ = BilletLengthSetting.objects.get_or_create(id=1)
    
    if request.method == 'POST':
        try:
            # We MUST explicitly convert the incoming POST strings into floats and ints
            settings_obj.cold_saw_cut = float(request.POST.get('cold_saw_cut', 0.0) or 0.0)
            settings_obj.cooling_bed_crop = float(request.POST.get('cooling_bed_crop', 0.4) or 0.4)
            settings_obj.div_sheer_loss = float(request.POST.get('div_sheer_loss', 0.0) or 0.0)
            settings_obj.base_crop_loss = float(request.POST.get('base_crop_loss', 0.0) or 0.0)
            settings_obj.section_wt_tolerance_pct = float(request.POST.get('section_wt_tolerance_pct', 0.5) or 0.5)
            settings_obj.hot_pieces_at_cb = int(request.POST.get('hot_pieces_at_cb', 1) or 1)
            settings_obj.save()
            messages.success(request, 'Billet Length Formula settings updated successfully.')
        except ValueError:
            messages.error(request, 'Invalid number format provided in settings.')
            
        return redirect('settings')
        
    return render(request, 'casting_plan/settings.html', {'settings': settings_obj})
