import time
import random
from datetime import datetime
import uuid
from models import db, Teacher, Classroom, Section, Subject, Setting, Schedule, ScheduleRun

def time_to_min(t_str):
    h, m = map(int, t_str.split(':'))
    return h * 60 + m

def min_to_time(m):
    return f"{m//60:02d}:{m%60:02d}"

def generate_schedule(phase='all', progress_callback=None, stop_check=None, diagnostic_mode=False, shuffle_sections=False):
    def update_progress(p, msg):
        if progress_callback:
            progress_callback(p, msg)
            
    def is_stopped():
        if stop_check and stop_check():
            return True
        return False

    start_time_proc = time.time()
    total_timeout = 600 # 10 minutes total limit
    update_progress(5, f"Initializing {phase.upper()} Data...")
    teachers = list(Teacher.query.all())
    classrooms = list(Classroom.query.all())
    sections = list(Section.query.all())
    subjects = list(Subject.query.all())
    settings = {s.key: s.value for s in Setting.query.all()}
    days = settings.get('active_days', 'Monday,Tuesday,Wednesday,Thursday,Friday').split(',')
    
    # PHASE FILTERING: Determine which sections we are generating for
    if phase == 'jhs':
        target_sections = [s for s in sections if s.department == 'JHS']
    elif phase == 'shs':
        target_sections = [s for s in sections if s.department == 'SHS']
    else:
        target_sections = sections

    # LOCK RULE: Load existing JHS schedules if generating SHS phase
    locked_jhs = []
    if phase == 'shs':
        active_run = ScheduleRun.query.filter_by(is_active=True).first()
        if active_run:
            locked_jhs = [s for s in Schedule.query.filter_by(run_id=active_run.id).all() if s.section.department == 'JHS']

    tba_teacher = Teacher.query.filter_by(name="TBA").first()
    if not tba_teacher:
        tba_teacher = Teacher(name="TBA", department="Both", grade_levels="All JHS", subjects="", max_hours_per_day=0, stay_window_hours=0)
        db.session.add(tba_teacher)
        db.session.commit()
    # Note: tba_teacher is NOT added to 'teachers' list to prevent silent fallback
    
    # Pre-calculate break slots
    sh_breaks = {
        'JHS_AM': (time_to_min(settings.get('jhs_am_break_start', '09:00')), time_to_min(settings.get('jhs_am_break_end', '09:30'))),
        'JHS_PM': (time_to_min(settings.get('jhs_pm_break_start', '15:00')), time_to_min(settings.get('jhs_pm_break_end', '15:30'))),
        'SHS_BREAK': (time_to_min(settings.get('shs_break_start', '09:30')), time_to_min(settings.get('shs_break_end', '10:00'))),
        'SHS_LUNCH': (time_to_min(settings.get('shs_lunch_start', '12:00')), time_to_min(settings.get('shs_lunch_end', '13:00'))),
        'JHS_AM_SPECIAL': (time_to_min(settings.get('jhs_am_special_break_start', '08:40')), time_to_min(settings.get('jhs_am_special_break_end', '09:10'))),
        'JHS_PM_SPECIAL': (time_to_min(settings.get('jhs_pm_special_break_start', '14:40')), time_to_min(settings.get('jhs_pm_special_break_end', '15:10'))),
        'SHS_SPECIAL': (time_to_min(settings.get('shs_special_break_start', '09:30')), time_to_min(settings.get('shs_special_break_end', '10:00')))
    }

    def get_section_shift(dept, grade_level):
        if dept == 'SHS': return 'FULL_DAY'
        gl = str(grade_level).upper().replace('GRADE', '').strip()
        is_am = settings.get(f'jhs_am_grade_{gl}') in ['active', 'on']
        is_pm = settings.get(f'jhs_pm_grade_{gl}') in ['active', 'on']
        if is_am and is_pm: return 'FULL_DAY'
        elif is_pm and not is_am: return 'PM'
        else: return 'AM'

    # Pre-allocate dynamic home rooms based on strict shift validation
    room_shift_owners = {r.id: set() for r in classrooms}
    
    # 1. Register static room owners
    for s in sections:
        if s.room_id and s.room_id in room_shift_owners:
            room_shift_owners[s.room_id].add(get_section_shift(s.department, s.grade_level))
            
    # 2. Assign fallback rooms to target sections missing a room
    for sec in target_sections:
        if not sec.room_id:
            sec_shift = get_section_shift(sec.department, sec.grade_level)
            for r in classrooms:
                if r.room_type == 'Room' and r.building in [sec.department, 'Both']:
                    owners = room_shift_owners[r.id]
                    if sec_shift == 'FULL_DAY' or 'FULL_DAY' in owners: continue
                    if sec_shift in owners: continue
                    # Room is available for this shift
                    sec.room_id = r.id
                    owners.add(sec_shift)
                    break

    def gl_match(req_gl, target_gl_str):
        if not target_gl_str: return True
        def norm(g): 
            if g is None: return ""
            return str(g).strip().upper().replace('GRADE', '').strip()
        req = norm(req_gl)
        if not req: return True
        targets = [norm(g) for g in str(target_gl_str).split(',') if norm(g)]
        
        if "ALL JHS" in str(target_gl_str).upper() and req in ['7', '8', '9', '10']: return True
        if "ALL SHS" in str(target_gl_str).upper() and req in ['11', '12']: return True
        
        # Flex match: check if req is part of any target (e.g. '11' matches '11STEM')
        for t in targets:
            if req == t or (len(req) >= 1 and req in t):
                return True
        return False

    def track_match(req_track, target_track_str):
        if not target_track_str or str(target_track_str).strip().upper() in ['NONE', '', 'ALL']: return True
        if not req_track: return False
        req = str(req_track).strip().upper()
        targets = [t.strip().upper() for t in str(target_track_str).split(',') if t.strip()]
        return req in targets

    def get_section_day_config(sec, day):
        # Check for Special Mode (e.g. Friday tweaks)
        dept = sec.department
        spec_enabled = settings.get(f'{dept.lower()}_special_enabled') == 'yes'
        spec_days = settings.get(f'{dept.lower()}_special_days', '').split(',')
        is_special_day = spec_enabled and day in spec_days
        
        # Base shift config
        if sec.department == 'SHS':
            s_start = time_to_min(settings.get('shs_start', '07:00'))
            s_end = time_to_min(settings.get('shs_end', '17:00'))
            if is_special_day:
                breaks = [sh_breaks['SHS_SPECIAL'], sh_breaks['SHS_LUNCH']]
            else:
                breaks = [sh_breaks['SHS_BREAK'], sh_breaks['SHS_LUNCH']]
        else:
            # Consistent JHS shift detection: default to AM unless PM-only is confirmed
            gl_clean = str(sec.grade_level).upper().replace('GRADE', '').strip()
            is_am_sec = settings.get(f'jhs_am_grade_{gl_clean}') in ['active', 'on']
            is_pm_sec = settings.get(f'jhs_pm_grade_{gl_clean}') in ['active', 'on']
            
            if is_pm_sec and not is_am_sec:
                s_start = time_to_min(settings.get('jhs_pm_start', '12:00'))
                s_end = time_to_min(settings.get('jhs_pm_end', '18:00'))
                breaks = [sh_breaks['JHS_PM_SPECIAL'] if is_special_day else sh_breaks['JHS_PM']]
            else:
                s_start = time_to_min(settings.get('jhs_am_start', '06:00'))
                s_end = time_to_min(settings.get('jhs_am_end', '12:00'))
                breaks = [sh_breaks['JHS_AM_SPECIAL'] if is_special_day else sh_breaks['JHS_AM']]
        spec_dur = int(settings.get(f'{dept.lower()}_special_duration', '40')) if spec_enabled else None
        extra_sub = settings.get(f'{dept.lower()}_special_extra_subject', '').strip() if spec_enabled else None
        
        return {
            'start': s_start, 'end': s_end, 'breaks': breaks,
            'is_special': is_special_day, 'spec_dur': spec_dur, 'extra_sub': extra_sub
        }

    # Prepare Section-Subject Requirements
    sections_requirements = {}
    for sec in target_sections:
        reqs = []
        dept = sec.department
        spec_extra_name = settings.get(f'{dept.lower()}_special_extra_subject', '').strip()
        
        for sub in subjects:
            if sub.department == dept:
                # Normal match
                if gl_match(sec.grade_level, sub.grade_level):
                    if dept != 'SHS' or track_match(sec.track, sub.track):
                        # Filter out the "Extra" subject from regular pool if it exists
                        if sub.name.strip() != spec_extra_name:
                            reqs.append(sub)
        
        # Add the Extra subject separately so we can handle its unique placement
        sections_requirements[sec.id] = reqs
        
    # --- VIRTUAL SUBJECT PROVISIONING ---
    # Ensure Special subjects (from settings) exist in DB so they have IDs
    # Since we added is_system, we can hide them from UI later
    for dept_code in ['jhs', 'shs']:
        sp_enabled = settings.get(f'{dept_code}_special_enabled') == 'yes'
        sp_name = settings.get(f'{dept_code}_special_extra_subject', '').strip()
        if sp_enabled and sp_name:
            existing = Subject.query.filter_by(name=sp_name, department=dept_code.upper()).first()
            if not existing:
                new_sub = Subject(
                    name=sp_name, 
                    department=dept_code.upper(), 
                    duration_mins=int(settings.get(f'{dept_code}_special_duration', '40')),
                    meetings_per_week=0, # Special only
                    is_system=True
                )
                db.session.add(new_sub)
                db.session.commit()
                subjects.append(new_sub) # Add to our local list
            elif existing not in subjects:
                subjects.append(existing)
            if existing not in subjects:
                subjects.append(existing)

    best_solution = None
    best_failed_count = 999
    best_details = []

    # --- Backtracking Logic ---
    def solve_day_sequence(sec, day, cursor, remaining_subs, t_busy, r_busy, cfg, sub_to_teachers, start_time, target_order=None, diagnostics=None):
        if is_stopped(): return None
        if not remaining_subs: return []
        
        # --- HARD BREAK ENFORCEMENT ---
        # If cursor is inside a break, jump to the end of the break
        # Use a loop to handle potential back-to-back breaks
        jumped = True
        while jumped:
            jumped = False
            for b_start, b_end in cfg['breaks']:
                if b_start <= cursor < b_end:
                    cursor = b_end
                    jumped = True
                    break

        # 1. FEASIBILITY & PRUNING
        total_dur = sum(cfg['spec_dur'] if cfg['is_special'] else s.duration_mins for s in remaining_subs)
        if cursor + total_dur > cfg['end']:
            if diagnostics is not None:
                diagnostics.update({
                    'classification': "Time Window Saturation",
                    'remaining_shift_cap': cfg['end'] - cursor,
                    'required_duration': total_dur
                })
            return None

        # PERFORMANCE LOCK: 1.0s limit per section-day
        if time.time() - start_time > 1.0: return None

        # --- SUBJECT REORDERING (Search across remaining pool) ---
        # MRV-inspired ordering for the search branch
        counts = {}
        for s in remaining_subs:
            # We treat subjects with fewer teacher options as higher priority
            weight = target_order.index(s.id) / 100.0 if (target_order and s.id in target_order) else 0
            c = len(sub_to_teachers.get(s.id, {})) + weight
            if c not in counts: counts[c] = []
            counts[c].append(s)
        
        idx_list = []
        for c in sorted(counts.keys()):
            for s in counts[c]: idx_list.append(remaining_subs.index(s))

        # 2. ATTEMPT PLACEMENT AT CURRENT CURSOR
        for i in idx_list:
            sub = remaining_subs[i]
            dur = cfg['spec_dur'] if cfg['is_special'] else sub.duration_mins
            
            # HARD BREAK OVERLAP CHECK: Subject cannot intersect any break
            overlaps_break = any(not (cursor + dur <= b_s or cursor >= b_e) for b_s, b_e in cfg['breaks'])
            if overlaps_break: 
                continue # Cannot place here, must allow cursor to eventually shift past this block

            c_teachers = sorted(list(sub_to_teachers.get(sub.id, [])), key=lambda x: len(t_busy[x.id][day]))
            
            # Room selection
            if sub.requires_lab:
                c_rooms = [r for r in classrooms if r.room_type == 'Laboratory' and r.building in [sec.department, 'Both']]
            else:
                # Use strictly the assigned room (either static or dynamically assigned earlier)
                c_rooms = [r for r in classrooms if r.id == sec.room_id] if sec.room_id else []

            slots = set(range(cursor, cursor + dur, 5))
            for t in c_teachers:
                if not slots.isdisjoint(t_busy[t.id][day]): continue
                
                # Hard Constraints
                load_ok = (len(t_busy[t.id][day]) * 5 + dur) <= (t.max_hours_per_day * 60)
                stay_ok = True
                if t_busy[t.id][day]:
                    f_s = min(t_busy[t.id][day])
                    l_e = max(t_busy[t.id][day]) + 5
                    stay_ok = (max(l_e, cursor + dur) - min(f_s, cursor)) <= (t.stay_window_hours * 60)
                
                if not (load_ok and stay_ok): continue
                
                for r in c_rooms:
                    if not slots.isdisjoint(r_busy[r.id][day]): continue

                    # Success: Recurse
                    t_busy[t.id][day].update(slots)
                    r_busy[r.id][day].update(slots)
                    
                    res = solve_day_sequence(sec, day, cursor + dur, remaining_subs[:i] + remaining_subs[i+1:], t_busy, r_busy, cfg, sub_to_teachers, start_time, target_order, diagnostics)
                    if res is not None:
                        return [{
                            'section_id': sec.id, 'subject_id': sub.id, 'teacher_id': t.id, 'room_id': r.id, 
                            'day': day, 'start': min_to_time(cursor), 'end': min_to_time(cursor + dur),
                            'is_soft_break_override': False # Now always False as breaks are hard
                        }] + res
                    
                    t_busy[t.id][day].difference_update(slots)
                    r_busy[r.id][day].difference_update(slots)

        # 3. SLIDING CURSOR (Time Shifting)
        # Saturated Shift Detection: Only shift if there is at least 5 mins of slack remaining
        # total_dur is already calculated above
        slack = (cfg['end'] - cursor) - total_dur
        
        if slack >= 5:
            new_cursor = cursor + 5
            return solve_day_sequence(sec, day, new_cursor, remaining_subs, t_busy, r_busy, cfg, sub_to_teachers, start_time, target_order, diagnostics)
        else:
            # ZERO-SLACK SATURATION: We hit a resource wall with no time left to shift
            if diagnostics is not None:
                if not diagnostics.get('classification'):
                    diagnostics.update({
                        'classification': "Zero-Slack Saturation",
                        'remaining_shift_cap': cfg['end'] - cursor,
                        'required_duration': total_dur
                    })
            return None

    # Pre-map teachers per (subject_id, grade_level)
    teachers_by_sub_gl = {}
    for sub in subjects:
        teachers_by_sub_gl[sub.id] = {}
        for gl_val in ['7', '8', '9', '10', '11', '12']:
            # Eligibility Rule:
            # For JHS phase: Anyone qualified for JHS
            # For SHS phase: Anyone qualified for SHS OR flagged as Hybrid
            # Subject expertise is a hard requirement for all candidates
            subject_qualified = [t for t in teachers if sub.name.strip() in [s_name.strip() for s_name in (t.subjects or '').split(',')]]
            
            if sub.department == 'SHS':
                # For SHS: Must be qualified for SHS OR marked as Hybrid
                # If they are 'Both' or 'Hybrid', they are eligible even if gl_match fails (as fallback)
                eligible = [t for t in subject_qualified if (gl_match(gl_val, t.grade_levels) or t.department == 'Both' or getattr(t, 'is_hybrid', False)) and (t.department in ['SHS', 'Both'] or getattr(t, 'is_hybrid', False))]
            else:
                # For JHS: Must be qualified for JHS OR 'Both'
                eligible = [t for t in subject_qualified if (gl_match(gl_val, t.grade_levels) or t.department == 'Both') and (t.department in ['JHS', 'Both'])]
            
            # Removed TBA fallback
            teachers_by_sub_gl[sub.id][gl_val] = eligible

    # Single Deterministic Pass (Minimum Working Model)
    max_attempts = 1
    for attempt in range(max_attempts):
        if is_stopped():
            update_progress(100, "Generation cancelled by user.")
            break
        
        # Consistent progress tracking
        prog = 10
        elapsed = int(time.time() - start_time_proc)
        msg = f"Stabilization Pass: Placing subjects (Phase: {phase.upper()})"
        if progress_callback:
            progress_callback(prog, f"{msg} | time:{elapsed}")
            
        current_sch, current_failed = [], []
        t_busy = {t.id: {d: set() for d in days} for t in teachers}
        r_busy = {r.id: {d: set() for d in days} for r in classrooms}

        # LOCK RULE: Pre-populate busy slots from locked JHS run if in SHS phase
        if phase == 'shs' and locked_jhs:
            for l_s in locked_jhs:
                if l_s.teacher_id in t_busy and l_s.day_of_week in days:
                    s_m, e_m = time_to_min(l_s.start_time), time_to_min(l_s.end_time)
                    t_busy[l_s.teacher_id][l_s.day_of_week].update(range(s_m, e_m, 5))
                if l_s.room_id in r_busy and l_s.day_of_week in days:
                    s_m, e_m = time_to_min(l_s.start_time), time_to_min(l_s.end_time)
                    r_busy[l_s.room_id][l_s.day_of_week].update(range(s_m, e_m, 5))
        
        def get_hierarchy_keys(sec):
            gl = str(sec.grade_level).upper().replace('GRADE', '').strip()
            # 1. SHIFT (Highest Container) - JHS AM/PM lookup
            is_pm = settings.get(f'jhs_pm_grade_{gl}') in ['active', 'on']
            shift_rank = 1 if is_pm else 0
            
            # 2. DOMAIN (JHS before SHS within same shift)
            domain_rank = 0 if sec.department == 'JHS' else 1
            
            # 3. BLOCK (Balanced Static: A=7-8, B=9-10 | A=11, B=12)
            if sec.department == 'JHS':
                block_rank = 0 if gl in ['7', '8'] else 1
            else: # SHS
                block_rank = 0 if gl == '11' else 1
            
            # 4. PRIORITY (Section A first)
            priority_rank = 0 if sec.is_section_a else 1
            
            return (shift_rank, domain_rank, block_rank, priority_rank, sec.id)

        # Apply deterministic hierarchical hierarchy
        shuffled_sections = sorted(target_sections, key=get_hierarchy_keys)
        if shuffle_sections:
            random.shuffle(shuffled_sections)
            
        # Note: We do NOT random.shuffle at the section level anymore by default to maintain the strict block hierarchy requested.
        section_target_order = {s.id: None for s in target_sections}
        for sec in shuffled_sections:
            sec_subs = sections_requirements[sec.id]
            sec_teachers_map = {}
            for s in sec_subs:
                eligible = teachers_by_sub_gl[s.id][str(sec.grade_level)]
                if eligible:
                    # Sort eligible teachers by least busy overall to maintain balance
                    def teacher_rank(t):
                        busy_slots = sum(len(t_busy[t.id][d]) for d in days)
                        # PRIORITY RULE: Exact grade-level matches come first.
                        # Dual-qualified faculty acting as fallbacks receive a large penalty to their rank.
                        penalty = 0 if gl_match(str(sec.grade_level), t.grade_levels) else 100000
                        return busy_slots + penalty + random.uniform(0, 0.1)
                    
                    sec_teachers_map[s.id] = sorted(eligible, key=teacher_rank)
                else:
                    sec_teachers_map[s.id] = []
            
            for d in days:
                cfg = get_section_day_config(sec, d)
                daily_pool = []
                for s in sec_subs:
                    freq = s.meetings_per_week
                    is_on_day = False
                    if freq >= 5: is_on_day = True
                    elif freq == 4: is_on_day = (d in ['Monday', 'Tuesday', 'Wednesday', 'Thursday'])
                    elif freq == 3: is_on_day = (d in ['Monday', 'Wednesday', 'Friday'])
                    elif freq == 2: is_on_day = (d in ['Tuesday', 'Thursday'])
                    elif freq == 1: is_on_day = (d == 'Wednesday')
                    
                    if is_on_day: daily_pool.append(s)
                
                if cfg['is_special'] and cfg['extra_sub']:
                    extra_obj = next((s for s in subjects if s.name.strip() == cfg['extra_sub'] and s.department == sec.department), None)
                    if extra_obj:
                        pos = settings.get(f'{sec.department.lower()}_special_position', 'first')
                        if pos == 'first': daily_pool.insert(0, extra_obj)
                        else: daily_pool.append(extra_obj)
                        
                        force_t = settings.get(f'{sec.department.lower()}_special_teacher', 'adviser')
                        if force_t == 'adviser':
                            adv = Teacher.query.get(sec.adviser_id) if sec.adviser_id else None
                            sec_teachers_map[extra_obj.id] = [adv] if adv else []
                        else:
                            target_t = Teacher.query.get(int(force_t)) if force_t.isdigit() else None
                            sec_teachers_map[extra_obj.id] = [target_t] if target_t else []

                if section_target_order.get(sec.id):
                    t_order = section_target_order[sec.id]
                    sp_pos = settings.get(f'{sec.department.lower()}_special_position', 'first')
                    def sort_priority(s):
                        if cfg['is_special'] and s.is_system:
                            return -1 if sp_pos == 'first' else 10000
                        return t_order.index(s.id) if s.id in t_order else 999
                    daily_pool.sort(key=sort_priority)
                else:
                    # BOTTLENECK-FIRST (MCV) HEURISTIC:
                    # Sort by teacher scarcity (ascending) when generating the first-day master sequence.
                    # This ensures single-expert subjects are 'locked in' before generalist slots occupy the shift.
                    gl_str = str(sec.grade_level)
                    def get_scarcity(s_id):
                        return len(teachers_by_sub_gl.get(s_id, {}).get(gl_str, []))
                    
                    # Deterministic Tie-breaking (Section-Day Rotating Priority)
                    rng = random.Random(sec.id + hash(d))
                    daily_pool.sort(key=lambda s: (get_scarcity(s.id), rng.random()))

                # STABILIZATION: Single Linear Backtracking Pass
                diag = {} if diagnostic_mode else None
                day_res = solve_day_sequence(sec, d, cfg['start'], daily_pool, t_busy, r_busy, cfg, sec_teachers_map, time.time(), section_target_order.get(sec.id), diag)
                
                if day_res is not None: 
                    current_sch.extend(day_res)
                    if section_target_order.get(sec.id) is None:
                        day_res_sorted = sorted(day_res, key=lambda x: time_to_min(x['start']))
                        section_target_order[sec.id] = [r['subject_id'] for r in day_res_sorted]
                else: 
                    if diagnostic_mode:
                        # REPORT AND HALT
                        diag.update({
                            'section_id': sec.id, 'section_name': sec.name, 'day': d,
                            'remaining_subs': [s.name for s in daily_pool],
                            'last_success': current_sch[-1]['start'] if current_sch else "None"
                        })
                        import json
                        if progress_callback:
                            progress_callback(prog, f"[DIAGNOSTIC] {json.dumps(diag)}")
                        raise Exception("DIAGNOSTIC_STOP")
                    
                    current_failed.append(f"{sec.name} on {d}: Infeasible ({len(daily_pool)} subs)")

        if len(current_failed) < best_failed_count:
            best_solution, best_failed_count, best_details = current_sch, len(current_failed), current_failed
            if best_failed_count == 0: break

    # Save Run
    update_progress(90, "Finalizing and saving to database...")
    school_year = settings.get('school_year', 'S.Y. 2024-2025')
    ScheduleRun.query.update({ScheduleRun.is_active: False})
    
    duration = round(time.time() - start_time_proc, 4)
    new_run = ScheduleRun(
        duration=duration, conflicts=best_failed_count, 
        conflict_log=' | '.join(best_details[:50]), is_active=True, school_year=school_year
    )
    db.session.add(new_run)
    db.session.flush()
    
    for sch in (best_solution or []):
        db.session.add(Schedule(
            section_id=sch['section_id'], subject_id=sch['subject_id'], 
            teacher_id=sch['teacher_id'], room_id=sch['room_id'], 
            day_of_week=sch['day'], start_time=sch['start'], 
            end_time=sch['end'], run_id=new_run.id,
            is_soft_break_override=sch.get('is_soft_break_override', False)
        ))
    
    # LOCK RULE: Preserve JHS assignments in the new unified run if in SHS phase
    if phase == 'shs' and locked_jhs:
        for l_s in locked_jhs:
            db.session.add(Schedule(
                section_id=l_s.section_id, subject_id=l_s.subject_id,
                teacher_id=l_s.teacher_id, room_id=l_s.room_id,
                day_of_week=l_s.day_of_week, start_time=l_s.start_time,
                end_time=l_s.end_time, run_id=new_run.id
            ))
    
    db.session.commit()
    update_progress(100, "Generation Complete!")
    return (best_failed_count == 0), duration, best_failed_count, "Success" if best_failed_count == 0 else "Partial"

