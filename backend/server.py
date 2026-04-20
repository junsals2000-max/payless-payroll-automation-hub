"""
Payless Automation Hub — Flask Backend
Payless Kitchen Cabinets & Bath Makeover
Central server for all automation modules.
"""

from __future__ import annotations
import sys
from pathlib import Path

# ── Path setup ─────────────────────────────────────────────────────────────────
BACKEND_DIR = Path(__file__).parent          # automation-hub/backend/
HUB_DIR     = BACKEND_DIR.parent            # automation-hub/
DATA_DIR    = HUB_DIR / 'data'
DATA_DIR.mkdir(exist_ok=True)

sys.path.insert(0, str(BACKEND_DIR))        # so database/auth import cleanly

from flask import Flask, jsonify, request, Response, send_from_directory, g
from flask_cors import CORS
import json, threading, asyncio, uuid, subprocess, time, os, io, zipfile, tempfile, base64, re as _re
import shutil, traceback
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash

from database import (
    init_db, get_user_by_username, get_user_by_id,
    get_all_users, create_user, update_user, delete_user,
    get_user_setting, set_user_setting, migrate_commission_config,
)
from auth import create_token, require_auth, require_admin

# ── Constants ──────────────────────────────────────────────────────────────────

HISTORY_FILE    = DATA_DIR / 'run_history.json'
COMM_EST_DIR         = HUB_DIR / 'automations' / 'commission_estimate'
ONSITE_RECON_DIR     = HUB_DIR / 'automations' / 'onsite_recon'
ONSITE_CONFIG_PATH   = DATA_DIR / 'onsite_recon_config.json'
BANK_CAT_DIR         = HUB_DIR / 'automations' / 'bank_categorizer'
BANK_CAT_CONFIG_PATH = DATA_DIR / 'bank_categorizer_config.json'
OLD_COMM_CONFIG = HUB_DIR / 'commission_config.json'   # legacy — migrated on first run

CDP_PORT            = 9222
CHROME_PATH         = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
CHROME_SRC_DIR      = Path.home() / 'Library' / 'Application Support' / 'Google' / 'Chrome'
CHROME_SRC_PROFILE  = 'Profile 2'
CHROME_AUTO_DIR     = str(Path.home() / '.payless-automation-chrome')

# ── Flask app ──────────────────────────────────────────────────────────────────

app = Flask(__name__, static_folder=str(HUB_DIR))
CORS(app, supports_credentials=True)

jobs         = {}   # job_id → queue.Queue
cancel_flags = {}   # job_id → threading.Event
job_loops    = {}   # job_id → asyncio.AbstractEventLoop


# ── Startup ────────────────────────────────────────────────────────────────────

def startup():
    init_db()
    admin = next((u for u in get_all_users() if u['is_admin']), None)
    if admin:
        migrate_commission_config(OLD_COMM_CONFIG, admin['id'])


# ── Helpers ────────────────────────────────────────────────────────────────────

def load_history():
    if HISTORY_FILE.exists():
        return json.loads(HISTORY_FILE.read_text())
    return []


def save_history(entry: dict):
    history = load_history()
    history.insert(0, entry)
    HISTORY_FILE.write_text(json.dumps(history[:50], indent=2))


def _load_json_config(path: Path, default=None):
    """Load a JSON config file from disk. Returns `default` if the file doesn't exist.
    Keys that start with '_' (comments/metadata) are stripped automatically."""
    if default is None:
        default = {}
    if not path.exists():
        return default
    raw = json.loads(path.read_text())
    return {k: v for k, v in raw.items() if not k.startswith('_')}


def _save_json_config(path: Path, data: dict):
    """Write a dict to a JSON config file on disk."""
    path.write_text(json.dumps(data, indent=2))


def _run_script(cmd: list, timeout: int = 120):
    """Run an automation script in a subprocess.
    Returns (result, error_message). On success error_message is None.
    On failure result is None and error_message contains the reason."""
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        if result.returncode != 0:
            error = (result.stderr or result.stdout or 'Processing failed.').strip()
            return None, error
        return result, None
    except subprocess.TimeoutExpired:
        return None, f'Processing timed out (>{timeout}s).'


def cdp_is_alive() -> bool:
    import requests as req
    try:
        req.get(f'http://localhost:{CDP_PORT}/json/version', timeout=1)
        return True
    except Exception:
        return False


# ── Static files ───────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return send_from_directory(str(HUB_DIR), 'index.html')


@app.route('/<path:filename>')
def static_files(filename):
    return send_from_directory(str(HUB_DIR), filename)


# ── Auth endpoints (public) ────────────────────────────────────────────────────

@app.route('/api/auth/login', methods=['POST'])
def login():
    data     = request.get_json(silent=True) or {}
    username = (data.get('username') or '').strip()
    password = data.get('password') or ''

    if not username or not password:
        return jsonify({'error': 'Username and password are required.'}), 400

    user = get_user_by_username(username)
    if not user or not check_password_hash(user['password_hash'], password):
        return jsonify({'error': 'Incorrect username or password.'}), 401

    user['allowed_modules'] = json.loads(user['allowed_modules']) if isinstance(user['allowed_modules'], str) else user['allowed_modules']
    token = create_token(user)
    return jsonify({'token': token, 'username': user['username'], 'is_admin': bool(user['is_admin'])})


@app.route('/api/auth/change-password', methods=['POST'])
@require_auth
def change_password():
    data     = request.get_json(silent=True) or {}
    cur_pw   = data.get('current_password') or ''
    new_pw   = data.get('new_password')     or ''

    if not cur_pw or not new_pw:
        return jsonify({'error': 'Current and new passwords are required.'}), 400
    if len(new_pw) < 6:
        return jsonify({'error': 'New password must be at least 6 characters.'}), 400

    user = get_user_by_id(g.user_id)
    if not user or not check_password_hash(user['password_hash'], cur_pw):
        return jsonify({'error': 'Current password is incorrect.'}), 401

    update_user(user_id=g.user_id,
                password_hash=generate_password_hash(new_pw, method='pbkdf2:sha256'))
    return jsonify({'ok': True})


@app.route('/api/auth/me')
@require_auth
def auth_me():
    user = get_user_by_id(g.user_id)
    if not user:
        return jsonify({'error': 'User not found'}), 404
    modules = json.loads(user['allowed_modules']) if isinstance(user['allowed_modules'], str) else user['allowed_modules']
    return jsonify({
        'user_id':         user['id'],
        'username':        user['username'],
        'is_admin':        bool(user['is_admin']),
        'allowed_modules': modules,
    })


# ── Module registry ────────────────────────────────────────────────────────────

MODULES = [
    {'id':'rehash',              'name':'Sales Report',              'icon':'📊','color':'#1e5799','status':'active',      'schedule':'Every Monday', 'sources':['Jive (GoTo)','LeadPerfection','SharePoint Excel'],'description':'Pulls weekly Jive call data and LeadPerfection PDFs, then writes all metrics into the Sales Report on SharePoint.'},
    {'id':'coordinator',         'name':'Project Coordinator Report','icon':'📋','color':'#7c3aed','status':'coming_soon', 'schedule':'Weekly',       'sources':['SharePoint'],                                       'description':'Automates project status tracking and coordinator reporting across active jobs.'},
    {'id':'marketing',           'name':'Marketing Report',          'icon':'📣','color':'#db2777','status':'coming_soon', 'schedule':'Weekly',       'sources':['SharePoint'],                                       'description':'Compiles lead sources, campaign performance, and cost-per-appointment metrics.'},
    {'id':'scheduling',          'name':'Scheduling Report',         'icon':'📅','color':'#0891b2','status':'coming_soon', 'schedule':'Weekly',       'sources':['LeadPerfection','SharePoint'],                       'description':'Tracks appointment scheduling activity, confirmation rates, and cancellations.'},
    {'id':'recruiting',          'name':'Recruiting Report',         'icon':'🤝','color':'#059669','status':'coming_soon', 'schedule':'Weekly',       'sources':['SharePoint'],                                       'description':'Monitors hiring pipeline, applicant volume, and onboarding progress.'},
    {'id':'commission',          'name':'Sales Commission',          'icon':'💵','color':'#d97706','status':'active',      'schedule':'Per client',   'sources':['Buildertrend','Commission PDF','SharePoint Excel'],  'description':'Processes commission entries per client — searches Buildertrend, extracts job & PDF data, matches sales rep, and inserts into SharePoint Excel.'},
    {'id':'commission_estimate', 'name':'Commission Estimate',       'icon':'📈','color':'#0f766e','status':'active',      'schedule':'Per estimate', 'sources':['Estimate Details PDF','Commission Sheet PDF'],       'description':'Upload Estimate Details and Commission Sheet PDFs from the CRM — calculates corrected commissions and returns annotated PDFs ready for review.'},
    {'id':'onsite_recon',        'name':'Onsite Payroll Report',     'icon':'💼','color':'#7c3aed','status':'active',      'schedule':'Per payroll run','sources':['Gusto Payroll CSV'],                                'description':'Upload a Gusto payroll CSV to generate a formatted, color-coded Excel report with department breakdown, totals, and variance analysis.'},
    {'id':'bank_categorizer',    'name':'5160 Report',               'icon':'🏦','color':'#0369a1','status':'active',      'schedule':'Per statement', 'sources':['Bank XLSX','Wise PDF (optional)','WU XLSX (optional)'], 'description':'Upload your bank statement and optionally Wise PDF and Western Union history to generate a color-coded categorized Excel report.'},
]


@app.route('/api/modules')
@require_auth
def get_modules():
    history  = load_history()
    allowed  = g.allowed_modules
    enriched = []
    for m in MODULES:
        if '*' not in allowed and m['id'] not in allowed:
            continue
        mod  = dict(m)
        runs = [h for h in history if h.get('module') == m['id']]
        mod['last_run'] = runs[0] if runs else None
        enriched.append(mod)
    return jsonify(enriched)


# ── Admin — user management ────────────────────────────────────────────────────

@app.route('/api/admin/users', methods=['GET'])
@require_admin
def admin_list_users():
    return jsonify(get_all_users())


@app.route('/api/admin/users', methods=['POST'])
@require_admin
def admin_create_user():
    data     = request.get_json(silent=True) or {}
    username = (data.get('username') or '').strip()
    password = data.get('password') or ''
    modules  = data.get('allowed_modules', [])

    if not username or not password:
        return jsonify({'error': 'Username and password are required.'}), 400
    if len(password) < 6:
        return jsonify({'error': 'Password must be at least 6 characters.'}), 400
    if get_user_by_username(username):
        return jsonify({'error': f"Username '{username}' is already taken."}), 409

    create_user(username=username, password_hash=generate_password_hash(password, method='pbkdf2:sha256'),
                is_admin=False, allowed_modules=modules)
    return jsonify({'ok': True, 'username': username}), 201


@app.route('/api/admin/users/<int:user_id>', methods=['PATCH'])
@require_admin
def admin_update_user(user_id):
    if user_id == g.user_id:
        return jsonify({'error': 'You cannot edit your own account from the admin panel.'}), 400
    data    = request.get_json(silent=True) or {}
    modules = data.get('allowed_modules')
    pw      = data.get('password')
    if pw and len(pw) < 6:
        return jsonify({'error': 'Password must be at least 6 characters.'}), 400
    update_user(user_id=user_id, allowed_modules=modules,
                password_hash=generate_password_hash(pw, method='pbkdf2:sha256') if pw else None)
    return jsonify({'ok': True})


@app.route('/api/admin/users/<int:user_id>', methods=['DELETE'])
@require_admin
def admin_delete_user(user_id):
    if user_id == g.user_id:
        return jsonify({'error': 'You cannot delete your own account.'}), 400
    delete_user(user_id)
    return jsonify({'ok': True})


@app.route('/api/admin/modules')
@require_admin
def admin_modules():
    return jsonify([{'id': m['id'], 'name': m['name'], 'icon': m['icon']} for m in MODULES])


# ── Onsite Recon — employee config ────────────────────────────────────────────


@app.route('/api/onsite_recon/employees', methods=['GET'])
@require_auth
def onsite_get_employees():
    config = _load_json_config(ONSITE_CONFIG_PATH, default={'departments': {}, 'nickname_overrides': {}})
    # Flatten into list: [{name, department}, ...]
    employees = []
    for dept, names in config.get('departments', {}).items():
        for name in names:
            employees.append({'name': name, 'department': dept})
    employees.sort(key=lambda e: (e['department'], e['name']))
    return jsonify({
        'employees': employees,
        'departments': ['INSTALLERS','PC-WAREHOUSE','SALES TEAM','OFFICE','PC-OFFICE','CALL CENTER','EVENT TEAM'],
    })


@app.route('/api/onsite_recon/employees', methods=['POST'])
@require_auth
def onsite_save_employees():
    data      = request.get_json(silent=True) or {}
    employees = data.get('employees', [])   # [{name, department}, ...]

    config = _load_json_config(ONSITE_CONFIG_PATH, default={'departments': {}, 'nickname_overrides': {}})
    dept_map = {}
    for emp in employees:
        dept = emp.get('department', '').strip()
        name = emp.get('name', '').strip()
        if dept and name:
            dept_map.setdefault(dept, []).append(name)

    config['departments'] = dept_map
    _save_json_config(ONSITE_CONFIG_PATH, config)
    return jsonify({'ok': True, 'count': len(employees)})


@app.route('/api/run/onsite_recon', methods=['POST'])
@require_auth
def run_onsite_recon():
    csv_file = request.files.get('payroll_csv')
    if not csv_file:
        return jsonify({'error': 'Gusto payroll CSV is required.'}), 400

    tmp_dir = Path(tempfile.mkdtemp())
    try:
        orig_ext  = Path(csv_file.filename).suffix.lower() or '.csv'
        csv_path  = tmp_dir / f'payroll{orig_ext}'
        xlsx_name = f"Onsite Report - {datetime.now().strftime('%b %d, %Y')}.xlsx"
        xlsx_path = tmp_dir / xlsx_name

        csv_file.save(str(csv_path))

        venv_python = HUB_DIR / '.venv' / 'bin' / 'python3'
        script      = ONSITE_RECON_DIR / 'process_payroll.py'

        cmd = [
            str(venv_python), str(script),
            '--input',  str(csv_path),
            '--output', str(xlsx_path),
            '--depts',  str(ONSITE_CONFIG_PATH),
        ]

        result, err = _run_script(cmd)
        if err:
            return jsonify({'error': err}), 500
        if not xlsx_path.exists():
            return jsonify({'error': 'No output file was generated.'}), 500

        # Parse unmatched employee warnings from stdout
        unmatched = []
        for line in (result.stdout or '').splitlines():
            s = line.strip()
            if s.startswith('- ') and 'UNMATCHED' not in s and '⚠️' not in s:
                unmatched.append(s[2:].strip())

        xlsx_b64 = base64.b64encode(xlsx_path.read_bytes()).decode()
        save_history({'module': 'onsite_recon', 'ran_at': datetime.now().isoformat(), 'status': 'success'})

        return jsonify({'ok': True, 'xlsx_b64': xlsx_b64, 'filename': xlsx_name, 'unmatched': unmatched})

    except Exception as e:
        return jsonify({'error': str(e), 'detail': traceback.format_exc()}), 500
    finally:
        shutil.rmtree(str(tmp_dir), ignore_errors=True)


# ── 5160 Report — employee config ─────────────────────────────────────────────

BANK_CAT_GROUPS = [
    "PC_PAYROLL", "SC_PAYROLL", "OFFICE_PAYROLL", "ADV_MARKETING",
    "ADV_ZELLE", "RECRUITER", "BRAND_AMBASSADOR", "EVENT",
    "REMOTE_SALES", "MATERIAL", "OVERHEAD",
]

@app.route('/api/bank_categorizer/employees', methods=['GET'])
@require_auth
def bank_cat_get_employees():
    emp_map = _load_json_config(BANK_CAT_CONFIG_PATH)
    employees = [
        {'key': k, 'label': v.get('label', ''), 'group': v.get('group', '')}
        for k, v in emp_map.items()
    ]
    employees.sort(key=lambda e: (e['group'], e['key']))
    return jsonify({'employees': employees, 'groups': BANK_CAT_GROUPS})


@app.route('/api/bank_categorizer/employees', methods=['POST'])
@require_auth
def bank_cat_save_employees():
    data      = request.get_json(silent=True) or {}
    employees = data.get('employees', [])   # [{key, label, group}, ...]

    emp_map = {}
    for emp in employees:
        key   = (emp.get('key') or '').strip().lower()
        label = (emp.get('label') or '').strip()
        group = (emp.get('group') or '').strip()
        if key and label and group:
            emp_map[key] = {'label': label, 'group': group}

    _save_json_config(BANK_CAT_CONFIG_PATH, emp_map)
    return jsonify({'ok': True, 'count': len(emp_map)})


@app.route('/api/run/bank_categorizer', methods=['POST'])
@require_auth
def run_bank_categorizer():
    bank_file = request.files.get('bank_xlsx')
    if not bank_file:
        return jsonify({'error': 'Bank statement XLSX is required.'}), 400

    wise_file = request.files.get('wise_pdf')
    wu_file   = request.files.get('wu_xlsx')

    tmp_dir = Path(tempfile.mkdtemp())
    try:
        bank_path = tmp_dir / bank_file.filename
        bank_file.save(str(bank_path))

        wise_path = wu_path = None
        if wise_file:
            wise_path = tmp_dir / wise_file.filename
            wise_file.save(str(wise_path))
        if wu_file:
            wu_path = tmp_dir / wu_file.filename
            wu_file.save(str(wu_path))

        xlsx_name = f"5160 Report - {datetime.now().strftime('%b %d, %Y')}.xlsx"
        xlsx_path = tmp_dir / xlsx_name

        venv_python = HUB_DIR / '.venv' / 'bin' / 'python3'
        script      = BANK_CAT_DIR / 'categorize.py'

        cmd = [
            str(venv_python), str(script),
            '--bank',   str(bank_path),
            '--empmap', str(BANK_CAT_CONFIG_PATH),
            '--output', str(xlsx_path),
        ]
        if wise_path: cmd += ['--wise', str(wise_path)]
        if wu_path:   cmd += ['--wu',   str(wu_path)]

        result, err = _run_script(cmd)
        if err:
            return jsonify({'error': err}), 500
        if not xlsx_path.exists():
            return jsonify({'error': 'No output file was generated.'}), 500

        # Parse UNMATCHED_SUGGESTION lines from stdout
        suggestions = []
        for line in (result.stdout or '').splitlines():
            s = line.strip()
            if s.startswith('UNMATCHED_SUGGESTION:'):
                payload = s[len('UNMATCHED_SUGGESTION:'):].strip()
                entry = {}
                for part in payload.split('|'):
                    part = part.strip()
                    if '=' in part:
                        k, v = part.split('=', 1)
                        entry[k.strip()] = v.strip()
                if entry.get('raw_name'):
                    suggestions.append({
                        'source':          entry.get('source', ''),
                        'raw_name':        entry.get('raw_name', ''),
                        'suggested_label': entry.get('suggested_label', ''),
                        'suggested_group': entry.get('suggested_group', ''),
                    })

        xlsx_b64 = base64.b64encode(xlsx_path.read_bytes()).decode()
        save_history({'module': 'bank_categorizer', 'ran_at': datetime.now().isoformat(), 'status': 'success'})

        return jsonify({
            'ok':          True,
            'xlsx_b64':    xlsx_b64,
            'filename':    xlsx_name,
            'suggestions': suggestions,
        })

    except Exception as e:
        return jsonify({'error': str(e), 'detail': traceback.format_exc()}), 500
    finally:
        shutil.rmtree(str(tmp_dir), ignore_errors=True)


# ── Chrome management ──────────────────────────────────────────────────────────

@app.route('/api/chrome/status')
@require_auth
def chrome_status():
    return jsonify({'ready': cdp_is_alive()})


def _sync_sessions_to_auto_dir():
    src_profile  = CHROME_SRC_DIR / CHROME_SRC_PROFILE
    auto_profile = Path(CHROME_AUTO_DIR) / CHROME_SRC_PROFILE
    first_time   = not auto_profile.exists()
    auto_profile.mkdir(parents=True, exist_ok=True)

    for fname in ('Cookies','Cookies-journal','Preferences','Secure Preferences',
                  'Web Data','Web Data-journal','Visited Links','Favicons','Top Sites',
                  'TransportSecurity','Network Persistent State'):
        src = src_profile / fname
        if src.exists():
            try:
                import shutil as _sh
                _sh.copy2(str(src), str(auto_profile / fname))
            except Exception:
                pass

    ls_src = CHROME_SRC_DIR / 'Local State'
    if ls_src.exists():
        try:
            import shutil as _sh
            _sh.copy2(str(ls_src), str(Path(CHROME_AUTO_DIR) / 'Local State'))
        except Exception:
            pass

    if not first_time:
        return

    for dname in ('Local Storage','Session Storage','Extension State',
                  'Local Extension Settings','IndexedDB','Sync Data'):
        src_dir = src_profile / dname
        dst_dir = auto_profile / dname
        if src_dir.exists():
            try:
                if dst_dir.exists():
                    shutil.rmtree(str(dst_dir))
                shutil.copytree(str(src_dir), str(dst_dir),
                                ignore=shutil.ignore_patterns('*.tmp'))
            except Exception:
                pass


@app.route('/api/chrome/launch', methods=['POST'])
@require_auth
def launch_chrome():
    if cdp_is_alive():
        return jsonify({'ok': True, 'already_running': True})
    _sync_sessions_to_auto_dir()
    subprocess.Popen([CHROME_PATH, f'--remote-debugging-port={CDP_PORT}',
                      f'--user-data-dir={CHROME_AUTO_DIR}',
                      f'--profile-directory={CHROME_SRC_PROFILE}',
                      '--no-first-run', '--no-default-browser-check', '--new-window',
                      '--remote-allow-origins=*'])
    for _ in range(60):
        time.sleep(0.5)
        if cdp_is_alive():
            return jsonify({'ok': True, 'already_running': False})
    return jsonify({'ok': False, 'error': 'Chrome did not respond in time'}), 500


# ── Run automation ─────────────────────────────────────────────────────────────

@app.route('/api/run/<module_id>', methods=['POST'])
@require_auth
def run_module(module_id):
    import queue as Q
    job_id = str(uuid.uuid4())
    q      = Q.Queue()
    stop   = threading.Event()
    jobs[job_id]         = q
    cancel_flags[job_id] = stop

    body         = request.get_json(silent=True) or {}
    body['user_id'] = g.user_id   # pass user_id into worker

    def worker():
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        job_loops[job_id] = loop
        try:
            rehash_dir = HUB_DIR.parent / 'data-scraping-automation'
            sys.path.insert(0, str(rehash_dir))

            if module_id == 'rehash':
                from automation.runner import run_full_automation
                config = json.loads((rehash_dir / 'config.json').read_text())
                loop.run_until_complete(run_full_automation(q, config, stop))

            elif module_id == 'commission':
                from automation.commission_runner import run_commission
                reps        = get_user_setting(body['user_id'], 'commission', 'reps') or []
                comm_config = {'reps': reps}
                client_name = body.get('client_name', '').strip()
                if not client_name:
                    q.put({'type': 'error', 'msg': 'No client name provided.'})
                    return
                loop.run_until_complete(run_commission(q, comm_config, client_name, stop))

            else:
                q.put({'type': 'error', 'msg': f'Unknown module: {module_id}'})
                return

            save_history({'module': module_id, 'ran_at': datetime.now().isoformat(),
                          'status': 'success', 'job_id': job_id})
        except Exception as e:
            import traceback
            q.put({'type': 'error', 'msg': f'Server error: {e}\n{traceback.format_exc()}'})
            save_history({'module': module_id, 'ran_at': datetime.now().isoformat(),
                          'status': 'error', 'error': str(e), 'job_id': job_id})
        finally:
            loop.close()
            cancel_flags.pop(job_id, None)
            job_loops.pop(job_id, None)

    threading.Thread(target=worker, daemon=True).start()
    return jsonify({'job_id': job_id})


@app.route('/api/cancel/<job_id>', methods=['POST'])
@require_auth
def cancel_job(job_id):
    stop = cancel_flags.get(job_id)
    if stop:
        stop.set()
        q = jobs.get(job_id)
        if q:
            q.put({'type': 'error', 'msg': '⛔ Run cancelled by user.'})
        return jsonify({'ok': True})
    return jsonify({'ok': False, 'error': 'Job not found or already finished'})


@app.route('/api/status/<job_id>')
@require_auth
def job_status(job_id):
    def generate():
        import queue as Q
        q = jobs.get(job_id)
        if not q:
            yield f"data: {json.dumps({'type':'error','msg':'Job not found'})}\n\n"
            return
        while True:
            try:
                msg = q.get(timeout=60)
                yield f"data: {json.dumps(msg)}\n\n"
                if msg.get('type') in ('done', 'error'):
                    jobs.pop(job_id, None)
                    break
            except Q.Empty:
                yield f"data: {json.dumps({'type':'ping'})}\n\n"

    return Response(generate(), mimetype='text/event-stream',
                    headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})


@app.route('/api/history')
@require_auth
def get_history():
    return jsonify(load_history())


# ── Rehash employees & config ──────────────────────────────────────────────────

def _rehash_config_path():
    return HUB_DIR.parent / 'data-scraping-automation' / 'config.json'


@app.route('/api/rehash/employees', methods=['GET'])
@require_auth
def get_rehash_employees():
    try:
        return jsonify(json.loads(_rehash_config_path().read_text()).get('employees', []))
    except Exception:
        return jsonify([])


@app.route('/api/rehash/employees', methods=['POST'])
@require_auth
def add_rehash_employee():
    data = request.json
    if not data.get('name'):
        return jsonify({'error': 'Name required'}), 400
    p = _rehash_config_path()
    config = json.loads(p.read_text())
    emp = {'id': data['name'].lower().replace(' ', '_'), 'name': data['name'],
           'jive_url': data.get('jive_url', ''), 'excel_url': data.get('excel_url', ''),
           'lp_name': data.get('lp_name', data['name'])}
    config['employees'].append(emp)
    p.write_text(json.dumps(config, indent=2))
    return jsonify(emp), 201


@app.route('/api/rehash/employees/<emp_id>', methods=['PATCH'])
@require_auth
def update_rehash_employee(emp_id):
    data = request.json
    p    = _rehash_config_path()
    config = json.loads(p.read_text())
    for emp in config['employees']:
        if emp['id'] == emp_id:
            for field in ('name', 'lp_name', 'jive_url', 'excel_url'):
                if field in data:
                    emp[field] = data[field]
            p.write_text(json.dumps(config, indent=2))
            return jsonify(emp)
    return jsonify({'error': 'Not found'}), 404


@app.route('/api/rehash/employees/<emp_id>', methods=['DELETE'])
@require_auth
def delete_rehash_employee(emp_id):
    p = _rehash_config_path()
    config = json.loads(p.read_text())
    config['employees'] = [e for e in config['employees'] if e['id'] != emp_id]
    p.write_text(json.dumps(config, indent=2))
    return jsonify({'ok': True})


@app.route('/api/rehash/config', methods=['GET'])
@require_auth
def get_rehash_config():
    try:
        return jsonify(json.loads(_rehash_config_path().read_text()).get('sharepoint', {}))
    except Exception:
        return jsonify({})


@app.route('/api/rehash/config', methods=['PATCH'])
@require_auth
def update_rehash_config():
    data = request.json
    p    = _rehash_config_path()
    config = json.loads(p.read_text())
    sp = config.setdefault('sharepoint', {})
    for field in ('demo_sheet_url', 'demo_sheet', 'file_url'):
        if field in data:
            sp[field] = data[field]
    p.write_text(json.dumps(config, indent=2))
    return jsonify(sp)


# ── Commission reps (per-user) ─────────────────────────────────────────────────

@app.route('/api/commission/reps', methods=['GET'])
@require_auth
def get_commission_reps():
    return jsonify(get_user_setting(g.user_id, 'commission', 'reps') or [])


@app.route('/api/commission/reps', methods=['POST'])
@require_auth
def add_commission_rep():
    data = request.json
    if not data.get('name'):
        return jsonify({'error': 'Name required'}), 400
    reps = get_user_setting(g.user_id, 'commission', 'reps') or []
    rep  = {'id': data['name'].lower().replace(' ', '_'),
            'name': data['name'], 'excel_url': data.get('excel_url', '')}
    reps.append(rep)
    set_user_setting(g.user_id, 'commission', 'reps', reps)
    return jsonify(rep), 201


@app.route('/api/commission/reps/<rep_id>', methods=['PATCH'])
@require_auth
def update_commission_rep(rep_id):
    data = request.json
    reps = get_user_setting(g.user_id, 'commission', 'reps') or []
    for rep in reps:
        if rep['id'] == rep_id:
            for field in ('name', 'excel_url'):
                if field in data:
                    rep[field] = data[field]
            set_user_setting(g.user_id, 'commission', 'reps', reps)
            return jsonify(rep)
    return jsonify({'error': 'Not found'}), 404


@app.route('/api/commission/reps/<rep_id>', methods=['DELETE'])
@require_auth
def delete_commission_rep(rep_id):
    reps = get_user_setting(g.user_id, 'commission', 'reps') or []
    reps = [r for r in reps if r['id'] != rep_id]
    set_user_setting(g.user_id, 'commission', 'reps', reps)
    return jsonify({'ok': True})


# ── Commission Estimate ────────────────────────────────────────────────────────

@app.route('/api/commission-estimate/process', methods=['POST'])
@require_auth
def process_commission_estimate():
    est_file  = request.files.get('estimate_pdf')
    comm_file = request.files.get('commission_pdf')

    if not est_file or not comm_file:
        return jsonify({'error': 'Both PDF files are required'}), 400

    finance_fee    = request.form.get('finance_fee', '').strip()
    finance_lender = request.form.get('finance_lender', '').strip()

    tmp_dir = Path(tempfile.mkdtemp())
    try:
        est_path  = tmp_dir / est_file.filename
        comm_path = tmp_dir / comm_file.filename
        est_file.save(str(est_path))
        comm_file.save(str(comm_path))

        out_dir = tmp_dir / 'output'
        out_dir.mkdir()

        venv_python = HUB_DIR / '.venv' / 'bin' / 'python3'
        script      = COMM_EST_DIR / 'commission_annotator.py'

        cmd = [str(venv_python), str(script),
               '--estimate',   str(est_path),
               '--commission', str(comm_path),
               '--output-dir', str(out_dir)]
        if finance_fee:    cmd += ['--ff',      finance_fee]
        if finance_lender: cmd += ['--ff-name', finance_lender]

        result, err = _run_script(cmd)
        if err:
            return jsonify({'error': err}), 500

        output_pdfs = list(out_dir.glob('*.pdf'))
        if not output_pdfs:
            return jsonify({'error': 'No output PDFs were generated.'}), 500

        stdout  = result.stdout or ''
        alerts, is_sean, summary = [], False, {}

        for line in stdout.splitlines():
            s = line.strip()
            if 'inform Anne' in s or ('⚠️' in s and '%' in s):
                clean = s.lstrip('⚠️ -').strip()
                if clean and clean not in alerts:
                    alerts.append(clean)
            rep_m = _re.search(r'Rep\s*:\s*(.+)', s)
            if rep_m:
                summary['rep_name'] = rep_m.group(1).strip()
                if 'sean' in rep_m.group(1).lower():
                    is_sean = True
            def _x(label, key):
                m = _re.search(rf'{label}\s*:\s*([^\n]+)', s, _re.IGNORECASE)
                if m: summary[key] = m.group(1).strip()
            _x('Profit Margin',       'profit_margin')
            _x('% Greenline',         'pct_greenline')
            _x('Commission %',        'commission_pct')
            _x(r'Est\. Commission',   'est_commission')
            _x(r'Greenline(?!\s*%)',  'greenline')

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for pdf in output_pdfs:
                zf.write(str(pdf), pdf.name)
        zip_buf.seek(0)

        save_history({'module': 'commission_estimate', 'ran_at': datetime.now().isoformat(), 'status': 'success'})

        return jsonify({'ok': True, 'zip_b64': base64.b64encode(zip_buf.read()).decode(),
                        'alerts': alerts, 'is_sean': is_sean, 'summary': summary})

    except subprocess.TimeoutExpired:
        return jsonify({'error': 'Processing timed out (>120 s).'}), 500
    except Exception as e:
        return jsonify({'error': str(e), 'detail': traceback.format_exc()}), 500
    finally:
        shutil.rmtree(str(tmp_dir), ignore_errors=True)


# ── Entry point ────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    startup()
    print('\n  ✓ Payless Automation Hub → http://localhost:5050\n')
    app.run(port=5050, debug=False)
