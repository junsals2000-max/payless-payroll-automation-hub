"""
Payless Automation Hub — Flask Backend
Payless Kitchen Cabinets & Bath Makeover
Central server for all automation modules.
"""

from flask import Flask, jsonify, request, Response, send_from_directory, send_file
from flask_cors import CORS
import json, threading, asyncio, uuid, subprocess, time, os, io, zipfile, tempfile, base64, re as _re
from pathlib import Path
from datetime import datetime

BASE_DIR    = Path(__file__).parent
HISTORY_FILE = BASE_DIR / 'run_history.json'
COMM_EST_DIR = Path.home() / 'Desktop' / 'Automation' / 'Commission Estimate Automation'
CDP_PORT    = 9222
CHROME_PATH = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
# Source profile: Junjun (junsals2000@gmail.com) — has all signed-in sessions
CHROME_SRC_DIR      = Path.home() / 'Library' / 'Application Support' / 'Google' / 'Chrome'
CHROME_SRC_PROFILE  = 'Profile 2'
# Automation Chrome uses a DEDICATED non-default directory so --remote-debugging-port works.
# Chrome silently disables CDP when --user-data-dir is the default profile path.
CHROME_AUTO_DIR     = str(Path.home() / '.payless-automation-chrome')

app = Flask(__name__, static_folder=str(BASE_DIR))
CORS(app)

jobs        = {}   # job_id → queue.Queue
cancel_flags = {}  # job_id → threading.Event  (set to stop a running job)
job_loops    = {}  # job_id → asyncio.AbstractEventLoop


# ── Helpers ────────────────────────────────────────────────────────────────────

def load_history():
    if HISTORY_FILE.exists():
        return json.loads(HISTORY_FILE.read_text())
    return []

def save_history(entry):
    history = load_history()
    history.insert(0, entry)
    history = history[:50]   # keep last 50
    HISTORY_FILE.write_text(json.dumps(history, indent=2))

def cdp_is_alive():
    import requests as req
    try:
        req.get(f'http://localhost:{CDP_PORT}/json/version', timeout=1)
        return True
    except Exception:
        return False


# ── Static files ───────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return send_from_directory(str(BASE_DIR), 'index.html')

@app.route('/<path:filename>')
def static_files(filename):
    return send_from_directory(str(BASE_DIR), filename)


# ── Module registry ────────────────────────────────────────────────────────────

MODULES = [
    {
        'id':          'rehash',
        'name':        'Sales Report',
        'description': 'Pulls weekly Jive call data and LeadPerfection PDFs, then writes all metrics into the Sales Report on SharePoint.',
        'icon':        '📊',
        'color':       '#1e5799',
        'status':      'active',
        'schedule':    'Every Monday',
        'sources':     ['Jive (GoTo)', 'LeadPerfection', 'SharePoint Excel'],
    },
    {
        'id':          'coordinator',
        'name':        'Project Coordinator Report',
        'description': 'Automates project status tracking and coordinator reporting across active jobs.',
        'icon':        '📋',
        'color':       '#7c3aed',
        'status':      'coming_soon',
        'schedule':    'Weekly',
        'sources':     ['SharePoint'],
    },
    {
        'id':          'marketing',
        'name':        'Marketing Report',
        'description': 'Compiles lead sources, campaign performance, and cost-per-appointment metrics.',
        'icon':        '📣',
        'color':       '#db2777',
        'status':      'coming_soon',
        'schedule':    'Weekly',
        'sources':     ['SharePoint'],
    },
    {
        'id':          'scheduling',
        'name':        'Scheduling Report',
        'description': 'Tracks appointment scheduling activity, confirmation rates, and cancellations.',
        'icon':        '📅',
        'color':       '#0891b2',
        'status':      'coming_soon',
        'schedule':    'Weekly',
        'sources':     ['LeadPerfection', 'SharePoint'],
    },
    {
        'id':          'recruiting',
        'name':        'Recruiting Report',
        'description': 'Monitors hiring pipeline, applicant volume, and onboarding progress.',
        'icon':        '🤝',
        'color':       '#059669',
        'status':      'coming_soon',
        'schedule':    'Weekly',
        'sources':     ['SharePoint'],
    },
    {
        'id':          'commission',
        'name':        'Sales Commission',
        'description': 'Processes commission entries per client — searches Buildertrend, extracts job & PDF data, matches sales rep, and inserts into SharePoint Excel.',
        'icon':        '💵',
        'color':       '#d97706',
        'status':      'active',
        'schedule':    'Per client',
        'sources':     ['Buildertrend', 'Commission PDF', 'SharePoint Excel'],
    },
    {
        'id':          'commission_estimate',
        'name':        'Commission Estimate',
        'description': 'Upload Estimate Details and Commission Sheet PDFs from the CRM — calculates corrected commissions and returns annotated PDFs ready for review.',
        'icon':        '📈',
        'color':       '#0f766e',
        'status':      'active',
        'schedule':    'Per estimate',
        'sources':     ['Estimate Details PDF', 'Commission Sheet PDF'],
    },
]

@app.route('/api/modules')
def get_modules():
    history = load_history()
    enriched = []
    for m in MODULES:
        mod = dict(m)
        # Find last run for this module
        runs = [h for h in history if h.get('module') == m['id']]
        mod['last_run'] = runs[0] if runs else None
        enriched.append(mod)
    return jsonify(enriched)


# ── Chrome management ──────────────────────────────────────────────────────────

@app.route('/api/chrome/status')
def chrome_status():
    return jsonify({'ready': cdp_is_alive()})

def _sync_sessions_to_auto_dir():
    """
    Sync sign-in sessions from Profile 2 (junsals2000@gmail.com) into the
    automation Chrome dir so no re-login or OTP is needed.

    First launch: full copy of all session directories (~4 s).
    Subsequent launches: only Cookies is refreshed (<0.1 s) since the rest
    of the profile is already in place.
    """
    import shutil

    src_profile  = CHROME_SRC_DIR / CHROME_SRC_PROFILE           # .../Chrome/Profile 2
    auto_profile = Path(CHROME_AUTO_DIR) / CHROME_SRC_PROFILE    # ~/.payless.../Profile 2
    first_time   = not auto_profile.exists()
    auto_profile.mkdir(parents=True, exist_ok=True)

    # --- always refresh cookies so sessions stay current ---
    for fname in (
        'Cookies', 'Cookies-journal',
        'Preferences', 'Secure Preferences',
        'Web Data', 'Web Data-journal',
        'Visited Links', 'Favicons', 'Top Sites',
        'TransportSecurity', 'Network Persistent State',
    ):
        src = src_profile / fname
        if src.exists():
            try:
                shutil.copy2(str(src), str(auto_profile / fname))
            except Exception:
                pass

    # --- top-level Local State (profile registry) ---
    ls_src = CHROME_SRC_DIR / 'Local State'
    if ls_src.exists():
        try:
            shutil.copy2(str(ls_src), str(Path(CHROME_AUTO_DIR) / 'Local State'))
        except Exception:
            pass

    if not first_time:
        return   # skip heavy directory copy on subsequent launches

    # --- first-time only: copy session directories ---
    for dname in (
        'Local Storage',
        'Session Storage',
        'Extension State',
        'Local Extension Settings',
        'IndexedDB',
        'Sync Data',
    ):
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
def launch_chrome():
    """
    Open the automation Chrome window using a dedicated non-default user-data-dir
    so that --remote-debugging-port is honoured (Chrome blocks CDP on the default dir).
    Syncs the full Profile 2 first so the user is already signed in — no OTP needed.
    """
    if cdp_is_alive():
        return jsonify({'ok': True, 'already_running': True})

    _sync_sessions_to_auto_dir()

    subprocess.Popen([
        CHROME_PATH,
        f'--remote-debugging-port={CDP_PORT}',
        f'--user-data-dir={CHROME_AUTO_DIR}',
        f'--profile-directory={CHROME_SRC_PROFILE}',   # = 'Profile 2'
        '--no-first-run',
        '--no-default-browser-check',
        '--new-window',
    ])

    for _ in range(60):          # wait up to 30 s
        time.sleep(0.5)
        if cdp_is_alive():
            return jsonify({'ok': True, 'already_running': False})

    return jsonify({'ok': False, 'error': 'Chrome did not respond in time'}), 500


# ── Run automation ─────────────────────────────────────────────────────────────

@app.route('/api/run/<module_id>', methods=['POST'])
def run_module(module_id):
    import queue as Q
    job_id = str(uuid.uuid4())
    q      = Q.Queue()
    stop   = threading.Event()
    jobs[job_id]         = q
    cancel_flags[job_id] = stop

    # Read request body HERE (in the Flask thread) before spawning worker
    body = request.get_json(silent=True) or {}

    def worker():
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        job_loops[job_id] = loop
        try:
            import sys
            sys.path.insert(0, str(BASE_DIR.parent / 'rehash-automation'))

            if module_id == 'rehash':
                from automation.runner import run_full_automation
                config = json.loads(
                    (BASE_DIR.parent / 'rehash-automation' / 'config.json').read_text()
                )
                loop.run_until_complete(run_full_automation(q, config, stop))

            elif module_id == 'commission':
                from automation.commission_runner import run_commission
                comm_config  = load_commission_config()
                client_name  = body.get('client_name', '').strip()
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
            if not stop.is_set():  # don't log history for manual cancels
                save_history({'module': module_id, 'ran_at': datetime.now().isoformat(),
                              'status': 'error', 'error': str(e), 'job_id': job_id})
        finally:
            loop.close()
            cancel_flags.pop(job_id, None)
            job_loops.pop(job_id, None)

    threading.Thread(target=worker, daemon=True).start()
    return jsonify({'job_id': job_id})


@app.route('/api/cancel/<job_id>', methods=['POST'])
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
def get_history():
    return jsonify(load_history())


@app.route('/api/rehash/employees', methods=['GET'])
def get_rehash_employees():
    try:
        config = json.loads((BASE_DIR.parent / 'rehash-automation' / 'config.json').read_text())
        return jsonify(config.get('employees', []))
    except Exception as e:
        return jsonify([])

@app.route('/api/rehash/employees', methods=['POST'])
def add_rehash_employee():
    data = request.json
    if not data.get('name'):
        return jsonify({'error': 'Name required'}), 400
    config_path = BASE_DIR.parent / 'rehash-automation' / 'config.json'
    config = json.loads(config_path.read_text())
    emp_id = data['name'].lower().replace(' ', '_')
    emp = {
        'id': emp_id,
        'name': data['name'],
        'jive_url': data.get('jive_url', ''),
        'excel_url': data.get('excel_url', ''),
        'lp_name': data.get('lp_name', data['name'])
    }
    config['employees'].append(emp)
    config_path.write_text(json.dumps(config, indent=2))
    return jsonify(emp), 201

@app.route('/api/rehash/employees/<emp_id>', methods=['PATCH'])
def update_rehash_employee(emp_id):
    data = request.json
    config_path = BASE_DIR.parent / 'rehash-automation' / 'config.json'
    config = json.loads(config_path.read_text())
    for emp in config['employees']:
        if emp['id'] == emp_id:
            for field in ('name', 'lp_name', 'jive_url', 'excel_url'):
                if field in data:
                    emp[field] = data[field]
            config_path.write_text(json.dumps(config, indent=2))
            return jsonify(emp)
    return jsonify({'error': 'Not found'}), 404


# ── Rehash SharePoint config ────────────────────────────────────────────────────

@app.route('/api/rehash/config', methods=['GET'])
def get_rehash_config():
    try:
        config = json.loads((BASE_DIR.parent / 'rehash-automation' / 'config.json').read_text())
        return jsonify(config.get('sharepoint', {}))
    except Exception:
        return jsonify({})

@app.route('/api/rehash/config', methods=['PATCH'])
def update_rehash_config():
    data = request.json
    config_path = BASE_DIR.parent / 'rehash-automation' / 'config.json'
    config = json.loads(config_path.read_text())
    sp = config.setdefault('sharepoint', {})
    for field in ('demo_sheet_url', 'demo_sheet', 'file_url'):
        if field in data:
            sp[field] = data[field]
    config_path.write_text(json.dumps(config, indent=2))
    return jsonify(sp)

@app.route('/api/rehash/employees/<emp_id>', methods=['DELETE'])
def delete_rehash_employee(emp_id):
    config_path = BASE_DIR.parent / 'rehash-automation' / 'config.json'
    config = json.loads(config_path.read_text())
    config['employees'] = [e for e in config['employees'] if e['id'] != emp_id]
    config_path.write_text(json.dumps(config, indent=2))
    return jsonify({'ok': True})


# ── Commission config ───────────────────────────────────────────────────────────

COMMISSION_CONFIG_FILE = BASE_DIR / 'commission_config.json'

def load_commission_config():
    if COMMISSION_CONFIG_FILE.exists():
        return json.loads(COMMISSION_CONFIG_FILE.read_text())
    return {'reps': []}

def save_commission_config(cfg):
    COMMISSION_CONFIG_FILE.write_text(json.dumps(cfg, indent=2))

@app.route('/api/commission/reps', methods=['GET'])
def get_commission_reps():
    return jsonify(load_commission_config().get('reps', []))

@app.route('/api/commission/reps', methods=['POST'])
def add_commission_rep():
    data = request.json
    if not data.get('name'):
        return jsonify({'error': 'Name required'}), 400
    cfg = load_commission_config()
    rep = {
        'id':        data['name'].lower().replace(' ', '_'),
        'name':      data['name'],
        'excel_url': data.get('excel_url', ''),
    }
    cfg['reps'].append(rep)
    save_commission_config(cfg)
    return jsonify(rep), 201

@app.route('/api/commission/reps/<rep_id>', methods=['PATCH'])
def update_commission_rep(rep_id):
    data = request.json
    cfg  = load_commission_config()
    for rep in cfg['reps']:
        if rep['id'] == rep_id:
            for field in ('name', 'excel_url'):
                if field in data:
                    rep[field] = data[field]
            save_commission_config(cfg)
            return jsonify(rep)
    return jsonify({'error': 'Not found'}), 404

@app.route('/api/commission/reps/<rep_id>', methods=['DELETE'])
def delete_commission_rep(rep_id):
    cfg = load_commission_config()
    cfg['reps'] = [r for r in cfg['reps'] if r['id'] != rep_id]
    save_commission_config(cfg)
    return jsonify({'ok': True})


# ── Commission Estimate (PDF upload → annotated PDF download) ──────────────────

@app.route('/api/commission-estimate/process', methods=['POST'])
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

        venv_python = BASE_DIR / '.venv' / 'bin' / 'python3'
        script      = COMM_EST_DIR / 'commission_annotator.py'

        cmd = [str(venv_python), str(script),
               '--estimate',   str(est_path),
               '--commission', str(comm_path),
               '--output-dir', str(out_dir)]

        if finance_fee:
            cmd += ['--ff', finance_fee]
        if finance_lender:
            cmd += ['--ff-name', finance_lender]

        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

        if result.returncode != 0:
            err = (result.stderr or result.stdout or 'Processing failed').strip()
            return jsonify({'error': err}), 500

        output_pdfs = list(out_dir.glob('*.pdf'))
        if not output_pdfs:
            return jsonify({'error': 'No output PDFs were generated. Check that the input files are valid CRM PDFs.'}), 500

        # ── Parse stdout for alerts, Sean flag, and summary values ────────────
        stdout = result.stdout or ''
        alerts  = []
        is_sean = False
        summary = {}

        for line in stdout.splitlines():
            s = line.strip()
            # Alerts section
            if 'inform Anne' in s or ('⚠️' in s and '%' in s):
                clean = s.lstrip('⚠️ -').strip()
                if clean and clean not in alerts:
                    alerts.append(clean)
            # Sean detection
            rep_m = _re.search(r'Rep\s*:\s*(.+)', s)
            if rep_m and 'sean' in rep_m.group(1).lower():
                is_sean = True
            # Summary values
            def _extract(label, key):
                m = _re.search(rf'{label}\s*:\s*([^\n]+)', s, _re.IGNORECASE)
                if m: summary[key] = m.group(1).strip()
            _extract('Profit Margin', 'profit_margin')
            _extract('% Greenline',   'pct_greenline')
            _extract('Commission %',  'commission_pct')
            _extract('Est\. Commission', 'est_commission')
            _extract('Greenline(?!\s*%)', 'greenline')

        # Bundle into zip → base64 so we can return JSON + metadata together
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for pdf in output_pdfs:
                zf.write(str(pdf), pdf.name)
        zip_buf.seek(0)
        zip_b64 = base64.b64encode(zip_buf.read()).decode('utf-8')

        save_history({'module': 'commission_estimate', 'ran_at': datetime.now().isoformat(),
                      'status': 'success'})

        return jsonify({
            'ok':      True,
            'zip_b64': zip_b64,
            'alerts':  alerts,
            'is_sean': is_sean,
            'summary': summary,
        })

    except subprocess.TimeoutExpired:
        return jsonify({'error': 'Processing timed out (>120 s). Check that the PDFs are not corrupted.'}), 500
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'detail': traceback.format_exc()}), 500
    finally:
        import shutil
        shutil.rmtree(str(tmp_dir), ignore_errors=True)


if __name__ == '__main__':
    print('\n  Payless Automation Hub → http://localhost:5000\n')
    app.run(port=5050, debug=False)
