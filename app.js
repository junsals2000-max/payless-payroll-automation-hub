/* ══════════════════════════════════════════════
   Payless Automation Hub — app.js
   ══════════════════════════════════════════════ */

const API = 'http://localhost:5050';

// Modules that don't need Chrome automation
const CHROME_FREE_MODULES = new Set(['commission_estimate']);

// Commission Estimate uses its own configurable API URL (stored in localStorage)
// so it works from Vercel when the local server is tunnelled or on the same LAN.
function getCeApiBase() {
  return (localStorage.getItem('ce_api_url') || 'https://scanning-headfirst-posing.ngrok-free.dev').replace(/\/$/, '');
}
function saveCeApiBase(url) {
  localStorage.setItem('ce_api_url', url.trim().replace(/\/$/, ''));
}

// ── Init ───────────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
  setGreeting();
  setDateChip();
  loadDashboard();
  pollChromeStatus();
  updateChromeModalText();
});

// ── User profile (localStorage) ────────────────────────────────────────────────

function getUserProfile() {
  return {
    name:  localStorage.getItem('user_name')  || 'there',
    email: localStorage.getItem('user_email') || '',
  };
}

function openSettings() {
  const { name, email } = getUserProfile();
  const nameEl  = document.getElementById('settingsName');
  const emailEl = document.getElementById('settingsEmail');
  if (nameEl)  nameEl.value  = name  === 'there' ? '' : name;
  if (emailEl) emailEl.value = email;
  document.getElementById('settingsOverlay').style.display = 'flex';
}

function closeSettings() {
  document.getElementById('settingsOverlay').style.display = 'none';
}

function saveSettings() {
  const name  = document.getElementById('settingsName')?.value.trim();
  const email = document.getElementById('settingsEmail')?.value.trim();
  if (name)  localStorage.setItem('user_name',  name);
  if (email) localStorage.setItem('user_email', email);
  closeSettings();
  setGreeting();
  updateChromeModalText();
}

function updateChromeModalText() {
  const { name, email } = getUserProfile();
  const desc = document.getElementById('modalUserDesc');
  if (!desc) return;
  if (name && email) {
    desc.textContent = `${name}'s sessions (${email})`;
  } else if (email) {
    desc.textContent = `your sessions (${email})`;
  } else if (name) {
    desc.textContent = `${name}'s sessions`;
  } else {
    desc.textContent = 'your sessions';
  }
}

function setGreeting() {
  const hour  = new Date().getHours();
  const { name } = getUserProfile();
  const greet = hour < 12 ? 'Good morning' : hour < 17 ? 'Good afternoon' : 'Good evening';
  const firstName = name.split(' ')[0];
  document.getElementById('greetingLine').textContent = `${greet}, ${firstName} 👋`;
}

function setDateChip() {
  const d = document.getElementById('dateChip');
  if (d) d.textContent = new Date().toLocaleDateString('en-US', {
    weekday: 'long', month: 'long', day: 'numeric', year: 'numeric'
  });
}


// ── View routing ───────────────────────────────────────────────────────────────

function showView(name) {
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active-view'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  const view = document.getElementById(`view-${name}`);
  const nav  = document.getElementById(`nav-${name}`);
  if (view) view.classList.add('active-view');
  if (nav)  nav.classList.add('active');
}


// ── Dashboard ──────────────────────────────────────────────────────────────────

async function loadDashboard() {
  try {
    const [modules, history] = await Promise.all([
      fetch(`${API}/api/modules`).then(r => r.json()),
      fetch(`${API}/api/history`).then(r => r.json()),
    ]);

    renderModuleGrid(modules);
    renderActivity(history);
    updateStats(history);
  } catch {
    document.getElementById('moduleGrid').innerHTML =
      '<p style="color:var(--text-3);font-size:13px">Could not connect to server. Is the hub running?</p>';
  }
}

function renderModuleGrid(modules) {
  const grid = document.getElementById('moduleGrid');
  grid.innerHTML = modules.map(m => {
    const isActive = m.status === 'active';
    const isSoon   = m.status === 'coming_soon';
    const lastRun  = m.last_run
      ? `Last run: ${timeAgo(m.last_run.ran_at)}`
      : 'Never run';

    const statusLabel = isSoon ? 'Coming Soon' : isActive ? 'Active' : 'Active';
    const statusClass = isSoon ? 'status-soon' : 'status-active';

    const sourceTags = (m.sources || []).map(s =>
      `<span class="source-tag">${s}</span>`
    ).join('');

    return `
      <div class="module-card ${isSoon ? 'module-soon' : ''}"
           onclick="${isActive ? `openModule('${m.id}')` : ''}">
        <div class="module-stripe" style="background:${m.color}"></div>
        <div class="module-body">
          <div class="module-top">
            <div class="module-icon-wrap" style="background:${m.color}22">
              <span>${m.icon}</span>
            </div>
            <span class="status-badge ${statusClass}">${statusLabel}</span>
          </div>
          <div class="module-name">${m.name}</div>
          <div class="module-desc">${m.description}</div>
          <div class="module-sources">${sourceTags}</div>
          <div class="module-meta">
            <span class="module-last-run">${lastRun}</span>
            ${isActive ? `<button class="btn-open" onclick="event.stopPropagation();openModule('${m.id}')">Open →</button>` : ''}
          </div>
        </div>
      </div>
    `;
  }).join('');
}

function renderActivity(history) {
  const list = document.getElementById('activityList');
  if (!history.length) return;

  list.innerHTML = history.slice(0, 8).map(h => {
    const ok = h.status === 'success';
    const moduleLabel = h.module === 'rehash'               ? 'Sales Report'
                      : h.module === 'commission'          ? 'Sales Commission'
                      : h.module === 'commission_estimate' ? 'Commission Estimate'
                      : h.module;
    return `
      <div class="activity-row">
        <div class="activity-dot ${ok ? 'success' : 'error'}"></div>
        <div class="activity-info">
          <div class="activity-name">${moduleLabel}</div>
          <div class="activity-time">${timeAgo(h.ran_at)}</div>
        </div>
        <span class="activity-status ${ok ? 'success' : 'error'}">${ok ? 'Success' : 'Failed'}</span>
      </div>
    `;
  }).join('');
}

function updateStats(history) {
  const success = history.filter(h => h.status === 'success').length;
  document.getElementById('statTotal').textContent   = history.length;
  document.getElementById('statSuccess').textContent = success;
}


// ── Module view ────────────────────────────────────────────────────────────────

let activeModule   = null;
let currentJobSrc  = null;
let currentJobId   = null;
let chromeReady    = false;

async function openModule(moduleId) {
  activeModule = moduleId;

  // Restore Chrome pill unless switching to a chrome-free module
  const pill = document.getElementById('chromePill');
  if (pill) pill.style.display = CHROME_FREE_MODULES.has(moduleId) ? 'none' : '';

  // Set sidebar active
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  const nav = document.getElementById(`nav-${moduleId}`);
  if (nav) nav.classList.add('active');

  showView('module');

  if (moduleId === 'rehash')               await renderRehashModule();
  if (moduleId === 'commission')           await renderCommissionModule();
  if (moduleId === 'commission_estimate')  renderCommEstimateModule();
}

async function renderRehashModule() {
  const view = document.getElementById('view-module');
  view.innerHTML = `
    <div class="module-view-header">
      <button class="back-btn" onclick="goBack()">← Back</button>
      <div>
        <div class="mv-title">📊 Rehash Report</div>
        <div class="mv-sub">Weekly Jive + LeadPerfection → SharePoint Excel</div>
      </div>
    </div>

    <div class="module-view-body">

      <!-- LEFT: run panel + log -->
      <div style="display:flex;flex-direction:column;gap:16px">

        <div class="run-panel">
          <div class="run-panel-header">
            <span class="run-panel-title">Run Automation</span>
            <span class="status-badge status-active">Active</span>
          </div>

          <div class="run-steps">
            <div class="run-step" id="rs-chrome">
              <div class="step-circle">1</div>
              <div class="step-info">
                <div class="step-label">Automation Chrome</div>
                <div class="step-hint" id="rs-chrome-hint">Close main Chrome first, then click Connect</div>
              </div>
              <button class="btn-chrome" id="rs-chrome-btn" onclick="handleChromeConnect()">Connect</button>
            </div>

            <div class="run-step" id="rs-run">
              <div class="step-circle">2</div>
              <div class="step-info">
                <div class="step-label">Run the report</div>
                <div class="step-hint">Jive → LeadPerfection → Excel</div>
              </div>
            </div>
          </div>

          <div style="display:flex;gap:8px;margin:0 22px 20px">
            <button class="btn-run-module" id="runModuleBtn" onclick="runModule()" disabled style="flex:1;margin:0">
              ▶ Run Rehash Report
            </button>
            <button class="btn-stop" id="stopModuleBtn" onclick="stopRun()" style="display:none">
              ⏹ Stop
            </button>
          </div>
        </div>

        <div class="log-wrap">
          <div class="log-header">
            <span class="log-title">Live Log</span>
            <span class="log-pill" id="logPill">Idle</span>
          </div>
          <div class="log-body" id="moduleLog">
            <span class="log-idle">Waiting to run…</span>
          </div>
        </div>
      </div>

      <!-- RIGHT: info + employees -->
      <div class="side-panel" id="sidePanelRehash"></div>
    </div>
  `;

  // Reflect current Chrome state
  syncChromeUI();
  await loadRehashSidePanel();
}

async function loadRehashSidePanel() {
  const panel = document.getElementById('sidePanelRehash');
  if (!panel) return;

  try {
    const cfg = await fetch(`${API}/api/modules`).then(r => r.json());
    const mod = cfg.find(m => m.id === 'rehash');

    // Fetch employee list directly from hub API
    let employees = [];
    try {
      employees = await fetch(`${API}/api/rehash/employees`).then(r => r.json());
    } catch {}

    const lastWeek   = getLastWeekLabel();
    const spCfg      = await fetch(`${API}/api/rehash/config`).then(r => r.json()).catch(() => ({}));
    const demoUrl    = spCfg.demo_sheet_url || '';

    panel.innerHTML = `
      <div class="info-card">
        <div class="info-card-title">Report Details</div>
        <div class="info-row"><span class="info-key">Week</span><span class="info-value">${lastWeek}</span></div>
        <div class="info-row"><span class="info-key">Report</span><span class="info-value">Q2 Rehash Report</span></div>
        <div class="info-row"><span class="info-key">Sources</span><span class="info-value">Jive + LeadPerfection</span></div>
        <div class="info-row"><span class="info-key">Output</span><span class="info-value">SharePoint Excel</span></div>
        <div class="info-row"><span class="info-key">Schedule</span><span class="info-value">Every Monday</span></div>
      </div>

      <!-- # of Demo Sheet URL -->
      <div class="info-card">
        <div class="info-card-title">SharePoint Settings</div>
        <div style="margin-bottom:6px">
          <div style="font-size:11px;color:var(--text-3);margin-bottom:4px;font-weight:500"># of Demo Sheet Link</div>
          <div style="display:flex;gap:6px;align-items:center">
            <input id="demoSheetUrlInput"
              value="${demoUrl}"
              placeholder="Paste 'Copy link to this sheet' URL…"
              style="${inputStyle};flex:1" />
            <button onclick="saveDemoSheetUrl()"
              style="background:var(--navy);color:#fff;border:none;border-radius:6px;padding:5px 10px;font-size:11px;font-weight:600;cursor:pointer;white-space:nowrap">
              Save
            </button>
          </div>
          <div id="demoUrlStatus" style="font-size:10px;color:var(--text-3);margin-top:3px">
            ${demoUrl ? '✓ URL saved' : 'No URL set — Step 4 will skip the Demo sheet'}
          </div>
        </div>
      </div>

      <div class="info-card">
        <div class="info-card-title" style="display:flex;justify-content:space-between;align-items:center">
          <span>Employees (${employees.length})</span>
          <button class="btn-open" onclick="toggleAddEmpForm()" id="addEmpToggle">+ Add</button>
        </div>

        <!-- Add employee form (hidden by default) -->
        <div id="addEmpForm" style="display:none;margin-bottom:12px;padding:12px;background:#f8fafc;border:1px solid var(--border);border-radius:8px">
          <div style="display:flex;flex-direction:column;gap:7px">
            <input id="newEmpName"  placeholder="Full Name"                        style="${inputStyle}" />
            <input id="newEmpUrl"   placeholder="Excel Link (Copy link to sheet)"  style="${inputStyle}" />
            <input id="newEmpJive"  placeholder="Jive URL"                         style="${inputStyle}" />
            <input id="newEmpLp"    placeholder="Name in LeadPerfection"           style="${inputStyle}" />
          </div>
          <button onclick="addRehashEmployee()" style="margin-top:8px;width:100%;background:var(--navy);color:#fff;border:none;border-radius:7px;padding:7px;font-size:12px;font-weight:600;cursor:pointer">
            Save Employee
          </button>
        </div>

        <div class="emp-list-mini" id="empListMini">
          ${renderEmpList(employees)}
        </div>
      </div>
    `;
  } catch {}
}

function goBack() {
  // Restore Chrome pill when leaving chrome-free modules
  const pill = document.getElementById('chromePill');
  if (pill) pill.style.display = '';
  showView('dashboard');
  loadDashboard();
}


// ── Employee manager (Rehash side panel) ──────────────────────────────────────

const inputStyle = 'width:100%;padding:6px 10px;border:1px solid var(--border);border-radius:6px;font-size:12px;outline:none';

function renderEmpList(employees) {
  if (!employees.length) {
    return '<p style="font-size:12px;color:var(--text-3);padding:4px 0">No employees yet. Click + Add above.</p>';
  }
  return employees.map(e => {
    const ini       = e.name.split(' ').map(p => p[0]).join('').slice(0, 2).toUpperCase();
    const lpDisplay = e.lp_name && e.lp_name !== e.name ? e.lp_name : '—';
    const urlSet    = e.excel_url && e.excel_url.startsWith('http');
    const urlLabel  = urlSet ? '✓ Excel link saved' : '⚠ No Excel link';
    const urlColor  = urlSet ? '#16a34a' : '#f59e0b';
    return `
      <div class="emp-mini" style="flex-wrap:wrap;gap:4px">
        <div class="emp-av">${ini}</div>
        <div style="flex:1;min-width:0">
          <div class="emp-n">${e.name}</div>
          <div style="font-size:10px;color:${urlColor};margin-top:1px">${urlLabel}</div>
          <div style="display:flex;align-items:center;gap:4px;margin-top:3px">
            <span style="font-size:10px;color:var(--text-3);white-space:nowrap">LP:</span>
            <span style="font-size:10px;color:var(--text-2);flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${lpDisplay}</span>
            <button onclick="openEmpEdit('${e.id}')"
              style="background:none;border:none;color:var(--text-3);font-size:10px;cursor:pointer;padding:0 2px;flex-shrink:0"
              title="Edit employee">✏️</button>
          </div>
          <div id="emp-edit-${e.id}" style="display:none;margin-top:6px;padding:8px;background:#f8fafc;border:1px solid var(--border);border-radius:8px">
            <div style="display:flex;flex-direction:column;gap:5px">
              <input id="edit-name-${e.id}"  value="${e.name}"             placeholder="Full Name"                       style="${inputStyle}" />
              <input id="edit-url-${e.id}"   value="${e.excel_url || ''}"  placeholder="Excel Link (Copy link to sheet)" style="${inputStyle}" />
              <input id="edit-lp-${e.id}"    value="${e.lp_name || ''}"    placeholder="Name in LeadPerfection"          style="${inputStyle}" />
              <input id="edit-jive-${e.id}"  value="${e.jive_url || ''}"   placeholder="Jive URL"                        style="${inputStyle}" />
            </div>
            <div style="display:flex;gap:4px;margin-top:6px">
              <button onclick="saveEmpEdit('${e.id}')"
                style="flex:1;background:var(--navy);color:#fff;border:none;border-radius:5px;padding:4px;font-size:11px;font-weight:600;cursor:pointer">Save</button>
              <button onclick="cancelEmpEdit('${e.id}')"
                style="flex:1;background:none;border:1px solid var(--border);border-radius:5px;padding:4px;font-size:11px;cursor:pointer">Cancel</button>
            </div>
          </div>
        </div>
        <button onclick="removeRehashEmployee('${e.id}')"
          style="background:none;border:1px solid #fca5a5;border-radius:5px;color:#ef4444;font-size:10px;padding:2px 7px;cursor:pointer;align-self:start">
          ✕
        </button>
      </div>`;
  }).join('');
}

function openEmpEdit(empId) {
  document.getElementById(`emp-edit-${empId}`).style.display = 'block';
  document.getElementById(`edit-name-${empId}`)?.focus();
}

function cancelEmpEdit(empId) {
  document.getElementById(`emp-edit-${empId}`).style.display = 'none';
}

async function saveEmpEdit(empId) {
  const name = document.getElementById(`edit-name-${empId}`)?.value.trim();
  const url  = document.getElementById(`edit-url-${empId}`)?.value.trim();
  const lp   = document.getElementById(`edit-lp-${empId}`)?.value.trim();
  const jive = document.getElementById(`edit-jive-${empId}`)?.value.trim();
  if (!name) { alert('Name is required'); return; }
  await fetch(`${API}/api/rehash/employees/${empId}`, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, excel_url: url, lp_name: lp, jive_url: jive })
  });
  const employees = await fetch(`${API}/api/rehash/employees`).then(r => r.json());
  const listEl = document.getElementById('empListMini');
  if (listEl) listEl.innerHTML = renderEmpList(employees);
}

function toggleAddEmpForm() {
  const f = document.getElementById('addEmpForm');
  if (f) f.style.display = f.style.display === 'none' ? 'block' : 'none';
}

async function addRehashEmployee() {
  const name = document.getElementById('newEmpName')?.value.trim();
  const url  = document.getElementById('newEmpUrl')?.value.trim();
  const jive = document.getElementById('newEmpJive')?.value.trim();
  const lp   = document.getElementById('newEmpLp')?.value.trim();
  if (!name) { alert('Name is required'); return; }

  await fetch(`${API}/api/rehash/employees`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, excel_url: url, jive_url: jive, lp_name: lp || name })
  });

  // Refresh list
  const employees = await fetch(`${API}/api/rehash/employees`).then(r => r.json());
  const listEl = document.getElementById('empListMini');
  if (listEl) listEl.innerHTML = renderEmpList(employees);

  // Update title count
  const titleEl = document.querySelector('#addEmpToggle')?.closest('.info-card')?.querySelector('.info-card-title span');
  if (titleEl) titleEl.textContent = `Employees (${employees.length})`;

  // Clear & hide form
  ['newEmpName','newEmpUrl','newEmpJive','newEmpLp'].forEach(id => {
    const el = document.getElementById(id); if (el) el.value = '';
  });
  toggleAddEmpForm();
}

async function removeRehashEmployee(empId) {
  if (!confirm('Remove this employee?')) return;
  await fetch(`${API}/api/rehash/employees/${empId}`, { method: 'DELETE' });
  const employees = await fetch(`${API}/api/rehash/employees`).then(r => r.json());
  const listEl = document.getElementById('empListMini');
  if (listEl) listEl.innerHTML = renderEmpList(employees);
}

async function saveDemoSheetUrl() {
  const input  = document.getElementById('demoSheetUrlInput');
  const status = document.getElementById('demoUrlStatus');
  if (!input) return;
  const url = input.value.trim();
  try {
    await fetch(`${API}/api/rehash/config`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ demo_sheet_url: url })
    });
    if (status) {
      status.textContent = url ? '✓ URL saved' : 'No URL set — Step 4 will skip the Demo sheet';
      status.style.color = url ? '#16a34a' : 'var(--text-3)';
    }
  } catch (e) {
    if (status) { status.textContent = 'Save failed — server error'; status.style.color = '#ef4444'; }
  }
}


// ── Commission Module ──────────────────────────────────────────────────────────

const COMM_STEPS = [
  { id: 'search',    num: '1',   label: 'Search Client',          hint: 'Buildertrend lookup by client name' },
  { id: 'job',       num: '2',   label: 'Extract Job Details',    hint: 'Inv# · Title · Sold Date · Contract Price' },
  { id: 'pdf',       num: '3',   label: 'Extract Commission PDF', hint: 'Greenline · % GL · Commission · Finance' },
  { id: 'rep',       num: '4',   label: 'Sales Rep Mapping',      hint: 'Match rep to Excel spreadsheet' },
  { id: 'validate',  num: '5',   label: 'Final Data Validation',  hint: 'Ensure all 9 fields are present' },
  { id: 'duplicate', num: '5.5', label: 'Duplicate Check',        hint: 'Check for existing entry in Excel' },
  { id: 'insert',    num: '6',   label: 'Insert Row into Excel',  hint: 'Columns A–I mapped and written' },
  { id: 'formulas',  num: '7',   label: 'Update SUM Formulas',    hint: 'Extend formula range to include new row' },
];

let commStepStates  = {};
let commExtracted   = {};

async function renderCommissionModule() {
  const view = document.getElementById('view-module');
  view.innerHTML = `
    <div class="module-view-header">
      <button class="back-btn" onclick="goBack()">← Back</button>
      <div>
        <div class="mv-title">💵 Sales Commission</div>
        <div class="mv-sub">Buildertrend · Commission PDF → SharePoint Excel</div>
      </div>
    </div>

    <div class="module-view-body">

      <!-- LEFT -->
      <div style="display:flex;flex-direction:column;gap:16px">

        <!-- Input card -->
        <div class="run-panel">
          <div class="run-panel-header" style="padding:13px 22px">
            <span class="run-panel-title">Client Lookup</span>
            <span class="status-badge status-active">Active</span>
          </div>
          <div style="padding:14px 22px 18px;display:flex;flex-direction:column;gap:10px">
            <div>
              <label class="comm-field-label">Client Name</label>
              <input id="comm-client-name" placeholder="e.g. Smith, John"
                class="comm-input" oninput="commInputChanged()" />
            </div>
            <div style="display:flex;gap:8px">
              <button class="btn-run-module" id="commRunBtn"
                onclick="commStartRun()" disabled
                style="margin:0;flex:1;font-size:14px;padding:11px">
                ▶ Search &amp; Process Commission
              </button>
              <button class="btn-stop" id="commResetBtn"
                onclick="commReset()" style="display:none;padding:10px 16px;font-size:12px">
                ↺ Reset
              </button>
            </div>
          </div>
        </div>

        <!-- Step tracker -->
        <div class="log-wrap" style="overflow:visible">
          <div class="log-header">
            <span class="log-title">Processing Steps</span>
            <span class="log-pill" id="commStatusPill">Idle</span>
          </div>
          <div style="padding:20px 22px 8px" id="commStepperWrap">
            ${buildCommStepper()}
          </div>
        </div>

        <!-- Data preview (hidden until data arrives) -->
        <div class="info-card" id="commDataCard" style="display:none">
          <div class="info-card-title" style="display:flex;align-items:center;justify-content:space-between">
            <span>Extracted Data</span>
            <span id="commValidBadge" class="status-badge" style="display:none"></span>
          </div>
          <div class="comm-data-grid" id="commDataGrid"></div>
        </div>

      </div>

      <!-- RIGHT -->
      <div class="side-panel" id="commSidePanel"></div>
    </div>
  `;

  commReset(true);
  await loadCommSidePanel();
}

function buildCommStepper() {
  return COMM_STEPS.map((step, i) => {
    const isLast = i === COMM_STEPS.length - 1;
    return `
      <div class="comm-step-row pending" id="csr-${step.id}">
        <div class="comm-step-left">
          <div class="comm-step-node" id="csn-${step.id}">${step.num}</div>
          ${!isLast ? `<div class="comm-step-line" id="csl-${step.id}"></div>` : ''}
        </div>
        <div class="comm-step-content" id="csc-${step.id}">
          <div class="comm-step-label">${step.label}</div>
          <div class="comm-step-hint" id="csh-${step.id}">${step.hint}</div>
        </div>
      </div>`;
  }).join('');
}

function setCommStep(stepId, state, hintText) {
  commStepStates[stepId] = state;
  const row  = document.getElementById(`csr-${stepId}`);
  const node = document.getElementById(`csn-${stepId}`);
  const hint = document.getElementById(`csh-${stepId}`);
  const line = document.getElementById(`csl-${stepId}`);
  const step = COMM_STEPS.find(s => s.id === stepId);
  if (!row || !node) return;

  row.className = `comm-step-row ${state}`;
  node.textContent = state === 'done' ? '✓' : state === 'error' ? '✗' : state === 'warning' ? '!' : step?.num || '?';
  if (hint && hintText) hint.textContent = hintText;
  if (line) line.style.background = state === 'done' ? 'var(--green)' : state === 'error' ? 'var(--red)' : state === 'warning' ? 'var(--amber)' : '';

  // Inject duplicate alert inline
  if (stepId === 'duplicate' && state === 'warning') {
    const content = document.getElementById(`csc-${stepId}`);
    if (content && !document.getElementById('commDupAlert')) {
      const alert = document.createElement('div');
      alert.className = 'comm-dup-alert';
      alert.id = 'commDupAlert';
      alert.innerHTML = `
        <div style="display:flex;gap:10px;align-items:flex-start">
          <span style="font-size:20px;line-height:1">⚠️</span>
          <div>
            <div style="font-weight:700;font-size:13px;color:#92400e;margin-bottom:4px">Duplicate Record Found</div>
            <div style="font-size:12px;color:#78350f" id="commDupDetail">${hintText || 'An existing entry matches this client / invoice number.'}</div>
            <div style="display:flex;gap:8px;margin-top:12px">
              <button class="comm-dup-btn proceed" onclick="commDupDecide('proceed')">✓ Proceed — Insert Anyway</button>
              <button class="comm-dup-btn skip"    onclick="commDupDecide('skip')">✗ Skip — Don't Insert</button>
            </div>
          </div>
        </div>`;
      content.appendChild(alert);
    }
  }
}

function commDupDecide(decision) {
  const alert = document.getElementById('commDupAlert');
  if (decision === 'skip') {
    if (alert) alert.innerHTML = '<span style="color:#dc2626;font-weight:600;font-size:12px">✗ Skipped — entry not inserted.</span>';
    setCommStep('duplicate', 'error', 'Skipped by user — entry not inserted');
    commFinish(false, 'Skipped');
  } else {
    if (alert) alert.innerHTML = '<span style="color:#16a34a;font-weight:600;font-size:12px">✓ Proceeding to insert…</span>';
    setCommStep('duplicate', 'done', 'Duplicate acknowledged — proceeding');
    // Step 6 & 7 will be wired in automation phase
  }
}

function commInputChanged() {
  const name = document.getElementById('comm-client-name')?.value.trim();
  const btn  = document.getElementById('commRunBtn');
  if (btn) btn.disabled = !name;
}

async function commStartRun() {
  const btn   = document.getElementById('commRunBtn');
  const reset = document.getElementById('commResetBtn');
  const pill  = document.getElementById('commStatusPill');
  if (btn)   { btn.disabled = true; btn.textContent = '⏳ Processing…'; btn.className = 'btn-run-module running'; btn.style.margin = '0'; }
  if (reset) reset.style.display = 'block';
  if (pill)  { pill.textContent = 'Running…'; pill.className = 'log-pill running'; }

  const clientName = document.getElementById('comm-client-name')?.value.trim();

  try {
    const res = await fetch(`${API}/api/run/commission`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ client_name: clientName }),
    });
    const { job_id } = await res.json();

    const evtSrc = new EventSource(`${API}/api/status/${job_id}`);

    evtSrc.onmessage = (e) => {
      const data = JSON.parse(e.data);
      if (data.type === 'ping') return;

      if (data.type === 'step') {
        setCommStep(data.step, data.state, data.hint);
        if (data.state === 'warning' && data.step === 'duplicate') {
          if (pill) { pill.textContent = 'Awaiting input'; pill.className = 'log-pill'; }
        }
      }

      if (data.type === 'log') {
        // Surface important log lines as step hints (optional)
      }

      if (data.type === 'done') {
        evtSrc.close();
        commFinish(true, 'Done');
      }

      if (data.type === 'error') {
        evtSrc.close();
        // Find the last active step and mark it error
        const activeStep = COMM_STEPS.find(s => commStepStates[s.id] === 'active');
        if (activeStep) setCommStep(activeStep.id, 'error', data.msg?.split('\n')[0]);
        commFinish(false, 'Error');
        if (pill) { pill.textContent = 'Error'; pill.className = 'log-pill error'; }
        alert(`Error:\n${data.msg}`);
      }
    };

    evtSrc.onerror = () => {
      evtSrc.close();
      commFinish(false, 'Lost connection');
    };

  } catch (e) {
    commFinish(false, 'Error');
    alert(`Failed to start: ${e.message}`);
  }
}

function commFinish(success, label) {
  const btn  = document.getElementById('commRunBtn');
  const pill = document.getElementById('commStatusPill');
  if (btn) {
    btn.disabled = false; btn.style.margin = '0';
    btn.className   = 'btn-run-module ' + (success ? 'done' : 'errored');
    btn.textContent = success ? '✓ Complete — Run Another' : '✗ Skipped — Reset to try again';
  }
  if (pill) { pill.textContent = label || (success ? 'Done' : 'Error'); pill.className = 'log-pill ' + (success ? 'done' : 'error'); }
}

function commReset(silent) {
  commStepStates = {}; commExtracted = {};
  const wrap = document.getElementById('commStepperWrap');
  if (wrap) wrap.innerHTML = buildCommStepper();
  const card = document.getElementById('commDataCard');
  if (card) card.style.display = 'none';
  const pill = document.getElementById('commStatusPill');
  if (pill) { pill.textContent = 'Idle'; pill.className = 'log-pill'; }
  const btn = document.getElementById('commRunBtn');
  if (btn) { btn.className = 'btn-run-module'; btn.textContent = '▶ Search & Process Commission'; btn.style.margin = '0'; commInputChanged(); }
  const resetBtn = document.getElementById('commResetBtn');
  if (resetBtn) resetBtn.style.display = 'none';
}

function updateCommDataGrid() {
  const card = document.getElementById('commDataCard');
  const grid = document.getElementById('commDataGrid');
  if (!card || !grid) return;
  const fields = [
    { key: 'inv_number',    label: 'Inv #',                  col: 'A' },
    { key: 'client',        label: 'Client',                 col: 'B' },
    { key: 'order_date',    label: 'Order Date',             col: 'C' },
    { key: 'invoice_total', label: 'Invoice Total',          col: 'D', money: true },
    { key: 'greenline',     label: 'Greenline',              col: 'E', money: true },
    { key: 'pct_gl',        label: '% of GL',                col: 'F', pct: true },
    { key: 'comm_rate',     label: 'Comm Rate',              col: 'G', pct: true },
    { key: 'total_comm',    label: 'Total Anticipated Comm', col: 'H', money: true },
    { key: 'finance_cont',  label: 'Finance / Contractor',   col: 'I', money: true },
  ];
  const hasAny = fields.some(f => commExtracted[f.key] !== undefined);
  if (!hasAny) { card.style.display = 'none'; return; }
  card.style.display = 'block';
  grid.innerHTML = fields.map(f => {
    const raw = commExtracted[f.key];
    let val = raw;
    if (raw === undefined || raw === null) { val = '—'; }
    else if (f.money) val = '$' + Number(raw).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    else if (f.pct)   val = (raw * 100).toFixed(1) + '%';
    const filled = raw !== undefined && raw !== null;
    return `<div class="comm-data-row">
      <span class="comm-data-col">Col ${f.col}</span>
      <span class="comm-data-key">${f.label}</span>
      <span class="comm-data-val ${filled ? 'filled' : ''}">${val}</span>
    </div>`;
  }).join('');
  const badge = document.getElementById('commValidBadge');
  if (badge && hasAny) { badge.style.display = ''; badge.className = 'status-badge status-active'; badge.textContent = '✓ Valid'; }
}

// ── Commission side panel ──────────────────────────────────────────────────────

async function loadCommSidePanel() {
  const panel = document.getElementById('commSidePanel');
  if (!panel) return;
  let reps = [];
  try { reps = await fetch(`${API}/api/commission/reps`).then(r => r.json()); } catch {}
  panel.innerHTML = `
    <div class="info-card">
      <div class="info-card-title">Module Details</div>
      <div class="info-row"><span class="info-key">Trigger</span><span class="info-value">Per client</span></div>
      <div class="info-row"><span class="info-key">Source 1</span><span class="info-value">Buildertrend</span></div>
      <div class="info-row"><span class="info-key">Source 2</span><span class="info-value">Commission PDF</span></div>
      <div class="info-row"><span class="info-key">Output</span><span class="info-value">SharePoint Excel</span></div>
      <div class="info-row"><span class="info-key">Columns</span><span class="info-value">A – I</span></div>
    </div>

    <div class="info-card">
      <div class="info-card-title">Output Columns</div>
      ${[['A','Inv #'],['B','Client'],['C','Order Date'],['D','Invoice Total'],['E','Greenline'],['F','% of GL'],['G','Comm Rate'],['H','Total Anticipated Comm'],['I','Finance / Contractor']].map(([col,name])=>`
        <div class="info-row">
          <span class="info-key" style="font-family:monospace;background:#f1f5f9;padding:1px 6px;border-radius:4px;font-size:10px;flex-shrink:0">Col ${col}</span>
          <span class="info-value">${name}</span>
        </div>`).join('')}
    </div>

    <div class="info-card">
      <div class="info-card-title" style="display:flex;justify-content:space-between;align-items:center">
        <span>Sales Rep Mapping (${reps.length})</span>
        <button class="btn-open" onclick="toggleCommRepForm()">+ Add</button>
      </div>
      <div id="commAddRepForm" style="display:none;margin-bottom:12px;padding:12px;background:#f8fafc;border:1px solid var(--border);border-radius:8px">
        <div style="display:flex;flex-direction:column;gap:7px">
          <input id="newRepName" placeholder="Sales Rep Name"                style="${inputStyle}" />
          <input id="newRepUrl"  placeholder="Excel Link (Copy link to sheet)" style="${inputStyle}" />
        </div>
        <button onclick="addCommRep()" style="margin-top:8px;width:100%;background:var(--navy);color:#fff;border:none;border-radius:7px;padding:7px;font-size:12px;font-weight:600;cursor:pointer">Save Rep</button>
      </div>
      <div id="commRepList">${renderCommRepList(reps)}</div>
    </div>
  `;
}

function renderCommRepList(reps) {
  if (!reps.length) return '<p style="font-size:12px;color:var(--text-3);padding:4px 0">No reps configured. Click + Add above.</p>';
  return reps.map(r => {
    const ini    = r.name.split(' ').map(p=>p[0]).join('').slice(0,2).toUpperCase();
    const urlSet = r.excel_url && r.excel_url.startsWith('http');
    return `
      <div class="emp-mini" style="flex-wrap:wrap;gap:4px">
        <div class="emp-av" style="background:linear-gradient(135deg,#d97706,#f59e0b)">${ini}</div>
        <div style="flex:1;min-width:0">
          <div class="emp-n">${r.name}</div>
          <div style="font-size:10px;color:${urlSet?'#16a34a':'#f59e0b'};margin-top:1px">${urlSet?'✓ Excel link saved':'⚠ No Excel link'}</div>
          <div style="margin-top:3px">
            <button onclick="openCommRepEdit('${r.id}')"
              style="background:none;border:none;color:var(--text-3);font-size:10px;cursor:pointer;padding:0">✏️ Edit</button>
          </div>
          <div id="comm-rep-edit-${r.id}" style="display:none;margin-top:6px;padding:8px;background:#f8fafc;border:1px solid var(--border);border-radius:8px">
            <div style="display:flex;flex-direction:column;gap:5px">
              <input id="cedit-name-${r.id}" value="${r.name}"             placeholder="Name"       style="${inputStyle}" />
              <input id="cedit-url-${r.id}"  value="${r.excel_url||''}"    placeholder="Excel Link" style="${inputStyle}" />
            </div>
            <div style="display:flex;gap:4px;margin-top:6px">
              <button onclick="saveCommRepEdit('${r.id}')" style="flex:1;background:var(--navy);color:#fff;border:none;border-radius:5px;padding:4px;font-size:11px;font-weight:600;cursor:pointer">Save</button>
              <button onclick="document.getElementById('comm-rep-edit-${r.id}').style.display='none'" style="flex:1;background:none;border:1px solid var(--border);border-radius:5px;padding:4px;font-size:11px;cursor:pointer">Cancel</button>
            </div>
          </div>
        </div>
        <button onclick="deleteCommRep('${r.id}')" style="background:none;border:1px solid #fca5a5;border-radius:5px;color:#ef4444;font-size:10px;padding:2px 7px;cursor:pointer;align-self:start">✕</button>
      </div>`;
  }).join('');
}

function toggleCommRepForm() {
  const f = document.getElementById('commAddRepForm');
  if (f) f.style.display = f.style.display === 'none' ? 'block' : 'none';
}
function openCommRepEdit(id) { const el = document.getElementById(`comm-rep-edit-${id}`); if(el) el.style.display='block'; }

async function addCommRep() {
  const name = document.getElementById('newRepName')?.value.trim();
  const url  = document.getElementById('newRepUrl')?.value.trim();
  if (!name) { alert('Name is required'); return; }
  await fetch(`${API}/api/commission/reps`, { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({name, excel_url: url}) });
  const reps = await fetch(`${API}/api/commission/reps`).then(r=>r.json());
  const el = document.getElementById('commRepList'); if(el) el.innerHTML = renderCommRepList(reps);
  ['newRepName','newRepUrl'].forEach(id=>{ const e=document.getElementById(id); if(e) e.value=''; });
  toggleCommRepForm();
}
async function saveCommRepEdit(id) {
  const name = document.getElementById(`cedit-name-${id}`)?.value.trim();
  const url  = document.getElementById(`cedit-url-${id}`)?.value.trim();
  if (!name) { alert('Name required'); return; }
  await fetch(`${API}/api/commission/reps/${id}`, { method:'PATCH', headers:{'Content-Type':'application/json'}, body: JSON.stringify({name, excel_url: url}) });
  const reps = await fetch(`${API}/api/commission/reps`).then(r=>r.json());
  const el = document.getElementById('commRepList'); if(el) el.innerHTML = renderCommRepList(reps);
}
async function deleteCommRep(id) {
  if (!confirm('Remove this rep?')) return;
  await fetch(`${API}/api/commission/reps/${id}`, { method:'DELETE' });
  const reps = await fetch(`${API}/api/commission/reps`).then(r=>r.json());
  const el = document.getElementById('commRepList'); if(el) el.innerHTML = renderCommRepList(reps);
}

// ── Chrome connection ──────────────────────────────────────────────────────────

async function pollChromeStatus() {
  if (!CHROME_FREE_MODULES.has(activeModule)) {
    try {
      const res  = await fetch(`${API}/api/chrome/status`);
      const data = await res.json();
      chromeReady = data.ready;
      updateChromeUI();
    } catch {}
  }
  setTimeout(pollChromeStatus, 4000);
}

function updateChromeUI() {
  const pill  = document.getElementById('chromePill');
  const dot   = document.getElementById('chromeDot');
  const label = document.getElementById('chromeLabel');

  if (chromeReady) {
    pill?.classList.add('connected');
    if (label) label.textContent = 'Chrome Connected';
  } else {
    pill?.classList.remove('connected');
    if (label) label.textContent = 'Automation Chrome';
  }

  syncChromeUI();
}

function syncChromeUI() {
  const step   = document.getElementById('rs-chrome');
  const hint   = document.getElementById('rs-chrome-hint');
  const btn    = document.getElementById('rs-chrome-btn');
  const runBtn = document.getElementById('runModuleBtn');

  if (!step) return;

  if (chromeReady) {
    step.className   = 'run-step step-done';
    if (hint) { hint.textContent = 'Connected — your sessions are active'; hint.className = 'step-hint ok'; }
    if (btn)  { btn.textContent = '✓ Connected'; btn.className = 'btn-chrome ok'; btn.disabled = true; }
    const rs2 = document.getElementById('rs-run');
    if (rs2) rs2.className = 'run-step step-active';
    if (runBtn) runBtn.disabled = false;
  } else {
    step.className   = 'run-step step-active';
    if (hint) { hint.textContent = 'Close main Chrome first, then click Connect'; hint.className = 'step-hint'; }
    if (btn)  { btn.textContent = 'Connect'; btn.className = 'btn-chrome'; btn.disabled = false; }
    if (runBtn) runBtn.disabled = true;
  }
}

function handleChromeConnect() {
  document.getElementById('chromeOverlay').style.display = 'flex';
}

function closeModal() {
  document.getElementById('chromeOverlay').style.display = 'none';
}

async function doLaunchChrome() {
  const btn = document.getElementById('launchModalBtn');
  btn.disabled = true;
  btn.textContent = 'Launching…';

  try {
    const res  = await fetch(`${API}/api/chrome/launch`, { method: 'POST' });
    const data = await res.json();

    if (data.ok) {
      chromeReady = true;
      updateChromeUI();
      closeModal();
    } else {
      btn.textContent = 'Retry';
      btn.disabled = false;
      alert(data.error || 'Could not launch Chrome.');
    }
  } catch (e) {
    btn.textContent = 'Retry';
    btn.disabled = false;
    alert(`Error: ${e.message}`);
  }
}


// ── Run module ─────────────────────────────────────────────────────────────────

async function runModule() {
  if (!activeModule) return;

  const btn  = document.getElementById('runModuleBtn');
  const log  = document.getElementById('moduleLog');
  const pill = document.getElementById('logPill');

  btn.disabled = true;
  btn.className = 'btn-run-module running';
  btn.textContent = '⏳ Running…';
  log.innerHTML = '';
  pill.textContent = 'Running…';
  pill.className = 'log-pill running';

  if (currentJobSrc) { currentJobSrc.close(); currentJobSrc = null; }

  // Show Stop button
  const stopBtn = document.getElementById('stopModuleBtn');
  if (stopBtn) stopBtn.style.display = 'block';

  try {
    const res = await fetch(`${API}/api/run/${activeModule}`, { method: 'POST' });
    const { job_id } = await res.json();
    currentJobId  = job_id;

    currentJobSrc = new EventSource(`${API}/api/status/${job_id}`);

    currentJobSrc.onmessage = (e) => {
      const data = JSON.parse(e.data);
      if (data.type === 'ping') return;

      if (data.type === 'log')   appendModuleLog(log, data.msg);
      if (data.type === 'done')  finishRun(btn, pill, log, true);
      if (data.type === 'error') { appendModuleLog(log, data.msg, 'error'); finishRun(btn, pill, log, false); }
    };

    currentJobSrc.onerror = () => {
      appendModuleLog(log, 'Lost connection to server.', 'error');
      finishRun(btn, pill, log, false);
    };
  } catch (e) {
    appendModuleLog(log, `Error: ${e.message}`, 'error');
    finishRun(btn, pill, log, false);
  }
}

function appendModuleLog(container, msg, type = '') {
  const line = document.createElement('span');
  line.className = 'log-line';
  if (type === 'error' || msg.toLowerCase().includes('error')) line.classList.add('error');
  else if (type === 'success' || msg.startsWith('✅'))          line.classList.add('success');
  else if (msg.toLowerCase().includes('warning'))               line.classList.add('warn');
  else if (msg.startsWith('━━━'))                               line.classList.add('step');
  line.textContent = msg;
  container.appendChild(line);
  container.appendChild(document.createElement('br'));
  container.scrollTop = container.scrollHeight;
}

async function stopRun() {
  if (!currentJobId) return;
  const stopBtn = document.getElementById('stopModuleBtn');
  if (stopBtn) stopBtn.disabled = true;
  try {
    await fetch(`${API}/api/cancel/${currentJobId}`, { method: 'POST' });
  } catch {}
}

function finishRun(btn, pill, log, success) {
  if (currentJobSrc) { currentJobSrc.close(); currentJobSrc = null; }
  currentJobId = null;
  const stopBtn = document.getElementById('stopModuleBtn');
  if (stopBtn) { stopBtn.style.display = 'none'; stopBtn.disabled = false; }
  btn.disabled = false;
  if (success) {
    btn.className   = 'btn-run-module done';
    btn.textContent = '✓ Complete — Run Again';
    pill.textContent = 'Done';
    pill.className   = 'log-pill done';
  } else {
    btn.className   = 'btn-run-module errored';
    btn.textContent = '✗ Failed — Try Again';
    pill.textContent = 'Error';
    pill.className   = 'log-pill error';
  }
}


// ── Commission Estimate Module ─────────────────────────────────────────────────

function renderCommEstimateModule() {
  // Chrome not needed — hide the sidebar pill to avoid confusion
  const pill = document.getElementById('chromePill');
  if (pill) pill.style.display = 'none';

  const view = document.getElementById('view-module');
  view.innerHTML = `
    <div class="module-view-header">
      <button class="back-btn" onclick="goBack()">← Back</button>
      <div>
        <div class="mv-title">📈 Commission Estimate</div>
        <div class="mv-sub">Upload CRM PDFs → Annotated PDF Download</div>
      </div>
    </div>

    <div class="module-view-body">

      <!-- LEFT: upload + process -->
      <div style="display:flex;flex-direction:column;gap:16px">

        <div class="run-panel">
          <div class="run-panel-header" style="padding:13px 22px">
            <span class="run-panel-title">Upload PDFs</span>
            <span class="status-badge status-active">Active</span>
          </div>
          <div style="padding:14px 22px 20px;display:flex;flex-direction:column;gap:14px">

            <!-- Estimate Details drop zone -->
            <div>
              <div style="font-size:11px;font-weight:600;color:var(--text-2);margin-bottom:6px;text-transform:uppercase;letter-spacing:.4px">Estimate Details PDF</div>
              <div class="ce-drop-zone" id="ceDropEst" onclick="document.getElementById('ceFileEst').click()"
                   ondragover="ceDragOver(event,'ceDropEst')" ondragleave="ceDragLeave('ceDropEst')" ondrop="ceDrop(event,'ceDropEst','ceFileEst')">
                <span id="ceDropEstLabel">📄 Click or drag to upload</span>
              </div>
              <input type="file" id="ceFileEst" accept=".pdf" style="display:none" onchange="ceFileSelected('ceFileEst','ceDropEst','ceDropEstLabel')" />
            </div>

            <!-- Commission Sheet drop zone -->
            <div>
              <div style="font-size:11px;font-weight:600;color:var(--text-2);margin-bottom:6px;text-transform:uppercase;letter-spacing:.4px">Commission Sheet PDF</div>
              <div class="ce-drop-zone" id="ceDropComm" onclick="document.getElementById('ceFileComm').click()"
                   ondragover="ceDragOver(event,'ceDropComm')" ondragleave="ceDragLeave('ceDropComm')" ondrop="ceDrop(event,'ceDropComm','ceFileComm')">
                <span id="ceDropCommLabel">📄 Click or drag to upload</span>
              </div>
              <input type="file" id="ceFileComm" accept=".pdf" style="display:none" onchange="ceFileSelected('ceFileComm','ceDropComm','ceDropCommLabel')" />
            </div>

            <!-- Optional fields -->
            <div style="display:flex;gap:10px">
              <div style="flex:1">
                <div style="font-size:11px;font-weight:600;color:var(--text-2);margin-bottom:4px;text-transform:uppercase;letter-spacing:.4px">Finance Fee <span style="font-weight:400;text-transform:none;letter-spacing:0">(optional)</span></div>
                <input id="ceFinanceFee" type="number" step="0.01" placeholder="e.g. 1799.13" style="${inputStyle}" />
              </div>
              <div style="flex:1">
                <div style="font-size:11px;font-weight:600;color:var(--text-2);margin-bottom:4px;text-transform:uppercase;letter-spacing:.4px">Lender <span style="font-weight:400;text-transform:none;letter-spacing:0">(optional)</span></div>
                <input id="ceFinanceLender" type="text" placeholder="e.g. Lendhi" style="${inputStyle}" />
              </div>
            </div>

            <button class="btn-run-module" id="ceProcessBtn" onclick="ceProcess()" disabled style="margin:0;font-size:14px;padding:11px">
              ▶ Process &amp; Download PDFs
            </button>
          </div>
        </div>

        <!-- Status card -->
        <div class="log-wrap" id="ceStatusWrap" style="display:none">
          <div class="log-header">
            <span class="log-title">Status</span>
            <span class="log-pill" id="ceStatusPill">Processing…</span>
          </div>
          <div class="log-body" id="ceStatusLog" style="min-height:60px;display:flex;align-items:center;padding:16px 20px">
            <span class="log-idle" id="ceStatusMsg">Uploading and processing PDFs…</span>
          </div>
        </div>

        <!-- Download card (shown after success) -->
        <div class="run-panel" id="ceDownloadCard" style="display:none">
          <div class="run-panel-header" style="padding:13px 22px">
            <span class="run-panel-title" style="color:#0f766e">✓ Processing Complete</span>
          </div>
          <div style="padding:14px 22px 20px;display:flex;flex-direction:column;gap:10px">
            <p style="font-size:13px;color:var(--text-2);margin:0">Your annotated PDFs are ready. The ZIP contains both the Estimate Details and Commission Sheet with corrections marked in red.</p>
            <button class="btn-run-module done" id="ceDownloadBtn" onclick="ceDownload()" style="margin:0;padding:11px">
              ⬇ Download Annotated PDFs
            </button>
            <button onclick="ceReset()" style="background:none;border:1px solid var(--border);border-radius:8px;padding:9px;font-size:13px;cursor:pointer;color:var(--text-2)">
              ↺ Process Another Estimate
            </button>
          </div>
        </div>

      </div>

      <!-- RIGHT: info panel -->
      <div class="side-panel">
        <div class="info-card">
          <div class="info-card-title">How It Works</div>
          <div class="info-row"><span class="info-key">Step 1</span><span class="info-value">Upload both CRM PDFs</span></div>
          <div class="info-row"><span class="info-key">Step 2</span><span class="info-value">Enter finance fee if applicable</span></div>
          <div class="info-row"><span class="info-key">Step 3</span><span class="info-value">Click Process — takes ~5 seconds</span></div>
          <div class="info-row"><span class="info-key">Step 4</span><span class="info-value">Download the annotated ZIP</span></div>
        </div>
        <div class="info-card">
          <div class="info-card-title">What Gets Annotated</div>
          <div class="info-row"><span class="info-key">Greenline</span><span class="info-value">Recalculated &amp; corrected</span></div>
          <div class="info-row"><span class="info-key">% GL</span><span class="info-value">Corrected in red</span></div>
          <div class="info-row"><span class="info-key">Comm Rate</span><span class="info-value">Tier-based correction</span></div>
          <div class="info-row"><span class="info-key">Est Commission</span><span class="info-value">Recalculated</span></div>
          <div class="info-row"><span class="info-key">Profit Margin</span><span class="info-value">Added to summary block</span></div>
          <div class="info-row"><span class="info-key">Contractors</span><span class="info-value">Highlighted in yellow</span></div>
        </div>
        <div class="info-card">
          <div class="info-card-title">Alerts</div>
          <div style="font-size:12px;color:var(--text-2);line-height:1.6">
            Remarks are added to the last page when:<br>
            · Profit margin &lt; 60% → notify Anne<br>
            · % Greenline &lt; 80% → notify Anne
          </div>
        </div>
        <div class="info-card">
          <div class="info-card-title">Server Connection</div>
          <div style="font-size:11px;color:var(--text-3);margin-bottom:8px;line-height:1.5">
            PDF processing requires your local Flask server.<br>
            Default: <code style="background:#f1f5f9;padding:1px 4px;border-radius:3px">http://localhost:5050</code>
          </div>
          <div style="display:flex;gap:6px;align-items:center">
            <input id="ceApiUrlInput"
              value="${getCeApiBase()}"
              placeholder="http://localhost:5050"
              style="${inputStyle};flex:1;font-size:11px" />
            <button onclick="saveCeServerUrl()"
              style="background:var(--navy);color:#fff;border:none;border-radius:6px;padding:5px 10px;font-size:11px;font-weight:600;cursor:pointer;white-space:nowrap">
              Save
            </button>
          </div>
          <div id="ceApiUrlStatus" style="font-size:10px;color:var(--text-3);margin-top:4px">
            ${getCeApiBase() !== 'http://localhost:5050' ? '✓ Custom URL saved' : 'Using default localhost'}
          </div>
        </div>
      </div>

    </div>
  `;

  ceCheckReady();
}

let ceZipBlob = null;

function ceCheckReady() {
  const est  = document.getElementById('ceFileEst')?.files[0];
  const comm = document.getElementById('ceFileComm')?.files[0];
  const btn  = document.getElementById('ceProcessBtn');
  if (btn) btn.disabled = !(est && comm);
}

function ceDragOver(e, zoneId) {
  e.preventDefault();
  document.getElementById(zoneId)?.classList.add('ce-drop-active');
}

function ceDragLeave(zoneId) {
  document.getElementById(zoneId)?.classList.remove('ce-drop-active');
}

function ceDrop(e, zoneId, inputId) {
  e.preventDefault();
  ceDragLeave(zoneId);
  const file = e.dataTransfer.files[0];
  if (!file || !file.name.endsWith('.pdf')) return;
  const input = document.getElementById(inputId);
  const dt = new DataTransfer(); dt.items.add(file); input.files = dt.files;
  const labelId = zoneId + 'Label';
  const zone  = document.getElementById(zoneId);
  const label = document.getElementById(labelId);
  if (label) label.textContent = '✅ ' + file.name;
  if (zone)  zone.classList.add('ce-drop-filled');
  ceCheckReady();
}

function ceFileSelected(inputId, zoneId, labelId) {
  const input = document.getElementById(inputId);
  const file  = input?.files[0];
  if (!file) return;
  const label = document.getElementById(labelId);
  const zone  = document.getElementById(zoneId);
  if (label) label.textContent = '✅ ' + file.name;
  if (zone)  zone.classList.add('ce-drop-filled');
  ceCheckReady();
}

async function ceProcess() {
  const estFile  = document.getElementById('ceFileEst')?.files[0];
  const commFile = document.getElementById('ceFileComm')?.files[0];
  if (!estFile || !commFile) return;

  const btn = document.getElementById('ceProcessBtn');
  if (btn) { btn.disabled = true; btn.textContent = '⏳ Processing…'; btn.className = 'btn-run-module running'; btn.style.margin = '0'; }

  const statusWrap = document.getElementById('ceStatusWrap');
  const statusPill = document.getElementById('ceStatusPill');
  const statusMsg  = document.getElementById('ceStatusMsg');
  if (statusWrap) statusWrap.style.display = 'block';
  if (statusPill) { statusPill.textContent = 'Processing…'; statusPill.className = 'log-pill running'; }
  if (statusMsg)  statusMsg.textContent = 'Uploading and processing PDFs — this takes about 5–10 seconds…';

  const fee    = document.getElementById('ceFinanceFee')?.value.trim();
  const lender = document.getElementById('ceFinanceLender')?.value.trim();

  const form = new FormData();
  form.append('estimate_pdf',   estFile);
  form.append('commission_pdf', commFile);
  if (fee)    form.append('finance_fee',    fee);
  if (lender) form.append('finance_lender', lender);

  try {
    const res = await fetch(`${getCeApiBase()}/api/commission-estimate/process`, { method: 'POST', body: form });

    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: 'Unknown server error' }));
      throw new Error(err.error || 'Server error');
    }

    const payload = await res.json();

    // Decode base64 zip → Blob for download
    const zipBytes  = Uint8Array.from(atob(payload.zip_b64), c => c.charCodeAt(0));
    ceZipBlob = new Blob([zipBytes], { type: 'application/zip' });

    if (statusPill) { statusPill.textContent = 'Done'; statusPill.className = 'log-pill done'; }
    if (statusMsg)  statusMsg.textContent = '✅ PDFs processed successfully!';
    if (btn) btn.style.display = 'none';

    // ── Show alerts if any ────────────────────────────────────────────────────
    ceShowAlerts(payload.alerts || [], payload.is_sean || false, payload.summary || {});

    const downloadCard = document.getElementById('ceDownloadCard');
    if (downloadCard) downloadCard.style.display = 'block';

  } catch (e) {
    if (statusPill) { statusPill.textContent = 'Error'; statusPill.className = 'log-pill error'; }
    const isBlocked = e.message === 'Load failed' || e.message === 'Failed to fetch';
    const hint = isBlocked
      ? '✗ Cannot reach the server. If you\'re on Vercel (HTTPS), open the app at http://localhost:5050 instead — browsers block HTTPS → HTTP calls. Or use a tunnelled HTTPS URL in Server Connection below.'
      : '✗ ' + e.message;
    if (statusMsg)  statusMsg.textContent = hint;
    if (btn) { btn.disabled = false; btn.textContent = '▶ Process & Download PDFs'; btn.className = 'btn-run-module'; btn.style.margin = '0'; }
  }
}

function ceDownload() {
  if (!ceZipBlob) return;
  const url = URL.createObjectURL(ceZipBlob);
  const a   = document.createElement('a');
  a.href    = url;
  a.download = 'commission_estimate_processed.zip';
  a.click();
  URL.revokeObjectURL(url);
}

function ceReset() {
  ceZipBlob = null;
  renderCommEstimateModule();
}

function ceShowAlerts(alerts, isSean, summary) {
  // Remove any existing alert card
  document.getElementById('ceAlertCard')?.remove();
  if (!alerts.length && !isSean) return;

  const card = document.createElement('div');
  card.id = 'ceAlertCard';
  card.className = 'run-panel';
  card.style.cssText = 'border:2px solid #f59e0b;background:#fffbeb';

  let html = `<div class="run-panel-header" style="padding:13px 22px;background:#fef3c7">
    <span class="run-panel-title" style="color:#92400e">⚠️ Action Required</span>
  </div>
  <div style="padding:14px 22px 18px;display:flex;flex-direction:column;gap:10px">`;

  // Alert rows (inform Anne)
  if (alerts.length) {
    html += `<div style="display:flex;flex-direction:column;gap:6px">`;
    alerts.forEach(a => {
      html += `<div style="display:flex;align-items:flex-start;gap:8px;background:#fef9c3;border:1px solid #fde047;border-radius:8px;padding:10px 12px">
        <span style="font-size:18px;line-height:1.2">🔔</span>
        <div>
          <div style="font-size:13px;font-weight:700;color:#78350f">${a}</div>
          <div style="font-size:11px;color:#92400e;margin-top:2px">Please inform Anne before proceeding.</div>
        </div>
      </div>`;
    });
    html += `</div>`;
  }

  // Sean tier note
  if (isSean) {
    html += `<div style="display:flex;align-items:flex-start;gap:8px;background:#eff6ff;border:1px solid #bfdbfe;border-radius:8px;padding:10px 12px">
      <span style="font-size:18px;line-height:1.2">ℹ️</span>
      <div>
        <div style="font-size:13px;font-weight:700;color:#1e40af">Sean Stern Commission Tiers Applied</div>
        <div style="font-size:11px;color:#1e3a8a;margin-top:2px">
          Rates used: ≥100% GL → 12% · ≥90% → 11% · ≥80% → 10% · ≥75% → 9% · &lt;75% → 5%
        </div>
      </div>
    </div>`;
  }

  // Summary values
  if (Object.keys(summary).length) {
    html += `<div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-top:4px">`;
    const fields = [
      ['profit_margin',  'Profit Margin'],
      ['pct_greenline',  '% Greenline'],
      ['commission_pct', 'Commission %'],
      ['est_commission', 'Est. Commission'],
    ];
    fields.forEach(([key, label]) => {
      if (!summary[key]) return;
      const isAlert = (key === 'profit_margin' && parseInt(summary[key]) < 60)
                   || (key === 'pct_greenline'  && parseInt(summary[key]) < 80);
      html += `<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;padding:7px 10px">
        <div style="font-size:10px;color:#64748b;font-weight:600;text-transform:uppercase;letter-spacing:.3px">${label}</div>
        <div style="font-size:14px;font-weight:700;color:${isAlert ? '#dc2626' : '#0f172a'};margin-top:2px">${summary[key]}</div>
      </div>`;
    });
    html += `</div>`;
  }

  html += `</div>`;
  card.innerHTML = html;

  // Insert before the download card
  const downloadCard = document.getElementById('ceDownloadCard');
  if (downloadCard) downloadCard.before(card);
  else document.getElementById('ceStatusWrap')?.after(card);
}

function saveCeServerUrl() {
  const input  = document.getElementById('ceApiUrlInput');
  const status = document.getElementById('ceApiUrlStatus');
  if (!input) return;
  const url = input.value.trim();
  if (!url) return;
  saveCeApiBase(url);
  if (status) {
    status.textContent = url === 'http://localhost:5050' ? 'Using default localhost' : '✓ Custom URL saved';
    status.style.color = '#16a34a';
    setTimeout(() => { if (status) status.style.color = 'var(--text-3)'; }, 2000);
  }
}


// ── Utilities ──────────────────────────────────────────────────────────────────

function getLastWeekLabel() {
  const today = new Date();
  const d = today.getDay();
  const lastMon = new Date(today); lastMon.setDate(today.getDate() - (d === 0 ? 13 : d + 6));
  const lastSun = new Date(lastMon); lastSun.setDate(lastMon.getDate() + 6);
  const f = d => d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
  return `${f(lastMon)} – ${f(lastSun)}`;
}

function timeAgo(iso) {
  const diff = Math.floor((Date.now() - new Date(iso)) / 1000);
  if (diff < 60)   return 'Just now';
  if (diff < 3600) return `${Math.floor(diff/60)}m ago`;
  if (diff < 86400)return `${Math.floor(diff/3600)}h ago`;
  return new Date(iso).toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
}
