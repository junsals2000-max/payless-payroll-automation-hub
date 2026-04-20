/* ══════════════════════════════════════════════
   Payless Automation Hub — app.js
   ══════════════════════════════════════════════ */

// ── Modules that don't need Chrome ────────────────────────────────────────────
const CHROME_FREE_MODULES = new Set(['commission_estimate', 'onsite_recon', 'bank_categorizer']);

// ── API base URL (configurable, stored in localStorage) ───────────────────────
function getApiBase() {
  return (localStorage.getItem('api_base_url') || 'http://localhost:5050').replace(/\/$/, '');
}
function saveApiBase(url) {
  localStorage.setItem('api_base_url', url.trim().replace(/\/$/, ''));
}

// ── Auth token helpers ─────────────────────────────────────────────────────────
function getAuthToken()       { return localStorage.getItem('auth_token'); }
function setAuthToken(t)      { localStorage.setItem('auth_token', t); }
function clearAuthToken()     { localStorage.removeItem('auth_token'); }

// ── Authenticated fetch wrapper ────────────────────────────────────────────────
function authFetch(url, options = {}) {
  const token = getAuthToken();
  return fetch(url, {
    ...options,
    headers: {
      ...(options.headers || {}),
      ...(token ? { 'Authorization': `Bearer ${token}` } : {}),
    },
  });
}

// ── Custom modal engine (replaces alert / confirm / prompt) ───────────────────

function _modal(icon, title, message, actions, extra = '') {
  document.getElementById('customModalIcon').textContent    = icon;
  document.getElementById('customModalTitle').textContent   = title;
  document.getElementById('customModalMessage').innerHTML   = message;
  document.getElementById('customModalExtra').innerHTML     = extra;
  document.getElementById('customModalActions').innerHTML   = actions;
  document.getElementById('customModalOverlay').style.display = 'flex';
}

function closeCustomModal() {
  document.getElementById('customModalOverlay').style.display = 'none';
  document.getElementById('customModalExtra').innerHTML = '';
}

// Close on backdrop click
document.addEventListener('click', e => {
  if (e.target.id === 'customModalOverlay') closeCustomModal();
});

/** showAlert(title, message, icon?) */
function showAlert(title, message, icon = 'ℹ️') {
  _modal(icon, title, message,
    `<button class="btn-primary" onclick="closeCustomModal()">Got it</button>`);
}

/** showSuccess(title, message) */
function showSuccess(title, message) {
  _modal('✅', title, message,
    `<button class="btn-primary" onclick="closeCustomModal()">OK</button>`);
}

/** showError(title, message) */
function showError(title, message) {
  _modal('❌', title, message,
    `<button class="btn-primary" onclick="closeCustomModal()">Close</button>`);
}

/** showConfirm(icon, title, message, confirmLabel, onConfirm, danger?) */
function showConfirm(icon, title, message, confirmLabel, onConfirm, danger = false) {
  const id  = '_cfn_' + Date.now();
  window[id] = () => { closeCustomModal(); onConfirm(); delete window[id]; };
  const confirmBtn = danger
    ? `<button class="btn-modal-danger" onclick="window['${id}']()">${confirmLabel}</button>`
    : `<button class="btn-primary"      onclick="window['${id}']()">${confirmLabel}</button>`;
  _modal(icon, title, message,
    `<button class="btn-secondary" onclick="closeCustomModal()">Cancel</button>${confirmBtn}`);
}

/** showFormModal — custom form with inputs, calls onConfirm(data) */
function showFormModal(icon, title, fields, confirmLabel, onConfirm) {
  const id = '_frm_' + Date.now();
  const extraHtml = fields.map(f => {
    if (f.type === 'select') {
      const opts = f.options.map(o => `<option value="${o}">${o}</option>`).join('');
      return `<label class="custom-modal-label">${f.label}</label>
              <select id="${f.id}" class="custom-modal-select">${opts}</select>`;
    }
    return `<label class="custom-modal-label">${f.label}</label>
            <input id="${f.id}" class="custom-modal-input" type="${f.inputType||'text'}"
                   placeholder="${f.placeholder||''}" />`;
  }).join('');

  window[id] = () => {
    const data = {};
    for (const f of fields) {
      const el = document.getElementById(f.id);
      data[f.id] = el ? el.value.trim() : '';
    }
    for (const f of fields) {
      if (f.required && !data[f.id]) {
        document.getElementById(f.id).focus();
        document.getElementById(f.id).style.borderColor = 'var(--red)';
        return;
      }
    }
    closeCustomModal();
    onConfirm(data);
    delete window[id];
  };

  _modal(icon, title, '', `
    <button class="btn-secondary" onclick="closeCustomModal()">Cancel</button>
    <button class="btn-primary"   onclick="window['${id}']()">${confirmLabel}</button>`,
    extraHtml);

  // Focus first field
  setTimeout(() => { const f = fields[0]; if (f) document.getElementById(f.id)?.focus(); }, 50);
}

// ── Current user (set after login) ────────────────────────────────────────────
let currentUser = null;   // { user_id, username, is_admin, allowed_modules }


// ── Init ───────────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
  setDateChip();
  checkAuth();
});

async function checkAuth() {
  const token = getAuthToken();
  if (!token) { showLoginScreen(); return; }

  // Token exists — silently validate before showing anything
  try {
    const res = await authFetch(`${getApiBase()}/api/auth/me`);
    if (res.status === 401) {
      clearAuthToken();
      showLoginScreen();
      return;
    }
    currentUser = await res.json();
    showHub();
  } catch {
    // Server unreachable — show login with offline message
    showLoginScreen('offline');
  }
}

// ── Login screen ───────────────────────────────────────────────────────────────

function showLoginScreen(state) {
  document.getElementById('loginOverlay').style.display = 'flex';
  document.getElementById('hubBody').style.display = 'none';

  if (state === 'offline') {
    const err = document.getElementById('loginError');
    if (err) {
      err.textContent = 'The Server is Offline, please contact Jun.';
      err.style.display = 'block';
    }
  }
}

function hideLoginScreen() {
  document.getElementById('loginOverlay').style.display = 'none';
  document.getElementById('hubBody').style.display = 'flex';
}

async function doLogin() {
  const usernameEl = document.getElementById('loginUsername');
  const passwordEl = document.getElementById('loginPassword');
  const errorEl    = document.getElementById('loginError');
  const btn        = document.getElementById('loginBtn');

  const username = usernameEl?.value.trim();
  const password = passwordEl?.value;

  if (!username || !password) {
    if (errorEl) { errorEl.textContent = 'Please enter your username and password.'; errorEl.style.display = 'block'; }
    return;
  }

  btn.disabled = true;
  btn.textContent = 'Signing in…';
  if (errorEl) errorEl.style.display = 'none';

  try {
    const res  = await fetch(`${getApiBase()}/api/auth/login`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password }),
    });
    const data = await res.json();

    if (!res.ok) {
      if (errorEl) { errorEl.textContent = data.error || 'Login failed.'; errorEl.style.display = 'block'; }
      btn.disabled = false;
      btn.textContent = 'Sign In';
      return;
    }

    setAuthToken(data.token);
    currentUser = { username: data.username, is_admin: data.is_admin };

    // Fetch full user info
    const meRes = await authFetch(`${getApiBase()}/api/auth/me`);
    if (meRes.ok) currentUser = await meRes.json();

    hideLoginScreen();
    showHub();

  } catch {
    if (errorEl) {
      errorEl.textContent = 'The Server is Offline, please contact Jun.';
      errorEl.style.display = 'block';
    }
    btn.disabled = false;
    btn.textContent = 'Sign In';
  }
}

function doLogout() {
  clearAuthToken();
  currentUser = null;
  activeModule = null;
  window.location.hash = '';
  closeSettings();
  showLoginScreen();
}

// Allow Enter key on login form
document.addEventListener('keydown', e => {
  if (e.key === 'Enter' && document.getElementById('loginOverlay')?.style.display !== 'none') {
    doLogin();
  }
});

// ── Hub initialization ─────────────────────────────────────────────────────────

function showHub() {
  hideLoginScreen();
  buildNav();
  setGreeting();
  loadDashboard();
  pollChromeStatus();

  // Restore module from URL hash
  const hash = window.location.hash.replace('#', '');
  if (hash) {
    openModule(hash);
  } else {
    showView('dashboard');
  }
}

// ── Greeting ───────────────────────────────────────────────────────────────────

function setGreeting() {
  const hour  = new Date().getHours();
  const name  = currentUser?.username || 'there';
  const greet = hour < 12 ? 'Good morning' : hour < 17 ? 'Good afternoon' : 'Good evening';
  const emoji = hour < 12 ? '☀️' : hour < 17 ? '👋' : '🌙';
  const el    = document.getElementById('greetingLine');
  if (el) el.textContent = `${greet}, ${name} ${emoji}`;
}

function setDateChip() {
  const d = document.getElementById('dateChip');
  if (d) d.textContent = new Date().toLocaleDateString('en-US', {
    weekday: 'long', month: 'long', day: 'numeric', year: 'numeric',
  });
}

// ── Dynamic nav ────────────────────────────────────────────────────────────────

const NAV_MODULES = [
  { id: 'rehash',              icon: '📊', label: 'Sales Report',         active: true  },
  { id: 'coordinator',         icon: '📋', label: 'Project Coordinator',  active: false },
  { id: 'marketing',           icon: '📣', label: 'Marketing Report',     active: false },
  { id: 'scheduling',          icon: '📅', label: 'Scheduling Report',    active: false },
  { id: 'recruiting',          icon: '🤝', label: 'Recruiting Report',    active: false },
  { id: 'commission',          icon: '💵', label: 'Sales Commission',     active: true  },
  { id: 'commission_estimate', icon: '📈', label: 'Commission Estimate',  active: true  },
  { id: 'onsite_recon',        icon: '💼', label: 'Onsite Payroll Report', active: true  },
  { id: 'bank_categorizer',    icon: '🏦', label: '5160 Report',           active: true  },
];

function buildNav() {
  const nav      = document.getElementById('sidebarNav');
  const allowed  = currentUser?.allowed_modules || [];
  const isAdmin  = currentUser?.is_admin;
  const all      = allowed.includes('*');

  // Module nav items
  let html = `<button class="nav-item active" onclick="goBack()" id="nav-dashboard"><span class="nav-icon">⊞</span> Dashboard</button>
    <div class="nav-section-label">Modules</div>`;

  for (const m of NAV_MODULES) {
    if (!all && !allowed.includes(m.id)) continue;
    if (m.active) {
      html += `<button class="nav-item" onclick="openModule('${m.id}')" id="nav-${m.id}"><span class="nav-icon">${m.icon}</span> ${m.label}</button>`;
    } else {
      html += `<button class="nav-item nav-disabled" id="nav-${m.id}"><span class="nav-icon">${m.icon}</span> ${m.label}<span class="nav-badge">Soon</span></button>`;
    }
  }

  // Admin nav item (superadmin only)
  if (isAdmin) {
    html += `<div class="nav-section-label">System</div>
      <button class="nav-item" onclick="showView('admin')" id="nav-admin"><span class="nav-icon">🛡</span> Admin</button>`;
  }

  nav.innerHTML = html;

  // Rebuild sidebar footer with logout
  const footer = document.getElementById('sidebarFooter');
  if (footer) {
    footer.innerHTML = `
      <div class="chrome-pill" id="chromePill">
        <span class="chrome-dot" id="chromeDot"></span>
        <span id="chromeLabel">Automation Chrome</span>
      </div>
      <div style="display:flex;gap:6px;justify-content:center">
        <button class="settings-btn" onclick="openSettings()" title="Settings" style="flex:1;text-align:center;display:flex;flex-direction:column;align-items:center;gap:4px;padding:10px 6px">
          <span style="font-size:18px;line-height:1">⚙</span>
          <span style="font-size:10px;font-weight:600;letter-spacing:.3px;opacity:.8">Settings</span>
        </button>
        <button class="settings-btn" onclick="doLogout()" title="Log out" style="flex:1;text-align:center;color:#f87171;display:flex;flex-direction:column;align-items:center;gap:4px;padding:10px 6px">
          <span style="font-size:18px;line-height:1">⏻</span>
          <span style="font-size:10px;font-weight:600;letter-spacing:.3px;opacity:.8">Log Out</span>
        </button>
      </div>`;
  }
}

// ── Settings modal ─────────────────────────────────────────────────────────────

function openSettings() {
  const urlEl = document.getElementById('settingsApiUrl');
  if (urlEl) urlEl.value = getApiBase();
  const userEl = document.getElementById('settingsUsername');
  if (userEl) userEl.textContent = currentUser?.username || '';
  // Clear password fields every time the modal opens
  ['settingsCurPw','settingsNewPw','settingsConfPw'].forEach(id => {
    const el = document.getElementById(id); if (el) el.value = '';
  });
  const pwErr = document.getElementById('settingsPwError');
  if (pwErr) pwErr.style.display = 'none';
  document.getElementById('settingsOverlay').style.display = 'flex';
}

async function changeOwnPassword() {
  const cur  = document.getElementById('settingsCurPw')?.value;
  const nw   = document.getElementById('settingsNewPw')?.value;
  const conf = document.getElementById('settingsConfPw')?.value;
  const errEl = document.getElementById('settingsPwError');

  const showErr = msg => { errEl.textContent = msg; errEl.style.display = 'block'; };
  errEl.style.display = 'none';

  if (!cur)              return showErr('Please enter your current password.');
  if (!nw || nw.length < 6) return showErr('New password must be at least 6 characters.');
  if (nw !== conf)       return showErr('New passwords do not match.');

  try {
    const res  = await authFetch(`${getApiBase()}/api/auth/change-password`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ current_password: cur, new_password: nw }),
    });
    const data = await res.json();
    if (!res.ok) return showErr(data.error || 'Failed to update password.');
    closeSettings();
    showSuccess('Password Updated', 'Your password has been changed successfully.');
  } catch(e) {
    showErr('Could not reach the server.');
  }
}

function closeSettings() {
  document.getElementById('settingsOverlay').style.display = 'none';
}

function saveSettings() {
  const url = document.getElementById('settingsApiUrl')?.value.trim();
  if (url) saveApiBase(url);
  const status = document.getElementById('settingsUrlStatus');
  if (status) {
    status.textContent = '✓ Saved';
    status.style.color = '#16a34a';
    setTimeout(() => { if (status) status.textContent = ''; }, 2000);
  }
}

// Chrome modal uses logged-in username
function updateChromeModalText() {
  const desc = document.getElementById('modalUserDesc');
  if (desc && currentUser) desc.textContent = `${currentUser.username}'s sessions`;
}


// ── View routing ───────────────────────────────────────────────────────────────

function showView(name) {
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active-view'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  const view = document.getElementById(`view-${name}`);
  const nav  = document.getElementById(`nav-${name}`);
  if (view) view.classList.add('active-view');
  if (nav)  nav.classList.add('active');
  if (name === 'admin') renderAdminPanel();
}


// ── Dashboard ──────────────────────────────────────────────────────────────────

async function loadDashboard() {
  renderHeroBanner();
  try {
    const [modules, history] = await Promise.all([
      authFetch(`${getApiBase()}/api/modules`).then(r => r.json()),
      authFetch(`${getApiBase()}/api/history`).then(r => r.json()),
    ]);
    renderModuleGrid(modules);
    renderActivity(history);
    updateStats(modules, history);
  } catch {
    document.getElementById('moduleGrid').innerHTML =
      '<p style="color:var(--text-3);font-size:13px">Could not connect to server. Check that the hub is running.</p>';
  }
}

function renderHeroBanner() {
  const existing = document.getElementById('dashHero');
  if (existing) return;
  const hour  = new Date().getHours();
  const emoji = hour < 12 ? '☀️' : hour < 17 ? '⚡' : '🌙';
  const label = hour < 12 ? 'Morning Shift' : hour < 17 ? 'Afternoon Shift' : 'Evening Shift';
  const hero  = document.createElement('div');
  hero.id     = 'dashHero';
  hero.className = 'dash-hero';
  hero.innerHTML = `
    <div class="dash-hero-text">
      <div class="dash-hero-title">${emoji} Welcome to Automation Hub</div>
      <div class="dash-hero-sub">Your automations are ready to run. Pick a module below to get started.</div>
    </div>
    <div class="dash-hero-badge">
      <div class="dash-hero-dot"></div>
      ${label} · ${new Date().toLocaleDateString('en-US',{weekday:'short',month:'short',day:'numeric'})}
    </div>`;
  const topbar = document.querySelector('#view-dashboard .topbar');
  if (topbar) topbar.insertAdjacentElement('afterend', hero);
}

function renderModuleGrid(modules) {
  const grid = document.getElementById('moduleGrid');
  grid.innerHTML = modules.map(m => {
    const isActive = m.status === 'active';
    const isSoon   = m.status === 'coming_soon';
    const lastRun  = m.last_run ? `Last run: ${timeAgo(m.last_run.ran_at)}` : 'Never run';
    const statusLabel = isSoon ? 'Coming Soon' : 'Active';
    const statusClass = isSoon ? 'status-soon' : 'status-active';
    const sourceTags  = (m.sources || []).map(s => `<span class="source-tag">${s}</span>`).join('');
    return `
      <div class="module-card ${isSoon ? 'module-soon' : ''}" onclick="${isActive ? `openModule('${m.id}')` : ''}">
        <div class="module-stripe" style="background:${m.color}"></div>
        <div class="module-body">
          <div class="module-top">
            <div class="module-icon-wrap" style="background:${m.color}22"><span>${m.icon}</span></div>
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
      </div>`;
  }).join('');
}

function renderActivity(history) {
  const list = document.getElementById('activityList');
  if (!history.length) return;
  list.innerHTML = history.slice(0, 8).map(h => {
    const ok  = h.status === 'success';
    const lbl = h.module === 'rehash'               ? 'Sales Report'
              : h.module === 'commission'            ? 'Sales Commission'
              : h.module === 'commission_estimate'   ? 'Commission Estimate'
              : h.module === 'onsite_recon'          ? 'Onsite Payroll Report'
              : h.module === 'bank_categorizer'      ? '5160 Report'
              : h.module;
    return `
      <div class="activity-row">
        <div class="activity-dot ${ok ? 'success' : 'error'}"></div>
        <div class="activity-info">
          <div class="activity-name">${lbl}</div>
          <div class="activity-time">${timeAgo(h.ran_at)}</div>
        </div>
        <span class="activity-status ${ok ? 'success' : 'error'}">${ok ? 'Success' : 'Failed'}</span>
      </div>`;
  }).join('');
}

function updateStats(modules, history) {
  const success = history.filter(h => h.status === 'success').length;
  const active  = modules.filter(m => m.status === 'active').length;
  document.getElementById('statTotal').textContent   = history.length;
  document.getElementById('statSuccess').textContent = success;
  document.getElementById('statModules').textContent = modules.length;
  document.getElementById('statActive').textContent  = active;
}

function timeAgo(iso) {
  const diff = (Date.now() - new Date(iso)) / 1000;
  if (diff < 60)   return 'just now';
  if (diff < 3600) return `${Math.floor(diff/60)}m ago`;
  if (diff < 86400)return `${Math.floor(diff/3600)}h ago`;
  return `${Math.floor(diff/86400)}d ago`;
}


// ── Module hero banner ─────────────────────────────────────────────────────────

const HERO_THEMES = [
  {
    photo: 'https://images.unsplash.com/photo-1556909114-f6e7ad7d3136?auto=format&fit=crop&w=1400&q=80',
    gradient: 'linear-gradient(135deg,#0b2a4a,#1a5276)',
    orb1: '#3b9eda', orb2: '#1e5799',
    quote: '"Where great kitchens come to life."',
  },
  {
    photo: 'https://images.unsplash.com/photo-1484154218962-a197022b5858?auto=format&fit=crop&w=1400&q=80',
    gradient: 'linear-gradient(135deg,#1a2f1a,#2d5a27)',
    orb1: '#52b788', orb2: '#1b4332',
    quote: '"Crafted spaces, lasting impressions."',
  },
  {
    photo: 'https://images.unsplash.com/photo-1556909172-54557c7e4fb7?auto=format&fit=crop&w=1400&q=80',
    gradient: 'linear-gradient(135deg,#2c1810,#6b3a2a)',
    orb1: '#d4956a', orb2: '#7b3f20',
    quote: '"Warm wood, clean lines, beautiful living."',
  },
  {
    photo: 'https://images.unsplash.com/photo-1552321554-5fefe8c9ef14?auto=format&fit=crop&w=1400&q=80',
    gradient: 'linear-gradient(135deg,#0d2137,#1a4a6e)',
    orb1: '#7ec8e3', orb2: '#0e4d6e',
    quote: '"Spa-like bathrooms, every single day."',
  },
  {
    photo: 'https://images.unsplash.com/photo-1507652313-4de6a7a7eb0b?auto=format&fit=crop&w=1400&q=80',
    gradient: 'linear-gradient(135deg,#1f1f2e,#3a3a5c)',
    orb1: '#a78bfa', orb2: '#4c1d95',
    quote: '"Modern design, timeless quality."',
  },
  {
    photo: 'https://images.unsplash.com/photo-1565538008-cf8e6c3fbbcc?auto=format&fit=crop&w=1400&q=80',
    gradient: 'linear-gradient(135deg,#1a2a1a,#3d5c3d)',
    orb1: '#86efac', orb2: '#166534',
    quote: '"Bringing vision and craftsmanship together."',
  },
  {
    photo: 'https://images.unsplash.com/photo-1560185893-a55b0e6d4fa8?auto=format&fit=crop&w=1400&q=80',
    gradient: 'linear-gradient(135deg,#2a1f0e,#5c3d1a)',
    orb1: '#fbbf24', orb2: '#92400e',
    quote: '"Every detail tells your story."',
  },
  {
    photo: 'https://images.unsplash.com/photo-1615873968403-89e068629265?auto=format&fit=crop&w=1400&q=80',
    gradient: 'linear-gradient(135deg,#0a1628,#162847)',
    orb1: '#60a5fa', orb2: '#1e3a8a',
    quote: '"Precision tools for beautiful results."',
  },
];

let _heroIdx = Math.floor(Math.random() * HERO_THEMES.length);

function moduleHeroHTML(icon, title, subtitle) {
  const theme = HERO_THEMES[_heroIdx % HERO_THEMES.length];
  _heroIdx++;

  const imgId  = 'mhPhoto_' + Date.now();
  const heroId = 'mhWrap_'  + Date.now();

  // Orb positions randomised per render
  const ox1 = Math.random()*60, oy1 = Math.random()*100;
  const ox2 = 40 + Math.random()*60, oy2 = Math.random()*100;

  return `
    <div class="mod-hero" id="${heroId}" style="background:${theme.gradient}">
      <div class="mod-hero-orb" style="width:180px;height:180px;left:${ox1}%;top:${oy1}%;background:${theme.orb1}"></div>
      <div class="mod-hero-orb" style="width:140px;height:140px;left:${ox2}%;top:${oy2}%;background:${theme.orb2}"></div>
      <img id="${imgId}" class="mod-hero-photo hidden" src="${theme.photo}"
           onload="this.classList.remove('hidden')"
           onerror="this.remove()"
           alt="" aria-hidden="true" />
      <div class="mod-hero-overlay">
        <div class="mod-hero-left">
          <button class="mod-hero-back" onclick="goBack()">← Back</button>
          <div class="mod-hero-title">${icon} ${title}</div>
          <div class="mod-hero-sub">${subtitle}</div>
        </div>
        <div class="mod-hero-right">
          <div class="mod-hero-badge">🏠 Payless Kitchens &amp; Baths</div>
          <div class="mod-hero-quote">${theme.quote}</div>
        </div>
      </div>
    </div>`;
}

// ── Module routing ─────────────────────────────────────────────────────────────

let activeModule  = null;
let currentJobSrc = null;
let currentJobId  = null;
let chromeReady   = false;

async function openModule(moduleId) {
  activeModule = moduleId;
  window.location.hash = moduleId;   // persist on refresh

  const pill = document.getElementById('chromePill');
  if (pill) pill.style.display = CHROME_FREE_MODULES.has(moduleId) ? 'none' : '';

  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  const nav = document.getElementById(`nav-${moduleId}`);
  if (nav) nav.classList.add('active');

  showView('module');

  if (moduleId === 'rehash')              await renderRehashModule();
  if (moduleId === 'commission')          await renderCommissionModule();
  if (moduleId === 'commission_estimate') renderCommEstimateModule();
  if (moduleId === 'onsite_recon')        renderOnsiteReconView();
  if (moduleId === 'bank_categorizer')    renderBankCategorizerView();
}

function goBack() {
  const pill = document.getElementById('chromePill');
  if (pill) pill.style.display = '';
  activeModule = null;
  window.location.hash = '';
  showView('dashboard');
  loadDashboard();
}

// Handle browser back/forward
window.addEventListener('hashchange', () => {
  const hash = window.location.hash.replace('#', '');
  if (!hash && activeModule) goBack();
  else if (hash && hash !== activeModule && currentUser) openModule(hash);
});


// ── Rehash module ──────────────────────────────────────────────────────────────

async function renderRehashModule() {
  const view = document.getElementById('view-module');
  view.innerHTML = `
    ${moduleHeroHTML('📊','Sales Report','Weekly Jive + LeadPerfection → SharePoint Excel')}

    <div class="module-view-body">
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
            <button class="btn-run-module" id="runModuleBtn" onclick="runModule()" disabled style="flex:1;margin:0">▶ Run Sales Report</button>
            <button class="btn-stop" id="stopModuleBtn" onclick="stopRun()" style="display:none">⏹ Stop</button>
          </div>
        </div>

        <div class="log-wrap">
          <div class="log-header">
            <span class="log-title">Live Log</span>
            <span class="log-pill" id="logPill">Idle</span>
          </div>
          <div class="log-body-compact" id="rehashStatusBar" style="display:none">
            <span id="rehashStatusMsg"></span>
          </div>
          <div class="log-body" id="moduleLog">
            <span class="log-idle">Waiting to run…</span>
          </div>
        </div>
      </div>

      <div class="side-panel" id="sidePanelRehash"></div>
    </div>`;

  syncChromeUI();
  await loadRehashSidePanel();
}

async function loadRehashSidePanel() {
  const panel = document.getElementById('sidePanelRehash');
  if (!panel) return;
  try {
    const [employees, spCfg] = await Promise.all([
      authFetch(`${getApiBase()}/api/rehash/employees`).then(r => r.json()).catch(() => []),
      authFetch(`${getApiBase()}/api/rehash/config`).then(r => r.json()).catch(() => ({})),
    ]);
    const demoUrl  = spCfg.demo_sheet_url || '';
    const lastWeek = getLastWeekLabel();

    panel.innerHTML = `
      <div class="info-card">
        <div class="info-card-title">Report Details</div>
        <div class="info-row"><span class="info-key">Week</span><span class="info-value">${lastWeek}</span></div>
        <div class="info-row"><span class="info-key">Sources</span><span class="info-value">Jive + LeadPerfection</span></div>
        <div class="info-row"><span class="info-key">Output</span><span class="info-value">SharePoint Excel</span></div>
        <div class="info-row"><span class="info-key">Schedule</span><span class="info-value">Every Monday</span></div>
      </div>

      <div class="info-card">
        <div class="info-card-title">SharePoint Settings</div>
        <div style="font-size:11px;color:var(--text-3);margin-bottom:4px;font-weight:500"># of Demo Sheet Link</div>
        <div style="display:flex;gap:6px;align-items:center">
          <input id="demoSheetUrlInput" value="${demoUrl}" placeholder="Paste 'Copy link to this sheet' URL…" class="hub-input" style="flex:1" />
          <button onclick="saveDemoSheetUrl()" style="background:var(--navy);color:#fff;border:none;border-radius:6px;padding:5px 10px;font-size:11px;font-weight:600;cursor:pointer;white-space:nowrap">Save</button>
        </div>
        <div id="demoUrlStatus" style="font-size:10px;color:var(--text-3);margin-top:3px">${demoUrl ? '✓ URL saved' : 'No URL set — Step 4 will skip the Demo sheet'}</div>
      </div>

      <div class="info-card">
        <div class="info-card-title" style="display:flex;justify-content:space-between;align-items:center">
          <span>Employees (${employees.length})</span>
          <button class="btn-open" onclick="toggleAddEmpForm()" id="addEmpToggle">+ Add</button>
        </div>
        <div id="addEmpForm" style="display:none;margin-bottom:12px;padding:12px;background:#f8fafc;border:1px solid var(--border);border-radius:8px">
          <div style="display:flex;flex-direction:column;gap:7px">
            <input id="newEmpName" placeholder="Full Name"                        class="hub-input" />
            <input id="newEmpUrl"  placeholder="Excel Link (Copy link to sheet)"  class="hub-input" />
            <input id="newEmpJive" placeholder="Jive URL"                         class="hub-input" />
            <input id="newEmpLp"   placeholder="Name in LeadPerfection"           class="hub-input" />
          </div>
          <button onclick="addRehashEmployee()" style="margin-top:8px;width:100%;background:var(--navy);color:#fff;border:none;border-radius:7px;padding:7px;font-size:12px;font-weight:600;cursor:pointer">Save Employee</button>
        </div>
        <div class="emp-list-mini" id="empListMini">${renderEmpList(employees)}</div>
      </div>`;
  } catch {}
}

// .hub-input is defined in styles.css — replaces the old inline inputStyle string.
// Use class="hub-input" on any small text/password input inside panels.

function getLastWeekLabel() {
  const now  = new Date();
  const day  = now.getDay();
  const end  = new Date(now); end.setDate(now.getDate() - day);
  const start= new Date(end); start.setDate(end.getDate() - 6);
  const fmt  = d => d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
  return `${fmt(start)} – ${fmt(end)}`;
}

function renderEmpList(employees) {
  if (!employees.length) return '<p style="font-size:12px;color:var(--text-3);padding:4px 0">No employees yet. Click + Add above.</p>';
  return employees.map(e => {
    const ini      = e.name.split(' ').map(p=>p[0]).join('').slice(0,2).toUpperCase();
    const lpDisplay= e.lp_name && e.lp_name !== e.name ? e.lp_name : '—';
    const urlSet   = e.excel_url && e.excel_url.startsWith('http');
    return `
      <div class="emp-mini" style="flex-wrap:wrap;gap:4px">
        <div class="emp-av">${ini}</div>
        <div style="flex:1;min-width:0">
          <div class="emp-n">${e.name}</div>
          <div style="font-size:10px;color:${urlSet?'#16a34a':'#f59e0b'};margin-top:1px">${urlSet?'✓ Excel link saved':'⚠ No Excel link'}</div>
          <div style="display:flex;align-items:center;gap:4px;margin-top:3px">
            <span style="font-size:10px;color:var(--text-3);white-space:nowrap">LP:</span>
            <span style="font-size:10px;color:var(--text-2);flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${lpDisplay}</span>
            <button onclick="openEmpEdit('${e.id}')" style="background:none;border:none;color:var(--text-3);font-size:10px;cursor:pointer;padding:0 2px;flex-shrink:0" title="Edit">✏️</button>
          </div>
          <div id="emp-edit-${e.id}" style="display:none;margin-top:6px;padding:8px;background:#f8fafc;border:1px solid var(--border);border-radius:8px">
            <div style="display:flex;flex-direction:column;gap:5px">
              <input id="edit-name-${e.id}" value="${e.name}"            placeholder="Full Name"       class="hub-input" />
              <input id="edit-url-${e.id}"  value="${e.excel_url||''}"   placeholder="Excel Link"      class="hub-input" />
              <input id="edit-lp-${e.id}"   value="${e.lp_name||''}"     placeholder="LP Name"         class="hub-input" />
              <input id="edit-jive-${e.id}" value="${e.jive_url||''}"    placeholder="Jive URL"        class="hub-input" />
            </div>
            <div style="display:flex;gap:4px;margin-top:6px">
              <button onclick="saveEmpEdit('${e.id}')" style="flex:1;background:var(--navy);color:#fff;border:none;border-radius:5px;padding:4px;font-size:11px;font-weight:600;cursor:pointer">Save</button>
              <button onclick="cancelEmpEdit('${e.id}')" style="flex:1;background:none;border:1px solid var(--border);border-radius:5px;padding:4px;font-size:11px;cursor:pointer">Cancel</button>
            </div>
          </div>
        </div>
        <button onclick="removeRehashEmployee('${e.id}')" style="background:none;border:1px solid #fca5a5;border-radius:5px;color:#ef4444;font-size:10px;padding:2px 7px;cursor:pointer;align-self:start">✕</button>
      </div>`;
  }).join('');
}

function openEmpEdit(id)   { const el = document.getElementById(`emp-edit-${id}`); if(el) el.style.display='block'; document.getElementById(`edit-name-${id}`)?.focus(); }
function cancelEmpEdit(id) { const el = document.getElementById(`emp-edit-${id}`); if(el) el.style.display='none'; }

async function saveEmpEdit(id) {
  const name = document.getElementById(`edit-name-${id}`)?.value.trim();
  const url  = document.getElementById(`edit-url-${id}`)?.value.trim();
  const lp   = document.getElementById(`edit-lp-${id}`)?.value.trim();
  const jive = document.getElementById(`edit-jive-${id}`)?.value.trim();
  if (!name) { showAlert('Required', 'Name is required.'); return; }
  await authFetch(`${getApiBase()}/api/rehash/employees/${id}`, { method:'PATCH', headers:{'Content-Type':'application/json'}, body: JSON.stringify({name, excel_url:url, lp_name:lp, jive_url:jive}) });
  const employees = await authFetch(`${getApiBase()}/api/rehash/employees`).then(r=>r.json());
  const el = document.getElementById('empListMini'); if(el) el.innerHTML = renderEmpList(employees);
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
  if (!name) { showAlert('Required', 'Name is required.'); return; }
  await authFetch(`${getApiBase()}/api/rehash/employees`, { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({name, excel_url:url, jive_url:jive, lp_name:lp||name}) });
  const employees = await authFetch(`${getApiBase()}/api/rehash/employees`).then(r=>r.json());
  const el = document.getElementById('empListMini'); if(el) el.innerHTML = renderEmpList(employees);
  ['newEmpName','newEmpUrl','newEmpJive','newEmpLp'].forEach(id => { const e=document.getElementById(id); if(e) e.value=''; });
  toggleAddEmpForm();
}

function removeRehashEmployee(id) {
  showConfirm('🗑️', 'Remove Employee', 'Are you sure you want to remove this employee?', 'Remove', async () => {
    await authFetch(`${getApiBase()}/api/rehash/employees/${id}`, { method:'DELETE' });
    const employees = await authFetch(`${getApiBase()}/api/rehash/employees`).then(r=>r.json());
    const el = document.getElementById('empListMini'); if(el) el.innerHTML = renderEmpList(employees);
  }, true);
}

async function saveDemoSheetUrl() {
  const input  = document.getElementById('demoSheetUrlInput');
  const status = document.getElementById('demoUrlStatus');
  if (!input) return;
  const url = input.value.trim();
  try {
    await authFetch(`${getApiBase()}/api/rehash/config`, { method:'PATCH', headers:{'Content-Type':'application/json'}, body: JSON.stringify({demo_sheet_url:url}) });
    if (status) { status.textContent = url ? '✓ URL saved' : 'No URL set — Step 4 will skip the Demo sheet'; status.style.color = url ? '#16a34a' : 'var(--text-3)'; }
  } catch {
    if (status) { status.textContent = 'Save failed — server error'; status.style.color = '#ef4444'; }
  }
}


// ── Commission module ──────────────────────────────────────────────────────────

const COMM_STEPS = [
  { id:'search',    num:'1',   label:'Search Client',          hint:'Buildertrend lookup by client name' },
  { id:'job',       num:'2',   label:'Extract Job Details',    hint:'Inv# · Title · Sold Date · Contract Price' },
  { id:'pdf',       num:'3',   label:'Extract Commission PDF', hint:'Greenline · % GL · Commission · Finance' },
  { id:'rep',       num:'4',   label:'Sales Rep Mapping',      hint:'Match rep to Excel spreadsheet' },
  { id:'validate',  num:'5',   label:'Final Data Validation',  hint:'Ensure all 9 fields are present' },
  { id:'duplicate', num:'5.5', label:'Duplicate Check',        hint:'Check for existing entry in Excel' },
  { id:'insert',    num:'6',   label:'Insert Row into Excel',  hint:'Columns A–I mapped and written' },
  { id:'formulas',  num:'7',   label:'Update SUM Formulas',    hint:'Extend formula range to include new row' },
];

let commStepStates = {};
let commExtracted  = {};

async function renderCommissionModule() {
  const view = document.getElementById('view-module');
  view.innerHTML = `
    ${moduleHeroHTML('💵','Sales Commission','Buildertrend · Commission PDF → SharePoint Excel')}

    <div class="module-view-body">
      <div style="display:flex;flex-direction:column;gap:16px">

        <div class="run-panel">
          <div class="run-panel-header" style="padding:13px 22px">
            <span class="run-panel-title">Client Lookup</span>
            <span class="status-badge status-active">Active</span>
          </div>
          <div style="padding:14px 22px 18px;display:flex;flex-direction:column;gap:10px">
            <div>
              <label class="comm-field-label">Client Name</label>
              <input id="comm-client-name" placeholder="e.g. Smith, John" class="comm-input" oninput="commInputChanged()" />
            </div>
            <div style="display:flex;gap:8px">
              <button class="btn-run-module" id="commRunBtn" onclick="commStartRun()" disabled style="margin:0;flex:1;font-size:14px;padding:11px">▶ Search &amp; Process Commission</button>
              <button class="btn-stop" id="commResetBtn" onclick="commReset()" style="display:none;padding:10px 16px;font-size:12px">↺ Reset</button>
            </div>
          </div>
        </div>

        <div class="log-wrap" style="overflow:visible">
          <div class="log-header">
            <span class="log-title">Processing Steps</span>
            <span class="log-pill" id="commStatusPill">Idle</span>
          </div>
          <div style="padding:20px 22px 8px" id="commStepperWrap">${buildCommStepper()}</div>
          <div class="log-body-compact" id="commStatusMsg" style="display:none">
            <span id="commStatusMsgText"></span>
          </div>
        </div>

        <div class="info-card" id="commDataCard" style="display:none">
          <div class="info-card-title" style="display:flex;align-items:center;justify-content:space-between">
            <span>Extracted Data</span>
            <span id="commValidBadge" class="status-badge" style="display:none"></span>
          </div>
          <div class="comm-data-grid" id="commDataGrid"></div>
        </div>
      </div>

      <div class="side-panel" id="commSidePanel"></div>
    </div>`;

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
  row.className    = `comm-step-row ${state}`;
  node.textContent = state==='done'?'✓':state==='error'?'✗':state==='warning'?'!':step?.num||'?';
  if (hint && hintText) hint.textContent = hintText;
  if (line) line.style.background = state==='done'?'var(--green)':state==='error'?'var(--red)':state==='warning'?'var(--amber)':'';

  if (stepId === 'duplicate' && state === 'warning') {
    const content = document.getElementById(`csc-${stepId}`);
    if (content && !document.getElementById('commDupAlert')) {
      const alertEl = document.createElement('div');
      alertEl.className = 'comm-dup-alert';
      alertEl.id = 'commDupAlert';
      alertEl.innerHTML = `
        <div style="display:flex;gap:10px;align-items:flex-start">
          <span style="font-size:20px;line-height:1">⚠️</span>
          <div>
            <div style="font-weight:700;font-size:13px;color:#92400e;margin-bottom:4px">Duplicate Record Found</div>
            <div style="font-size:12px;color:#78350f" id="commDupDetail">${hintText||'An existing entry matches this client.'}</div>
            <div style="display:flex;gap:8px;margin-top:12px">
              <button class="comm-dup-btn proceed" onclick="commDupDecide('proceed')">✓ Proceed — Insert Anyway</button>
              <button class="comm-dup-btn skip"    onclick="commDupDecide('skip')">✗ Skip — Don't Insert</button>
            </div>
          </div>
        </div>`;
      content.appendChild(alertEl);
    }
  }
}

function commDupDecide(decision) {
  const alertEl = document.getElementById('commDupAlert');
  if (decision === 'skip') {
    if (alertEl) alertEl.innerHTML = '<span style="color:#dc2626;font-weight:600;font-size:12px">✗ Skipped — entry not inserted.</span>';
    setCommStep('duplicate','error','Skipped by user');
    commFinish(false,'Skipped');
  } else {
    if (alertEl) alertEl.innerHTML = '<span style="color:#16a34a;font-weight:600;font-size:12px">✓ Proceeding to insert…</span>';
    setCommStep('duplicate','done','Duplicate acknowledged');
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
  if (btn)   { btn.disabled=true; btn.textContent='⏳ Processing…'; btn.className='btn-run-module running'; btn.style.margin='0'; }
  if (reset) reset.style.display = 'block';
  if (pill)  { pill.textContent='Running…'; pill.className='log-pill running'; }
  const bar  = document.getElementById('commStatusMsg');
  const txt  = document.getElementById('commStatusMsgText');
  if (bar) bar.style.display = 'flex';
  if (txt) txt.textContent   = '⏳ Looking up client and processing commission…';

  const clientName = document.getElementById('comm-client-name')?.value.trim();

  try {
    const res = await authFetch(`${getApiBase()}/api/run/commission`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ client_name: clientName }),
    });
    const { job_id } = await res.json();
    const evtSrc = new EventSource(`${getApiBase()}/api/status/${job_id}?token=${getAuthToken()}`);

    evtSrc.onmessage = (e) => {
      const data = JSON.parse(e.data);
      if (data.type === 'ping') return;
      if (data.type === 'step') {
        setCommStep(data.step, data.state, data.hint);
        if (data.state === 'warning' && data.step === 'duplicate')
          if (pill) { pill.textContent='Awaiting input'; pill.className='log-pill'; }
      }
      if (data.type === 'done')  { evtSrc.close(); commFinish(true,'Done'); }
      if (data.type === 'error') {
        evtSrc.close();
        const activeStep = COMM_STEPS.find(s => commStepStates[s.id]==='active');
        if (activeStep) setCommStep(activeStep.id,'error',data.msg?.split('\n')[0]);
        commFinish(false,'Error');
        if (pill) { pill.textContent='Error'; pill.className='log-pill error'; }
        showError('Commission Error', data.msg?.split('\n')[0] || 'An error occurred.');
      }
    };
    evtSrc.onerror = () => { evtSrc.close(); commFinish(false,'Lost connection'); };
  } catch (e) {
    commFinish(false,'Error');
    showError('Failed to Start', e.message);
  }
}

function commFinish(success, label) {
  const btn  = document.getElementById('commRunBtn');
  const pill = document.getElementById('commStatusPill');
  const txt  = document.getElementById('commStatusMsgText');
  if (btn) { btn.disabled=false; btn.style.margin='0'; btn.className='btn-run-module '+(success?'done':'errored'); btn.textContent=success?'✓ Complete — Run Another':'✗ Skipped — Reset to try again'; }
  if (pill) { pill.textContent=label||(success?'Done':'Error'); pill.className='log-pill '+(success?'done':'error'); }
  if (txt)  txt.textContent = success?'✅ Commission processed successfully!':'❌ Processing stopped — check steps above.';
}

function commReset(silent) {
  commStepStates = {}; commExtracted = {};
  const wrap = document.getElementById('commStepperWrap'); if(wrap) wrap.innerHTML=buildCommStepper();
  const card = document.getElementById('commDataCard');    if(card) card.style.display='none';
  const pill = document.getElementById('commStatusPill'); if(pill) { pill.textContent='Idle'; pill.className='log-pill'; }
  const bar  = document.getElementById('commStatusMsg');  if(bar)  bar.style.display='none';
  const btn  = document.getElementById('commRunBtn');     if(btn)  { btn.className='btn-run-module'; btn.textContent='▶ Search & Process Commission'; btn.style.margin='0'; commInputChanged(); }
  const rst  = document.getElementById('commResetBtn');   if(rst)  rst.style.display='none';
}

async function loadCommSidePanel() {
  const panel = document.getElementById('commSidePanel');
  if (!panel) return;
  let reps = [];
  try { reps = await authFetch(`${getApiBase()}/api/commission/reps`).then(r=>r.json()); } catch {}
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
      <div class="info-card-title" style="display:flex;justify-content:space-between;align-items:center">
        <span>Sales Rep Mapping (${reps.length})</span>
        <button class="btn-open" onclick="toggleCommRepForm()">+ Add</button>
      </div>
      <div id="commAddRepForm" style="display:none;margin-bottom:12px;padding:12px;background:#f8fafc;border:1px solid var(--border);border-radius:8px">
        <div style="display:flex;flex-direction:column;gap:7px">
          <input id="newRepName" placeholder="Sales Rep Name"                   class="hub-input" />
          <input id="newRepUrl"  placeholder="Excel Link (Copy link to sheet)"  class="hub-input" />
        </div>
        <button onclick="addCommRep()" style="margin-top:8px;width:100%;background:var(--navy);color:#fff;border:none;border-radius:7px;padding:7px;font-size:12px;font-weight:600;cursor:pointer">Save Rep</button>
      </div>
      <div id="commRepList">${renderCommRepList(reps)}</div>
    </div>`;
}

function renderCommRepList(reps) {
  if (!reps.length) return '<p style="font-size:12px;color:var(--text-3);padding:4px 0">No reps configured. Click + Add above.</p>';
  return reps.map(r => {
    const ini  = r.name.split(' ').map(p=>p[0]).join('').slice(0,2).toUpperCase();
    const ok   = r.excel_url && r.excel_url.startsWith('http');
    return `
      <div class="emp-mini" style="flex-wrap:wrap;gap:4px">
        <div class="emp-av" style="background:linear-gradient(135deg,#d97706,#f59e0b)">${ini}</div>
        <div style="flex:1;min-width:0">
          <div class="emp-n">${r.name}</div>
          <div style="font-size:10px;color:${ok?'#16a34a':'#f59e0b'};margin-top:1px">${ok?'✓ Excel link saved':'⚠ No Excel link'}</div>
          <div style="margin-top:3px">
            <button onclick="openCommRepEdit('${r.id}')" style="background:none;border:none;color:var(--text-3);font-size:10px;cursor:pointer;padding:0">✏️ Edit</button>
          </div>
          <div id="comm-rep-edit-${r.id}" style="display:none;margin-top:6px;padding:8px;background:#f8fafc;border:1px solid var(--border);border-radius:8px">
            <div style="display:flex;flex-direction:column;gap:5px">
              <input id="cedit-name-${r.id}" value="${r.name}"          placeholder="Name"       class="hub-input" />
              <input id="cedit-url-${r.id}"  value="${r.excel_url||''}" placeholder="Excel Link" class="hub-input" />
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

function toggleCommRepForm()   { const f=document.getElementById('commAddRepForm'); if(f) f.style.display=f.style.display==='none'?'block':'none'; }
function openCommRepEdit(id)   { const el=document.getElementById(`comm-rep-edit-${id}`); if(el) el.style.display='block'; }

async function addCommRep() {
  const name=document.getElementById('newRepName')?.value.trim();
  const url =document.getElementById('newRepUrl')?.value.trim();
  if (!name) { showAlert('Required', 'Name is required.'); return; }
  await authFetch(`${getApiBase()}/api/commission/reps`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name,excel_url:url})});
  const reps=await authFetch(`${getApiBase()}/api/commission/reps`).then(r=>r.json());
  const el=document.getElementById('commRepList'); if(el) el.innerHTML=renderCommRepList(reps);
  ['newRepName','newRepUrl'].forEach(id=>{const e=document.getElementById(id);if(e)e.value='';});
  toggleCommRepForm();
}
async function saveCommRepEdit(id) {
  const name=document.getElementById(`cedit-name-${id}`)?.value.trim();
  const url =document.getElementById(`cedit-url-${id}`)?.value.trim();
  if (!name) { showAlert('Required', 'Name is required.'); return; }
  await authFetch(`${getApiBase()}/api/commission/reps/${id}`,{method:'PATCH',headers:{'Content-Type':'application/json'},body:JSON.stringify({name,excel_url:url})});
  const reps=await authFetch(`${getApiBase()}/api/commission/reps`).then(r=>r.json());
  const el=document.getElementById('commRepList'); if(el) el.innerHTML=renderCommRepList(reps);
}
function deleteCommRep(id) {
  showConfirm('🗑️', 'Remove Sales Rep', 'Are you sure you want to remove this rep?', 'Remove', async () => {
    await authFetch(`${getApiBase()}/api/commission/reps/${id}`,{method:'DELETE'});
    const reps=await authFetch(`${getApiBase()}/api/commission/reps`).then(r=>r.json());
    const el=document.getElementById('commRepList'); if(el) el.innerHTML=renderCommRepList(reps);
  }, true);
}


// ── Chrome connection ──────────────────────────────────────────────────────────

async function pollChromeStatus() {
  if (!CHROME_FREE_MODULES.has(activeModule)) {
    try {
      const res  = await authFetch(`${getApiBase()}/api/chrome/status`);
      const data = await res.json();
      chromeReady = data.ready;
      updateChromeUI();
    } catch {}
  }
  setTimeout(pollChromeStatus, 4000);
}

function updateChromeUI() {
  const pill  = document.getElementById('chromePill');
  const label = document.getElementById('chromeLabel');
  if (chromeReady) { pill?.classList.add('connected'); if(label) label.textContent='Chrome Connected'; }
  else             { pill?.classList.remove('connected'); if(label) label.textContent='Automation Chrome'; }
  syncChromeUI();
}

function syncChromeUI() {
  const step   = document.getElementById('rs-chrome');
  const hint   = document.getElementById('rs-chrome-hint');
  const btn    = document.getElementById('rs-chrome-btn');
  const runBtn = document.getElementById('runModuleBtn');
  if (!step) return;
  if (chromeReady) {
    step.className='run-step step-done';
    if(hint) { hint.textContent='Connected — your sessions are active'; hint.className='step-hint ok'; }
    if(btn)  { btn.textContent='✓ Connected'; btn.className='btn-chrome ok'; btn.disabled=true; }
    const rs2=document.getElementById('rs-run'); if(rs2) rs2.className='run-step step-active';
    if(runBtn) runBtn.disabled=false;
  } else {
    step.className='run-step step-active';
    if(hint) { hint.textContent='Close main Chrome first, then click Connect'; hint.className='step-hint'; }
    if(btn)  { btn.textContent='Connect'; btn.className='btn-chrome'; btn.disabled=false; }
    if(runBtn) runBtn.disabled=true;
  }
}

function handleChromeConnect() {
  updateChromeModalText();
  document.getElementById('chromeOverlay').style.display='flex';
}
function closeModal() { document.getElementById('chromeOverlay').style.display='none'; }

async function doLaunchChrome() {
  const btn=document.getElementById('launchModalBtn');
  btn.disabled=true; btn.textContent='Launching…';
  try {
    const res=await authFetch(`${getApiBase()}/api/chrome/launch`,{method:'POST'});
    const data=await res.json();
    if(data.ok) { chromeReady=true; updateChromeUI(); closeModal(); }
    else { btn.textContent='Retry'; btn.disabled=false; showError('Launch Failed', data.error||'Could not launch Chrome.'); }
  } catch(e) { btn.textContent='Retry'; btn.disabled=false; showError('Error', e.message); }
}


// ── Run module (Rehash) ────────────────────────────────────────────────────────

async function runModule() {
  if (!activeModule) return;
  const btn  = document.getElementById('runModuleBtn');
  const log  = document.getElementById('moduleLog');
  const pill = document.getElementById('logPill');

  btn.disabled=true; btn.className='btn-run-module running'; btn.textContent='⏳ Running…';
  log.innerHTML=''; pill.textContent='Running…'; pill.className='log-pill running';

  const bar=document.getElementById('rehashStatusBar'); const msg=document.getElementById('rehashStatusMsg');
  if(bar) bar.style.display='flex'; if(msg) msg.textContent='⏳ Report is running — live output below…';

  if (currentJobSrc) { currentJobSrc.close(); currentJobSrc=null; }
  const stopBtn=document.getElementById('stopModuleBtn'); if(stopBtn) stopBtn.style.display='block';

  try {
    const res=await authFetch(`${getApiBase()}/api/run/${activeModule}`,{method:'POST'});
    const {job_id}=await res.json();
    currentJobId=job_id;
    currentJobSrc=new EventSource(`${getApiBase()}/api/status/${job_id}?token=${getAuthToken()}`);

    currentJobSrc.onmessage=(e)=>{
      const data=JSON.parse(e.data);
      if(data.type==='ping') return;
      if(data.type==='log') {
        appendModuleLog(log,data.msg);
        if(msg) msg.textContent=data.msg.replace(/^[✅❌⚠️⏳🔄📋📊💾🔍]\s*/,'').slice(0,80)+(data.msg.length>80?'…':'');
      }
      if(data.type==='done')  finishRun(btn,pill,log,true);
      if(data.type==='error') { appendModuleLog(log,data.msg,'error'); finishRun(btn,pill,log,false); }
    };
    currentJobSrc.onerror=()=>{ appendModuleLog(log,'Lost connection to server.','error'); finishRun(btn,pill,log,false); };
  } catch(e) { appendModuleLog(log,`Error: ${e.message}`,'error'); finishRun(btn,pill,log,false); }
}

function appendModuleLog(container, msg, type='') {
  const line=document.createElement('span'); line.className='log-line';
  if(type==='error'||msg.toLowerCase().includes('error')) line.classList.add('error');
  else if(type==='success'||msg.startsWith('✅'))          line.classList.add('success');
  else if(msg.toLowerCase().includes('warning'))           line.classList.add('warn');
  else if(msg.startsWith('━━━'))                           line.classList.add('step');
  line.textContent=msg;
  container.appendChild(line); container.appendChild(document.createElement('br'));
  container.scrollTop=container.scrollHeight;
}

async function stopRun() {
  if (!currentJobId) return;
  const stopBtn=document.getElementById('stopModuleBtn'); if(stopBtn) stopBtn.disabled=true;
  try { await authFetch(`${getApiBase()}/api/cancel/${currentJobId}`,{method:'POST'}); } catch {}
}

function finishRun(btn, pill, log, success) {
  if(currentJobSrc) { currentJobSrc.close(); currentJobSrc=null; }
  currentJobId=null;
  const stopBtn=document.getElementById('stopModuleBtn'); if(stopBtn) { stopBtn.style.display='none'; stopBtn.disabled=false; }
  const msg=document.getElementById('rehashStatusMsg');
  btn.disabled=false;
  if(success) {
    btn.className='btn-run-module done'; btn.textContent='✓ Complete — Run Again';
    pill.textContent='Done'; pill.className='log-pill done';
    if(msg) msg.textContent='✅ Report completed successfully!';
  } else {
    btn.className='btn-run-module errored'; btn.textContent='✗ Failed — Try Again';
    pill.textContent='Error'; pill.className='log-pill error';
    if(msg) msg.textContent='❌ Report failed — check the log below for details.';
  }
}


// ── Commission Estimate module ─────────────────────────────────────────────────

let ceZipBlob     = null;
let ceZipFilename = 'commission_estimate_processed.zip';

function ceCheckReady() {
  const est  = document.getElementById('ceFileEst')?.files[0];
  const comm = document.getElementById('ceFileComm')?.files[0];
  const btn  = document.getElementById('ceProcessBtn');
  if (btn) btn.disabled = !(est && comm);
}

function ceDragOver(e, zoneId)  { e.preventDefault(); document.getElementById(zoneId)?.classList.add('ce-drop-active'); }
function ceDragLeave(zoneId)    { document.getElementById(zoneId)?.classList.remove('ce-drop-active'); }

function ceDrop(e, zoneId, inputId) {
  e.preventDefault();
  ceDragLeave(zoneId);
  const file = e.dataTransfer?.files[0];
  if (file && file.type === 'application/pdf') {
    const input = document.getElementById(inputId);
    const dt    = new DataTransfer();
    dt.items.add(file);
    input.files = dt.files;
    ceFileSelected(inputId, zoneId, `${zoneId}Label`);
  }
}

function ceFileSelected(inputId, zoneId, labelId) {
  const input = document.getElementById(inputId);
  const zone  = document.getElementById(zoneId);
  const label = document.getElementById(labelId);
  const file  = input?.files[0];
  if (file) {
    if (label) label.textContent = `✅ ${file.name}`;
    zone?.classList.add('ce-drop-filled');
    zone?.classList.remove('ce-drop-active');
  }
  ceCheckReady();
}

function renderCommEstimateModule() {
  const pill = document.getElementById('chromePill');
  if (pill) pill.style.display = 'none';

  const view = document.getElementById('view-module');
  view.innerHTML = `
    ${moduleHeroHTML('📈','Commission Estimate','Upload CRM PDFs → Annotated PDF Download')}

    <div class="module-view-body">
      <div style="display:flex;flex-direction:column;gap:16px">

        <div class="run-panel">
          <div class="run-panel-header" style="padding:13px 22px">
            <span class="run-panel-title">Upload PDFs</span>
            <span class="status-badge status-active">Active</span>
          </div>
          <div style="padding:14px 22px 20px;display:flex;flex-direction:column;gap:14px">

            <div>
              <div style="font-size:11px;font-weight:600;color:var(--text-2);margin-bottom:6px;text-transform:uppercase;letter-spacing:.4px">Estimate Details PDF</div>
              <div class="ce-drop-zone" id="ceDropEst" onclick="document.getElementById('ceFileEst').click()"
                   ondragover="ceDragOver(event,'ceDropEst')" ondragleave="ceDragLeave('ceDropEst')" ondrop="ceDrop(event,'ceDropEst','ceFileEst')">
                <span id="ceDropEstLabel">📄 Click or drag to upload</span>
              </div>
              <input type="file" id="ceFileEst" accept=".pdf" style="display:none" onchange="ceFileSelected('ceFileEst','ceDropEst','ceDropEstLabel')" />
            </div>

            <div>
              <div style="font-size:11px;font-weight:600;color:var(--text-2);margin-bottom:6px;text-transform:uppercase;letter-spacing:.4px">Commission Sheet PDF</div>
              <div class="ce-drop-zone" id="ceDropComm" onclick="document.getElementById('ceFileComm').click()"
                   ondragover="ceDragOver(event,'ceDropComm')" ondragleave="ceDragLeave('ceDropComm')" ondrop="ceDrop(event,'ceDropComm','ceFileComm')">
                <span id="ceDropCommLabel">📄 Click or drag to upload</span>
              </div>
              <input type="file" id="ceFileComm" accept=".pdf" style="display:none" onchange="ceFileSelected('ceFileComm','ceDropComm','ceDropCommLabel')" />
            </div>

            <div style="display:flex;gap:12px">
              <div style="flex:1">
                <div style="font-size:11px;font-weight:600;color:var(--text-2);margin-bottom:4px;text-transform:uppercase;letter-spacing:.4px">Finance Fee (optional)</div>
                <input id="ceFinanceFee" type="number" step="0.01" placeholder="e.g. 308.10" class="hub-input" />
              </div>
              <div style="flex:2">
                <div style="font-size:11px;font-weight:600;color:var(--text-2);margin-bottom:4px;text-transform:uppercase;letter-spacing:.4px">Lender Name (optional)</div>
                <input id="ceFinanceLender" type="text" placeholder="e.g. Dividend Finance" class="hub-input" />
              </div>
            </div>

            <button class="btn-run-module" id="ceProcessBtn" onclick="ceProcess()" disabled style="margin:0">
              ▶ Process PDFs
            </button>
          </div>
        </div>

        <!-- Status -->
        <div class="log-wrap" id="ceStatusWrap" style="display:none">
          <div class="log-header">
            <span class="log-title">Status</span>
            <span class="log-pill" id="ceStatusPill">Processing…</span>
          </div>
          <div class="log-body-compact" id="ceStatusLog">
            <span id="ceStatusMsg">Uploading and processing PDFs…</span>
          </div>
        </div>

        <!-- Download card -->
        <div class="log-wrap" id="ceDownloadCard" style="display:none">
          <div class="log-header">
            <span class="log-title">Download</span>
            <span class="log-pill done">Ready</span>
          </div>
          <div style="padding:16px 20px;display:flex;gap:10px">
            <button class="btn-run-module done" onclick="ceDownload()" style="margin:0;flex:1">⬇ Download Annotated PDFs</button>
            <button class="btn-stop" onclick="ceReset()" style="padding:10px 16px;font-size:12px">↺ New</button>
          </div>
        </div>

      </div>

      <!-- RIGHT: info + server config -->
      <div class="side-panel">
        <div class="info-card">
          <div class="info-card-title">How It Works</div>
          <div class="info-row"><span class="info-key">Step 1</span><span class="info-value">Upload both PDFs</span></div>
          <div class="info-row"><span class="info-key">Step 2</span><span class="info-value">Add Finance Fee if applicable</span></div>
          <div class="info-row"><span class="info-key">Step 3</span><span class="info-value">Click Process</span></div>
          <div class="info-row"><span class="info-key">Step 4</span><span class="info-value">Download annotated ZIP</span></div>
        </div>
        <div class="info-card">
          <div class="info-card-title">Server Connection</div>
          <div style="font-size:11px;color:var(--text-3);margin-bottom:6px">Current URL</div>
          <input id="ceApiUrlInput" value="${getApiBase()}" class="hub-input" style="margin-bottom:6px" />
          <button onclick="ceSaveApiUrl()" style="width:100%;background:var(--navy);color:#fff;border:none;border-radius:6px;padding:6px;font-size:12px;font-weight:600;cursor:pointer">Save URL</button>
          <div id="ceApiUrlStatus" style="font-size:10px;color:var(--text-3);margin-top:4px"></div>
        </div>
      </div>
    </div>`;
}

function ceSaveApiUrl() {
  const url    = document.getElementById('ceApiUrlInput')?.value.trim();
  const status = document.getElementById('ceApiUrlStatus');
  if (url) { saveApiBase(url); if(status) { status.textContent='✓ URL saved'; status.style.color='#16a34a'; setTimeout(()=>{if(status)status.textContent='';},2000); } }
}

async function ceProcess() {
  const estFile  = document.getElementById('ceFileEst')?.files[0];
  const commFile = document.getElementById('ceFileComm')?.files[0];
  if (!estFile || !commFile) { showAlert('Files Required', 'Please upload both PDF files first.'); return; }

  const btn = document.getElementById('ceProcessBtn');
  const statusWrap = document.getElementById('ceStatusWrap');
  const statusPill = document.getElementById('ceStatusPill');
  const statusMsg  = document.getElementById('ceStatusMsg');

  if (btn) btn.style.display = 'none';
  if (statusWrap) statusWrap.style.display = 'block';
  if (statusPill) { statusPill.textContent='Processing…'; statusPill.className='log-pill running'; }
  if (statusMsg)  statusMsg.textContent = 'Uploading and processing PDFs — this takes about 5–10 seconds…';

  // Clear previous alerts
  document.querySelectorAll('.ce-alert-card').forEach(el => el.remove());

  const formData = new FormData();
  formData.append('estimate_pdf',  estFile);
  formData.append('commission_pdf', commFile);
  const fee    = document.getElementById('ceFinanceFee')?.value.trim();
  const lender = document.getElementById('ceFinanceLender')?.value.trim();
  if (fee)    formData.append('finance_fee',    fee);
  if (lender) formData.append('finance_lender', lender);

  try {
    const res = await authFetch(`${getApiBase()}/api/commission-estimate/process`, { method:'POST', body: formData });

    if (!res.ok) {
      const err = await res.json().catch(()=>({error:'Unknown error'}));
      throw new Error(err.error || 'Server error');
    }

    const payload = await res.json();

    // Decode zip
    const zipBytes = Uint8Array.from(atob(payload.zip_b64), c => c.charCodeAt(0));
    ceZipBlob = new Blob([zipBytes], { type:'application/zip' });

    // Build filename: ClientName - RepName.zip
    const estFilename = estFile.name || '';
    const clientPart  = estFilename.replace(/\.pdf$/i,'').split(/[_]/)[0].trim().replace(/,\s*$/,'').trim();
    const repPart     = (payload.summary?.rep_name||'').trim();
    ceZipFilename = clientPart && repPart
      ? `${clientPart} - ${repPart}.zip`
      : clientPart ? `${clientPart} - Commission Estimate.zip`
      : 'commission_estimate_processed.zip';

    if (statusPill) { statusPill.textContent='Done'; statusPill.className='log-pill done'; }
    if (statusMsg)  statusMsg.textContent = '✅ PDFs processed successfully!';

    ceShowAlerts(payload.alerts||[], payload.is_sean||false, payload.summary||{});

    const downloadCard = document.getElementById('ceDownloadCard');
    if (downloadCard) downloadCard.style.display='block';

  } catch(e) {
    if (statusPill) { statusPill.textContent='Error'; statusPill.className='log-pill error'; }
    let hint = e.message;
    if (hint.includes('Failed to fetch') || hint.includes('NetworkError')) {
      hint = `Cannot reach server at ${getApiBase()}. Check your Server URL in the settings on the right.`;
    }
    if (statusMsg) statusMsg.textContent = `❌ ${hint}`;
    if (btn) btn.style.display='';
  }
}

function ceShowAlerts(alerts, isSean, summary) {
  const downloadCard = document.getElementById('ceDownloadCard');
  const statusWrap   = document.getElementById('ceStatusWrap');
  const insertAfter  = downloadCard || statusWrap;
  if (!insertAfter) return;

  // Clear existing
  document.querySelectorAll('.ce-alert-card').forEach(el => el.remove());

  // Amber alert cards (inform Anne situations)
  for (const msg of alerts) {
    const card = document.createElement('div');
    card.className = 'log-wrap ce-alert-card';
    card.style.cssText = 'border-left:4px solid #f59e0b;background:#fffbeb';
    card.innerHTML = `<div class="log-header" style="background:#fffbeb"><span class="log-title" style="color:#92400e">🔔 Alert</span></div>
      <div style="padding:12px 18px;font-size:13px;color:#78350f">${msg}</div>`;
    insertAfter.after(card);
  }

  // Blue info card for Sean tiers
  if (isSean) {
    const card = document.createElement('div');
    card.className = 'log-wrap ce-alert-card';
    card.style.cssText = 'border-left:4px solid #3b9eda;background:#eff6ff';
    card.innerHTML = `<div class="log-header" style="background:#eff6ff"><span class="log-title" style="color:#1e40af">ℹ️ Sean Stern Commission Tiers Applied</span></div>
      <div style="padding:12px 18px;font-size:12px;color:#1e3a8a;display:grid;grid-template-columns:1fr 1fr;gap:4px 16px">
        <span>≥ 42% margin → 12%</span><span>≥ 39% → 11%</span>
        <span>≥ 36% → 10%</span><span>≥ 32% → 9%</span>
        <span>below 32% → 5%</span>
      </div>`;
    insertAfter.after(card);
  }

  // Summary grid
  if (Object.keys(summary).length) {
    const thresholds = { profit_margin: 0.36, pct_greenline: 0.36 };
    const rows = [
      { key:'profit_margin', label:'Profit Margin' },
      { key:'pct_greenline', label:'% Greenline'  },
      { key:'commission_pct',label:'Commission %'  },
      { key:'est_commission',label:'Est. Commission'},
    ].filter(r => summary[r.key]);

    if (rows.length) {
      const card = document.createElement('div');
      card.className = 'log-wrap ce-alert-card';
      const rowsHtml = rows.map(r => {
        const val = summary[r.key];
        const num = parseFloat(val?.replace(/[%,$]/g,''));
        const low = thresholds[r.key] && !isNaN(num) && num < thresholds[r.key]*100;
        return `<div style="display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid var(--border)">
          <span style="font-size:12px;color:var(--text-2)">${r.label}</span>
          <span style="font-size:13px;font-weight:700;color:${low?'#ef4444':'var(--text)'}">${val}</span>
        </div>`;
      }).join('');
      card.innerHTML = `<div class="log-header"><span class="log-title">Summary</span></div>
        <div style="padding:8px 18px">${rowsHtml}</div>`;
      insertAfter.after(card);
    }
  }
}

function ceDownload() {
  if (!ceZipBlob) return;
  const url = URL.createObjectURL(ceZipBlob);
  const a   = document.createElement('a');
  a.href    = url;
  a.download= ceZipFilename;
  a.click();
  URL.revokeObjectURL(url);
}

function ceReset() {
  ceZipBlob     = null;
  ceZipFilename = 'commission_estimate_processed.zip';
  renderCommEstimateModule();
}


// ── Admin panel ────────────────────────────────────────────────────────────────

async function renderAdminPanel() {
  const view = document.getElementById('view-admin');
  if (!view) return;
  view.innerHTML = `
    <header class="topbar">
      <div>
        <h1 class="page-title">Admin</h1>
        <p class="page-sub">User management &amp; access control</p>
      </div>
    </header>
    <div style="padding:0 32px 32px;max-width:900px">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px">
        <span style="font-size:13px;font-weight:600;color:var(--text)">Users</span>
        <button onclick="adminShowCreateForm()" class="btn-open" style="padding:7px 14px;font-size:12px">+ New User</button>
      </div>

      <!-- Create user form (hidden) -->
      <div id="adminCreateForm" style="display:none;background:var(--card);border:1px solid var(--border);border-radius:12px;padding:20px;margin-bottom:16px">
        <div style="font-size:13px;font-weight:700;margin-bottom:14px">New User</div>
        <div style="display:flex;flex-direction:column;gap:10px">
          <div style="display:flex;gap:10px">
            <input id="adminNewUsername" placeholder="Username" class="hub-input" style="flex:1" />
            <input id="adminNewPassword" type="password" placeholder="Password (min 6)" class="hub-input" style="flex:1" />
          </div>
          <div id="adminModuleCheckboxes" style="display:flex;flex-wrap:wrap;gap:8px;padding:10px;background:#f8fafc;border:1px solid var(--border);border-radius:8px">
            <span style="font-size:11px;color:var(--text-3)">Loading modules…</span>
          </div>
          <div id="adminCreateError" style="display:none;font-size:12px;color:#ef4444"></div>
          <div style="display:flex;gap:8px">
            <button onclick="adminCreateUser()" style="background:var(--navy);color:#fff;border:none;border-radius:7px;padding:8px 18px;font-size:12px;font-weight:600;cursor:pointer">Create User</button>
            <button onclick="adminHideCreateForm()" style="background:none;border:1px solid var(--border);border-radius:7px;padding:8px 14px;font-size:12px;cursor:pointer">Cancel</button>
          </div>
        </div>
      </div>

      <div id="adminUserTable" style="background:var(--card);border:1px solid var(--border);border-radius:12px;overflow:hidden">
        <div style="padding:20px;text-align:center;color:var(--text-3);font-size:13px">Loading users…</div>
      </div>
    </div>`;

  await adminLoadData();
}

let adminAllModules = [];

async function adminLoadData() {
  try {
    const [users, modules] = await Promise.all([
      authFetch(`${getApiBase()}/api/admin/users`).then(r=>r.json()),
      authFetch(`${getApiBase()}/api/admin/modules`).then(r=>r.json()),
    ]);
    adminAllModules = modules;
    adminRenderUsers(users);
    adminRenderModuleCheckboxes('adminModuleCheckboxes', []);
  } catch(e) {
    const t = document.getElementById('adminUserTable');
    if(t) t.innerHTML = `<div style="padding:20px;color:#ef4444;font-size:13px">Failed to load: ${e.message}</div>`;
  }
}

function adminRenderUsers(users) {
  const table = document.getElementById('adminUserTable');
  if (!table) return;
  if (!users.length) {
    table.innerHTML = '<div style="padding:20px;text-align:center;color:var(--text-3);font-size:13px">No users yet.</div>';
    return;
  }
  table.innerHTML = `
    <table style="width:100%;border-collapse:collapse">
      <thead>
        <tr style="border-bottom:1px solid var(--border);background:#f8fafc">
          <th style="padding:10px 16px;text-align:left;font-size:11px;color:var(--text-3);font-weight:600;text-transform:uppercase;letter-spacing:.4px">Username</th>
          <th style="padding:10px 16px;text-align:left;font-size:11px;color:var(--text-3);font-weight:600;text-transform:uppercase;letter-spacing:.4px">Role</th>
          <th style="padding:10px 16px;text-align:left;font-size:11px;color:var(--text-3);font-weight:600;text-transform:uppercase;letter-spacing:.4px">Module Access</th>
          <th style="padding:10px 16px;text-align:left;font-size:11px;color:var(--text-3);font-weight:600;text-transform:uppercase;letter-spacing:.4px">Created</th>
          <th style="padding:10px 16px;text-align:left;font-size:11px;color:var(--text-3);font-weight:600;text-transform:uppercase;letter-spacing:.4px"></th>
        </tr>
      </thead>
      <tbody>
        ${users.map(u => {
          const isSelf     = u.id === currentUser?.user_id;
          const modList    = (u.allowed_modules||[]).includes('*') ? 'All modules' : (u.allowed_modules||[]).join(', ') || 'None';
          const roleLabel  = u.is_admin ? '🛡 Admin' : '👤 User';
          const createdDate= new Date(u.created_at).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'});
          return `
            <tr style="border-bottom:1px solid var(--border)" id="admin-row-${u.id}">
              <td style="padding:12px 16px;font-size:13px;font-weight:600">${u.username}${isSelf?' <span style="font-size:10px;color:var(--text-3)">(you)</span>':''}</td>
              <td style="padding:12px 16px;font-size:12px;color:var(--text-2)">${roleLabel}</td>
              <td style="padding:12px 16px;font-size:12px;color:var(--text-2);max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${modList}">${modList}</td>
              <td style="padding:12px 16px;font-size:12px;color:var(--text-3)">${createdDate}</td>
              <td style="padding:12px 16px">
                ${isSelf ? '' : `
                  <div style="display:flex;gap:6px">
                    <button onclick="adminEditUser(${u.id})" style="background:none;border:1px solid var(--border);border-radius:5px;padding:3px 8px;font-size:11px;cursor:pointer">Edit</button>
                    <button onclick="adminResetPw(${u.id})" style="background:none;border:1px solid #bfdbfe;border-radius:5px;color:#1d4ed8;padding:3px 8px;font-size:11px;cursor:pointer">Reset PW</button>
                    <button onclick="adminDeleteUser(${u.id},'${u.username}')" style="background:none;border:1px solid #fca5a5;border-radius:5px;color:#ef4444;padding:3px 8px;font-size:11px;cursor:pointer">Delete</button>
                  </div>`}
              </td>
            </tr>
            <tr id="admin-edit-row-${u.id}" style="display:none;background:#f8fafc">
              <td colspan="5" style="padding:16px">
                <div style="display:flex;flex-direction:column;gap:10px">
                  <div style="font-size:12px;font-weight:600;color:var(--text-2)">Edit Modules — ${u.username}</div>
                  <div id="adminEditModules-${u.id}" style="display:flex;flex-wrap:wrap;gap:8px;padding:10px;background:#fff;border:1px solid var(--border);border-radius:8px">
                    ${adminRenderModuleCheckboxes(`adminEditModules-${u.id}`, u.allowed_modules||[], true)}
                  </div>
                  <div style="display:flex;gap:8px">
                    <button onclick="adminSaveUser(${u.id})" style="background:var(--navy);color:#fff;border:none;border-radius:7px;padding:7px 16px;font-size:12px;font-weight:600;cursor:pointer">Save</button>
                    <button onclick="document.getElementById('admin-edit-row-${u.id}').style.display='none'" style="background:none;border:1px solid var(--border);border-radius:7px;padding:7px 12px;font-size:12px;cursor:pointer">Cancel</button>
                  </div>
                </div>
              </td>
            </tr>
            <tr id="admin-pw-row-${u.id}" style="display:none;background:#eff6ff">
              <td colspan="5" style="padding:16px">
                <div style="display:flex;flex-direction:column;gap:10px">
                  <div style="font-size:12px;font-weight:600;color:#1d4ed8">Reset Password — ${u.username}</div>
                  <input id="adminNewPw-${u.id}" type="password" placeholder="New password" class="hub-input" />
                  <input id="adminConfirmPw-${u.id}" type="password" placeholder="Confirm new password" class="hub-input" />
                  <div id="adminPwError-${u.id}" style="display:none;font-size:12px;color:var(--red);background:#fef2f2;border:1px solid #fecaca;border-radius:7px;padding:8px 12px"></div>
                  <div style="display:flex;gap:8px">
                    <button onclick="adminSaveResetPw(${u.id})" style="background:#1d4ed8;color:#fff;border:none;border-radius:7px;padding:7px 16px;font-size:12px;font-weight:600;cursor:pointer">Set Password</button>
                    <button onclick="document.getElementById('admin-pw-row-${u.id}').style.display='none'" style="background:none;border:1px solid var(--border);border-radius:7px;padding:7px 12px;font-size:12px;cursor:pointer">Cancel</button>
                  </div>
                </div>
              </td>
            </tr>`;
        }).join('')}
      </tbody>
    </table>`;
}

function adminRenderModuleCheckboxes(containerId, selected, returnHtml = false) {
  const all     = selected.includes('*');
  const html    = adminAllModules.map(m => {
    const checked = all || selected.includes(m.id) ? 'checked' : '';
    return `<label style="display:flex;align-items:center;gap:5px;font-size:12px;cursor:pointer;white-space:nowrap">
      <input type="checkbox" value="${m.id}" ${checked} style="cursor:pointer" /> ${m.icon} ${m.name}
    </label>`;
  }).join('');

  if (returnHtml) return html;

  const el = document.getElementById(containerId);
  if (el) el.innerHTML = html;
}

function adminShowCreateForm() {
  document.getElementById('adminCreateForm').style.display = 'block';
  adminRenderModuleCheckboxes('adminModuleCheckboxes', []);
}
function adminHideCreateForm() { document.getElementById('adminCreateForm').style.display = 'none'; }

function adminEditUser(userId) {
  const editRow = document.getElementById(`admin-edit-row-${userId}`);
  if (editRow) editRow.style.display = editRow.style.display === 'none' ? 'table-row' : 'none';
}

async function adminCreateUser() {
  const username = document.getElementById('adminNewUsername')?.value.trim();
  const password = document.getElementById('adminNewPassword')?.value;
  const errEl    = document.getElementById('adminCreateError');
  const modules  = [...document.querySelectorAll('#adminModuleCheckboxes input:checked')].map(el=>el.value);

  if (!username || !password) { if(errEl){errEl.textContent='Username and password required.';errEl.style.display='block';} return; }

  try {
    const res  = await authFetch(`${getApiBase()}/api/admin/users`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username,password,allowed_modules:modules})});
    const data = await res.json();
    if (!res.ok) { if(errEl){errEl.textContent=data.error||'Failed.';errEl.style.display='block';} return; }
    adminHideCreateForm();
    document.getElementById('adminNewUsername').value='';
    document.getElementById('adminNewPassword').value='';
    await adminLoadData();
  } catch(e) { if(errEl){errEl.textContent=e.message;errEl.style.display='block';} }
}

async function adminSaveUser(userId) {
  const modules = [...document.querySelectorAll(`#adminEditModules-${userId} input:checked`)].map(el=>el.value);
  try {
    const res  = await authFetch(`${getApiBase()}/api/admin/users/${userId}`,{method:'PATCH',headers:{'Content-Type':'application/json'},body:JSON.stringify({allowed_modules:modules})});
    const data = await res.json();
    if (!res.ok) { showError('Save Failed', data.error||'Could not save changes.'); return; }
    await adminLoadData();
  } catch(e) { showError('Error', e.message); }
}

function adminResetPw(userId) {
  // Close edit row if open, toggle the reset-pw row
  document.getElementById(`admin-edit-row-${userId}`).style.display = 'none';
  const row = document.getElementById(`admin-pw-row-${userId}`);
  row.style.display = row.style.display === 'none' ? 'table-row' : 'none';
  // Clear fields when opening
  if (row.style.display !== 'none') {
    document.getElementById(`adminNewPw-${userId}`).value = '';
    document.getElementById(`adminConfirmPw-${userId}`).value = '';
    document.getElementById(`adminPwError-${userId}`).style.display = 'none';
    document.getElementById(`adminNewPw-${userId}`).focus();
  }
}

async function adminSaveResetPw(userId) {
  const pw      = document.getElementById(`adminNewPw-${userId}`)?.value;
  const confirm = document.getElementById(`adminConfirmPw-${userId}`)?.value;
  const errEl   = document.getElementById(`adminPwError-${userId}`);

  if (!pw || pw.length < 6) {
    errEl.textContent = 'Password must be at least 6 characters.';
    errEl.style.display = 'block'; return;
  }
  if (pw !== confirm) {
    errEl.textContent = 'Passwords do not match.';
    errEl.style.display = 'block'; return;
  }
  errEl.style.display = 'none';

  try {
    const res  = await authFetch(`${getApiBase()}/api/admin/users/${userId}`,{method:'PATCH',headers:{'Content-Type':'application/json'},body:JSON.stringify({password:pw})});
    const data = await res.json();
    if (!res.ok) { errEl.textContent = data.error||'Failed.'; errEl.style.display='block'; return; }
    document.getElementById(`admin-pw-row-${userId}`).style.display = 'none';
    showSuccess('Password Updated', 'The password has been changed successfully.');
  } catch(e) { errEl.textContent = e.message; errEl.style.display='block'; }
}

function adminDeleteUser(userId, username) {
  showConfirm('🗑️', 'Delete User', `Are you sure you want to delete <strong>${username}</strong>?<br>This cannot be undone.`,
    'Delete', async () => {
      try {
        const res  = await authFetch(`${getApiBase()}/api/admin/users/${userId}`,{method:'DELETE'});
        const data = await res.json();
        if (!res.ok) { showError('Delete Failed', data.error||'Could not delete user.'); return; }
        await adminLoadData();
      } catch(e) { showError('Error', e.message); }
    }, true);
}


// ══════════════════════════════════════════════════════════════════════════════
//  ONSITE PAYROLL REPORT MODULE
// ══════════════════════════════════════════════════════════════════════════════

let onsiteEmployees  = [];   // [{name, department}, ...]
let onsiteDepts      = [];
let onsiteDropFile   = null;
let onsiteRunning    = false;
let onsiteSaveTimer  = null;

function renderOnsiteReconView() {
  const wrap = document.getElementById('view-module');
  wrap.innerHTML = `
    ${moduleHeroHTML('💼','Onsite Payroll Report','Upload a Gusto payroll file — get a formatted Excel report instantly.')}

    <div class="module-view-body">

      <!-- ── Left: Run panel ── -->
      <div style="display:flex;flex-direction:column;gap:16px">

        <div class="run-panel">
          <div class="run-panel-header">
            <span class="run-panel-title">Upload &amp; Generate</span>
          </div>
          <div style="padding:20px 22px;display:flex;flex-direction:column;gap:14px">

            <!-- CSV Drop Zone -->
            <div>
              <label class="comm-field-label">Gusto Payroll CSV</label>
              <div class="ce-drop-zone" id="onsiteCsvZone"
                   onclick="document.getElementById('onsiteCsvInput').click()"
                   ondragover="event.preventDefault();this.classList.add('ce-drop-active')"
                   ondragleave="this.classList.remove('ce-drop-active')"
                   ondrop="onsiteHandleDrop(event)">
                📄 Click or drag your Gusto file here (.xls, .xlsx, .csv)
              </div>
              <input id="onsiteCsvInput" type="file" accept=".csv,.xls,.xlsx" style="display:none"
                     onchange="onsiteHandleFileSelect(this.files[0])"/>
            </div>

            <button class="btn-run-module" id="onsiteRunBtn" onclick="onsiteRun()" disabled>
              Generate Report
            </button>
          </div>

          <!-- Status bar -->
          <div class="log-body-compact" id="onsiteStatus">
            <span>⏳ Upload a CSV to get started.</span>
          </div>
        </div>

        <!-- Download card (hidden until done) -->
        <div id="onsiteDownloadCard" style="display:none;background:var(--card);border:1px solid #86efac;border-radius:16px;padding:20px 22px">
          <div style="font-size:13px;font-weight:700;color:var(--green);margin-bottom:4px">✅ Report ready!</div>
          <div id="onsiteFilename" style="font-size:12px;color:var(--text-2);margin-bottom:14px"></div>
          <button class="btn-run-module" id="onsiteDownloadBtn" style="margin:0;width:100%"
                  onclick="onsiteDownload()">⬇ Download Excel Report</button>
          <div id="onsiteUnmatched" style="display:none;margin-top:12px;background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:12px 14px;font-size:12px;color:#92400e"></div>
        </div>

      </div>

      <!-- ── Right: Employee setup ── -->
      <div class="side-panel">
        <div class="info-card" style="padding:0;overflow:hidden">
          <div style="display:flex;align-items:center;justify-content:space-between;padding:16px 18px;border-bottom:1px solid var(--border)">
            <span class="info-card-title" style="margin:0">Employee Setup</span>
            <button onclick="onsiteAddEmployee()" style="font-size:11px;font-weight:700;padding:4px 10px;border-radius:6px;border:none;background:var(--navy);color:#fff;cursor:pointer">+ Add</button>
          </div>

          <!-- Search -->
          <div style="padding:10px 14px;border-bottom:1px solid var(--border)">
            <input id="onsiteSearch" class="comm-input" placeholder="Search employees…"
                   oninput="onsiteRenderList()" style="font-size:12px;padding:6px 10px"/>
          </div>

          <!-- Employee list -->
          <div id="onsiteEmpList" style="max-height:420px;overflow-y:auto">
            <div style="padding:20px;text-align:center;color:var(--text-3);font-size:12px">Loading…</div>
          </div>

          <!-- Save button -->
          <div style="padding:12px 14px;border-top:1px solid var(--border);display:flex;gap:8px;align-items:center">
            <button onclick="onsiteSaveEmployees()" style="flex:1;padding:8px;background:var(--navy);color:#fff;border:none;border-radius:8px;font-size:12px;font-weight:700;cursor:pointer">Save Changes</button>
            <span id="onsiteSaveStatus" style="font-size:11px;font-weight:600;color:var(--green)"></span>
          </div>
        </div>
      </div>

    </div>`;

  onsiteLoadEmployees();
}

// ── File handling ──────────────────────────────────────────────────────────────

function onsiteHandleFileSelect(file) {
  if (!file) return;
  onsiteDropFile = file;
  const zone = document.getElementById('onsiteCsvZone');
  zone.textContent = `✅ ${file.name}`;
  zone.classList.add('ce-drop-filled');
  zone.classList.remove('ce-drop-active');
  document.getElementById('onsiteRunBtn').disabled = false;
  document.getElementById('onsiteDownloadCard').style.display = 'none';
  onsiteSetStatus(`📄 Ready — ${file.name}`);
}

function onsiteHandleDrop(e) {
  e.preventDefault();
  const zone = document.getElementById('onsiteCsvZone');
  zone.classList.remove('ce-drop-active');
  const file = e.dataTransfer.files[0];
  if (file && /\.(csv|xls|xlsx)$/i.test(file.name)) {
    onsiteHandleFileSelect(file);
  } else {
    onsiteSetStatus('⚠️ Please drop a .xls, .xlsx, or .csv file.');
  }
}

function onsiteSetStatus(msg) {
  const el = document.getElementById('onsiteStatus');
  if (el) el.innerHTML = `<span>${msg}</span>`;
}

// ── Run ────────────────────────────────────────────────────────────────────────

let onsiteResultData = null;

async function onsiteRun() {
  if (!onsiteDropFile || onsiteRunning) return;
  onsiteRunning = true;

  const btn = document.getElementById('onsiteRunBtn');
  btn.disabled = true;
  btn.textContent = 'Generating…';
  btn.classList.add('running');
  document.getElementById('onsiteDownloadCard').style.display = 'none';
  onsiteSetStatus('⚙️ Processing payroll CSV…');

  try {
    const form = new FormData();
    form.append('payroll_csv', onsiteDropFile);

    const res  = await authFetch(`${getApiBase()}/api/run/onsite_recon`, { method: 'POST', body: form });
    const data = await res.json();

    if (!res.ok) {
      onsiteSetStatus(`❌ ${data.error || 'Processing failed.'}`);
      btn.classList.remove('running');
      btn.classList.add('errored');
      btn.textContent = 'Generate Report';
      btn.disabled = false;
      onsiteRunning = false;
      return;
    }

    onsiteResultData = data;
    onsiteSetStatus('✅ Report generated successfully.');
    btn.classList.remove('running');
    btn.classList.add('done');
    btn.textContent = '✓ Done';

    // Show download card
    const card = document.getElementById('onsiteDownloadCard');
    card.style.display = 'block';
    document.getElementById('onsiteFilename').textContent = data.filename;

    // Show unmatched warning if any
    if (data.unmatched && data.unmatched.length > 0) {
      const um = document.getElementById('onsiteUnmatched');
      um.style.display = 'block';
      um.innerHTML = `<strong>⚠️ Unmatched employees (assigned to "No Group"):</strong><br>${data.unmatched.map(n=>`• ${n}`).join('<br>')}`;
    }

  } catch(e) {
    onsiteSetStatus(`❌ ${e.message}`);
    btn.classList.remove('running');
    btn.textContent = 'Generate Report';
    btn.disabled = false;
  }
  onsiteRunning = false;
}

function onsiteDownload() {
  if (!onsiteResultData?.xlsx_b64) return;
  const bytes  = atob(onsiteResultData.xlsx_b64);
  const arr    = new Uint8Array(bytes.length);
  for (let i = 0; i < bytes.length; i++) arr[i] = bytes.charCodeAt(i);
  const blob = new Blob([arr], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = onsiteResultData.filename;
  a.click();
  URL.revokeObjectURL(url);
}

// ── Employee management ────────────────────────────────────────────────────────

async function onsiteLoadEmployees() {
  try {
    const res  = await authFetch(`${getApiBase()}/api/onsite_recon/employees`);
    const data = await res.json();
    onsiteEmployees = data.employees || [];
    onsiteDepts     = data.departments || [];
    onsiteRenderList();
  } catch(e) {
    document.getElementById('onsiteEmpList').innerHTML =
      `<div style="padding:16px;color:var(--red);font-size:12px">Failed to load employees.</div>`;
  }
}

function onsiteRenderList() {
  const query = (document.getElementById('onsiteSearch')?.value || '').toLowerCase();
  const list  = document.getElementById('onsiteEmpList');
  if (!list) return;

  const filtered = onsiteEmployees.filter(e =>
    e.name.toLowerCase().includes(query) || e.department.toLowerCase().includes(query)
  );

  if (!filtered.length) {
    list.innerHTML = `<div style="padding:16px;text-align:center;color:var(--text-3);font-size:12px">${query ? 'No matches.' : 'No employees yet.'}</div>`;
    return;
  }

  const deptOptions = onsiteDepts.map(d => `<option value="${d}">${d}</option>`).join('');

  list.innerHTML = filtered.map((emp, i) => {
    const realIdx = onsiteEmployees.indexOf(emp);
    const opts = onsiteDepts.map(d =>
      `<option value="${d}" ${d === emp.department ? 'selected' : ''}>${d}</option>`
    ).join('');
    return `
      <div style="display:flex;align-items:center;gap:8px;padding:8px 14px;border-bottom:1px solid var(--border)">
        <span style="flex:1;font-size:12px;font-weight:500;color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${emp.name}">${emp.name}</span>
        <select onchange="onsiteEmployees[${realIdx}].department=this.value;onsiteAutoSave()"
                style="font-size:11px;border:1px solid var(--border);border-radius:5px;padding:3px 6px;color:var(--text-2);background:#fff;cursor:pointer">${opts}</select>
        <button onclick="onsiteRemoveEmployee(${realIdx})"
                style="font-size:11px;color:var(--red);background:none;border:none;cursor:pointer;padding:0 2px" title="Remove">✕</button>
      </div>`;
  }).join('');
}

function onsiteAddEmployee() {
  showFormModal('👤', 'Add Employee', [
    { id: 'empName', label: 'Full Name', placeholder: 'e.g. John Smith', required: true },
    { id: 'empDept', label: 'Department', type: 'select', options: onsiteDepts, required: true },
  ], 'Add Employee', ({ empName, empDept }) => {
    const normalized = onsiteNormalizeName(empName);
    onsiteEmployees.push({ name: normalized, department: empDept });
    onsiteRenderList();
    onsiteAutoSave();
  });
}

function onsiteNormalizeName(name) {
  const SUFFIX_RE = /\b(JR|SR|III|II|IV|JUNIOR|SENIOR)\.?\b/gi;
  name = name.replace(/[*.]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
  name = name.replace(SUFFIX_RE, '').replace(/\s+/g, ' ').trim();

  if (name.includes(',')) {
    const [last, ...rest] = name.split(',');
    const first = rest.join('').trim().split(' ')[0] || '';
    return `${last.trim()},${first}`;
  }

  // "Firstname Lastname" → "Lastname,Firstname"
  const parts = name.split(' ');
  if (parts.length >= 2) {
    return `${parts[parts.length - 1]},${parts[0]}`;
  }
  return name;
}

function onsiteRemoveEmployee(idx) {
  onsiteEmployees.splice(idx, 1);
  onsiteRenderList();
  onsiteAutoSave();
}

async function onsiteSaveEmployees() {
  const status = document.getElementById('onsiteSaveStatus');
  if (status) { status.textContent = 'Saving…'; status.style.color = 'var(--text-3)'; }
  try {
    const res  = await authFetch(`${getApiBase()}/api/onsite_recon/employees`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ employees: onsiteEmployees }),
    });
    const data = await res.json();
    if (!res.ok) {
      if (status) { status.textContent = '✗ Save failed'; status.style.color = 'var(--red)'; }
      return;
    }
    if (status) {
      status.textContent = '✓ Saved';
      status.style.color = 'var(--green)';
      setTimeout(() => { if (status) status.textContent = ''; }, 2500);
    }
  } catch(e) {
    if (status) { status.textContent = '✗ Error'; status.style.color = 'var(--red)'; }
  }
}

// Auto-save with 800ms debounce — called after every change
function onsiteAutoSave() {
  if (onsiteSaveTimer) clearTimeout(onsiteSaveTimer);
  const status = document.getElementById('onsiteSaveStatus');
  if (status) { status.textContent = '…'; status.style.color = 'var(--text-3)'; }
  onsiteSaveTimer = setTimeout(() => onsiteSaveEmployees(), 800);
}


// ══════════════════════════════════════════════════════════════════════════════
//  5160 REPORT MODULE
// ══════════════════════════════════════════════════════════════════════════════

let bankEmployees   = [];   // [{key, label, group}, ...]
let bankGroups      = [];
let bankBankFile    = null;
let bankWiseFile    = null;
let bankWuFile      = null;
let bankRunning     = false;
let bankResultData  = null;
let bankSaveTimer   = null;

function renderBankCategorizerView() {
  const wrap = document.getElementById('view-module');
  wrap.innerHTML = `
    ${moduleHeroHTML('🏦','5160 Report','Upload your bank statement to generate a categorized Excel report.')}

    <div class="module-view-body">

      <!-- ── Left: Upload + Run ── -->
      <div style="display:flex;flex-direction:column;gap:16px">

        <div class="run-panel">
          <div class="run-panel-header">
            <span class="run-panel-title">Upload Files</span>
          </div>
          <div style="padding:20px 22px;display:flex;flex-direction:column;gap:16px">

            <!-- Bank Statement (required) -->
            <div>
              <label class="comm-field-label">
                Bank Statement <span style="color:var(--red)">*</span>
                <span style="font-size:10px;font-weight:400;color:var(--text-3);margin-left:6px">Required · .xlsx</span>
              </label>
              <div class="ce-drop-zone" id="bankBankZone"
                   onclick="document.getElementById('bankBankInput').click()"
                   ondragover="event.preventDefault();this.classList.add('ce-drop-active')"
                   ondragleave="this.classList.remove('ce-drop-active')"
                   ondrop="bankHandleDrop(event,'bank')">
                📊 Click or drag Bank Statement (.xlsx)
              </div>
              <input id="bankBankInput" type="file" accept=".xlsx,.xls" style="display:none"
                     onchange="bankHandleFile('bank',this.files[0])"/>
            </div>

            <!-- Wise PDF (optional) -->
            <div>
              <label class="comm-field-label">
                Wise PDF
                <span style="font-size:10px;font-weight:400;color:var(--text-3);margin-left:6px">Optional · used to identify Wise recipients</span>
              </label>
              <div class="ce-drop-zone" id="bankWiseZone"
                   onclick="document.getElementById('bankWiseInput').click()"
                   ondragover="event.preventDefault();this.classList.add('ce-drop-active')"
                   ondragleave="this.classList.remove('ce-drop-active')"
                   ondrop="bankHandleDrop(event,'wise')">
                📄 Click or drag Wise PDF (optional)
              </div>
              <input id="bankWiseInput" type="file" accept=".pdf" style="display:none"
                     onchange="bankHandleFile('wise',this.files[0])"/>
            </div>

            <!-- WU XLSX (optional) -->
            <div>
              <label class="comm-field-label">
                Western Union History
                <span style="font-size:10px;font-weight:400;color:var(--text-3);margin-left:6px">Optional · used to identify WU recipients</span>
              </label>
              <div class="ce-drop-zone" id="bankWuZone"
                   onclick="document.getElementById('bankWuInput').click()"
                   ondragover="event.preventDefault();this.classList.add('ce-drop-active')"
                   ondragleave="this.classList.remove('ce-drop-active')"
                   ondrop="bankHandleDrop(event,'wu')">
                💳 Click or drag WU History (.xlsx)
              </div>
              <input id="bankWuInput" type="file" accept=".xlsx,.xls" style="display:none"
                     onchange="bankHandleFile('wu',this.files[0])"/>
            </div>

            <button class="btn-run-module" id="bankRunBtn" onclick="bankRun()" disabled>
              Generate Report
            </button>
          </div>

          <!-- Status bar -->
          <div class="log-body-compact" id="bankStatus">
            <span>⏳ Upload the bank statement to get started.</span>
          </div>
        </div>

        <!-- Download card (hidden until done) -->
        <div id="bankDownloadCard" style="display:none;background:var(--card);border:1px solid #86efac;border-radius:16px;padding:20px 22px">
          <div style="font-size:13px;font-weight:700;color:var(--green);margin-bottom:4px">✅ Report ready!</div>
          <div id="bankFilename" style="font-size:12px;color:var(--text-2);margin-bottom:14px"></div>
          <button class="btn-run-module" style="margin:0;width:100%;margin-bottom:10px"
                  onclick="bankDownload()">⬇ Download Categorized Excel</button>
          <button onclick="bankResetFiles()" style="width:100%;background:none;border:1px solid var(--border);border-radius:8px;padding:8px;font-size:12px;color:var(--text-2);cursor:pointer">↩ Run Another</button>
        </div>

        <!-- Suggestions card (hidden until unmatched employees found) -->
        <div id="bankSuggestionsCard" style="display:none"></div>

      </div>

      <!-- ── Right: Employee map ── -->
      <div class="side-panel">
        <div class="info-card" style="padding:0;overflow:hidden">
          <div style="display:flex;align-items:center;justify-content:space-between;padding:16px 18px;border-bottom:1px solid var(--border)">
            <span class="info-card-title" style="margin:0">Employee Map</span>
            <button onclick="bankAddEmployee()" style="font-size:11px;font-weight:700;padding:4px 10px;border-radius:6px;border:none;background:var(--navy);color:#fff;cursor:pointer">+ Add</button>
          </div>

          <!-- Search -->
          <div style="padding:10px 14px;border-bottom:1px solid var(--border)">
            <input id="bankSearch" class="comm-input" placeholder="Search name or group…"
                   oninput="bankRenderList()" style="font-size:12px;padding:6px 10px"/>
          </div>

          <!-- Employee list -->
          <div id="bankEmpList" style="max-height:480px;overflow-y:auto">
            <div style="padding:20px;text-align:center;color:var(--text-3);font-size:12px">Loading…</div>
          </div>

          <!-- Save button -->
          <div style="padding:12px 14px;border-top:1px solid var(--border);display:flex;gap:8px;align-items:center">
            <button onclick="bankSaveEmployees()" style="flex:1;padding:8px;background:var(--navy);color:#fff;border:none;border-radius:8px;font-size:12px;font-weight:700;cursor:pointer">Save Changes</button>
            <span id="bankSaveStatus" style="font-size:11px;font-weight:600;color:var(--green)"></span>
          </div>
        </div>
      </div>

    </div>`;

  bankLoadEmployees();
}

// ── File handling ──────────────────────────────────────────────────────────────

function bankHandleFile(type, file) {
  if (!file) return;
  if (type === 'bank') {
    bankBankFile = file;
    const z = document.getElementById('bankBankZone');
    z.textContent = `✅ ${file.name}`;
    z.classList.add('ce-drop-filled');
    document.getElementById('bankDownloadCard').style.display = 'none';
    document.getElementById('bankSuggestionsCard').style.display = 'none';
  } else if (type === 'wise') {
    bankWiseFile = file;
    const z = document.getElementById('bankWiseZone');
    z.textContent = `✅ ${file.name}`;
    z.classList.add('ce-drop-filled');
  } else if (type === 'wu') {
    bankWuFile = file;
    const z = document.getElementById('bankWuZone');
    z.textContent = `✅ ${file.name}`;
    z.classList.add('ce-drop-filled');
  }
  document.getElementById('bankRunBtn').disabled = !bankBankFile;
  bankSetStatus(`📄 Ready${bankWiseFile ? ' · Wise attached' : ''}${bankWuFile ? ' · WU attached' : ''}`);
}

function bankHandleDrop(e, type) {
  e.preventDefault();
  const zoneId = type === 'bank' ? 'bankBankZone' : type === 'wise' ? 'bankWiseZone' : 'bankWuZone';
  document.getElementById(zoneId)?.classList.remove('ce-drop-active');
  const file = e.dataTransfer.files[0];
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  if (type === 'wise' && ext !== 'pdf') { bankSetStatus('⚠️ Wise file must be a PDF.'); return; }
  if ((type === 'bank' || type === 'wu') && !['xlsx','xls'].includes(ext)) { bankSetStatus('⚠️ Please use an .xlsx or .xls file.'); return; }
  bankHandleFile(type, file);
}

function bankSetStatus(msg) {
  const el = document.getElementById('bankStatus');
  if (el) el.innerHTML = `<span>${msg}</span>`;
}

// ── Run ────────────────────────────────────────────────────────────────────────

async function bankRun() {
  if (!bankBankFile || bankRunning) return;
  bankRunning = true;

  const btn = document.getElementById('bankRunBtn');
  btn.disabled = true;
  btn.textContent = 'Generating…';
  btn.classList.add('running');
  document.getElementById('bankDownloadCard').style.display = 'none';
  document.getElementById('bankSuggestionsCard').style.display = 'none';
  bankSetStatus('⚙️ Categorizing transactions…');

  try {
    const form = new FormData();
    form.append('bank_xlsx', bankBankFile);
    if (bankWiseFile) form.append('wise_pdf',  bankWiseFile);
    if (bankWuFile)   form.append('wu_xlsx',   bankWuFile);

    const res  = await authFetch(`${getApiBase()}/api/run/bank_categorizer`, { method: 'POST', body: form });
    const data = await res.json();

    if (!res.ok) {
      bankSetStatus(`❌ ${data.error || 'Processing failed.'}`);
      btn.classList.remove('running');
      btn.classList.add('errored');
      btn.textContent = 'Generate Report';
      btn.disabled = false;
      bankRunning = false;
      return;
    }

    bankResultData = data;
    bankSetStatus('✅ Report categorized successfully.');
    btn.classList.remove('running');
    btn.classList.add('done');
    btn.textContent = '✓ Done';

    // Show download card
    const card = document.getElementById('bankDownloadCard');
    card.style.display = 'block';
    document.getElementById('bankFilename').textContent = data.filename;

    // Show suggestions if any unmatched employees
    bankShowSuggestions(data.suggestions || []);

  } catch(e) {
    bankSetStatus(`❌ ${e.message}`);
    btn.classList.remove('running');
    btn.textContent = 'Generate Report';
    btn.disabled = false;
  }
  bankRunning = false;
}

function bankShowSuggestions(suggestions) {
  const card = document.getElementById('bankSuggestionsCard');
  if (!suggestions.length) { card.style.display = 'none'; return; }

  const groupOpts = bankGroups.map(g => `<option value="${g}">${g}</option>`).join('');

  card.style.display = 'block';
  card.innerHTML = `
    <div style="background:var(--card);border:1px solid #fde68a;border-radius:16px;overflow:hidden">
      <div style="display:flex;align-items:center;gap:10px;padding:16px 18px;background:#fffbeb;border-bottom:1px solid #fde68a">
        <span style="font-size:20px">⚠️</span>
        <div>
          <div style="font-size:13px;font-weight:700;color:#92400e">New Employees Detected</div>
          <div style="font-size:11px;color:#a16207;margin-top:2px">${suggestions.length} name${suggestions.length > 1 ? 's' : ''} found in the statement that aren't in your employee map yet.</div>
        </div>
      </div>
      <div style="padding:14px 18px;display:flex;flex-direction:column;gap:10px">
        ${suggestions.map((s, i) => `
          <div style="background:#fffbeb;border:1px solid #fef3c7;border-radius:10px;padding:12px 14px">
            <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
              <span style="font-size:11px;font-weight:700;padding:2px 8px;border-radius:99px;background:#fde68a;color:#78350f">${s.source}</span>
              <span style="font-size:12px;font-weight:600;color:var(--text)">"${s.raw_name}"</span>
            </div>
            <div style="display:flex;flex-direction:column;gap:5px">
              <div style="font-size:10px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.4px">Suggested entry — edit below</div>
              <div style="display:flex;gap:6px;flex-wrap:wrap">
                <input id="bsug-key-${i}" value="${s.raw_name.toLowerCase()}"
                       placeholder="Map key (lowercase name)"
                       style="flex:2;min-width:120px;font-size:11px;border:1px solid var(--border);border-radius:6px;padding:5px 8px;color:var(--text)" />
                <input id="bsug-label-${i}" value="${s.suggested_label}"
                       placeholder="Label (e.g. PC PAYROLL - Name)"
                       style="flex:3;min-width:140px;font-size:11px;border:1px solid var(--border);border-radius:6px;padding:5px 8px;color:var(--text)" />
                <select id="bsug-group-${i}"
                        style="flex:2;min-width:120px;font-size:11px;border:1px solid var(--border);border-radius:6px;padding:5px 8px;color:var(--text);background:#fff">
                  <option value="">Select group…</option>
                  ${bankGroups.map(g => `<option value="${g}" ${g === s.suggested_group ? 'selected' : ''}>${g}</option>`).join('')}
                </select>
              </div>
              <button onclick="bankAddSuggestion(${i})"
                      style="margin-top:4px;align-self:flex-start;background:var(--navy);color:#fff;border:none;border-radius:6px;padding:5px 14px;font-size:11px;font-weight:700;cursor:pointer">
                + Add to Employee Map
              </button>
            </div>
          </div>`).join('')}
      </div>
      <div style="padding:10px 18px;border-top:1px solid #fde68a;font-size:11px;color:#a16207;background:#fffbeb">
        💡 After adding, click <strong>Save Changes</strong> in the employee map on the right, then re-run the report.
      </div>
    </div>`;
}

function bankAddSuggestion(i) {
  const key   = document.getElementById(`bsug-key-${i}`)?.value.trim().toLowerCase();
  const label = document.getElementById(`bsug-label-${i}`)?.value.trim();
  const group = document.getElementById(`bsug-group-${i}`)?.value;
  if (!key || !label || !group) {
    showAlert('Missing Info', 'Please fill in all three fields (key, label, and group) before adding.');
    return;
  }
  // Check for duplicate
  if (bankEmployees.some(e => e.key === key)) {
    showAlert('Already Exists', `An entry for "${key}" already exists in the employee map.`);
    return;
  }
  bankEmployees.push({ key, label, group });
  bankRenderList();
  bankAutoSave();
  // Visually mark the suggestion row as added
  const btn = document.querySelector(`#bankSuggestionsCard button[onclick="bankAddSuggestion(${i})"]`);
  if (btn) { btn.textContent = '✓ Added'; btn.disabled = true; btn.style.background = '#16a34a'; }
}

function bankDownload() {
  if (!bankResultData?.xlsx_b64) return;
  const bytes = atob(bankResultData.xlsx_b64);
  const arr   = new Uint8Array(bytes.length);
  for (let i = 0; i < bytes.length; i++) arr[i] = bytes.charCodeAt(i);
  const blob = new Blob([arr], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = bankResultData.filename;
  a.click();
  URL.revokeObjectURL(url);
}

function bankResetFiles() {
  bankBankFile = bankWiseFile = bankWuFile = bankResultData = null;
  renderBankCategorizerView();
}

// ── Employee map management ────────────────────────────────────────────────────

async function bankLoadEmployees() {
  try {
    const res  = await authFetch(`${getApiBase()}/api/bank_categorizer/employees`);
    const data = await res.json();
    bankEmployees = data.employees || [];
    bankGroups    = data.groups    || [];
    bankRenderList();
  } catch(e) {
    const el = document.getElementById('bankEmpList');
    if (el) el.innerHTML = `<div style="padding:16px;color:var(--red);font-size:12px">Failed to load employee map.</div>`;
  }
}

function bankRenderList() {
  const query = (document.getElementById('bankSearch')?.value || '').toLowerCase();
  const list  = document.getElementById('bankEmpList');
  if (!list) return;

  const filtered = bankEmployees.filter(e =>
    e.key.includes(query) || e.label.toLowerCase().includes(query) || e.group.toLowerCase().includes(query)
  );

  if (!filtered.length) {
    list.innerHTML = `<div style="padding:16px;text-align:center;color:var(--text-3);font-size:12px">${query ? 'No matches.' : 'No employees yet.'}</div>`;
    return;
  }

  // Group by group for display
  const byGroup = {};
  for (const emp of filtered) {
    (byGroup[emp.group] = byGroup[emp.group] || []).push(emp);
  }

  list.innerHTML = Object.entries(byGroup).map(([grp, emps]) => `
    <div style="padding:6px 14px 2px;font-size:10px;font-weight:700;color:var(--text-3);text-transform:uppercase;letter-spacing:.5px;background:#f8fafc;border-bottom:1px solid var(--border)">${grp}</div>
    ${emps.map(emp => {
      const realIdx = bankEmployees.indexOf(emp);
      return `
        <div style="display:flex;align-items:flex-start;gap:6px;padding:8px 14px;border-bottom:1px solid var(--border)">
          <div style="flex:1;min-width:0">
            <div style="font-size:11px;font-weight:600;color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${emp.key}">${emp.key}</div>
            <input value="${emp.label}" onchange="bankEmployees[${realIdx}].label=this.value;bankAutoSave()"
                   style="margin-top:3px;width:100%;font-size:11px;border:1px solid var(--border);border-radius:5px;padding:3px 6px;color:var(--text-2);background:#fff"/>
            <select onchange="bankEmployees[${realIdx}].group=this.value;bankAutoSave()"
                    style="margin-top:3px;width:100%;font-size:10px;border:1px solid var(--border);border-radius:5px;padding:3px 6px;color:var(--text-2);background:#fff">
              ${bankGroups.map(g => `<option value="${g}" ${g===emp.group?'selected':''}>${g}</option>`).join('')}
            </select>
          </div>
          <button onclick="bankRemoveEmployee(${realIdx})"
                  style="font-size:11px;color:var(--red);background:none;border:none;cursor:pointer;padding:2px 4px;margin-top:2px" title="Remove">✕</button>
        </div>`;
    }).join('')}
  `).join('');
}

function bankAddEmployee() {
  showFormModal('🏦', 'Add Employee to Map', [
    { id: 'bKey',   label: 'Name Key (lowercase — as it appears in bank/Wise/WU)', placeholder: 'e.g. john doe',             required: true },
    { id: 'bLabel', label: 'Label (shown in Excel report)',                          placeholder: 'e.g. PC PAYROLL - John Doe', required: true },
    { id: 'bGroup', label: 'Group', type: 'select', options: bankGroups,            required: true },
  ], 'Add Employee', ({ bKey, bLabel, bGroup }) => {
    const key = bKey.trim().toLowerCase();
    if (bankEmployees.some(e => e.key === key)) {
      showAlert('Already Exists', `"${key}" already exists in the employee map.`);
      return;
    }
    bankEmployees.push({ key, label: bLabel.trim(), group: bGroup });
    bankRenderList();
    bankAutoSave();
  });
}

function bankRemoveEmployee(idx) {
  bankEmployees.splice(idx, 1);
  bankRenderList();
  bankAutoSave();
}

async function bankSaveEmployees() {
  const status = document.getElementById('bankSaveStatus');
  if (status) { status.textContent = 'Saving…'; status.style.color = 'var(--text-3)'; }
  try {
    const res  = await authFetch(`${getApiBase()}/api/bank_categorizer/employees`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ employees: bankEmployees }),
    });
    const data = await res.json();
    if (!res.ok) {
      if (status) { status.textContent = '✗ Save failed'; status.style.color = 'var(--red)'; }
      return;
    }
    if (status) {
      status.textContent = '✓ Saved';
      status.style.color = 'var(--green)';
      setTimeout(() => { if (status) status.textContent = ''; }, 2500);
    }
  } catch(e) {
    if (status) { status.textContent = '✗ Error'; status.style.color = 'var(--red)'; }
  }
}

function bankAutoSave() {
  if (bankSaveTimer) clearTimeout(bankSaveTimer);
  const status = document.getElementById('bankSaveStatus');
  if (status) { status.textContent = '…'; status.style.color = 'var(--text-3)'; }
  bankSaveTimer = setTimeout(() => bankSaveEmployees(), 800);
}
