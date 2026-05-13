'use strict';
const D = window.__DASH__;
const KEY = window.__KEY__;
const rendered = new Set();

/* ── Utilities ─────────────────────────────────────────────────────────── */
function esc(str) {
  if (str == null) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function fmt(val, dec = 1) {
  const n = parseFloat(val);
  return isNaN(n) ? '—' : n.toFixed(dec);
}
function fmtDate(iso) {
  if (!iso) return '—';
  try { return String(iso).slice(0, 10); } catch { return String(iso); }
}
function fmtCost(val) {
  const n = parseFloat(val);
  if (isNaN(n)) return '—';
  if (n >= 1e9) return '$' + (n/1e9).toFixed(1) + 'B';
  if (n >= 1e6) return '$' + (n/1e6).toFixed(1) + 'M';
  if (n >= 1e3) return '$' + (n/1e3).toFixed(1) + 'K';
  return '$' + n.toFixed(0);
}
function badge(cls, text) {
  return `<span class="badge badge-${esc(cls)}">${esc(text)}</span>`;
}
function changeBadge(t) {
  const m = { added:'green', deleted:'red', changed:'amber', unchanged:'gray' };
  return badge(m[t] || 'gray', t || '—');
}
function statusBadge(s) {
  const m = { TK_Complete:['green','Complete'], TK_Active:['blue','Active'], TK_NotStart:['gray','Not Started'] };
  const [cls, lbl] = m[s] || ['gray', s || '—'];
  return badge(cls, lbl);
}
function varCell(days) {
  if (!days && days !== 0) return '—';
  const d = parseFloat(days);
  if (d === 0) return '0d';
  const color = d > 0 ? 'var(--red)' : 'var(--green)';
  return `<span style="color:${color};font-weight:500">${d > 0 ? '+' : ''}${d.toFixed(1)}d</span>`;
}
function kpiCard(label, value, sub, mod = '') {
  return `<div class="kpi-card ${esc(mod)}">
    <div class="kpi-val">${esc(String(value ?? '—'))}</div>
    <div class="kpi-lbl">${esc(label)}</div>
    ${sub ? `<div class="kpi-sub">${esc(String(sub))}</div>` : ''}
  </div>`;
}

/* ── Sidebar ────────────────────────────────────────────────────────────── */
function initSidebar() {
  const sidebar = document.getElementById('sidebar');
  const btn = document.getElementById('collapse-btn');
  btn.addEventListener('click', () => {
    sidebar.classList.toggle('collapsed');
    btn.textContent = sidebar.classList.contains('collapsed') ? '›' : '‹';
    btn.setAttribute('aria-label', sidebar.classList.contains('collapsed') ? 'Expand sidebar' : 'Collapse sidebar');
  });
  document.querySelectorAll('.nav-item').forEach(item => {
    item.addEventListener('click', () => navigate(item.dataset.page));
    item.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); navigate(item.dataset.page); }
    });
  });
}

/* ── Navigation ─────────────────────────────────────────────────────────── */
const PAGE_TITLES = {
  overview: 'Overview', schedule: 'Schedule', milestones: 'Milestones',
  lookahead: 'Lookahead', ev: 'Earned Value', logic: 'Logic & Relationships'
};
const RENDERERS = {};

function navigate(page) {
  document.querySelectorAll('.nav-item').forEach(i => {
    const active = i.dataset.page === page;
    i.classList.toggle('active', active);
    if (active) i.setAttribute('aria-current', 'page');
    else i.removeAttribute('aria-current');
  });
  document.querySelectorAll('.page').forEach(p => {
    p.classList.remove('active');
    p.hidden = true;
  });
  const target = document.getElementById('page-' + page);
  if (target) { target.classList.add('active'); target.hidden = false; }
  const heading = document.getElementById('page-heading');
  if (heading) heading.textContent = PAGE_TITLES[page] || page;

  if (!rendered.has(page) && RENDERERS[page]) {
    rendered.add(page);
    RENDERERS[page]();
  }
}

/* ── Topbar ─────────────────────────────────────────────────────────────── */
function initTopbar() {
  const ai = D.ai_summary || {};
  const health = ai.schedule_health || 'unknown';
  const badge_el = document.getElementById('health-badge');
  if (!badge_el) return;
  const labels = { on_track: 'On Track', at_risk: 'At Risk', critical: 'Critical', unknown: '—' };
  const classes = { on_track: 'badge-green', at_risk: 'badge-amber', critical: 'badge-red', unknown: 'badge-gray' };
  badge_el.textContent = labels[health] || String(health);
  badge_el.className = 'health-badge badge ' + (classes[health] || 'badge-gray');
}

/* ── G2 chart helpers ───────────────────────────────────────────────────── */
function g2Donut(id, data, angleField, colorField, height = 260) {
  const el = document.getElementById(id);
  if (!el || !data.length || typeof G2 === 'undefined') return;
  el.style.height = height + 'px';
  const chart = new G2.Chart({ container: el, autoFit: true, height });
  chart.options({
    type: 'interval',
    coordinate: { type: 'theta', innerRadius: 0.55 },
    data,
    encode: { y: angleField, color: colorField },
    style: { stroke: '#fff', lineWidth: 2 },
    legend: { color: { position: 'right', rowPadding: 4 } },
    tooltip: { items: [{ channel: 'y', name: colorField }] },
  });
  chart.render();
}
function g2HBar(id, data, xField, yField, height = 280, fillColor = '#059669') {
  const el = document.getElementById(id);
  if (!el || !data.length || typeof G2 === 'undefined') return;
  el.style.height = height + 'px';
  const chart = new G2.Chart({ container: el, autoFit: true, height });
  chart.options({
    type: 'interval',
    coordinate: { transform: [{ type: 'transpose' }] },
    data,
    encode: { x: xField, y: yField },
    style: { fill: fillColor, radius: [0, 3, 3, 0] },
    axis: { x: { title: false, labelFormatter: d => String(d).slice(0, 20) }, y: { title: 'Days' } },
    sort: { by: yField, reverse: true },
  });
  chart.render();
}
function g2Line(id, data, xField, yField, colorField, height = 280) {
  const el = document.getElementById(id);
  if (!el || !data.length || typeof G2 === 'undefined') return;
  el.style.height = height + 'px';
  const chart = new G2.Chart({ container: el, autoFit: true, height });
  chart.options({
    type: 'line',
    data,
    encode: { x: xField, y: yField, color: colorField },
    style: { strokeWidth: 2 },
    axis: { y: { title: '% Complete', tickFormatter: d => d + '%' }, x: { title: false } },
    legend: { color: { position: 'top-right' } },
    tooltip: { items: [{ channel: 'y', name: yField, valueFormatter: d => d?.toFixed(1) + '%' }] },
  });
  chart.render();
}

/* ── Overview ────────────────────────────────────────────────────────────── */
RENDERERS.overview = function renderOverview() {
  const el = document.getElementById('page-overview');
  const s = D.summary || {};
  const ai = D.ai_summary || {};

  const maxDelay = s.max_delay_days || 0;
  const kpis = `<div class="kpi-row">
    ${kpiCard('Activities Changed', s.changed ?? '—', `of ${s.total_updated ?? '—'} total`, '--amber')}
    ${kpiCard('Max Delay', maxDelay + 'd', `avg ${fmt(s.avg_finish_variance_days)}d slip`, maxDelay > 30 ? '--red' : '--amber')}
    ${kpiCard('Critical Activities', s.critical_activities_updated ?? '—', 'zero or negative float', '--red')}
    ${kpiCard('Added / Deleted', `${s.added ?? 0} / ${s.deleted ?? 0}`, 'new / removed activities', '--blue')}
  </div>`;

  const charts = `<div class="grid-2">
    <div class="card">
      <div class="card-hd">Activity Change Distribution</div>
      <div class="card-bd"><div id="chart-donut"></div></div>
    </div>
    <div class="card">
      <div class="card-hd">Top 10 Delayed Activities (days)</div>
      <div class="card-bd"><div id="chart-delays"></div></div>
    </div>
  </div>`;

  const scurveHtml = (D.scurve?.baseline?.length || D.scurve?.updated?.length) ? `
  <div class="card grid-1">
    <div class="card-hd">S-Curve — Cumulative Progress</div>
    <div class="card-bd"><div id="chart-scurve"></div></div>
  </div>` : '';

  const proc = D.procurement || {};
  const procHtml = (proc.items && proc.items.length) ? `
  <div class="card grid-1">
    <div class="card-hd">Procurement Status</div>
    <div class="tbl-wrap">
      <table class="tbl"><thead><tr>
        <th>ID</th><th>Activity</th><th>Status</th><th>Lead Time</th>
      </tr></thead><tbody>${
        proc.items.slice(0, 12).map(p => `<tr>
          <td><code>${esc(p.task_code || '')}</code></td>
          <td>${esc(p.task_name || '')}</td>
          <td>${esc(p.status || '—')}</td>
          <td>${p.lead_time_days != null ? p.lead_time_days + 'd' : '—'}</td>
        </tr>`).join('')
      }</tbody></table>
    </div>
  </div>` : '';

  const aiHtml = ai.executive_summary ? `
  <div class="card grid-1">
    <div class="card-hd">AI Schedule Analysis</div>
    <div class="card-bd">
      <p class="ai-summary-text">${esc(ai.executive_summary)}</p>
      ${(ai.key_risks || []).length ? `<h4 class="ai-section-title">Key Risks</h4><ul class="ai-list">${(ai.key_risks||[]).map(r=>`<li>${esc(r)}</li>`).join('')}</ul>` : ''}
      ${(ai.recommended_actions || []).length ? `<h4 class="ai-section-title">Recommended Actions</h4><ul class="ai-list">${(ai.recommended_actions||[]).map(a=>`<li>${esc(a)}</li>`).join('')}</ul>` : ''}
    </div>
  </div>` : '';

  // Top changed/delayed activities table — always rendered so tests can find table rows
  const topActivities = (D.activity_variances || [])
    .filter(v => v.change_type !== 'unchanged')
    .sort((a, b) => Math.abs(b.finish_variance_days) - Math.abs(a.finish_variance_days))
    .slice(0, 20);
  const topRows = topActivities.map(v => `<tr>
    <td><code>${esc(v.task_code)}</code></td>
    <td>${esc(v.task_name)}</td>
    <td>${varCell(v.finish_variance_days)}</td>
    <td>${changeBadge(v.change_type)}</td>
  </tr>`).join('');
  const topTableHtml = `<div class="card grid-1">
    <div class="card-hd">Top Changed Activities</div>
    <div class="tbl-wrap"><table class="tbl" id="overview-activity-table">
      <thead><tr><th>ID</th><th>Activity Name</th><th>Finish Variance</th><th>Change</th></tr></thead>
      <tbody>${topRows || '<tr><td colspan="4" style="text-align:center;color:var(--muted);padding:1rem">No changed activities</td></tr>'}</tbody>
    </table></div>
  </div>`;

  el.innerHTML = kpis + charts + scurveHtml + topTableHtml + procHtml + aiHtml;

  setTimeout(() => {
    const changeData = [
      { type: 'Changed', count: s.changed || 0 },
      { type: 'Unchanged', count: s.unchanged || 0 },
      { type: 'Added', count: s.added || 0 },
      { type: 'Deleted', count: s.deleted || 0 },
    ].filter(d => d.count > 0);
    g2Donut('chart-donut', changeData, 'count', 'type', 260);

    const topDelayed = (s.top_delayed || []).slice(0, 10).map(a => ({
      name: (a.task_code + ' ' + a.task_name).slice(0, 35),
      days: a.finish_variance_days
    }));
    g2HBar('chart-delays', topDelayed, 'name', 'days', 280, '#d97706');

    if (scurveHtml) {
      const scData = [
        ...(D.scurve.baseline || []).map(r => ({ ...r, series: 'Baseline' })),
        ...(D.scurve.updated  || []).map(r => ({ ...r, series: 'Updated'  })),
      ];
      g2Line('chart-scurve', scData, 'period', 'cumulative_pct', 'series', 300);
    }
  }, 0);
};

/* ── Schedule ────────────────────────────────────────────────────────────── */
const PAGE_SIZE = 200;
let schedFiltered = [];
let schedPage = 0;

RENDERERS.schedule = function renderSchedule() {
  const el = document.getElementById('page-schedule');
  el.innerHTML = `
    <div class="toolbar">
      <input type="search" class="search-box" id="sched-search" placeholder="Search ID or activity name…" aria-label="Search activities">
      <select class="filter-sel" id="sched-status" aria-label="Filter by status">
        <option value="">All statuses</option>
        <option value="TK_NotStart">Not Started</option>
        <option value="TK_Active">Active</option>
        <option value="TK_Complete">Complete</option>
      </select>
      <select class="filter-sel" id="sched-change" aria-label="Filter by change type">
        <option value="">All changes</option>
        <option value="changed">Changed</option>
        <option value="added">Added</option>
        <option value="deleted">Deleted</option>
        <option value="unchanged">Unchanged</option>
      </select>
      <span class="count-chip" id="sched-count" aria-live="polite"></span>
    </div>
    <div class="card">
      <div class="tbl-wrap">
        <table class="tbl" role="grid">
          <thead><tr>
            <th>ID</th><th>Activity Name</th><th>WBS</th>
            <th>BL Finish</th><th>Updated Finish</th>
            <th>Start Var</th><th>Finish Var</th><th>Dur Var</th>
            <th>Change</th>
          </tr></thead>
          <tbody id="sched-tbody" aria-live="polite"></tbody>
        </table>
      </div>
    </div>
    <div class="load-more-bar" id="load-more-bar" style="display:none">
      <button class="btn" id="load-more-btn">Load 200 more</button>
    </div>`;

  schedFiltered = [...(D.activity_variances || [])];
  schedPage = 0;
  renderSchedRows();

  let t;
  document.getElementById('sched-search').addEventListener('input', () => { clearTimeout(t); t = setTimeout(applySchedFilter, 150); });
  document.getElementById('sched-status').addEventListener('change', applySchedFilter);
  document.getElementById('sched-change').addEventListener('change', applySchedFilter);
  document.getElementById('load-more-btn').addEventListener('click', () => { schedPage++; renderSchedRows(true); });
};

function applySchedFilter() {
  const q = (document.getElementById('sched-search').value || '').toLowerCase();
  const status = document.getElementById('sched-status').value;
  const change = document.getElementById('sched-change').value;
  schedFiltered = (D.activity_variances || []).filter(v => {
    if (q && !String(v.task_code).toLowerCase().includes(q) && !String(v.task_name).toLowerCase().includes(q)) return false;
    if (status && v.new_status !== status && v.old_status !== status) return false;
    if (change && v.change_type !== change) return false;
    return true;
  });
  schedPage = 0;
  renderSchedRows();
}

function renderSchedRows(append = false) {
  const tbody = document.getElementById('sched-tbody');
  const countEl = document.getElementById('sched-count');
  const bar = document.getElementById('load-more-bar');
  if (!tbody) return;

  const start = append ? schedPage * PAGE_SIZE : 0;
  const slice = schedFiltered.slice(start, start + PAGE_SIZE);

  const html = slice.map(v => `<tr>
    <td><code>${esc(v.task_code)}</code></td>
    <td>${esc(v.task_name)}</td>
    <td>${esc(v.wbs_name)}</td>
    <td>${esc(fmtDate(v.baseline_finish))}</td>
    <td>${esc(fmtDate(v.updated_finish))}</td>
    <td>${varCell(v.start_variance_days)}</td>
    <td>${varCell(v.finish_variance_days)}</td>
    <td>${varCell(v.duration_variance_days)}</td>
    <td>${changeBadge(v.change_type)}</td>
  </tr>`).join('');

  if (append) tbody.insertAdjacentHTML('beforeend', html);
  else { tbody.innerHTML = html; if (countEl) countEl.textContent = schedFiltered.length + ' activities'; }

  if (bar) bar.style.display = (schedPage + 1) * PAGE_SIZE < schedFiltered.length ? 'flex' : 'none';
}

/* ── Milestones ──────────────────────────────────────────────────────────── */
RENDERERS.milestones = function renderMilestones() {
  const el = document.getElementById('page-milestones');
  const miles = D.milestones || [];
  if (!miles.length) {
    el.innerHTML = `<div class="card"><div class="card-bd empty-state">No milestone data found in this schedule.</div></div>`;
    return;
  }
  const rows = miles.map(m => {
    const v = parseFloat(m.variance_days) || 0;
    const cls = v > 14 ? 'badge-red' : v > 0 ? 'badge-amber' : 'badge-green';
    const lbl = v > 14 ? 'At Risk' : v > 0 ? 'Slipped' : 'On Track';
    return `<tr>
      <td><code>${esc(m.task_code || '')}</code></td>
      <td>${esc(m.task_name || '')}</td>
      <td>${esc(fmtDate(m.baseline_finish))}</td>
      <td>${esc(fmtDate(m.updated_finish))}</td>
      <td>${varCell(v)}</td>
      <td><span class="badge ${esc(cls)}">${esc(lbl)}</span></td>
    </tr>`;
  }).join('');
  el.innerHTML = `<div class="card">
    <div class="card-hd">Milestones (${miles.length})</div>
    <div class="tbl-wrap"><table class="tbl">
      <thead><tr><th>ID</th><th>Name</th><th>BL Finish</th><th>Upd Finish</th><th>Variance</th><th>Status</th></tr></thead>
      <tbody>${rows}</tbody>
    </table></div>
  </div>`;
};

/* ── Lookahead ────────────────────────────────────────────────────────────── */
let laWindow = 30;
RENDERERS.lookahead = function renderLookahead() {
  const el = document.getElementById('page-lookahead');
  el.innerHTML = `
    <div class="toolbar">
      <span style="font-size:.875rem;font-weight:600;color:var(--muted)">Window:</span>
      ${[30,60,90].map(d => `<button class="btn lookahead-chip${d===30?' active':''}" data-days="${d}">${d} Days</button>`).join('')}
      <span class="count-chip" id="la-count" aria-live="polite"></span>
    </div>
    <div class="card">
      <div class="tbl-wrap"><table class="tbl">
        <thead><tr><th>ID</th><th>Activity Name</th><th>WBS</th><th>Start</th><th>Finish</th><th>Duration</th><th>Status</th></tr></thead>
        <tbody id="la-tbody" aria-live="polite"></tbody>
      </table></div>
    </div>`;
  document.querySelectorAll('.lookahead-chip').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.lookahead-chip').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      laWindow = parseInt(btn.dataset.days);
      renderLARows();
    });
  });
  renderLARows();
};
function renderLARows() {
  const all = D.lookahead || [];
  const filtered = all.filter(a => (a.days_to_start == null || a.days_to_start <= laWindow));
  const countEl = document.getElementById('la-count');
  if (countEl) countEl.textContent = filtered.length + ' activities';
  const tbody = document.getElementById('la-tbody');
  if (!tbody) return;
  tbody.innerHTML = filtered.length ? filtered.map(a => `<tr>
    <td><code>${esc(a.task_code || '')}</code></td>
    <td>${esc(a.task_name || '')}</td>
    <td>${esc(a.wbs_name || '')}</td>
    <td>${esc(fmtDate(a.planned_start))}</td>
    <td>${esc(fmtDate(a.planned_finish))}</td>
    <td>${a.duration_days != null ? esc(a.duration_days + 'd') : '—'}</td>
    <td>${statusBadge(a.status_code || '')}</td>
  </tr>`).join('') : `<tr><td colspan="7" style="text-align:center;color:var(--muted);padding:2rem">No activities starting in next ${laWindow} days</td></tr>`;
}

/* ── Earned Value ─────────────────────────────────────────────────────────── */
RENDERERS.ev = function renderEV() {
  const el = document.getElementById('page-ev');
  const ev = D.ev || {};
  if (!ev.has_cost_data) {
    el.innerHTML = `<div class="card"><div class="card-bd empty-state">
      No cost data found. EV metrics require resource cost assignments in the XER file.
    </div></div>`;
    return;
  }
  const u = ev.updated || {};
  const cpi = parseFloat(u.CPI);
  const spi = parseFloat(u.SPI);
  const kpis = `<div class="kpi-row">
    ${kpiCard('CPI', fmt(u.CPI, 2), 'Cost Performance Index', cpi >= 1 ? '--green' : '--red')}
    ${kpiCard('SPI', fmt(u.SPI, 2), 'Schedule Performance Index', spi >= 1 ? '--green' : '--red')}
    ${kpiCard('BAC', fmtCost(u.BAC), 'Budget at Completion', '--blue')}
    ${kpiCard('EAC', fmtCost(u.EAC), 'Estimate at Completion', parseFloat(u.EAC) <= parseFloat(u.BAC) ? '--green' : '--red')}
  </div>`;
  const evKpis = (D.kpis || []).filter(k => ['CPI','SPI','BCWP','ACWP','PV','EV','AC','CV','SV'].includes(k.name));
  const rows = evKpis.map(k => `<tr>
    <td>${esc(k.label || k.name)}</td>
    <td><strong>${esc(k.value != null ? String(k.value) : '—')}</strong></td>
    <td>${esc(k.benchmark || '—')}</td>
    <td>${k.status ? badge(k.status === 'good' ? 'green' : k.status === 'warning' ? 'amber' : 'red', k.status) : '—'}</td>
  </tr>`).join('');
  el.innerHTML = kpis + `<div class="card">
    <div class="card-hd">EV Metrics Detail</div>
    <div class="tbl-wrap"><table class="tbl">
      <thead><tr><th>Metric</th><th>Value</th><th>Benchmark</th><th>Status</th></tr></thead>
      <tbody>${rows || '<tr><td colspan="4" style="text-align:center;color:var(--muted);padding:2rem">No EV metrics available</td></tr>'}</tbody>
    </table></div>
  </div>`;
};

/* ── Logic / Relationships ────────────────────────────────────────────────── */
RENDERERS.logic = function renderLogic() {
  const el = document.getElementById('page-logic');
  const rels = D.relationship_variances || [];
  if (!rels.length) {
    el.innerHTML = `<div class="card"><div class="card-bd empty-state">No relationship changes detected between the two schedules.</div></div>`;
    return;
  }
  const added   = rels.filter(r => r.change_type === 'added').length;
  const deleted  = rels.filter(r => r.change_type === 'deleted').length;
  const changed  = rels.filter(r => r.change_type === 'changed').length;
  const rows = rels.map(r => `<tr>
    <td><code>${esc(r.pred_code)}</code></td>
    <td><code>${esc(r.succ_code)}</code></td>
    <td>${esc(r.old_pred_type || '—')}</td>
    <td>${esc(r.new_pred_type || '—')}</td>
    <td>${r.lag_change_hours ? varCell(r.lag_change_hours / 8) : '—'}</td>
    <td>${changeBadge(r.change_type)}</td>
  </tr>`).join('');
  el.innerHTML = `<div class="kpi-row" style="grid-template-columns:repeat(3,1fr)">
    ${kpiCard('Added', added, 'new relationships', '--green')}
    ${kpiCard('Deleted', deleted, 'removed relationships', '--red')}
    ${kpiCard('Changed', changed, 'modified relationships', '--amber')}
  </div>
  <div class="card">
    <div class="card-hd">Relationship Changes (${rels.length})</div>
    <div class="tbl-wrap"><table class="tbl">
      <thead><tr><th>Predecessor</th><th>Successor</th><th>Old Type</th><th>New Type</th><th>Lag Change</th><th>Change</th></tr></thead>
      <tbody>${rows}</tbody>
    </table></div>
  </div>`;
};

/* ── Chat FAB ─────────────────────────────────────────────────────────────── */
function initChat() {
  const fab    = document.getElementById('chat-fab');
  const modal  = document.getElementById('chat-modal');
  const close  = document.getElementById('chat-close');
  const input  = document.getElementById('chat-input');
  const send   = document.getElementById('chat-send');
  const msgs   = document.getElementById('chat-messages');
  if (!fab || !modal) return;

  function openChat() { modal.classList.add('open'); input.focus(); }
  function closeChat() { modal.classList.remove('open'); }

  fab.addEventListener('click', () => modal.classList.contains('open') ? closeChat() : openChat());
  fab.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fab.click(); } });
  close.addEventListener('click', closeChat);
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && modal.classList.contains('open')) closeChat(); });

  function appendMsg(role, text) {
    const div = document.createElement('div');
    div.className = `chat-msg chat-msg-${role}`;
    div.textContent = text;
    msgs.appendChild(div);
    msgs.scrollTop = msgs.scrollHeight;
    return div;
  }

  function sendMsg() {
    const msg = input.value.trim();
    if (!msg) return;
    appendMsg('user', msg);
    input.value = '';
    send.disabled = true;
    const thinking = appendMsg('assistant', '…');

    fetch('/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ key: KEY, message: msg })
    }).then(r => {
      const reader = r.body.getReader();
      const dec = new TextDecoder();
      let buf = '';
      thinking.textContent = '';

      function pump() {
        reader.read().then(({ done, value }) => {
          if (done) { send.disabled = false; return; }
          buf += dec.decode(value);
          const lines = buf.split('\n');
          buf = lines.pop();
          lines.forEach(line => {
            if (!line.startsWith('data: ')) return;
            const raw = line.slice(6).trim();
            if (raw === '[DONE]') { send.disabled = false; return; }
            try {
              const obj = JSON.parse(raw);
              if (obj.error) { thinking.textContent = '⚠ ' + obj.error; return; }
              thinking.textContent += obj.content || obj.text || '';
              msgs.scrollTop = msgs.scrollHeight;
            } catch {}
          });
          pump();
        }).catch(e => { thinking.textContent = '⚠ ' + e.message; send.disabled = false; });
      }
      pump();
    }).catch(e => { thinking.textContent = '⚠ Network error'; send.disabled = false; });
  }

  send.addEventListener('click', sendMsg);
  input.addEventListener('keydown', e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); sendMsg(); } });
}

/* ── Init ─────────────────────────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => {
  initSidebar();
  initTopbar();
  initChat();
  navigate('overview');
});
