/* dashboard.js — Schedule Comparison Dashboard */
'use strict';

const D = window.__DASH__;
const KEY = window.__KEY__;
const rendered = new Set();

/* ── Utility helpers ─────────────────────────────────────────────────────── */

function esc(str) {
  if (str == null) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function fmt(val, decimals = 1) {
  if (val == null || val === '') return '—';
  const n = parseFloat(val);
  if (isNaN(n)) return '—';
  return n.toFixed(decimals);
}

function fmtDate(iso) {
  if (!iso) return '—';
  try { return iso.slice(0, 10); } catch (e) { return iso; }
}

function plural(n, word) {
  return `${n} ${word}${n === 1 ? '' : 's'}`;
}

/* ── AntV G2 v5 chart helpers ────────────────────────────────────────────── */

function renderG2Bar(containerId, data, xField, yField, colorField, height = 280) {
  const container = document.getElementById(containerId);
  if (!container || !data.length) return;
  container.style.height = height + 'px';
  const chart = new G2.Chart({ container, autoFit: true, height });
  chart.options({
    type: 'interval',
    data,
    encode: { x: xField, y: yField, ...(colorField ? { color: colorField } : {}) },
    style: { radius: [4, 4, 0, 0] },
    axis: { y: { title: false }, x: { title: false } },
  });
  chart.render();
}

function renderG2HBar(containerId, data, xField, yField, height = 320) {
  const container = document.getElementById(containerId);
  if (!container || !data.length) return;
  container.style.height = height + 'px';
  const chart = new G2.Chart({ container, autoFit: true, height });
  chart.options({
    type: 'interval',
    coordinate: { transform: [{ type: 'transpose' }] },
    data,
    encode: { x: xField, y: yField },
    style: { radius: [0, 4, 4, 0] },
    axis: { y: { title: false }, x: { title: false } },
  });
  chart.render();
}

function renderG2Line(containerId, data, xField, yField, colorField, height = 280) {
  const container = document.getElementById(containerId);
  if (!container || !data.length) return;
  container.style.height = height + 'px';
  const chart = new G2.Chart({ container, autoFit: true, height });
  chart.options({
    type: 'line',
    data,
    encode: { x: xField, y: yField, color: colorField },
    style: { strokeWidth: 2 },
    axis: { y: { title: false }, x: { title: false } },
  });
  chart.render();
}

/* ── Tab routing ─────────────────────────────────────────────────────────── */

function activateTab(name) {
  // Update buttons with ARIA
  document.querySelectorAll('.tab-btn').forEach(btn => {
    const isActive = btn.dataset.tab === name;
    btn.classList.toggle('active', isActive);
    btn.setAttribute('aria-selected', isActive ? 'true' : 'false');
  });
  // Show/hide panels with ARIA
  document.querySelectorAll('.tab-panel').forEach(panel => {
    panel.classList.remove('active');
    panel.setAttribute('aria-hidden', 'true');
  });
  const panel = document.getElementById('panel-' + name);
  if (!panel) return;
  panel.classList.add('active');
  panel.setAttribute('aria-hidden', 'false');

  // Render if not yet done
  if (!rendered.has(name)) {
    rendered.add(name);
    const renderers = {
      overview:    renderOverview,
      schedule:    renderSchedule,
      kpis:        renderKPIs,
      lookahead:   renderLookahead,
      milestones:  renderMilestones,
      procurement: renderProcurement,
      ev:          renderEV,
      wbs:         renderWBS,
      logic:       renderLogic,
      chat:        renderChat,
    };
    if (renderers[name]) renderers[name](panel);
  }
}

/* ── Tab 1: Overview ─────────────────────────────────────────────────────── */

function renderOverview(el) {
  const S = D.summary || {};
  const ai = D.ai_summary || {};
  const health = ai.schedule_health || 'unknown';

  const delayed5 = (D.activity_variances || []).filter(a => (a.finish_variance_days || 0) > 5).length;
  const maxDelay = S.max_delay_days != null ? fmt(S.max_delay_days, 0) + 'd' : '—';
  const avgVar   = S.avg_finish_variance_days != null ? fmt(S.avg_finish_variance_days, 1) + 'd' : '—';
  const critical = D.kpis ? (D.kpis.critical_total || 0) : 0;

  // Build top-10 delayed table rows
  const delayed = (D.activity_variances || [])
    .filter(a => (a.finish_variance_days || 0) > 0)
    .sort((a, b) => (b.finish_variance_days || 0) - (a.finish_variance_days || 0))
    .slice(0, 10);

  let delayedRows = '';
  if (delayed.length === 0) {
    delayedRows = `<tr data-testid="activity-row"><td colspan="4" class="no-data">No delayed activities</td></tr>`;
  } else {
    delayedRows = delayed.map(a => `
      <tr data-testid="activity-row" class="${a.is_critical ? 'row-critical' : ''}">
        <td>${esc(a.task_code)}</td>
        <td>${esc(a.task_name)}</td>
        <td>${esc(a.wbs_name)}</td>
        <td class="num">${fmt(a.finish_variance_days, 0)}d</td>
      </tr>`).join('');
  }

  // AI risks
  const risks = (ai.key_risks || []);
  const riskHtml = risks.length
    ? `<ul class="ai-risks">${risks.map(r => `<li class="risk-item">${esc(r)}</li>`).join('')}</ul>`
    : '';

  // AI actions
  const actions = (ai.recommended_actions || []);
  const actionsHtml = actions.length
    ? `<div class="ai-actions"><strong>Recommended Actions</strong><ul>${actions.map(a => `<li>${esc(a)}</li>`).join('')}</ul></div>`
    : '';

  el.innerHTML = `
    <div class="summary-bar">
      <div class="summary-group">
        <span class="summary-group-label">Changes</span>
        <div class="summary-metrics">
          <div class="summary-metric"><span class="summary-val added">${S.added != null ? S.added : 0}</span><span class="summary-key">Added</span></div>
          <div class="summary-metric"><span class="summary-val deleted">${S.deleted != null ? S.deleted : 0}</span><span class="summary-key">Deleted</span></div>
          <div class="summary-metric"><span class="summary-val changed">${S.changed != null ? S.changed : 0}</span><span class="summary-key">Changed</span></div>
          <div class="summary-metric"><span class="summary-val muted">${S.unchanged != null ? S.unchanged : 0}</span><span class="summary-key">Unchanged</span></div>
        </div>
      </div>
      <div class="summary-divider"></div>
      <div class="summary-group">
        <span class="summary-group-label">Schedule Health</span>
        <div class="summary-metrics">
          <div class="summary-metric"><span class="summary-val ${delayed5 > 0 ? 'delayed' : 'ok'}">${delayed5}</span><span class="summary-key">Delayed &gt;5d</span></div>
          <div class="summary-metric"><span class="summary-val ${parseFloat(maxDelay) > 14 ? 'deleted' : parseFloat(maxDelay) > 0 ? 'changed' : 'ok'}">${maxDelay}</span><span class="summary-key">Max Delay</span></div>
          <div class="summary-metric"><span class="summary-val muted">${avgVar}</span><span class="summary-key">Avg Variance</span></div>
          <div class="summary-metric"><span class="summary-val ${critical > 0 ? 'deleted' : 'ok'}">${critical}</span><span class="summary-key">Critical</span></div>
        </div>
      </div>
    </div>

    <div class="ai-card">
      <div class="ai-header">
        <h2 class="section-title">AI Schedule Assessment</h2>
        <span class="health-badge health-${health}">${health.replace('_', ' ').replace(/\b\w/g, c => c.toUpperCase())}</span>
      </div>
      <p>${esc(ai.executive_summary || 'No AI summary available.')}</p>
      ${riskHtml}
      ${actionsHtml}
    </div>

    <h2 class="section-title">Top Delayed Activities</h2>
    <div class="chart-container" style="margin-bottom:24px;">
      <div id="chart-dist"></div>
    </div>

    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr>
            <th>Code</th><th>Name</th><th>WBS</th><th class="num">Finish Variance</th>
          </tr>
        </thead>
        <tbody>
          ${delayedRows}
        </tbody>
      </table>
    </div>
  `;

  // Render distribution chart on next tick
  setTimeout(() => {
    const distData = (D.chart_data && D.chart_data.dist_labels)
      ? D.chart_data.dist_labels.map((label, i) => ({
          label,
          value: D.chart_data.dist_values[i] || 0,
        }))
      : [
          { label: 'Added',     value: S.added     || 0 },
          { label: 'Deleted',   value: S.deleted   || 0 },
          { label: 'Changed',   value: S.changed   || 0 },
          { label: 'Unchanged', value: S.unchanged || 0 },
        ];

    renderG2Bar('chart-dist', distData, 'label', 'value', null, 220);
  }, 0);
}

/* ── Tab 2: Schedule ─────────────────────────────────────────────────────── */

function renderSchedule(el) {
  const variances = D.activity_variances || [];

  // Build top-20 delay chart data
  const top20 = variances
    .filter(a => (a.finish_variance_days || 0) > 0)
    .sort((a, b) => (b.finish_variance_days || 0) - (a.finish_variance_days || 0))
    .slice(0, 20);

  const PAGE = 200;

  function buildRows(items) {
    return items.map(a => `
    <tr data-testid="activity-row" class="row-${a.change_type || 'unchanged'}${a.is_critical ? ' row-critical' : ''}">
      <td>${esc(a.task_code)}</td>
      <td>${esc(a.task_name)}</td>
      <td>${esc(a.wbs_name)}</td>
      <td><span class="status-badge status-${a.change_type || 'unchanged'}">${esc(a.change_type || '—')}</span></td>
      <td class="num">${a.start_variance_days != null ? fmt(a.start_variance_days, 1) + 'd' : '—'}</td>
      <td class="num">${a.finish_variance_days != null ? fmt(a.finish_variance_days, 1) + 'd' : '—'}</td>
      <td class="num">${a.duration_variance_days != null ? fmt(a.duration_variance_days, 1) + 'd' : '—'}</td>
      <td class="num">${a.pct_complete_change != null ? fmt(a.pct_complete_change, 1) + '%' : '—'}</td>
      <td class="num">${a.float_change_hours != null ? fmt(a.float_change_hours, 1) + 'h' : '—'}</td>
      <td>${esc(a.new_status || a.old_status || '—')}</td>
      <td>${a.is_critical ? '<span class="tag-chip tag-critical">Critical</span>' : ''}</td>
    </tr>`).join('');
  }

  const initialRows = buildRows(variances.slice(0, PAGE));
  const hasMore = variances.length > PAGE;

  el.innerHTML = `
    <div class="filter-row">
      <label for="sched-search" class="sr-only">Search activities</label>
      <input type="search" class="search-input" id="sched-search" placeholder="Search code or name…" aria-label="Search activities">
      <label for="sched-filter" class="sr-only">Filter by change type</label>
      <select class="filter-select" id="sched-filter" aria-label="Filter by change type">
        <option value="">All Changes</option>
        <option value="added">Added</option>
        <option value="deleted">Deleted</option>
        <option value="changed">Changed</option>
        <option value="unchanged">Unchanged</option>
      </select>
    </div>

    <div class="chart-container" style="margin-bottom:24px;">
      <h2 class="section-title">Top 20 Delayed Activities</h2>
      <div id="chart-delay"></div>
    </div>

    <div class="table-wrap" id="sched-table-wrap">
      <table class="data-table" id="sched-table">
        <thead>
          <tr>
            <th>Code</th><th>Name</th><th>WBS</th><th>Change</th>
            <th class="num">Start Var</th><th class="num">Finish Var</th>
            <th class="num">Dur Var</th><th class="num">% Chg</th>
            <th class="num">Float Chg</th><th>Status</th><th>Critical</th>
          </tr>
        </thead>
        <tbody id="sched-tbody">
          ${initialRows}
        </tbody>
      </table>
    </div>
    ${hasMore ? `<div style="text-align:center;padding:16px;"><button class="btn-primary" id="sched-load-more">Load more (${variances.length - PAGE} remaining)</button></div>` : ''}
  `;

  // Wire up filtering
  const searchEl  = document.getElementById('sched-search');
  const filterEl  = document.getElementById('sched-filter');

  function applyFilter() {
    const q   = searchEl.value.toLowerCase();
    const typ = filterEl.value;
    Array.from(document.querySelectorAll('#sched-tbody tr[data-testid="activity-row"]')).forEach(row => {
      const text = row.textContent.toLowerCase();
      const matchQ   = !q   || text.includes(q);
      const matchTyp = !typ || row.classList.contains('row-' + typ);
      row.style.display = (matchQ && matchTyp) ? '' : 'none';
    });
  }
  let _debTimer;
  searchEl.addEventListener('input', () => { clearTimeout(_debTimer); _debTimer = setTimeout(applyFilter, 150); });
  filterEl.addEventListener('change', applyFilter);

  // Load more
  const loadMoreBtn = document.getElementById('sched-load-more');
  if (loadMoreBtn) {
    loadMoreBtn.addEventListener('click', () => {
      document.getElementById('sched-tbody').innerHTML = buildRows(variances);
      loadMoreBtn.parentElement.remove();
    });
  }

  // Render delay chart
  setTimeout(() => {
    const chartData = top20.map(a => ({
      code: (a.task_code || '').slice(0, 12),
      days: parseFloat(a.finish_variance_days) || 0,
    }));
    renderG2HBar('chart-delay', chartData, 'code', 'days', 300);
  }, 0);
}

/* ── Tab 3: KPIs ─────────────────────────────────────────────────────────── */

function renderKPIs(el) {
  const K = D.kpis || {};

  function colorForFloat(v) {
    if (v == null) return 'blue';
    if (v > 5) return 'red';
    if (v >= 1) return 'orange';
    return 'green';
  }
  function colorForPct(v) {
    if (v == null) return 'blue';
    if (v > 60) return 'green';
    if (v >= 30) return 'orange';
    return 'red';
  }
  function colorForSPI(v) {
    if (v == null) return 'blue';
    if (v >= 1.0) return 'green';
    if (v >= 0.8) return 'orange';
    return 'red';
  }
  function colorForDelay(v) {
    if (v == null) return 'blue';
    if (v <= 0) return 'green';
    if (v <= 14) return 'orange';
    return 'red';
  }

  const floatColor = colorForFloat(K.float_consumption_days);
  const pctColor   = colorForPct(K.pct_complete_weighted);
  const spiColor   = colorForSPI(K.spi_duration);
  const delayColor = colorForDelay(K.schedule_delay_days);

  const onTimeStart  = K.on_time_start_rate  != null ? fmt(K.on_time_start_rate * 100, 1)  + '%' : '—';
  const onTimeFinish = K.on_time_finish_rate != null ? fmt(K.on_time_finish_rate * 100, 1) + '%' : '—';

  const kpiChartData = [
    { metric: 'Float Consumed (d)', value: K.float_consumption_days  || 0 },
    { metric: 'SPI Duration',       value: K.spi_duration            || 0 },
    { metric: 'Delay (d)',          value: K.schedule_delay_days     || 0 },
    { metric: 'Near Critical',      value: K.near_critical_count     || 0 },
    { metric: 'Behind',             value: K.activities_behind       || 0 },
    { metric: 'Improved',           value: K.activities_improved     || 0 },
  ].filter(d => d.value !== 0);

  el.innerHTML = `
    <h2 class="section-title">Schedule Performance</h2>
    <div class="kpi-panel">
      <div class="kpi-row">
        <div class="kpi-name">Float Consumption</div>
        <div class="kpi-figure kpi-${floatColor}">${K.float_consumption_days != null ? fmt(K.float_consumption_days, 1) + 'd' : '—'}</div>
        <div class="kpi-context">avg days consumed since baseline</div>
        <div class="kpi-status kpi-${floatColor}">${K.float_consumption_days > 5 ? '↑ High' : K.float_consumption_days >= 1 ? '↑ Moderate' : '✓ Low'}</div>
      </div>
      <div class="kpi-row">
        <div class="kpi-name">Weighted % Complete</div>
        <div class="kpi-figure kpi-${pctColor}">${K.pct_complete_weighted != null ? fmt(K.pct_complete_weighted, 1) + '%' : '—'}</div>
        <div class="kpi-context">duration-weighted progress across all activities</div>
        <div class="kpi-status kpi-${pctColor}">${K.pct_complete_weighted > 60 ? '✓ On track' : K.pct_complete_weighted >= 30 ? '— Moderate' : '↓ Low'}</div>
      </div>
      <div class="kpi-row">
        <div class="kpi-name">SPI (Duration)</div>
        <div class="kpi-figure kpi-${spiColor}">${K.spi_duration != null ? fmt(K.spi_duration, 3) : '—'}</div>
        <div class="kpi-context">earned duration ÷ planned duration — ≥1.0 is on schedule</div>
        <div class="kpi-status kpi-${spiColor}">${K.spi_duration >= 1.0 ? '✓ On schedule' : K.spi_duration >= 0.8 ? '↓ Slipping' : '↓↓ Critical'}</div>
      </div>
      <div class="kpi-row">
        <div class="kpi-name">Schedule Delay</div>
        <div class="kpi-figure kpi-${delayColor}">${K.schedule_delay_days != null ? fmt(K.schedule_delay_days, 0) + 'd' : '—'}</div>
        <div class="kpi-context">net delay of updated plan-end vs baseline plan-end</div>
        <div class="kpi-status kpi-${delayColor}">${K.schedule_delay_days <= 0 ? '✓ No delay' : K.schedule_delay_days <= 14 ? '↑ Minor' : '↑↑ Significant'}</div>
      </div>
    </div>

    <div class="stats-grid" style="margin-top:24px;">
      <div class="stat-card">
        <div class="stat-num stat-num-blue">${K.avg_float_days != null ? fmt(K.avg_float_days, 1) + 'd' : '—'}</div>
        <div class="stat-label">Avg Float</div>
      </div>
      <div class="stat-card">
        <div class="stat-num stat-num-orange">${K.near_critical_count != null ? K.near_critical_count : '—'}</div>
        <div class="stat-label">Near Critical</div>
      </div>
      <div class="stat-card">
        <div class="stat-num stat-num-blue">${onTimeStart}</div>
        <div class="stat-label">On-Time Start</div>
      </div>
      <div class="stat-card">
        <div class="stat-num stat-num-blue">${onTimeFinish}</div>
        <div class="stat-label">On-Time Finish</div>
      </div>
      <div class="stat-card">
        <div class="stat-num stat-num-red">${K.activities_behind != null ? K.activities_behind : '—'}</div>
        <div class="stat-label">Behind</div>
      </div>
      <div class="stat-card">
        <div class="stat-num stat-num-green">${K.activities_improved != null ? K.activities_improved : '—'}</div>
        <div class="stat-label">Improved</div>
      </div>
      <div class="stat-card">
        <div class="stat-num stat-num-green">${K.activities_on_track != null ? K.activities_on_track : '—'}</div>
        <div class="stat-label">On Track</div>
      </div>
      <div class="stat-card">
        <div class="stat-num stat-num-blue">${K.total_activities != null ? K.total_activities : '—'}</div>
        <div class="stat-label">Total</div>
      </div>
    </div>

    <div class="chart-container" style="margin-top:24px;">
      <h2 class="section-title">KPI Overview</h2>
      <div id="chart-kpi"></div>
    </div>
  `;

  setTimeout(() => {
    if (kpiChartData.length) renderG2Bar('chart-kpi', kpiChartData, 'metric', 'value', null, 260);
  }, 0);
}

/* ── Tab 4: Lookahead ────────────────────────────────────────────────────── */

function renderLookahead(el) {
  const LA = D.lookahead || {};
  const counts = LA.counts || {};

  const windows = [
    { key: 'overdue',   label: 'Overdue',  items: LA.overdue   || [] },
    { key: 'two_week',  label: '2-Week',   items: LA.two_week  || [] },
    { key: 'four_week', label: '4-Week',   items: LA.four_week || [] },
    { key: 'six_week',  label: '6-Week',   items: LA.six_week  || [] },
  ];

  const defaultWindow = (LA.overdue && LA.overdue.length > 0) ? 'overdue' : 'two_week';

  const toggleHtml = windows.map(w => `
    <button class="toggle-btn${w.key === defaultWindow ? ' active' : ''}"
            data-window="${w.key}">
      ${w.label} <span class="badge-count">${w.items.length}</span>
    </button>`).join('');

  el.innerHTML = `
    <h2 class="section-title">Schedule Lookahead</h2>
    <div class="lookahead-toggle" id="la-toggle">
      ${toggleHtml}
    </div>
    <div id="la-table-wrap"></div>
  `;

  function renderWindow(key) {
    const win    = windows.find(w => w.key === key);
    const items  = win ? win.items : [];
    const wrap   = document.getElementById('la-table-wrap');

    // Group by WBS
    const groups = {};
    items.forEach(item => {
      const g = item.wbs_name || 'Ungrouped';
      if (!groups[g]) groups[g] = [];
      groups[g].push(item);
    });

    let rows = '';
    if (items.length === 0) {
      rows = `<tr><td colspan="10" class="no-data">No activities in this window</td></tr>`;
    } else {
      Object.entries(groups).forEach(([wbsName, groupItems]) => {
        rows += `<tr class="wbs-group-row"><td colspan="10">${esc(wbsName)}</td></tr>`;
        rows += groupItems.map(a => `
          <tr data-testid="activity-row" class="${a.is_critical ? 'row-critical' : ''}">
            <td>${esc(a.task_code)}</td>
            <td>${esc(a.task_name)}</td>
            <td>${esc(a.wbs_name)}</td>
            <td>${fmtDate(a.planned_start)}</td>
            <td>${fmtDate(a.planned_finish)}</td>
            <td class="num">${a.duration_days != null ? fmt(a.duration_days, 0) + 'd' : '—'}</td>
            <td class="num">${a.pct_complete != null ? fmt(a.pct_complete, 0) + '%' : '—'}</td>
            <td class="num">${a.float_days != null ? fmt(a.float_days, 1) + 'd' : '—'}</td>
            <td>${esc(a.status || '—')}</td>
            <td>${a.is_critical ? '<span class="tag-chip tag-critical">Critical</span>' : ''}</td>
          </tr>`).join('');
      });
    }

    wrap.innerHTML = `
      <table class="data-table">
        <thead>
          <tr>
            <th>Code</th><th>Name</th><th>WBS</th>
            <th>Start</th><th>Finish</th>
            <th class="num">Duration</th><th class="num">% Comp</th>
            <th class="num">Float</th><th>Status</th><th>Critical</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>`;
  }

  // Toggle handlers
  document.getElementById('la-toggle').addEventListener('click', e => {
    const btn = e.target.closest('.toggle-btn');
    if (!btn) return;
    document.querySelectorAll('#la-toggle .toggle-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    renderWindow(btn.dataset.window);
  });

  renderWindow(defaultWindow);
}

/* ── Tab 5: Milestones ───────────────────────────────────────────────────── */

function renderMilestones(el) {
  const milestones = D.milestones || [];

  // Count by status
  const counts = { complete: 0, on_track: 0, at_risk: 0, late: 0 };
  milestones.forEach(m => { if (counts[m.status] != null) counts[m.status]++; });

  // Sort: late first, at_risk, on_track, complete
  const order = { late: 0, at_risk: 1, on_track: 2, complete: 3, unknown: 4 };
  const sorted = [...milestones].sort((a, b) => (order[a.status] || 4) - (order[b.status] || 4));

  const rows = sorted.map(m => `
    <tr data-testid="activity-row" class="row-${m.change_type || 'unchanged'}">
      <td>${esc(m.task_code)}</td>
      <td>${esc(m.task_name)}</td>
      <td>${esc(m.wbs_name)}</td>
      <td>${fmtDate(m.baseline_finish)}</td>
      <td>${fmtDate(m.updated_finish)}</td>
      <td>${fmtDate(m.actual_finish)}</td>
      <td class="num">${m.finish_variance_days != null ? fmt(m.finish_variance_days, 0) + 'd' : '—'}</td>
      <td class="num">${m.pct_complete != null ? fmt(m.pct_complete, 0) + '%' : '—'}</td>
      <td><span class="status-badge status-${m.status || 'unknown'}">${esc(m.status || 'unknown')}</span></td>
    </tr>`).join('');

  el.innerHTML = `
    <h2 class="section-title">Milestones</h2>
    <div class="chips-row">
      <span class="tag-chip">Total: ${milestones.length}</span>
      <span class="tag-chip tag-complete">Complete: ${counts.complete}</span>
      <span class="tag-chip tag-on-track">On Track: ${counts.on_track}</span>
      <span class="tag-chip tag-at-risk">At Risk: ${counts.at_risk}</span>
      <span class="tag-chip tag-late">Late: ${counts.late}</span>
    </div>
    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr>
            <th>Code</th><th>Name</th><th>WBS</th>
            <th>Baseline Finish</th><th>Updated Finish</th><th>Actual Finish</th>
            <th class="num">Variance</th><th class="num">%</th><th>Status</th>
          </tr>
        </thead>
        <tbody>
          ${rows.length ? rows : '<tr><td colspan="9" class="no-data">No milestones found</td></tr>'}
        </tbody>
      </table>
    </div>
  `;
}

/* ── Tab 6: Procurement ──────────────────────────────────────────────────── */

function renderProcurement(el) {
  const P = D.procurement || {};
  const items = P.items || [];
  const detected = P.detected_wbs_nodes || [];
  const summary  = P.summary || {};

  if (detected.length === 0) {
    el.innerHTML = `
      <h2 class="section-title">Procurement</h2>
      <div class="info-card">
        <strong>No procurement WBS nodes detected</strong>
        <p>No WBS nodes matching typical procurement patterns were found in the schedules.</p>
      </div>`;
    return;
  }

  const wbsChips = detected.map(n => `<span class="tag-chip">${esc(n)}</span>`).join('');

  const rows = items.map(item => {
    const typeTag = item.is_long_lead
      ? '<span class="tag-chip tag-long-lead">Long Lead</span>'
      : item.is_short_lead
        ? '<span class="tag-chip tag-short-lead">Short Lead</span>'
        : '—';
    return `
    <tr data-testid="activity-row">
      <td>${esc(item.task_code)}</td>
      <td>${esc(item.task_name)}</td>
      <td>${esc(item.wbs_name)}</td>
      <td>${fmtDate(item.baseline_finish)}</td>
      <td>${fmtDate(item.updated_finish)}</td>
      <td class="num">${item.finish_variance_days != null ? fmt(item.finish_variance_days, 0) + 'd' : '—'}</td>
      <td class="num">${item.pct_complete != null ? fmt(item.pct_complete, 0) + '%' : '—'}</td>
      <td><span class="status-badge status-${item.status || 'unknown'}">${esc(item.status || '—')}</span></td>
      <td>${typeTag}</td>
    </tr>`;
  }).join('');

  el.innerHTML = `
    <h2 class="section-title">Procurement</h2>
    <div class="chips-row" style="margin-bottom:12px;">
      <strong>Detected WBS Nodes:</strong> ${wbsChips}
    </div>
    <div class="stats-grid" style="margin-bottom:24px;">
      <div class="stat-card">
        <div class="stat-num stat-num-blue">${summary.total || 0}</div>
        <div class="stat-label">Total</div>
      </div>
      <div class="stat-card">
        <div class="stat-num stat-num-green">${summary.complete || 0}</div>
        <div class="stat-label">Complete</div>
      </div>
      <div class="stat-card">
        <div class="stat-num stat-num-orange">${summary.in_progress || 0}</div>
        <div class="stat-label">In Progress</div>
      </div>
      <div class="stat-card">
        <div class="stat-num stat-num-muted">${summary.not_started || 0}</div>
        <div class="stat-label">Not Started</div>
      </div>
      <div class="stat-card">
        <div class="stat-num stat-num-red">${summary.late || 0}</div>
        <div class="stat-label">Late</div>
      </div>
    </div>
    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr>
            <th>Code</th><th>Name</th><th>WBS</th>
            <th>Baseline Finish</th><th>Updated Finish</th>
            <th class="num">Variance</th><th class="num">%</th>
            <th>Status</th><th>Type</th>
          </tr>
        </thead>
        <tbody>
          ${rows.length ? rows : '<tr><td colspan="9" class="no-data">No procurement items found</td></tr>'}
        </tbody>
      </table>
    </div>
  `;
}

/* ── Tab 7: Earned Value ─────────────────────────────────────────────────── */

function renderEV(el) {
  const ev = D.ev || {};
  const K  = D.kpis || {};

  if (!ev.has_cost_data) {
    el.innerHTML = `
      <h2 class="section-title">Earned Value</h2>
      <div class="info-card">
        <strong>No cost data found in the XER files</strong>
        <p>EV analysis requires cost-loaded activities. Showing duration-based schedule metrics instead.</p>
      </div>
      <div class="kpi-grid" style="margin-top:24px;">
        <div class="kpi-card">
          <div class="kpi-val kpi-blue">${K.spi_duration != null ? fmt(K.spi_duration, 2) : '—'}</div>
          <div class="kpi-label">SPI (Duration)</div>
          <div class="kpi-desc">Duration-based schedule performance index</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-val kpi-blue">${K.pct_complete_weighted != null ? fmt(K.pct_complete_weighted, 1) + '%' : '—'}</div>
          <div class="kpi-label">Weighted % Complete</div>
          <div class="kpi-desc">Duration-weighted progress</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-val kpi-blue">${K.schedule_delay_days != null ? fmt(K.schedule_delay_days, 0) + 'd' : '—'}</div>
          <div class="kpi-label">Schedule Delay</div>
          <div class="kpi-desc">Net delay vs baseline</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-val kpi-blue">${K.float_consumption_days != null ? fmt(K.float_consumption_days, 1) + 'd' : '—'}</div>
          <div class="kpi-label">Float Consumption</div>
          <div class="kpi-desc">Average float consumed</div>
        </div>
      </div>`;
    return;
  }

  const upd = ev.updated || {};
  const bas = ev.baseline || {};

  const fmtCost = v => v != null ? '$' + Number(v).toLocaleString(undefined, { maximumFractionDigits: 0 }) : '—';
  const fmtIdx  = v => v != null ? fmt(v, 3) : '—';

  el.innerHTML = `
    <h2 class="section-title">Earned Value Analysis</h2>
    <div class="kpi-grid">
      <div class="kpi-card">
        <div class="kpi-val kpi-blue">${fmtCost(bas.BAC || upd.BAC)}</div>
        <div class="kpi-label">BAC</div>
        <div class="kpi-desc">Budget at Completion</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-val kpi-blue">${fmtCost(upd.EV)}</div>
        <div class="kpi-label">EV</div>
        <div class="kpi-desc">Earned Value</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-val kpi-blue">${fmtCost(upd.AC)}</div>
        <div class="kpi-label">AC</div>
        <div class="kpi-desc">Actual Cost</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-val kpi-${upd.SPI >= 1 ? 'green' : upd.SPI >= 0.8 ? 'orange' : 'red'}">${fmtIdx(upd.SPI)}</div>
        <div class="kpi-label">SPI</div>
        <div class="kpi-desc">Schedule Performance Index</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-val kpi-${upd.CPI >= 1 ? 'green' : upd.CPI >= 0.8 ? 'orange' : 'red'}">${fmtIdx(upd.CPI)}</div>
        <div class="kpi-label">CPI</div>
        <div class="kpi-desc">Cost Performance Index</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-val kpi-blue">${fmtCost(upd.EAC)}</div>
        <div class="kpi-label">EAC</div>
        <div class="kpi-desc">Estimate at Completion</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-val kpi-${(upd.VAC || 0) >= 0 ? 'green' : 'red'}">${fmtCost(upd.VAC)}</div>
        <div class="kpi-label">VAC</div>
        <div class="kpi-desc">Variance at Completion</div>
      </div>
    </div>

    <div class="chart-container" style="margin-top:24px;">
      <h2 class="section-title">S-Curve</h2>
      <div id="chart-scurve"></div>
    </div>
  `;

  setTimeout(() => {
    const scurve = D.scurve || {};
    const baselineData = (scurve.baseline || []).map(p => ({ period: p.period, cost: p.cumulative_cost, series: 'Baseline' }));
    const updatedData  = (scurve.updated  || []).map(p => ({ period: p.period, cost: p.cumulative_cost, series: 'Updated'  }));
    const combined = [...baselineData, ...updatedData];
    if (combined.length) renderG2Line('chart-scurve', combined, 'period', 'cost', 'series', 280);
  }, 0);
}

/* ── Tab 8: WBS ──────────────────────────────────────────────────────────── */

function renderWBS(el) {
  const wbs = D.wbs_summary || [];

  if (wbs.length === 0) {
    el.innerHTML = `
      <h2 class="section-title">WBS Summary</h2>
      <div class="info-card"><p>No WBS data available.</p></div>`;
    return;
  }

  // Discover columns from first row
  const sampleKeys = Object.keys(wbs[0] || {});

  const headers = sampleKeys.map(k => `<th>${esc(k.replace(/_/g, ' '))}</th>`).join('');

  const rows = wbs.map(row => `
    <tr data-testid="activity-row">
      ${sampleKeys.map(k => `<td>${esc(row[k])}</td>`).join('')}
    </tr>`).join('');

  el.innerHTML = `
    <h2 class="section-title">WBS Summary</h2>
    <div class="table-wrap">
      <table class="data-table">
        <thead><tr>${headers}</tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

/* ── Tab 9: Logic (Relationships) ────────────────────────────────────────── */

function renderLogic(el) {
  const rels = D.relationship_variances || [];

  const countAdded   = rels.filter(r => r.change_type === 'added').length;
  const countDeleted = rels.filter(r => r.change_type === 'deleted').length;
  const countChanged = rels.filter(r => r.change_type === 'changed').length;

  const rows = rels.map(r => `
    <tr data-testid="activity-row" class="row-${r.change_type || 'unchanged'}">
      <td>${esc(r.pred_code)}</td>
      <td>${esc(r.succ_code)}</td>
      <td><span class="status-badge status-${r.change_type || 'unchanged'}">${esc(r.change_type || '—')}</span></td>
      <td>${esc(r.old_pred_type || '—')}</td>
      <td>${esc(r.new_pred_type || '—')}</td>
      <td class="num">${r.lag_change_hours != null ? fmt(r.lag_change_hours, 1) + 'h' : '—'}</td>
    </tr>`).join('');

  el.innerHTML = `
    <h2 class="section-title">Relationship / Logic Changes</h2>
    <div class="chips-row" style="margin-bottom:16px;">
      <span class="tag-chip tag-added">Added: ${countAdded}</span>
      <span class="tag-chip tag-deleted">Deleted: ${countDeleted}</span>
      <span class="tag-chip tag-changed">Changed: ${countChanged}</span>
      <span class="tag-chip">Total: ${rels.length}</span>
    </div>
    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr>
            <th>Pred Code</th><th>Succ Code</th><th>Change</th>
            <th>Old Type</th><th>New Type</th><th class="num">Lag Change (h)</th>
          </tr>
        </thead>
        <tbody>
          ${rows.length ? rows : '<tr><td colspan="6" class="no-data">No relationship changes</td></tr>'}
        </tbody>
      </table>
    </div>
  `;
}

/* ── Tab 10: AI Chat ─────────────────────────────────────────────────────── */

function renderChat(el) {
  const projectName = (D.project && D.project.updated_name) || 'this schedule';

  const suggestions = [
    `What are the top schedule risks?`,
    `Which activities are most delayed?`,
    `What is the overall schedule health?`,
  ];

  const suggestionChips = suggestions.map(s =>
    `<button class="suggestion-chip" data-question="${esc(s)}">${esc(s)}</button>`
  ).join('');

  el.innerHTML = `
    <h2 class="section-title">AI Schedule Assistant</h2>
    <div class="chat-container">
      <div class="chat-messages" id="chat-messages" aria-live="polite" aria-atomic="false">
        <div class="chat-msg msg-ai">
          <strong>Assistant:</strong> I'm analyzing <em>${esc(projectName)}</em>. Ask me anything about the schedule comparison.
        </div>
      </div>
      <div class="suggestions" id="chat-suggestions">
        ${suggestionChips}
      </div>
      <div class="chat-input-row">
        <input type="text" class="chat-input" id="chat-input" placeholder="Ask about the schedule…" autocomplete="off">
        <button class="btn-primary" id="chat-send">Send</button>
      </div>
    </div>
  `;

  const messagesEl  = document.getElementById('chat-messages');
  const inputEl     = document.getElementById('chat-input');
  const sendBtn     = document.getElementById('chat-send');
  const suggestEl   = document.getElementById('chat-suggestions');

  function appendMessage(role, text) {
    const div = document.createElement('div');
    div.className = `chat-msg msg-${role}`;
    div.innerHTML = `<strong>${role === 'user' ? 'You' : 'Assistant'}:</strong> ${renderMarkdownLite(text)}`;
    messagesEl.appendChild(div);
    messagesEl.scrollTop = messagesEl.scrollHeight;
    return div;
  }

  function renderMarkdownLite(text) {
    return esc(text)
      .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
      .replace(/\*(.*?)\*/g, '<em>$1</em>')
      .replace(/`(.*?)`/g, '<code>$1</code>')
      .replace(/\n/g, '<br>');
  }

  async function sendMessage(message) {
    if (!message.trim()) return;
    inputEl.value = '';
    sendBtn.disabled = true;
    suggestEl.style.display = 'none';

    appendMessage('user', message);

    const thinkingDiv = appendMessage('ai', 'Thinking…');

    try {
      const response = await fetch('/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ key: KEY, message }),
      });

      if (!response.ok) {
        thinkingDiv.innerHTML = `<strong>Assistant:</strong> Error: ${response.statusText}`;
        return;
      }

      // Check if SSE or regular JSON
      const contentType = response.headers.get('Content-Type') || '';
      if (contentType.includes('text/event-stream') || contentType.includes('text/plain')) {
        // SSE streaming
        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let accumulated = '';
        thinkingDiv.innerHTML = '<strong>Assistant:</strong> ';

        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          const chunk = decoder.decode(value, { stream: true });
          // Parse SSE lines
          const lines = chunk.split('\n');
          for (const line of lines) {
            if (line.startsWith('data: ')) {
              const jsonStr = line.slice(6).trim();
              if (jsonStr === '[DONE]') break;
              try {
                const parsed = JSON.parse(jsonStr);
                const token = parsed.text || parsed.token || parsed.content || '';
                accumulated += token;
              } catch {
                // Raw text token
                accumulated += jsonStr;
              }
            }
          }
          thinkingDiv.innerHTML = `<strong>Assistant:</strong> ${renderMarkdownLite(accumulated)}`;
          messagesEl.scrollTop = messagesEl.scrollHeight;
        }

        if (!accumulated) {
          thinkingDiv.innerHTML = '<strong>Assistant:</strong> No response received.';
        }
      } else {
        // Regular JSON
        const data = await response.json();
        const text = data.response || data.text || data.message || JSON.stringify(data);
        thinkingDiv.innerHTML = `<strong>Assistant:</strong> ${renderMarkdownLite(text)}`;
      }
    } catch (err) {
      thinkingDiv.innerHTML = `<strong>Assistant:</strong> Error: ${esc(err.message)}`;
    } finally {
      sendBtn.disabled = false;
    }
  }

  sendBtn.addEventListener('click', () => sendMessage(inputEl.value));
  inputEl.addEventListener('keydown', e => { if (e.key === 'Enter') sendMessage(inputEl.value); });

  suggestEl.addEventListener('click', e => {
    const chip = e.target.closest('.suggestion-chip');
    if (chip) sendMessage(chip.dataset.question);
  });
}

/* ── Bootstrap ───────────────────────────────────────────────────────────── */

document.addEventListener('DOMContentLoaded', () => {
  // Tab click handlers
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const name = btn.dataset.tab;
      history.pushState(null, '', '#' + name);
      activateTab(name);
    });
  });

  // Hash change (browser back/forward)
  window.addEventListener('hashchange', () => {
    const name = location.hash.slice(1) || 'overview';
    activateTab(name);
  });

  // Initial render from hash (or default to overview)
  const initial = location.hash.slice(1) || 'overview';
  activateTab(initial);
});
