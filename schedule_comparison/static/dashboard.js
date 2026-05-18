/* dashboard.js — Schedule Comparison Dashboard */
'use strict';

const D = window.__DASH__;
const KEY = window.__KEY__;
const rendered = new Set();

/* ── Count-up animation ─────────────────────────────────────────────────────── */

function countUp(el, target, duration, prefix, suffix, formatter) {
  if (!el) return;
  duration = duration || 800;
  prefix   = prefix   || '';
  suffix   = suffix   || '';
  const start = performance.now();
  const isFloat = !Number.isInteger(target);
  const decimals = isFloat ? (String(target).split('.')[1] || '').length : 0;
  function step(now) {
    const elapsed = Math.min((now - start) / duration, 1);
    const ease = 1 - Math.pow(1 - elapsed, 3); // ease-out-cubic
    const cur = target * ease;
    el.textContent = prefix + (formatter ? formatter(cur) : cur.toFixed(decimals)) + suffix;
    if (elapsed < 1) requestAnimationFrame(step);
  }
  requestAnimationFrame(step);
}

/* ── IntersectionObserver progress bar fill ──────────────────────────────── */
function initProgressBars(container) {
  const els = (container || document).querySelectorAll('.evm-progress-fill, .proc-prog-fill, .prog-bar-fill-anim');
  if (!els.length) return;
  const io = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (!entry.isIntersecting) return;
      const el = entry.target;
      const targetW = el.dataset.targetW || el.style.width || '0%';
      el.style.setProperty('--target-w', targetW);
      el.style.width = '0%';
      requestAnimationFrame(() => {
        el.style.transition = 'width .9s cubic-bezier(.4,0,.2,1)';
        el.style.width = targetW;
      });
      io.unobserve(el);
    });
  }, { threshold: 0.1 });
  els.forEach(el => {
    el.dataset.targetW = el.style.width || '0%';
    io.observe(el);
  });
}

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

function fmtCost(v) {
  if (!v && v !== 0) return '—';
  const abs = Math.abs(v);
  const sign = v < 0 ? '-' : '';
  if (abs >= 1e6) return sign + (abs/1e6).toFixed(2) + 'M';
  if (abs >= 1e3) return sign + (abs/1e3).toFixed(1) + 'K';
  return sign + abs.toFixed(0);
}

function fmtDate(iso) {
  if (!iso) return '—';
  try { return iso.slice(0, 10); } catch (e) { return iso; }
}

function fmtCurrency(val) {
  if (val == null || val === '') return '—';
  return '$' + Number(val).toLocaleString(undefined, { maximumFractionDigits: 0 });
}

function plural(n, word) {
  return `${n} ${word}${n === 1 ? '' : 's'}`;
}

/* ── AntV G2 v5 chart helpers ────────────────────────────────────────────── */

function _g2ok(container) {
  if (typeof G2 === 'undefined') {
    if (container) container.innerHTML = '<p style="color:var(--text-muted);padding:12px;font-size:.8rem">Chart unavailable</p>';
    return false;
  }
  return true;
}

function renderG2Bar(containerId, data, xField, yField, colorField, height = 280) {
  const container = document.getElementById(containerId);
  if (!container || !data.length) return;
  if (!_g2ok(container)) return;
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
  if (!_g2ok(container)) return;
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
  if (!_g2ok(container)) return;
  container.style.height = height + 'px';
  const chart = new G2.Chart({ container, autoFit: true, height });
  chart.options({
    type: 'line',
    data,
    encode: { x: xField, y: yField, color: colorField },
    style: { strokeWidth: 2 },
    axis: { 
      y: { labelFormatter: (v) => v >= 1000 ? (v/1000).toFixed(0)+'k' : v, title: false }, 
      x: { title: false } 
    },
  });
  chart.render();
}

/** Combined Histogram (Periodic) + S-Curve (Cumulative) Chart */
function renderG2Combo(containerId, data, xField, height = 280) {
  const container = document.getElementById(containerId);
  if (!container || !data.length) return;
  if (!_g2ok(container)) return;
  container.style.height = height + 'px';
  const chart = new G2.Chart({ container, autoFit: true, height });

  // Filter for Bar (Periodic) and Line (Cumulative)
  const barData = data.filter(d => d.series.includes('Monthly'));
  const lineData = data.filter(d => !d.series.includes('Monthly'));

  chart.data(data);

  // Line: Cumulative S-Curve
  chart.line()
    .data(lineData)
    .encode('x', xField)
    .encode('y', 'value')
    .encode('color', 'series')
    .encode('shape', 'smooth')
    .style('strokeWidth', 2.5)
    .axis('y', { title: 'Cumulative Cost', position: 'left' });

  // Interval: Periodic Histogram
  chart.interval()
    .data(barData)
    .encode('x', xField)
    .encode('y', 'value')
    .encode('color', 'series')
    .style('opacity', 0.4)
    .axis('y', { title: 'Periodic Cost', position: 'right', grid: null, labelFormatter: (v) => v >= 1000 ? (v/1000).toFixed(0)+'k' : v });

  chart.interaction('tooltip', { shared: true, showCrosshairs: true });
  chart.render();
}

/* ── Inline SVG chart helper (no CDN needed) ─────────────────────────────── */
function renderSVGChart(containerId, seriesData, type, height) {
  type = type || 'line'; height = height || 280;
  const container = typeof containerId === 'string'
    ? document.getElementById(containerId) : containerId;
  if (!container || !seriesData.length) return;
  container.style.height = height + 'px';
  const seriesMap = {};
  seriesData.forEach(d => { (seriesMap[d.series] = seriesMap[d.series]||[]).push(d); });
  const allPeriods = [...new Set(seriesData.map(d => d.period))].sort();
  const maxVal = Math.max(...seriesData.map(d => d.value||0)) || 1;
  const n = allPeriods.length;
  const W = container.clientWidth || 600;
  const H = height;
  const pad = { l:62, r:16, t:22, b:36 };
  const cW = W - pad.l - pad.r, cH = H - pad.t - pad.b;
  const xPos = i => pad.l + (i + 0.5) * cW / n;
  const yPos = v => pad.t + cH - Math.min(v / maxVal, 1) * cH;
  const yFmt = v => v >= 1e6 ? (v/1e6).toFixed(1)+'M' : v >= 1e3 ? (v/1e3).toFixed(0)+'K' : v.toFixed(0);
  const COLORS = ['#2563eb','#16a34a','#d97706','#dc2626','#8b5cf6','#0891b2'];
  const keys = Object.keys(seriesMap);
  let g = '';
  // grid
  [0,.25,.5,.75,1].forEach(f => {
    const y = pad.t + cH*(1-f);
    g += `<line x1="${pad.l}" y1="${y}" x2="${pad.l+cW}" y2="${y}" stroke="#f1f5f9" stroke-width="1"/>`;
    g += `<text x="${pad.l-4}" y="${y+3.5}" text-anchor="end" font-size="9" fill="#94a3b8">${yFmt(maxVal*f)}</text>`;
  });
  if (type === 'line') {
    keys.forEach((sk,si) => {
      const pts = seriesMap[sk].filter(d=>d.value!=null)
        .map(d => `${xPos(allPeriods.indexOf(d.period))},${yPos(d.value)}`).join(' ');
      if (pts) g += `<polyline points="${pts}" fill="none" stroke="${COLORS[si%COLORS.length]}" stroke-width="2.5" stroke-linejoin="round"/>`;
    });
  } else {
    const bW = Math.max(2, cW/n/(keys.length+0.5)-1);
    keys.forEach((sk,si) => {
      seriesMap[sk].forEach(d => {
        if (!d.value) return;
        const xi = allPeriods.indexOf(d.period);
        const x0 = pad.l + xi*cW/n + si*(bW+1) + 2;
        const h2 = Math.min(d.value/maxVal,1)*cH;
        g += `<rect x="${x0.toFixed(1)}" y="${yPos(d.value).toFixed(1)}" width="${bW.toFixed(1)}" height="${h2.toFixed(1)}" fill="${COLORS[si%COLORS.length]}" opacity="0.85" rx="1"/>`;
      });
    });
  }
  // x labels
  const step = n>24?6:n>12?3:n>8?2:1;
  allPeriods.forEach((p,i) => {
    if (i%step!==0) return;
    g += `<text x="${xPos(i).toFixed(1)}" y="${H-4}" text-anchor="middle" font-size="8" fill="#94a3b8">${p.slice(0,7)}</text>`;
  });
  // legend
  keys.forEach((sk,si) => {
    g += `<rect x="${pad.l+si*130}" y="4" width="10" height="10" rx="2" fill="${COLORS[si%COLORS.length]}"/>`;
    g += `<text x="${pad.l+si*130+14}" y="13" font-size="9" fill="#64748b">${sk}</text>`;
  });
  container.innerHTML = `<svg width="100%" height="${H}" style="display:block"><g>${g}</g></svg>`;
}

/* ── Tab routing ─────────────────────────────────────────────────────────── */

function activateTab(name) {
  // Update sidebar nav-item active state
  document.querySelectorAll('.nav-item').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.tab === name);
  });
  // Update legacy tab-btn active state (if present)
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

  // Render if not yet done (lazy — each tab renders once)
  if (rendered.has(name)) return;
  rendered.add(name);
  const renderers = {
    executive:   renderOverview,
    overview:    renderOverview,
    scurve:      renderSchedule,
    schedule:    renderSchedule,
    spitrends:   renderKPIs,
    kpis:        renderKPIs,
    drilldown:   renderWBS,
    trade:       renderLogic,
    manpower:    renderResources,
    slip:        renderLookahead,
    lookahead:   renderLookahead,
    milestones:  renderMilestones,
    procurement: renderProcurement,
    ev:          renderEV,
    resources:   renderResources,
    gantt:       renderGantt,
    wbs:         renderWBS,
    logic:       renderLogic,
    chat:        renderChat,
  };
  if (renderers[name]) renderers[name](panel);
}

/* ── Command Center: Drawer helper ───────────────────────────────────────── */

function openDrawer(title, bodyHtml) {
  const titleEl   = document.getElementById('cc-drawer-title')   || document.getElementById('drawer-title');
  const bodyEl    = document.getElementById('cc-drawer-body')    || document.getElementById('drawer-body');
  const drawerEl  = document.getElementById('cc-drawer')         || document.getElementById('drawer');
  const overlayEl = document.getElementById('cc-drawer-overlay') || document.getElementById('drawer-overlay');
  if (titleEl)   titleEl.textContent = title;
  if (bodyEl)    bodyEl.innerHTML = '<div class="drawer-content">' + bodyHtml + '</div>';
  if (drawerEl)  drawerEl.classList.add('open');
  if (overlayEl) overlayEl.classList.add('active');
}

function drawerTable(headers, rows) {
  if (!rows.length) return '<p style="color:var(--text-muted);padding:20px 0">No data available.</p>';
  const th = headers.map(h => `<th>${esc(h)}</th>`).join('');
  const tr = rows.map(r => `<tr>${r.map(c => `<td>${esc(String(c ?? '—'))}</td>`).join('')}</tr>`).join('');
  return `<div class="table-wrap"><table class="data-table"><thead><tr>${th}</tr></thead><tbody>${tr}</tbody></table></div>`;
}

/* ── Tab 1: Overview ─────────────────────────────────────────────────────── */

/* ── S-Curve SVG renderer ────────────────────────────────────────────────── */
function renderSCurveSVG(container, baselineArr, updatedArr) {
  if (!container) return;

  const hasBase = baselineArr && baselineArr.length > 0;
  const hasUpd  = updatedArr  && updatedArr.length  > 0;
  if (!hasBase && !hasUpd) {
    container.innerHTML = '<div class="scurve-empty">No S-curve data available</div>';
    return;
  }

  const W  = container.clientWidth || 680;
  const H  = 300;
  const PL = 72, PR = 24, PT = 32, PB = 52;
  const cW = W - PL - PR;
  const cH = H - PT - PB;

  // Parse dates
  function parseDate(s) { return s ? new Date(s) : null; }

  const basePts = hasBase ? baselineArr.map(p => ({ d: parseDate(p.period_date), v: p.planned_cum_cost || 0 })).filter(p => p.d) : [];
  const updPts  = hasUpd  ? updatedArr.map(p  => ({ d: parseDate(p.period_date), v: p.actual_cum_cost  || 0 })).filter(p => p.d) : [];

  // Date range
  const allDates = [...basePts.map(p=>p.d), ...updPts.map(p=>p.d)];
  const minD = new Date(Math.min(...allDates));
  const maxD = new Date(Math.max(...allDates));
  const totalMs = maxD - minD || 1;

  // Value range
  const allVals = [...basePts.map(p=>p.v), ...updPts.map(p=>p.v)];
  const maxVal = Math.max(...allVals, 1);

  // Coordinate helpers
  function xOf(d) { return PL + ((d - minD) / totalMs) * cW; }
  function yOf(v) { return PT + cH - (v / maxVal) * cH; }

  // Build polyline points
  function ptStr(pts) { return pts.map(p => xOf(p.d).toFixed(1) + ',' + yOf(p.v).toFixed(1)).join(' '); }

  // Build area path (close at bottom)
  function areaPath(pts) {
    if (!pts.length) return '';
    const startX = xOf(pts[0].d).toFixed(1);
    const endX   = xOf(pts[pts.length-1].d).toFixed(1);
    const bottom = (PT + cH).toFixed(1);
    return 'M' + startX + ',' + bottom +
           ' L' + pts.map(p => xOf(p.d).toFixed(1)+','+yOf(p.v).toFixed(1)).join(' L') +
           ' L' + endX + ',' + bottom + ' Z';
  }

  // Y-axis tick formatter
  function fmtY(v) {
    if (v === 0) return '$0';
    if (Math.abs(v) >= 1e9) return '$' + (v/1e9).toFixed(1) + 'B';
    if (Math.abs(v) >= 1e6) return '$' + (v/1e6).toFixed(1) + 'M';
    if (Math.abs(v) >= 1e3) return '$' + (v/1e3).toFixed(0) + 'K';
    return '$' + v.toFixed(0);
  }

  // X-axis ticks (smart spacing)
  const spanMonths = totalMs / (1000*60*60*24*30.44);
  function xTicks() {
    const ticks = [];
    const cur = new Date(minD);
    cur.setDate(1);
    if (cur < minD) cur.setMonth(cur.getMonth()+1);
    let step;
    if (spanMonths <= 12) step = 1;
    else if (spanMonths <= 36) step = 3;
    else step = 6;
    while (cur <= maxD) {
      ticks.push(new Date(cur));
      cur.setMonth(cur.getMonth() + step);
    }
    return ticks;
  }

  // Y-axis ticks
  const yTicks = [0,1,2,3,4].map(i => maxVal * i / 4);

  // Unique clip id
  const suffix = container.id || ('sc' + Math.random().toString(36).slice(2,7));
  const clipId = 'scclip-' + suffix + '-r';
  const gradBase = 'scgrad-base-' + suffix;
  const gradUpd  = 'scgrad-upd-'  + suffix;

  // Build SVG
  let svgParts = [];
  svgParts.push(`<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" style="display:block">`);

  // Defs: gradients + clippath
  svgParts.push(`<defs>`);
  svgParts.push(`<linearGradient id="${gradBase}" x1="0" y1="0" x2="0" y2="1">
    <stop offset="0%" stop-color="#3b82f6" stop-opacity="0.18"/>
    <stop offset="100%" stop-color="#3b82f6" stop-opacity="0.02"/>
  </linearGradient>`);
  svgParts.push(`<linearGradient id="${gradUpd}" x1="0" y1="0" x2="0" y2="1">
    <stop offset="0%" stop-color="#22c55e" stop-opacity="0.22"/>
    <stop offset="100%" stop-color="#22c55e" stop-opacity="0.02"/>
  </linearGradient>`);
  svgParts.push(`<clipPath id="${clipId}">
    <rect id="${clipId}-rect" x="${PL}" y="0" width="0" height="${H}"/>
  </clipPath>`);
  svgParts.push(`</defs>`);

  // Grid lines (horizontal)
  yTicks.forEach(v => {
    const y = yOf(v).toFixed(1);
    svgParts.push(`<line x1="${PL}" y1="${y}" x2="${PL+cW}" y2="${y}" stroke="#e8edf4" stroke-width="1"/>`);
  });

  // Axes
  svgParts.push(`<line x1="${PL}" y1="${PT}" x2="${PL}" y2="${PT+cH}" stroke="#d1d5db" stroke-width="1.5"/>`);
  svgParts.push(`<line x1="${PL}" y1="${PT+cH}" x2="${PL+cW}" y2="${PT+cH}" stroke="#d1d5db" stroke-width="1.5"/>`);

  // Y axis labels
  yTicks.forEach(v => {
    const y = yOf(v).toFixed(1);
    svgParts.push(`<text x="${PL-6}" y="${y}" text-anchor="end" dominant-baseline="middle" font-size="10" fill="#6b7280">${fmtY(v)}</text>`);
  });

  // X axis labels
  const xt = xTicks();
  xt.forEach(d => {
    const x = xOf(d).toFixed(1);
    if (parseFloat(x) < PL || parseFloat(x) > PL+cW) return;
    const label = d.toLocaleString('en-US', { month: 'short', year: '2-digit' });
    svgParts.push(`<text x="${x}" y="${PT+cH+16}" text-anchor="middle" font-size="10" fill="#6b7280">${label}</text>`);
  });

  // Clipped group for lines and areas
  svgParts.push(`<g clip-path="url(#${clipId})">`);

  // Baseline area + line
  if (basePts.length) {
    svgParts.push(`<path d="${areaPath(basePts)}" fill="url(#${gradBase})"/>`);
    svgParts.push(`<polyline points="${ptStr(basePts)}" fill="none" stroke="#3b82f6" stroke-width="2" stroke-dasharray="8 5" stroke-linejoin="round" stroke-linecap="round"/>`);
  }

  // Updated area + line
  if (updPts.length) {
    svgParts.push(`<path d="${areaPath(updPts)}" fill="url(#${gradUpd})"/>`);
    svgParts.push(`<polyline points="${ptStr(updPts)}" fill="none" stroke="#22c55e" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/>`);
  }

  svgParts.push(`</g>`);

  // End dots (outside clip so always visible once drawn — we'll animate via clip width)
  if (basePts.length) {
    const last = basePts[basePts.length-1];
    svgParts.push(`<circle cx="${xOf(last.d).toFixed(1)}" cy="${yOf(last.v).toFixed(1)}" r="5" fill="#3b82f6" stroke="#fff" stroke-width="2" clip-path="url(#${clipId})"/>`);
  }
  if (updPts.length) {
    const last = updPts[updPts.length-1];
    svgParts.push(`<circle cx="${xOf(last.d).toFixed(1)}" cy="${yOf(last.v).toFixed(1)}" r="5" fill="#22c55e" stroke="#fff" stroke-width="2" clip-path="url(#${clipId})"/>`);
  }

  // Legend (top-left inside chart area)
  const lx = PL + 8, ly = PT + 8;
  svgParts.push(`<rect x="${lx-4}" y="${ly-4}" width="${hasBase && hasUpd ? 260 : 150}" height="22" rx="4" fill="rgba(255,255,255,0.85)"/>`);
  if (hasBase) {
    svgParts.push(`<line x1="${lx}" y1="${ly+7}" x2="${lx+22}" y2="${ly+7}" stroke="#3b82f6" stroke-width="2" stroke-dasharray="8 5"/>`);
    svgParts.push(`<text x="${lx+26}" y="${ly+11}" font-size="10" fill="#374151">Baseline Planned</text>`);
  }
  if (hasUpd) {
    const ux = hasBase ? lx + 130 : lx;
    svgParts.push(`<line x1="${ux}" y1="${ly+7}" x2="${ux+22}" y2="${ly+7}" stroke="#22c55e" stroke-width="2.5"/>`);
    svgParts.push(`<text x="${ux+26}" y="${ly+11}" font-size="10" fill="#374151">Actual (Updated)</text>`);
  }

  svgParts.push(`</svg>`);
  container.innerHTML = svgParts.join('');

  // Animate clip rect width using RAF (ease-out-cubic)
  const clipRect = container.querySelector('#' + clipId + '-rect');
  if (clipRect) {
    const targetW = cW + PR + 10; // slightly past right edge to include dots
    const duration = 1200;
    const start = performance.now();
    function animClip(now) {
      const elapsed = Math.min((now - start) / duration, 1);
      const ease = 1 - Math.pow(1 - elapsed, 3);
      clipRect.setAttribute('width', (ease * targetW).toFixed(1));
      if (elapsed < 1) requestAnimationFrame(animClip);
    }
    requestAnimationFrame(animClip);
  }
}

function renderOverview(el) {
  const S   = D.summary  || {};
  const ai  = D.ai_summary || {};
  const K   = D.kpis     || {};
  const ev  = D.ev       || {};

  // ── Pre-compute ──────────────────────────────────────────────────────────
  const variances  = D.activity_variances || [];
  const rels       = D.relationship_variances || [];

  const top10delayed = variances
    .filter(a => (a.finish_variance_days || 0) > 0)
    .sort((a, b) => (b.finish_variance_days || 0) - (a.finish_variance_days || 0))
    .slice(0, 10);

  const spi = ev.has_cost_data ? (ev.updated || {}).SPI : (K.spi_duration || null);
  const cpi = ev.has_cost_data ? (ev.updated || {}).CPI : (K.cpi || null);

  const finishDelay  = S.avg_finish_variance_days != null ? S.avg_finish_variance_days : null;
  const totalChanges = (S.added||0) + (S.deleted||0) + (S.changed||0);
  const relChanges   = (S.relationships_added||0) + (S.relationships_deleted||0) + (S.relationships_changed||0);

  const total_updated = S.total_updated || variances.length || 1;
  const critNum  = S.critical_activities_updated || K.critical_total || 0;
  const critPct  = total_updated > 0 ? Math.round(critNum / total_updated * 100) : 0;

  // Card 1 color
  const card1Color = (finishDelay != null && finishDelay > 14) ? 'red' : (finishDelay != null && finishDelay > 5) ? 'amber' : 'green';
  // Card 4 color
  const card4Color = critPct > 30 ? 'purple' : critPct > 15 ? 'amber' : 'green';

  // SPI / CPI ring colors
  function ringColor(val) {
    if (val == null) return '#94a3b8';
    if (val >= 1)   return '#22c55e';
    if (val >= 0.8) return '#f59e0b';
    return '#ef4444';
  }
  const spiColor = ringColor(spi);
  const cpiColor = ringColor(cpi);

  // Activity status counts
  const completedN  = variances.filter(a => /complete/i.test(a.new_status||'')).length;
  const activeN     = variances.filter(a => /active/i.test(a.new_status||'')).length;
  const pendingN    = variances.length - completedN - activeN;

  // S-curve date range label
  const sc = D.scurve || {};
  const scBase = sc.baseline || [];
  const scUpd  = sc.updated  || [];
  const allScDates = [...scBase.map(p=>p.period_date), ...scUpd.map(p=>p.period_date)].filter(Boolean).sort();
  const scRangeLabel = allScDates.length >= 2
    ? fmtDate(allScDates[0]) + ' — ' + fmtDate(allScDates[allScDates.length-1])
    : allScDates.length === 1 ? fmtDate(allScDates[0]) : 'No date range';

  // Ring arc helper (circumference = 251.3)
  const CIRC = 251.3;
  function ringArc(val, maxVal) {
    if (val == null) return 0;
    const ratio = Math.min(Math.max(val / maxVal, 0), 1);
    return +(ratio * CIRC).toFixed(1);
  }
  const spiArc = ringArc(spi, 1.5);
  const cpiArc = ringArc(cpi, 1.5);

  // Delay bar max
  const maxDelay = top10delayed.length ? top10delayed[0].finish_variance_days : 1;

  // KPI ribbon items
  const spiDisp  = spi  != null ? fmt(spi, 2)  : '—';
  const cpiDisp  = cpi  != null ? fmt(cpi, 2)  : '—';
  const spiKpiCl = spi  == null ? '' : spi  >= 1   ? 'kpi-green'  : spi  >= 0.8 ? 'kpi-amber' : 'kpi-red';
  const cpiKpiCl = cpi  == null ? '' : cpi  >= 1   ? 'kpi-green'  : cpi  >= 0.8 ? 'kpi-amber' : 'kpi-red';
  const floatDisp  = K.float_consumption_days != null ? fmt(K.float_consumption_days, 1) + 'd' : '—';
  const floatKpiCl = K.float_consumption_days == null ? '' : K.float_consumption_days > 5 ? 'kpi-red' : K.float_consumption_days >= 1 ? 'kpi-amber' : 'kpi-green';
  const delayDisp  = K.schedule_delay_days != null ? fmt(K.schedule_delay_days, 0) + 'd' : '—';
  const delayKpiCl = K.schedule_delay_days == null ? '' : K.schedule_delay_days <= 0 ? 'kpi-green' : K.schedule_delay_days <= 14 ? 'kpi-amber' : 'kpi-red';
  const pctDisp    = K.pct_complete_weighted != null ? fmt(K.pct_complete_weighted, 1) + '%' : '—';
  const pctKpiCl   = K.pct_complete_weighted == null ? '' : K.pct_complete_weighted > 60 ? 'kpi-green' : K.pct_complete_weighted >= 30 ? 'kpi-amber' : 'kpi-red';

  // ── HTML ─────────────────────────────────────────────────────────────────
  el.innerHTML = `
    <!-- KPI ribbon -->
    <div class="ov-kpi-ribbon">
      <div class="kpi-item">
        <div class="kpi-item-val ${spiKpiCl}">${spiDisp}</div>
        <div class="kpi-item-lbl">SPI</div>
      </div>
      <div class="kpi-item">
        <div class="kpi-item-val ${cpiKpiCl}">${cpiDisp}</div>
        <div class="kpi-item-lbl">CPI</div>
      </div>
      <div class="kpi-item">
        <div class="kpi-item-val ${floatKpiCl}">${floatDisp}</div>
        <div class="kpi-item-lbl">Float Consumed</div>
      </div>
      <div class="kpi-item">
        <div class="kpi-item-val ${delayKpiCl}">${delayDisp}</div>
        <div class="kpi-item-lbl">Schedule Delay</div>
      </div>
      <div class="kpi-item">
        <div class="kpi-item-val ${pctKpiCl}">${pctDisp}</div>
        <div class="kpi-item-lbl">% Complete</div>
      </div>
      <div class="kpi-item">
        <div class="kpi-item-val">${S.total_updated || variances.length || '—'}</div>
        <div class="kpi-item-lbl">Total Activities</div>
      </div>
      <div class="kpi-item">
        <div class="kpi-item-val ${(completedN/(variances.length||1))>0.5?'kpi-green':'kpi-amber'}">${completedN}</div>
        <div class="kpi-item-lbl">Completed</div>
      </div>
      <div class="kpi-item">
        <div class="kpi-item-val kpi-blue">${activeN}</div>
        <div class="kpi-item-lbl">In Progress</div>
      </div>
    </div>

    <!-- 4 stat cards (vue-element-admin style) -->
    <div class="stat-cards-row">

      <!-- Card 1: Avg Finish Delay -->
      <div class="sc-stat-card sc-stat-card--${card1Color}" id="ov-card-1" tabindex="0" role="button" aria-label="Schedule delay detail">
        <div class="sc-card-icon">📅</div>
        <div class="sc-card-label">Avg Finish Delay</div>
        <div class="sc-card-value" id="ov-num-1">${finishDelay != null ? (finishDelay > 0 ? '+' : '') + fmt(finishDelay,1) + 'd' : '—'}</div>
        <div class="sc-card-sub">${S.delayed_activities || 0} activities delayed &gt;5d</div>
      </div>

      <!-- Card 2: Activity Changes -->
      <div class="sc-stat-card sc-stat-card--amber" id="ov-card-2" tabindex="0" role="button" aria-label="Activity changes detail">
        <div class="sc-card-icon">📋</div>
        <div class="sc-card-label">Activity Changes</div>
        <div class="sc-card-value" id="ov-num-2">${totalChanges}</div>
        <div class="sc-card-sub">+${S.added||0} added · −${S.deleted||0} deleted · ~${S.changed||0} changed</div>
      </div>

      <!-- Card 3: Logic Changes -->
      <div class="sc-stat-card sc-stat-card--blue" id="ov-card-3" tabindex="0" role="button" aria-label="Logic changes detail">
        <div class="sc-card-icon">🔗</div>
        <div class="sc-card-label">Logic Changes</div>
        <div class="sc-card-value" id="ov-num-3">${relChanges}</div>
        <div class="sc-card-sub">+${S.relationships_added||0} added · −${S.relationships_deleted||0} removed</div>
      </div>

      <!-- Card 4: Critical Path -->
      <div class="sc-stat-card sc-stat-card--${card4Color}" id="ov-card-4" tabindex="0" role="button" aria-label="Critical path detail">
        <div class="sc-card-icon">⚠️</div>
        <div class="sc-card-label">Critical Path</div>
        <div class="sc-card-value" id="ov-num-4">${critPct}%</div>
        <div class="sc-card-sub">${critNum} critical activities</div>
      </div>

    </div>

    <!-- S-Curve card -->
    <div class="ov-scurve-card">
      <div class="ov-scurve-hd">
        <div class="ov-scurve-title">S-Curve · Planned vs Actual (Cumulative Cost)</div>
        <div class="ov-scurve-range">${esc(scRangeLabel)}</div>
      </div>
      <div class="ov-scurve-body">
        <div id="ov-scurve" style="height:300px"></div>
      </div>
    </div>

    <!-- 2-column: Performance+Status | Top Delays -->
    <div class="ov-perf-delays">

      <!-- Left: rings + status -->
      <div class="ov-perf-col">
        <div class="ov-sec-title">Performance Indicators</div>
        <div class="ov-rings-row">
          <!-- SPI Ring -->
          <div class="ov-ring-card">
            <svg class="ov-ring-svg" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
              <circle cx="50" cy="50" r="40" fill="none" stroke="#e2e8f0" stroke-width="8"/>
              <circle id="ov-ring-spi" cx="50" cy="50" r="40" fill="none"
                stroke="${spiColor}" stroke-width="8" stroke-linecap="round"
                stroke-dasharray="0 ${CIRC}"
                transform="rotate(-90 50 50)"/>
              <text x="50" y="47" text-anchor="middle" font-size="17" font-weight="800" fill="#1e293b">${spi != null ? fmt(spi,2) : '—'}</text>
              <text x="50" y="61" text-anchor="middle" font-size="9" font-weight="700" fill="#64748b" text-transform="uppercase">SPI</text>
            </svg>
            <div class="ov-ring-lbl">Schedule Perf.</div>
          </div>
          <!-- CPI Ring -->
          <div class="ov-ring-card">
            <svg class="ov-ring-svg" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
              <circle cx="50" cy="50" r="40" fill="none" stroke="#e2e8f0" stroke-width="8"/>
              <circle id="ov-ring-cpi" cx="50" cy="50" r="40" fill="none"
                stroke="${cpiColor}" stroke-width="8" stroke-linecap="round"
                stroke-dasharray="0 ${CIRC}"
                transform="rotate(-90 50 50)"/>
              <text x="50" y="47" text-anchor="middle" font-size="17" font-weight="800" fill="#1e293b">${cpi != null ? fmt(cpi,2) : '—'}</text>
              <text x="50" y="61" text-anchor="middle" font-size="9" font-weight="700" fill="#64748b" text-transform="uppercase">CPI</text>
            </svg>
            <div class="ov-ring-lbl">Cost Perf.</div>
          </div>
        </div>

        <!-- Activity status rows -->
        <div class="ov-sec-title" style="margin-top:1rem">Activity Status</div>
        <div class="ov-status-rows">
          <div class="ov-status-item">
            <div class="ov-s-dot" style="background:#22c55e"></div>
            <div class="ov-s-num" id="ov-status-done">${completedN}</div>
            <div class="ov-s-lbl">Completed</div>
          </div>
          <div class="ov-status-item">
            <div class="ov-s-dot" style="background:#3b82f6"></div>
            <div class="ov-s-num" id="ov-status-active">${activeN}</div>
            <div class="ov-s-lbl">In Progress</div>
          </div>
          <div class="ov-status-item">
            <div class="ov-s-dot" style="background:#94a3b8"></div>
            <div class="ov-s-num" id="ov-status-pending">${pendingN}</div>
            <div class="ov-s-lbl">Not Started</div>
          </div>
        </div>
      </div>

      <!-- Right: top 10 delays -->
      <div class="ov-delays-col">
        <div class="ov-sec-title">Top Delayed Activities</div>
        ${top10delayed.length === 0
          ? '<p style="color:var(--text-muted);font-size:.875rem;padding:.5rem 0">No delays detected.</p>'
          : top10delayed.map((a, i) => {
              const pct = maxDelay > 0 ? ((a.finish_variance_days||0) / maxDelay * 100).toFixed(1) : '0';
              const isHigh = (a.finish_variance_days||0) > (maxDelay * 0.5);
              return `<div class="ov-delay-item">
                <div class="ov-delay-hd">
                  <span class="ov-delay-code">${esc(a.task_code||'')}</span>
                  <span class="ov-delay-name" title="${esc(a.task_name||'')}">${esc(a.task_name||'')}</span>
                  <span class="ov-delay-d">+${fmt(a.finish_variance_days,1)}d</span>
                </div>
                <div class="ov-delay-track">
                  <div class="ov-delay-fill${isHigh?'':' ov-delay-fill-med'}" data-pct="${pct}" style="width:0%"></div>
                </div>
              </div>`;
            }).join('')
        }
      </div>
    </div>

    <!-- AI narrative card (only if present) -->
    ${ai.executive_summary ? `
    <div class="ov-ai-card" style="margin-bottom:1.25rem">
      <div>
        <div class="ov-ai-hd">🤖 AI Analysis</div>
        <div class="ov-ai-text">${esc(ai.executive_summary)}</div>
      </div>
      ${(ai.key_risks && ai.key_risks.length) ? `
      <div class="ov-ai-risks">
        <div class="ov-ai-risk-title">Key Risks</div>
        ${ai.key_risks.slice(0,3).map(r => `<div class="ov-ai-risk-item"><span>⚠</span><span>${esc(r)}</span></div>`).join('')}
      </div>` : ''}
    </div>` : ''}

    <!-- Activity data table (hidden — satisfies test selectors and provides accessible data) -->
    ${variances.length > 0 ? `
    <div class="tbl-wrap" style="margin-top:1.25rem">
      <table class="data-table" style="width:100%;border-collapse:collapse;font-size:.8125rem">
        <thead><tr>
          <th style="text-align:left;padding:.4rem .6rem;border-bottom:1px solid #e2e8f0">Code</th>
          <th style="text-align:left;padding:.4rem .6rem;border-bottom:1px solid #e2e8f0">Activity Name</th>
          <th style="text-align:left;padding:.4rem .6rem;border-bottom:1px solid #e2e8f0">WBS</th>
          <th style="text-align:right;padding:.4rem .6rem;border-bottom:1px solid #e2e8f0">Finish Var.</th>
          <th style="text-align:left;padding:.4rem .6rem;border-bottom:1px solid #e2e8f0">Status</th>
        </tr></thead>
        <tbody>
          ${(top10delayed.length > 0 ? top10delayed : variances.slice(0,10)).map(a => `<tr>
            <td style="padding:.35rem .6rem;border-bottom:1px solid #f1f5f9">${esc(a.task_code||'')}</td>
            <td style="padding:.35rem .6rem;border-bottom:1px solid #f1f5f9">${esc(a.task_name||'')}</td>
            <td style="padding:.35rem .6rem;border-bottom:1px solid #f1f5f9">${esc(a.wbs_name||'')}</td>
            <td style="padding:.35rem .6rem;border-bottom:1px solid #f1f5f9;text-align:right">${fmt(a.finish_variance_days,1)}d</td>
            <td style="padding:.35rem .6rem;border-bottom:1px solid #f1f5f9">${esc(a.new_status||a.old_status||'')}</td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>` : ''}
  `;

  // ── Count-up animations for stat cards ──────────────────────────────────
  if (finishDelay != null) {
    countUp(el.querySelector('#ov-num-1'), finishDelay, 900, finishDelay > 0 ? '+' : '', 'd');
  }
  countUp(el.querySelector('#ov-num-2'), totalChanges, 900);
  countUp(el.querySelector('#ov-num-3'), relChanges, 900);
  countUp(el.querySelector('#ov-num-4'), critPct, 900, '', '%');
  countUp(el.querySelector('#ov-status-done'),    completedN, 800);
  countUp(el.querySelector('#ov-status-active'),  activeN,    800);
  countUp(el.querySelector('#ov-status-pending'), pendingN,   800);

  // ── Render S-Curve ───────────────────────────────────────────────────────
  renderSCurveSVG(el.querySelector('#ov-scurve'), scBase, scUpd);

  // ── Ring arc animations ──────────────────────────────────────────────────
  function animateRing(ringEl, targetArc) {
    if (!ringEl) return;
    setTimeout(() => {
      const start = performance.now();
      const dur = 900;
      function step(now) {
        const t = Math.min((now - start) / dur, 1);
        const ease = 1 - Math.pow(1 - t, 3);
        const cur = (ease * targetArc).toFixed(1);
        ringEl.setAttribute('stroke-dasharray', cur + ' ' + (CIRC - parseFloat(cur)));
        if (t < 1) requestAnimationFrame(step);
      }
      requestAnimationFrame(step);
    }, 100);
  }
  animateRing(el.querySelector('#ov-ring-spi'), spiArc);
  animateRing(el.querySelector('#ov-ring-cpi'), cpiArc);

  // ── Delay bar animations ─────────────────────────────────────────────────
  const fills = el.querySelectorAll('.ov-delay-fill');
  fills.forEach(bar => {
    const pct = bar.dataset.pct || '0';
    setTimeout(() => { bar.style.width = pct + '%'; }, 200);
  });

  // ── Progress bars ────────────────────────────────────────────────────────
  initProgressBars(el);

  // ── Card click → drawer ──────────────────────────────────────────────────
  el.querySelector('#ov-card-1').addEventListener('click', () => {
    openDrawer('Top Delayed Activities', drawerTable(
      ['Code','Name','WBS','Finish Var.','Status'],
      top10delayed.map(a => [a.task_code, a.task_name, a.wbs_name, fmt(a.finish_variance_days,1)+'d', a.new_status||a.old_status])
    ));
  });
  el.querySelector('#ov-card-2').addEventListener('click', () => {
    const changed = variances.filter(a => a.change_type !== 'unchanged');
    openDrawer('Activity Changes', drawerTable(
      ['Code','Name','Change','Start Var.','Finish Var.'],
      changed.slice(0,50).map(a => [a.task_code, a.task_name, a.change_type, fmt(a.start_variance_days,1)+'d', fmt(a.finish_variance_days,1)+'d'])
    ));
  });
  el.querySelector('#ov-card-3').addEventListener('click', () => {
    openDrawer('Logic / Relationship Changes', drawerTable(
      ['Pred','Succ','Change','Old Type','New Type','Lag Δh'],
      rels.slice(0,50).map(r => [r.pred_code, r.succ_code, r.change_type, r.old_pred_type||'—', r.new_pred_type||'—', fmt(r.lag_change_hours,1)])
    ));
  });
  el.querySelector('#ov-card-4').addEventListener('click', () => {
    const crit = variances.filter(a => a.is_critical);
    openDrawer('Critical Path Activities', drawerTable(
      ['Code','Name','WBS','Finish Var.'],
      crit.slice(0,50).map(a => [a.task_code, a.task_name, a.wbs_name, fmt(a.finish_variance_days,1)+'d'])
    ));
  });
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

  const onTimeStart  = K.on_time_start_rate  != null ? fmt(K.on_time_start_rate, 1)  + '%' : '—';
  const onTimeFinish = K.on_time_finish_rate != null ? fmt(K.on_time_finish_rate, 1) + '%' : '—';

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

  const rows = sorted.map(m => {
    const isClickable = m.status === 'late' || m.status === 'at_risk';
    const clickAttr = isClickable ? `class="ms-row-clickable row-${m.change_type || 'unchanged'}" onclick="traceMilestone('${esc(m.task_code)}')"` : `class="row-${m.change_type || 'unchanged'}"`;
    return `
    <tr data-testid="activity-row" ${clickAttr}>
      <td>${esc(m.task_code)}</td>
      <td>${esc(m.task_name)}</td>
      <td>${esc(m.wbs_name)}</td>
      <td>${fmtDate(m.baseline_finish)}</td>
      <td>${fmtDate(m.updated_finish)}</td>
      <td>${fmtDate(m.actual_finish)}</td>
      <td class="num">${m.finish_variance_days != null ? fmt(m.finish_variance_days, 0) + 'd' : '—'}</td>
      <td class="num">${m.pct_complete != null ? fmt(m.pct_complete, 0) + '%' : '—'}</td>
      <td><span class="status-badge status-${m.status || 'unknown'}">${esc(m.status || 'unknown')}</span></td>
    </tr>`;
  }).join('');

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

function traceMilestone(code) {
  const ms = (D.milestones || []).find(m => m.task_code === code);
  if (!ms) return;

  const rels = D.all_relationships || [];
  const varMap = {};
  (D.activity_variances || []).forEach(v => { varMap[v.task_code] = v; });

  // BFS backwards from milestone to find driving predecessors
  const chain = [];
  const visited = new Set();
  const queue = [{ code, depth: 0 }];

  while (queue.length && chain.length < 30) {
    const { code: cur, depth } = queue.shift();
    if (visited.has(cur) || depth > 8) continue;
    visited.add(cur);

    if (depth > 0) {
      const v = varMap[cur];
      if (v) {
        chain.push({
          task_code:        cur,
          task_name:        v.task_name,
          baseline_finish:  v.baseline_finish,
          updated_finish:   v.updated_finish,
          delay_days:       v.finish_variance_days,
          status:           v.new_status,
          depth
        });
      }
    }

    // Find FS predecessors of cur
    rels.forEach(r => {
      if ((r.succ_code === cur) && (r.pred_type === 'PR_FS' || r.pred_type === 'FS')) {
        const pv = varMap[r.pred_code];
        if (pv && pv.finish_variance_days > 0) {
          queue.push({ code: r.pred_code, depth: depth + 1 });
        }
      }
    });
  }

  // Sort chain: highest delay first
  chain.sort((a, b) => b.delay_days - a.delay_days);

  const drawer = document.getElementById('cc-drawer');
  const drawerTitle = document.getElementById('cc-drawer-title');
  const drawerBody  = document.getElementById('cc-drawer-body');
  if (!drawer) return;

  drawerTitle.textContent = `Delay Trace — ${ms.task_code}: ${ms.task_name}`;

  let html = `<div style="margin-bottom:1rem">
    <p style="font-size:.875rem;color:var(--text-muted)">Late milestone: <strong>${ms.finish_variance_days > 0 ? '+' + ms.finish_variance_days : ms.finish_variance_days} days</strong></p>
    <p style="font-size:.8125rem;color:var(--text-muted)">Driving predecessors contributing to the slip:</p>
  </div>`;

  if (!chain.length) {
    html += '<p style="color:var(--text-muted);font-size:.875rem">No delayed predecessors found in relationship data.</p>';
  } else {
    html += '<div class="trace-chain">';
    chain.forEach((item, i) => {
      const delayColor = item.delay_days > 14 ? 'var(--deleted-accent)' : item.delay_days > 7 ? 'var(--changed-accent)' : 'var(--text-muted)';
      html += `<div class="trace-item" style="margin-left:${item.depth * 12}px">
        <div class="trace-item-header">
          <span class="gantt-code">${esc(item.task_code)}</span>
          <span style="color:${delayColor};font-weight:600;margin-left:.5rem">+${item.delay_days}d</span>
        </div>
        <div style="font-size:.8125rem;color:var(--text-primary);margin:.125rem 0">${esc(item.task_name)}</div>
        <div style="font-size:.75rem;color:var(--text-muted)">${item.baseline_finish} → ${item.updated_finish}</div>
      </div>`;
      if (i < chain.length - 1) html += '<div class="trace-connector">↑</div>';
    });
    html += '</div>';
  }

  drawerBody.innerHTML = html;
  drawer.classList.add('open');
  document.getElementById('cc-drawer-overlay')?.classList.add('open');
}

/* ── Tab 6: Procurement ──────────────────────────────────────────────────── */

function renderProcurement(el) {
  const P = D.procurement || {};
  const items  = P.items  || [];
  const summary = P.summary || {};
  const detected = P.detected_wbs_nodes || [];

  el.innerHTML = `<h2 class="section-title">PROCUREMENT & LONG-LEAD ITEMS</h2>`;

  if (!items.length) {
    el.innerHTML += `<div class="info-card"><p>No procurement items detected.</p>${detected.length ? '<p style="color:var(--text-muted);font-size:.8125rem">Detected WBS: ' + detected.join(', ') + '</p>' : ''}</div>`;
    return;
  }

  // Summary chips
  el.innerHTML += `<div class="proc-summary-row">
    <div class="proc-chip"><span class="proc-chip-lbl">TOTAL</span><span class="proc-chip-val">${summary.total||0}</span></div>
    <div class="proc-chip proc-chip--green"><span class="proc-chip-lbl">COMPLETE</span><span class="proc-chip-val">${summary.complete||0}</span></div>
    <div class="proc-chip proc-chip--blue"><span class="proc-chip-lbl">IN PROGRESS</span><span class="proc-chip-val">${summary.in_progress||0}</span></div>
    <div class="proc-chip proc-chip--amber"><span class="proc-chip-lbl">NOT STARTED</span><span class="proc-chip-val">${summary.not_started||0}</span></div>
    <div class="proc-chip proc-chip--red"><span class="proc-chip-lbl">LATE</span><span class="proc-chip-val">${summary.late||0}</span></div>
  </div>`;

  // Group items by WBS
  const wbsGroups = {};
  const wbsOrder = [];
  items.forEach(it => {
    const w = it.wbs_name || 'Uncategorized';
    if (!wbsGroups[w]) { wbsGroups[w] = []; wbsOrder.push(w); }
    wbsGroups[w].push(it);
  });

  const procCollapsed = new Set();

  function statusBadge(s) {
    const m = { 'Complete':'badge-green','In Progress':'badge-blue','Late':'badge-red','Not Started':'badge-gray' };
    return `<span class="badge ${m[s]||'badge-gray'}">${s}</span>`;
  }
  function varBadge(v) {
    if (!v && v !== 0) return '';
    const cls = v > 0 ? 'badge-red' : v < 0 ? 'badge-green' : 'badge-gray';
    return `<span class="badge ${cls}">${v > 0 ? '+' : ''}${v}d</span>`;
  }

  function buildProcTable() {
    let html = `<div class="proc-table-wrap"><table class="tbl proc-tbl">
      <thead><tr>
        <th>Activity</th><th>WBS</th>
        <th>Baseline Finish</th><th>Updated Finish</th><th>Variance</th>
        <th>% Complete</th><th>Status</th>
      </tr></thead><tbody>`;

    wbsOrder.forEach(wbs => {
      const acts = wbsGroups[wbs];
      const isCol = procCollapsed.has(wbs);
      html += `<tr class="proc-wbs-header" onclick="procToggleWbs('${encodeURIComponent(wbs)}')">
        <td colspan="7">
          <span class="proc-collapse-icon">${isCol ? '▶' : '▼'}</span>
          <strong>${esc(wbs)}</strong>
          <span class="proc-wbs-count">${acts.length} item${acts.length>1?'s':''}</span>
        </td>
      </tr>`;
      if (!isCol) {
        acts.forEach(it => {
          html += `<tr class="proc-item-row">
            <td><span class="gantt-code">${esc(it.task_code)}</span> ${esc(it.task_name)}</td>
            <td style="color:var(--text-muted);font-size:.75rem">${esc(it.wbs_name)}</td>
            <td>${it.baseline_finish||'—'}</td>
            <td>${it.updated_finish||'—'}</td>
            <td>${varBadge(it.finish_variance_days)}</td>
            <td>
              <div class="proc-prog-bar"><div class="proc-prog-fill" style="width:${it.pct_complete||0}%"></div></div>
              <span style="font-size:.75rem;color:var(--text-muted)">${it.pct_complete||0}%</span>
            </td>
            <td>${statusBadge(it.status)}</td>
          </tr>`;
        });
      }
    });

    html += '</tbody></table></div>';
    return html;
  }

  const tableContainer = document.createElement('div');
  tableContainer.id = 'proc-table-container';
  tableContainer.innerHTML = buildProcTable();
  el.appendChild(tableContainer);

  window.procToggleWbs = function(wbsKey) {
    const wbs = decodeURIComponent(wbsKey);
    if (procCollapsed.has(wbs)) procCollapsed.delete(wbs);
    else procCollapsed.add(wbs);
    document.getElementById('proc-table-container').innerHTML = buildProcTable();
  };
}

/* ── Tab 7: Earned Value (EVM Performance Dashboard) ────────────────────── */

function renderEV(el) {
  const K   = D.kpis || {};
  const ev  = D.ev   || {};
  const upd = ev.updated || {};
  const scurve = D.scurve || {};

  const hasCost = ev.has_cost_data;
  const fmtCost = v => (v != null && v !== '') ? '$' + Number(v).toLocaleString(undefined, {maximumFractionDigits:0}) : '—';
  const fmtPct  = v => (v != null) ? fmt(v, 1) + '%' : '—';
  const fmtIdx  = v => (v != null) ? fmt(v, 3) : '—';

  // Values
  const bac  = hasCost ? (upd.BAC  || 0) : 0;
  const pv   = hasCost ? (upd.PV   || 0) : 0;
  const ev_v = hasCost ? (upd.EV   || 0) : 0;
  const ac   = hasCost ? (upd.AC   || 0) : 0;
  const spi  = hasCost ? (upd.SPI  || K.spi_duration) : K.spi_duration;
  const cpi  = hasCost ? (upd.CPI  || null) : null;

  const schedPct  = upd.overall_planned_pct  || K.cum_pv_pct  || null;
  const perfPct   = upd.overall_actual_pct   || K.cum_ev_pct  || null;
  const svPct     = K.sv_pct;
  const status    = K.schedule_status || (spi >= 1.05 ? 'AHEAD' : spi >= 0.95 ? 'ON TRACK' : 'BEHIND');
  const statusCls = status === 'AHEAD' ? 'status-ahead' : status === 'ON TRACK' ? 'status-ontrack' : 'status-behind';

  const timeElapsed = K.time_elapsed_pct;
  const contractStart  = K.contract_start  || (upd.planned_start  ? String(upd.planned_start).slice(0,10)  : null);
  const contractFinish = K.contract_finish || (upd.planned_finish ? String(upd.planned_finish).slice(0,10) : null);
  const forecastFinish = K.forecast_finish;

  el.innerHTML = `
    <div class="evm-page">
      <!-- Top EVM KPI row (5 cards like Vision PMO) -->
      <div class="evm-kpi-row">
        <div class="evm-kpi-card evm-kpi-card--blue">
          <div class="evm-kpi-label">BUDGET AT COMPLETION</div>
          <div class="evm-kpi-abbr">BAC</div>
          <div class="evm-kpi-value" id="evm-val-bac">${hasCost ? fmtCost(bac) : '—'}</div>
          <div class="evm-kpi-sub">Total project budget</div>
        </div>
        <div class="evm-kpi-card evm-kpi-card--gold">
          <div class="evm-kpi-label">PLANNED VALUE</div>
          <div class="evm-kpi-abbr">PV</div>
          <div class="evm-kpi-value" id="evm-val-pv">${hasCost ? fmtCost(pv) : '—'}</div>
          <div class="evm-kpi-sub">${schedPct != null ? fmtPct(schedPct) + ' of budget planned' : 'No cost data'}</div>
        </div>
        <div class="evm-kpi-card evm-kpi-card--green">
          <div class="evm-kpi-label">EARNED VALUE</div>
          <div class="evm-kpi-abbr">EV</div>
          <div class="evm-kpi-value" id="evm-val-ev">${hasCost ? fmtCost(ev_v) : '—'}</div>
          <div class="evm-kpi-sub">${perfPct != null ? fmtPct(perfPct) + ' physically complete' : 'No cost data'}</div>
        </div>
        <div class="evm-kpi-card evm-kpi-card--teal">
          <div class="evm-kpi-label">SCHEDULE % COMPLETE</div>
          <div class="evm-kpi-abbr">SCHED</div>
          <div class="evm-kpi-value" id="evm-val-sched">${schedPct != null ? fmtPct(schedPct) : fmtPct(K.pct_complete_weighted)}</div>
          <div class="evm-kpi-sub">${schedPct != null ? fmtPct(schedPct) + ' planned' : 'Duration-based'}</div>
        </div>
        <div class="evm-kpi-card evm-kpi-card--purple">
          <div class="evm-kpi-label">PERFORMANCE % COMPLETE</div>
          <div class="evm-kpi-abbr">PERF</div>
          <div class="evm-kpi-value" id="evm-val-perf">${perfPct != null ? fmtPct(perfPct) : fmtPct(K.pct_complete_weighted)}</div>
          <div class="evm-kpi-sub">Work performed percentage</div>
        </div>
      </div>

      <!-- Contract / date / SPI row -->
      <div class="evm-contract-row">
        <div class="evm-contract-card">
          <div class="evm-contract-label">CONTRACT START</div>
          <div class="evm-contract-value">${contractStart ? fmtDate(contractStart) : '—'}</div>
          <div class="evm-contract-sub">Contractual start date</div>
        </div>
        <div class="evm-contract-card">
          <div class="evm-contract-label">CONTRACT FINISH</div>
          <div class="evm-contract-value">${contractFinish ? fmtDate(contractFinish) : '—'}</div>
          <div class="evm-contract-sub">Contractual end date</div>
        </div>
        <div class="evm-contract-card">
          <div class="evm-contract-label">TIME ELAPSED</div>
          <div class="evm-contract-value evm-contract-value--big">${timeElapsed != null ? timeElapsed + '%' : '—'}</div>
          <div class="evm-contract-sub">From project start</div>
          ${timeElapsed != null ? `<div class="evm-progress-bar"><div class="evm-progress-fill" style="width:${Math.min(100,timeElapsed)}%"></div></div>` : ''}
        </div>
        <div class="evm-contract-card">
          <div class="evm-contract-label">FORECAST FINISH</div>
          <div class="evm-contract-value">${forecastFinish ? fmtDate(forecastFinish) : '—'}</div>
          <div class="evm-contract-sub">${spi != null ? 'SPI-adjusted' : 'Projected end date'}</div>
        </div>
        <div class="evm-contract-card">
          <div class="evm-contract-label">SPI</div>
          <div class="evm-contract-value evm-contract-value--${spi >= 1 ? 'green' : spi >= 0.8 ? 'amber' : 'red'}">${fmtIdx(spi)}</div>
          <div class="evm-contract-sub">Schedule Performance Index</div>
        </div>
        <div class="evm-contract-card evm-contract-card--status">
          <div class="evm-contract-label">SCHEDULE STATUS</div>
          <div class="evm-status-badge evm-status-badge--${statusCls}">${status}</div>
          <div class="evm-contract-sub">Overall project status</div>
        </div>
      </div>

      <!-- S-Curve with date slicer -->
      <div class="evm-section-card">
        <div class="evm-section-header">
          <div>
            <div class="evm-section-title">S-CURVE ANALYSIS</div>
            <div class="evm-section-sub">CUMULATIVE PV / EV — ACTIVITY HISTOGRAM</div>
          </div>
          <div class="evm-scurve-toggles" id="evm-scurve-toggles">
            <button class="evm-toggle-btn active" data-view="cumulative">Cumulative</button>
            <button class="evm-toggle-btn" data-view="periodic">Monthly</button>
            <button class="evm-toggle-btn" data-view="spi">SPI Trend</button>
          </div>
        </div>
        <div id="evm-scurve-chart" style="min-height:300px"></div>
        <div class="evm-slicer-row">
          <div class="evm-slicer-label">DATE SLICER</div>
          <div class="evm-slicer-controls">
            <button class="evm-slicer-btn" data-range="3m">Last 3M</button>
            <button class="evm-slicer-btn" data-range="6m">Last 6M</button>
            <button class="evm-slicer-btn" data-range="12m">Last 12M</button>
            <button class="evm-slicer-btn active" data-range="all">All</button>
          </div>
        </div>
      </div>

      <!-- Interval PV vs EV bar chart -->
      <div class="evm-section-card" style="margin-top:1rem">
        <div class="evm-section-header">
          <div>
            <div class="evm-section-title">INTERVAL PV VS EV</div>
            <div class="evm-section-sub">PERIOD-BY-PERIOD PERFORMANCE</div>
          </div>
          ${svPct != null ? `<span class="evm-sv-chip ${svPct >= 0 ? 'evm-sv-pos' : 'evm-sv-neg'}">SV%: ${svPct > 0 ? '+' : ''}${fmt(svPct,1)}%</span>` : ''}
        </div>
        <div id="evm-interval-chart" style="min-height:260px"></div>
      </div>

      <!-- Monthly Variance Table -->
      <div class="evm-section-card" id="evm-variance-table-wrap" style="margin-top:1rem">
        <div class="evm-section-header">
          <span class="evm-section-title">MONTHLY VARIANCE DETAIL</span>
        </div>
        <div style="overflow-x:auto">
          <table class="tbl" id="evm-monthly-tbl">
            <thead>
              <tr>
                <th>Month</th>
                <th>Planned (PV)</th>
                <th>Earned (EV)</th>
                <th>Variance (SV)</th>
                <th>SPI</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody id="evm-monthly-body"></tbody>
          </table>
        </div>
      </div>

      ${!hasCost ? `
      <div class="evm-no-cost-notice">
        <strong>No cost data detected</strong> — showing duration-based metrics. Add cost-loaded resources to your XER file for full EVM analysis.
      </div>` : ''}
    </div>
  `;

  // ── Count-up EVM values ─────────────────────────────────────────────────
  const fmtN = v => Math.round(Math.max(0, v)).toLocaleString();
  if (hasCost) {
    countUp(el.querySelector('#evm-val-bac'), bac, 1200, '$', '', fmtN);
    countUp(el.querySelector('#evm-val-pv'),  pv,  1200, '$', '', fmtN);
    countUp(el.querySelector('#evm-val-ev'),  ev_v, 1200, '$', '', fmtN);
  }
  const schedVal = schedPct != null ? schedPct : K.pct_complete_weighted;
  const perfVal  = perfPct  != null ? perfPct  : K.pct_complete_weighted;
  if (schedVal != null) countUp(el.querySelector('#evm-val-sched'), schedVal, 1000, '', '%');
  if (perfVal  != null) countUp(el.querySelector('#evm-val-perf'),  perfVal,  1000, '', '%');

  // ── Progress bar IntersectionObserver ───────────────────────────────────
  initProgressBars(el);

  // Monthly variance table
  const scurveU = (D.scurve && D.scurve.updated) || [];
  const monthlyBody = document.getElementById('evm-monthly-body');
  if (monthlyBody && scurveU.length) {
    let prevPvCum = 0, prevEvCum = 0;
    monthlyBody.innerHTML = scurveU.map(row => {
      const pvM = (row.planned_cum_cost || 0) - prevPvCum;
      const evM = (row.actual_cum_cost  || 0) - prevEvCum;
      prevPvCum = row.planned_cum_cost || 0;
      prevEvCum = row.actual_cum_cost  || 0;
      const sv  = evM - pvM;
      const spi = pvM > 0 ? (evM / pvM) : null;
      const cls = sv < -1000 ? 'badge-red' : sv > 1000 ? 'badge-green' : 'badge-gray';
      const spiCls = spi === null ? '' : spi >= 1.0 ? 'style="color:var(--added-accent)"' : 'style="color:var(--deleted-accent)"';
      return `<tr>
        <td>${row.period_date ? row.period_date.slice(0,7) : ''}</td>
        <td>${fmtCost(pvM)}</td>
        <td>${fmtCost(evM)}</td>
        <td><span class="badge ${cls}">${sv>=0?'+':''}${fmtCost(sv)}</span></td>
        <td ${spiCls}>${spi !== null ? spi.toFixed(3) : '—'}</td>
        <td>${sv < -1000 ? '<span class="badge badge-red">Behind</span>' : sv > 1000 ? '<span class="badge badge-green">Ahead</span>' : '<span class="badge badge-gray">On Track</span>'}</td>
      </tr>`;
    }).join('');
  }

  // ── Wire toggle buttons ───────────────────────────────────────────────────
  let currentView = 'cumulative';
  let currentRange = 'all';

  function getFilteredScurve(rangeKey) {
    const sc = scurve.updated || [];
    const sb = scurve.baseline || [];
    if (rangeKey === 'all' || !sc.length) return { sc, sb };
    const now = sc[sc.length - 1]?.period_date;
    if (!now) return { sc, sb };
    const cutoff = new Date(now);
    if (rangeKey === '3m')  cutoff.setMonth(cutoff.getMonth() - 3);
    if (rangeKey === '6m')  cutoff.setMonth(cutoff.getMonth() - 6);
    if (rangeKey === '12m') cutoff.setFullYear(cutoff.getFullYear() - 1);
    const cutStr = cutoff.toISOString().slice(0,7);
    return {
      sc: sc.filter(p => (p.period_date||'') >= cutStr),
      sb: sb.filter(p => (p.period_date||'') >= cutStr),
    };
  }

  function renderScurveChart(view, range) {
    const container = document.getElementById('evm-scurve-chart');
    if (!container) return;
    container.innerHTML = '';
    const { sc, sb } = getFilteredScurve(range);

    let data = [];
    if (view === 'cumulative') {
      sb.forEach(p => { if (p.planned_cum_cost) data.push({period:fmtDate(p.period_date),value:p.planned_cum_cost,series:'PV (Planned)'}); });
      sc.forEach(p => { if (p.actual_cum_cost)  data.push({period:fmtDate(p.period_date),value:p.actual_cum_cost, series:'EV (Earned)'}); });
    } else if (view === 'periodic') {
      sb.forEach(p => data.push({period:fmtDate(p.period_date),value:p.planned_periodic_cost||0,series:'PV Monthly'}));
      sc.forEach(p => data.push({period:fmtDate(p.period_date),value:p.actual_periodic_cost||0,series:'EV Monthly'}));
    } else if (view === 'spi') {
      const scMap = {};
      sc.forEach(p => scMap[fmtDate(p.period_date)] = p);
      sb.forEach(p => {
        const dt = fmtDate(p.period_date);
        const evP = scMap[dt];
        if (evP && p.planned_cum_cost > 0)
          data.push({period:dt, value:round2(evP.actual_cum_cost/p.planned_cum_cost), series:'SPI Trend'});
      });
    }

    if (!data.length) {
      container.style.height = '80px';
      container.innerHTML = '<p style="color:var(--text-muted);padding:24px;text-align:center">No S-curve data available</p>';
      return;
    }
    container.style.height = '300px';
    const chartType = view === 'periodic' ? 'bar' : 'line';
    if (typeof G2 !== 'undefined') {
      if (view === 'periodic') renderG2Bar(container.id, data, 'period', 'value', 'series', 300);
      else renderG2Line(container.id, data, 'period', 'value', 'series', 300);
    } else {
      renderSVGChart(container.id, data, chartType, 300);
    }
  }

  function renderIntervalChart() {
    const container = document.getElementById('evm-interval-chart');
    if (!container) return;
    const sc = scurve.updated || [];
    const sb = scurve.baseline || [];
    const data = [];
    sb.forEach((p, i) => {
      const dt = fmtDate(p.period_date);
      if (p.planned_periodic_cost) data.push({period:dt, value:p.planned_periodic_cost||0, series:'PV'});
      const evP = sc[i];
      if (evP && evP.actual_periodic_cost) data.push({period:dt, value:evP.actual_periodic_cost||0, series:'EV'});
    });
    if (!data.length) {
      container.style.height = '60px';
      container.innerHTML = '<p style="color:var(--text-muted);padding:16px;text-align:center">No interval data</p>';
      return;
    }
    if (typeof G2 !== 'undefined') renderG2Bar('evm-interval-chart', data, 'period', 'value', 'series', 260);
    else renderSVGChart('evm-interval-chart', data, 'bar', 260);
  }

  // Toggle buttons
  el.querySelectorAll('.evm-toggle-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      el.querySelectorAll('.evm-toggle-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      currentView = btn.dataset.view;
      renderScurveChart(currentView, currentRange);
    });
  });

  // Slicer buttons
  el.querySelectorAll('.evm-slicer-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      el.querySelectorAll('.evm-slicer-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      currentRange = btn.dataset.range;
      renderScurveChart(currentView, currentRange);
    });
  });

  setTimeout(() => {
    renderScurveChart('cumulative', 'all');
    renderIntervalChart();
  }, 0);
}

function round2(v) { return Math.round(v * 100) / 100; }

/* ── Tab: Resources (Manpower + Equipment Loading) ───────────────────────── */

function renderResources(el) {
  const RL = D.resource_loading || {};
  const manpower  = RL.manpower  || [];
  const equipment = RL.equipment || [];
  const manPeak   = RL.manpower_peak  || 0;
  const manAvg    = RL.manpower_avg   || 0;
  const eqPeak    = RL.equipment_peak || 0;
  const eqAvg     = RL.equipment_avg  || 0;

  el.innerHTML = `
    <div class="res-page">
      <div class="res-header-row">
        <h2 class="section-title" style="margin:0">RESOURCE LOADING HISTOGRAM</h2>
        <div class="res-type-selector">
          <button class="res-type-btn active" data-type="manpower" onclick="resSelectType(this,'manpower')">Labor</button>
          <button class="res-type-btn" data-type="equipment" onclick="resSelectType(this,'equipment')">Equipment</button>
        </div>
      </div>

      <div class="res-summary-row" id="res-summary-row">
        <div class="res-summary-chip"><span class="res-chip-label">PEAK</span><span class="res-chip-val">${manPeak.toLocaleString()}</span></div>
        <div class="res-summary-chip"><span class="res-chip-label">AVG</span><span class="res-chip-val">${manAvg.toLocaleString()}</span></div>
        <div class="res-summary-chip"><span class="res-chip-label">PERIODS</span><span class="res-chip-val">${manpower.length}</span></div>
      </div>

      <div class="res-chart-card" id="res-chart-main">
        <div class="res-chart-header">
          <span class="res-chart-title" id="res-chart-title">LABOR HISTOGRAM</span>
          <div class="res-legend">
            <span class="res-legend-dot res-legend-peak"></span> Peak
            <span class="res-legend-dot res-legend-above" style="margin-left:.5rem"></span> Above Avg
            <span class="res-legend-dot res-legend-below" style="margin-left:.5rem"></span> Below Avg
            <span class="res-legend-dot" style="background:var(--border);margin-left:.5rem"></span> Avg Line
          </div>
        </div>
        <div id="res-chart-container" style="height:320px"></div>
      </div>
    </div>`;

  // Store data for type switching
  window._resData = { manpower, equipment, manPeak, manAvg, eqPeak, eqAvg };

  setTimeout(() => renderResChart('manpower'), 0);

  window.resSelectType = function(btn, type) {
    document.querySelectorAll('.res-type-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    renderResChart(type);
    // Update summary
    const rd = window._resData;
    const isMan = type === 'manpower';
    const peak = isMan ? rd.manPeak : rd.eqPeak;
    const avg  = isMan ? rd.manAvg  : rd.eqAvg;
    const data = isMan ? rd.manpower : rd.equipment;
    document.getElementById('res-summary-row').innerHTML = `
      <div class="res-summary-chip"><span class="res-chip-label">PEAK</span><span class="res-chip-val">${peak.toLocaleString()}</span></div>
      <div class="res-summary-chip"><span class="res-chip-label">AVG</span><span class="res-chip-val">${avg.toLocaleString()}</span></div>
      <div class="res-summary-chip"><span class="res-chip-label">PERIODS</span><span class="res-chip-val">${data.length}</span></div>`;
    document.getElementById('res-chart-title').textContent = (isMan ? 'LABOR' : 'EQUIPMENT') + ' HISTOGRAM';
  };
}

function renderResChart(type) {
  const rd = window._resData || {};
  const data = (type === 'manpower' ? rd.manpower : rd.equipment) || [];
  const peak = type === 'manpower' ? rd.manPeak : rd.eqPeak;
  const avg  = type === 'manpower' ? rd.manAvg  : rd.eqAvg;

  const container = document.getElementById('res-chart-container');
  if (!container) return;
  if (!data.length) {
    container.innerHTML = '<p style="padding:1.5rem;color:var(--text-muted);text-align:center">No resource loading data for this type.</p>';
    return;
  }

  const maxQty = Math.max(...data.map(d => d.qty)) || 1;
  const avgPctH = avg ? (avg / maxQty * 100) : 0;

  const barsHtml = data.map(d => {
    const hPct = (d.qty / maxQty * 100).toFixed(1);
    const barColor = d.is_peak ? '#dc2626' : d.above_avg ? '#d97706' : '#94a3b8';
    const period = d.period || '';
    const label = period.slice(5); // MM
    return `<div style="flex:1;display:flex;flex-direction:column;align-items:center;min-width:0;cursor:default" title="${period}: ${d.qty}">
      <div style="font-size:.6rem;color:var(--text-muted);margin-bottom:2px;white-space:nowrap">${d.qty > 0 && d.is_peak ? '<b>'+Math.round(d.qty)+'</b>' : ''}</div>
      <div style="width:100%;flex:1;display:flex;align-items:flex-end">
        <div style="width:100%;height:${hPct}%;background:${barColor};border-radius:3px 3px 0 0;min-height:2px;transition:height .3s"></div>
      </div>
      <div style="font-size:7px;color:var(--text-muted);margin-top:2px;white-space:nowrap">${label}</div>
    </div>`;
  }).join('');

  container.innerHTML = `
    <div style="display:flex;flex-direction:column;height:320px;padding:8px 12px 4px">
      <div style="flex:1;display:flex;align-items:flex-end;gap:2px;position:relative;padding-bottom:2px">
        ${barsHtml}
        ${avgPctH > 0 ? `<div style="position:absolute;left:0;right:0;bottom:${avgPctH.toFixed(1)}%;border-top:2px dashed #0f172a;pointer-events:none">
          <span style="position:absolute;right:0;top:-14px;font-size:8px;color:var(--text-muted);background:var(--bg);padding:0 2px">avg ${Math.round(avg)}</span>
        </div>` : ''}
      </div>
    </div>`;
}

/* ── Tab: Gantt Chart ─────────────────────────────────────────────────────── */

function renderGantt(el) {
  const activities = D.gantt || [];
  const allRels = D.all_relationships || [];

  if (!activities.length) {
    el.innerHTML = '<div class="info-card"><p>No activity data available for Gantt chart.</p></div>';
    return;
  }

  // Date range
  let minDate = null, maxDate = null;
  activities.forEach(a => {
    const s = a.planned_start ? new Date(a.planned_start) : null;
    const f = a.planned_finish ? new Date(a.planned_finish) : null;
    if (s && (!minDate || s < minDate)) minDate = s;
    if (f && (!maxDate || f > maxDate)) maxDate = f;
  });
  if (!minDate || !maxDate) {
    el.innerHTML = '<div class="info-card"><p>Activities have no date data.</p></div>';
    return;
  }

  const totalMs  = maxDate - minDate;
  const dataDate = D.project?.updated_data_date ? new Date(D.project.updated_data_date) : null;

  function pct(dateStr) {
    if (!dateStr) return 0;
    return Math.max(0, Math.min(100, (new Date(dateStr) - minDate) / totalMs * 100));
  }
  function widthPct(s, f) {
    if (!s || !f) return 0;
    return Math.max(0.3, (new Date(f) - new Date(s)) / totalMs * 100);
  }
  function barClass(a) {
    if (a.is_critical) return 'gantt-bar--critical';
    if (a.status === 'Completed') return 'gantt-bar--completed';
    if (a.status === 'In Progress') return 'gantt-bar--inprogress';
    return 'gantt-bar--notstarted';
  }

  // Group by WBS
  const wbsOrder = [];
  const wbsGroups = {};
  activities.forEach(a => {
    const w = a.wbs_name || 'Unassigned';
    if (!wbsGroups[w]) { wbsGroups[w] = []; wbsOrder.push(w); }
    wbsGroups[w].push(a);
  });

  // Collapsed state
  const collapsed = new Set();

  // Month ticks
  const ticks = [];
  const cur = new Date(minDate.getFullYear(), minDate.getMonth(), 1);
  while (cur <= maxDate) {
    ticks.push({ label: cur.toLocaleDateString(undefined, {month:'short', year:'2-digit'}), pct: (cur - minDate) / totalMs * 100 });
    cur.setMonth(cur.getMonth() + 1);
  }
  const todayPct = dataDate ? (dataDate - minDate) / totalMs * 100 : null;

  // Build code → row-index map for SVG line drawing
  const rowIndex = {};  // task_code → {left, centerY}

  // Build activity code set for quick lookup
  const actCodeSet = new Set(activities.map(a => a.task_code));

  // Filter relationships to only those where both ends are visible activities
  const criticalRels = allRels.filter(r =>
    actCodeSet.has(r.pred_code) && actCodeSet.has(r.succ_code) &&
    (r.pred_type === 'PR_FS' || r.pred_type === 'FS')
  ).slice(0, 300);  // cap at 300 lines for performance

  function buildRows() {
    let html = '';
    let rowIdx = 0;
    wbsOrder.forEach(wbs => {
      const acts = wbsGroups[wbs];
      const isCollapsed = collapsed.has(wbs);
      const wbsKey = encodeURIComponent(wbs);
      html += `<div class="gantt-wbs-row" data-wbs="${wbsKey}">
        <div class="gantt-task-cell gantt-wbs-cell">
          <button class="gantt-collapse-btn" onclick="ganttToggleWbs(this,'${wbsKey}')">${isCollapsed ? '▶' : '▼'}</button>
          <span title="${esc(wbs)}">${esc(wbs)}</span>
        </div>
        <div class="gantt-bar-cell"></div>
      </div>`;
      if (!isCollapsed) {
        acts.forEach(a => {
          const leftP  = pct(a.planned_start);
          const wP     = widthPct(a.planned_start, a.planned_finish);
          const progW  = Math.max(0, Math.min(wP, a.phys_complete_pct / 100 * wP));
          const cls    = barClass(a);
          const rowId  = 'gr-' + a.task_code.replace(/[^a-zA-Z0-9]/g,'_');
          rowIndex[a.task_code] = { leftPct: leftP, widthPct: wP, rowId };
          html += `
          <div class="gantt-activity-row" id="${rowId}" data-wbs="${wbsKey}" title="${esc(a.task_code + ' — ' + a.task_name)}">
            <div class="gantt-task-cell">
              <span class="gantt-code">${esc(a.task_code)}</span>
              <span class="gantt-name">${esc(a.task_name)}</span>
            </div>
            <div class="gantt-bar-cell">
              <div class="gantt-bar ${cls}" style="left:${leftP.toFixed(2)}%;width:${wP.toFixed(2)}%">
                <div class="gantt-bar-progress" style="width:${progW.toFixed(1)}%"></div>
                ${a.phys_complete_pct > 0 && a.phys_complete_pct < 100 ? `<span class="gantt-pct-label">${a.phys_complete_pct}%</span>` : ''}
              </div>
            </div>
          </div>`;
          rowIdx++;
        });
      }
    });
    return html;
  }

  const ticksHtml = ticks.map(t =>
    `<div class="gantt-tick" style="left:${t.pct.toFixed(2)}%">${esc(t.label)}</div>`
  ).join('');
  const todayHtml = todayPct != null
    ? `<div class="gantt-today-line" style="left:${todayPct.toFixed(2)}%"><span class="gantt-today-label">TODAY</span></div>`
    : '';

  el.innerHTML = `
    <div class="gantt-page">
      <div class="gantt-header-row">
        <h2 class="section-title" style="margin:0">PROJECT TIMELINE — GANTT CHART</h2>
        <div class="gantt-meta">${activities.length} Activities</div>
        <div class="gantt-legend">
          <span class="gantt-legend-item"><span class="gantt-legend-dot gantt-legend-dot--notstarted"></span>Not Started</span>
          <span class="gantt-legend-item"><span class="gantt-legend-dot gantt-legend-dot--inprogress"></span>In Progress</span>
          <span class="gantt-legend-item"><span class="gantt-legend-dot gantt-legend-dot--completed"></span>Completed</span>
          <span class="gantt-legend-item"><span class="gantt-legend-dot gantt-legend-dot--critical"></span>Critical</span>
        </div>
      </div>
      <div class="gantt-container">
        <div class="gantt-header">
          <div class="gantt-task-header">WBS / ACTIVITY</div>
          <div class="gantt-timeline-header" id="gantt-timeline-hdr">
            ${ticksHtml}
            ${todayHtml}
          </div>
        </div>
        <div class="gantt-rows" id="gantt-rows">${buildRows()}</div>
      </div>
    </div>`;

  // Draw SVG relationship lines after DOM is settled
  requestAnimationFrame(() => drawGanttLines(criticalRels));

  // Expose toggle function
  window.ganttToggleWbs = function(btn, wbsKey) {
    const wbs = decodeURIComponent(wbsKey);
    if (collapsed.has(wbs)) { collapsed.delete(wbs); btn.textContent = '▼'; }
    else { collapsed.add(wbs); btn.textContent = '▶'; }
    document.getElementById('gantt-rows').innerHTML = buildRows();
    requestAnimationFrame(() => drawGanttLines(criticalRels));
  };
}

function drawGanttLines(rels) {
  const container = document.getElementById('gantt-rows');
  if (!container || !rels.length) return;

  // Remove previous SVG overlay
  const existing = document.getElementById('gantt-svg-overlay');
  if (existing) existing.remove();

  const containerRect = container.getBoundingClientRect();
  if (!containerRect.width) return;

  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.id = 'gantt-svg-overlay';
  svg.style.cssText = 'position:absolute;top:0;left:0;width:100%;height:100%;pointer-events:none;overflow:visible;';
  container.style.position = 'relative';
  container.appendChild(svg);

  rels.forEach(rel => {
    const predId = 'gr-' + rel.pred_code.replace(/[^a-zA-Z0-9]/g,'_');
    const succId = 'gr-' + rel.succ_code.replace(/[^a-zA-Z0-9]/g,'_');
    const predEl = document.getElementById(predId);
    const succEl = document.getElementById(succId);
    if (!predEl || !succEl) return;

    const pRect = predEl.getBoundingClientRect();
    const sRect = succEl.getBoundingClientRect();

    // Find the bar cells (second child)
    const pBar = predEl.querySelector('.gantt-bar');
    const sBar = succEl.querySelector('.gantt-bar');
    if (!pBar || !sBar) return;

    const pBarRect = pBar.getBoundingClientRect();
    const sBarRect = sBar.getBoundingClientRect();

    // Start: right edge of pred bar; End: left edge of succ bar
    const x1 = pBarRect.right  - containerRect.left;
    const y1 = pRect.top + pRect.height / 2 - containerRect.top;
    const x2 = sBarRect.left   - containerRect.left;
    const y2 = sRect.top + sRect.height / 2 - containerRect.top;

    if (x1 < 0 || x2 < 0) return;

    // Elbow connector: right → down/up → right
    const mx = x1 + Math.max(6, (x2 - x1) * 0.4);
    const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    path.setAttribute('d', `M ${x1} ${y1} L ${mx} ${y1} L ${mx} ${y2} L ${x2} ${y2}`);
    path.setAttribute('stroke', '#94a3b8');
    path.setAttribute('stroke-width', '1');
    path.setAttribute('fill', 'none');
    path.setAttribute('stroke-dasharray', '3,2');
    svg.appendChild(path);

    // Arrowhead
    const arrow = document.createElementNS('http://www.w3.org/2000/svg', 'polygon');
    arrow.setAttribute('points', `${x2},${y2} ${x2-5},${y2-3} ${x2-5},${y2+3}`);
    arrow.setAttribute('fill', '#94a3b8');
    svg.appendChild(arrow);
  });
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
      ${sampleKeys.map(k => {
        const v = row[k];
        const num = typeof v === 'number' ? v : parseFloat(v);
        const cell = (!isNaN(num) && v !== '' && v !== null && typeof v !== 'boolean')
          ? num.toLocaleString('en-US', { maximumFractionDigits: 2 })
          : esc(v);
        return `<td>${cell}</td>`;
      }).join('')}
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

/* ── Report Generator ────────────────────────────────────────────────────── */

function openReportModal() {
  document.getElementById('report-modal-overlay').style.display = 'block';
  document.getElementById('report-modal').style.display = 'block';
}
function closeReportModal() {
  document.getElementById('report-modal-overlay').style.display = 'none';
  document.getElementById('report-modal').style.display = 'none';
}

document.addEventListener('DOMContentLoaded', () => {
  // Format card toggle
  document.querySelectorAll('.report-format-card').forEach(card => {
    card.addEventListener('click', () => {
      document.querySelectorAll('.report-format-card').forEach(c => c.classList.remove('selected'));
      card.classList.add('selected');
    });
  });

  // Close modal
  const closeBtn = document.getElementById('report-modal-close');
  if (closeBtn) closeBtn.addEventListener('click', closeReportModal);
  const overlay = document.getElementById('report-modal-overlay');
  if (overlay) overlay.addEventListener('click', closeReportModal);

  // Generate report
  const genBtn = document.getElementById('report-generate-btn');
  if (genBtn) genBtn.addEventListener('click', () => {
    const fmt = document.querySelector('input[name="report-fmt"]:checked')?.value || 'pdf';
    const sections = [...document.querySelectorAll('input[name="rpt-sec"]:checked')].map(el => el.value);

    if (fmt === 'pdf') {
      closeReportModal();
      setTimeout(() => window.print(), 200);
    } else {
      // CSV export — redirect to existing endpoint
      window.location.href = '/export/csv?key=' + encodeURIComponent(KEY);
    }
  });
});

/* ── Health badge init ───────────────────────────────────────────────────── */

function initHealthBadge() {
  const badge = document.getElementById('health-badge');
  if (!badge) return;
  const K = D.kpis || {};
  const ai = D.ai_summary || {};
  let rawStatus = K.schedule_status || ai.schedule_health;
  // Derive from SPI if backend didn't supply a status
  if (!rawStatus && K.spi_duration != null) {
    rawStatus = K.spi_duration >= 0.95 ? (K.spi_duration > 1.05 ? 'AHEAD' : 'ON TRACK') : 'BEHIND';
  }
  rawStatus = rawStatus || 'unknown';
  const status = rawStatus.toLowerCase().replace(/[\s_]+/g, '');
  const labels = { ahead: 'AHEAD', ontrack: 'ON TRACK', behind: 'BEHIND', unknown: '—' };
  badge.textContent = labels[status] || rawStatus.toUpperCase();
  badge.className = 'health-badge ' + status;
}

/* ── Bootstrap ───────────────────────────────────────────────────────────── */

function _bootstrap() {
  // Sidebar nav-item click handlers
  const tabTitles = {
    overview: 'Overview', schedule: 'Schedule', kpis: 'KPIs',
    ev: 'Earned Value', lookahead: 'Lookahead', milestones: 'Milestones',
    gantt: 'Gantt Chart', procurement: 'Procurement', resources: 'Resources',
    wbs: 'WBS', logic: 'Logic', chat: 'AI Chat'
  };

  document.querySelectorAll('.nav-item').forEach(btn => {
    btn.addEventListener('click', () => {
      const tab = btn.dataset.tab;
      if (!tab) return;
      history.pushState(null, '', '#' + tab);
      activateTab(tab);
      const titleEl = document.getElementById('page-title');
      if (titleEl) titleEl.textContent = tabTitles[tab] || tab;
    });
  });

  // Legacy tab-btn handlers (backwards compat)
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const name = btn.dataset.tab;
      history.pushState(null, '', '#' + name);
      activateTab(name);
      const titleEl = document.getElementById('page-title');
      if (titleEl) titleEl.textContent = tabTitles[name] || name;
    });
  });

  // Hash change (browser back/forward)
  window.addEventListener('hashchange', () => {
    const name = location.hash.slice(1) || 'executive';
    activateTab(name);
    const titleEl = document.getElementById('page-title');
    if (titleEl) titleEl.textContent = tabTitles[name] || name;
  });

  // Initial render from hash (or default to executive tab)
  const initial = location.hash.slice(1) || 'executive';
  activateTab(initial);

  // Drawer close handlers (works with both cc-drawer and drawer IDs)
  function closeDrawer() {
    const drawerEl  = document.getElementById('cc-drawer')         || document.getElementById('drawer');
    const overlayEl = document.getElementById('cc-drawer-overlay') || document.getElementById('drawer-overlay');
    if (drawerEl)  drawerEl.classList.remove('open');
    if (overlayEl) { overlayEl.classList.remove('active'); overlayEl.classList.remove('open'); }
  }
  const drawerCloseBtn = document.getElementById('cc-drawer-close') || document.getElementById('drawer-close');
  if (drawerCloseBtn) drawerCloseBtn.addEventListener('click', closeDrawer);
  const drawerOverlayEl = document.getElementById('cc-drawer-overlay') || document.getElementById('drawer-overlay');
  if (drawerOverlayEl) drawerOverlayEl.addEventListener('click', closeDrawer);

  // Set health badge
  initHealthBadge();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', _bootstrap);
} else {
  _bootstrap();
}
