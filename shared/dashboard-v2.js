/**
 * dashboard-v2.js — Shared dashboard orchestrator for ads dashboards v2
 * Handles data loading, month selection, KPI rendering, table rendering, and PDF export.
 * Each client dashboard (trees, harmonices) imports this and calls initDashboard with a render callback.
 */

import {
  formatCurrency, formatPct, formatCount, formatCompact,
  calcDelta, deltaHtml, formatMonthLabel, formatMonthFull, getCssVar
} from './utils.js';
import { exportToPdf } from './pdf-export.js';

let currentData = null;
let config = null;

// ──────────────────────────────────────────────
// Public: Initialize dashboard
// ──────────────────────────────────────────────

/**
 * @param {Object} cfg
 * @param {string} cfg.clientId
 * @param {string} cfg.clientName
 * @param {string} cfg.dataBasePath - e.g. '../data/trees'
 * @param {Function} cfg.onDataLoaded - callback(data) to render client-specific sections
 */
export async function initDashboard(cfg) {
  config = cfg;

  // Chart.js global defaults
  const fontFamily = getCssVar('--font-family') || "'Nunito', sans-serif";
  Chart.defaults.font.family = fontFamily;
  Chart.defaults.font.size = 13;
  Chart.defaults.color = getCssVar('--color-text-muted') || '#6B6B6B';
  Chart.defaults.plugins.legend.labels.usePointStyle = true;

  // Load manifest
  const manifest = await fetchJson(`${cfg.dataBasePath}/manifest.json`);
  if (!manifest) { showError('No se pudo cargar el manifiesto de datos.'); return; }

  // Month selector
  const selector = document.getElementById('month-selector');
  if (selector) {
    selector.innerHTML = '';
    manifest.months.slice().reverse().forEach(m => {
      const opt = document.createElement('option');
      opt.value = m;
      opt.textContent = formatMonthFull(m);
      selector.appendChild(opt);
    });
    selector.value = manifest.latestMonth;
    selector.addEventListener('change', () => loadMonth(selector.value));
  }

  // PDF export
  const btnPdf = document.getElementById('btn-export-pdf');
  if (btnPdf) {
    btnPdf.addEventListener('click', () => {
      if (currentData) exportToPdf(cfg.clientName, currentData.meta.month);
    });
  }

  // Load latest
  await loadMonth(manifest.latestMonth);
}

export async function fetchJson(url) {
  try {
    const res = await fetch(url);
    if (!res.ok) throw new Error(res.status);
    return await res.json();
  } catch (e) {
    console.error('Fetch error:', url, e);
    return null;
  }
}

async function loadMonth(month) {
  const data = await fetchJson(`${config.dataBasePath}/${month}.json`);
  if (!data) { showError(`No se pudieron cargar los datos de ${month}.`); return; }
  currentData = data;

  // Update date range display
  const rangeEl = document.getElementById('date-range');
  if (rangeEl && data.meta.dateRange) {
    const from = new Date(data.meta.dateRange.from);
    const to = new Date(data.meta.dateRange.to);
    const opts = { day: 'numeric', month: 'short', year: 'numeric' };
    rangeEl.textContent = `${from.toLocaleDateString('es-ES', opts)} — ${to.toLocaleDateString('es-ES', opts)}`;
  }

  // Update footer
  const footerEl = document.getElementById('last-updated');
  if (footerEl && data.meta.generatedAt) {
    const d = new Date(data.meta.generatedAt);
    footerEl.textContent = `Última actualización: ${d.toLocaleDateString('es-ES', { day: 'numeric', month: 'long', year: 'numeric', hour: '2-digit', minute: '2-digit' })}`;
  }

  // Call client-specific renderer
  if (config.onDataLoaded) config.onDataLoaded(data);
}

function showError(msg) {
  const el = document.getElementById('dashboard-content');
  if (el) el.innerHTML = `<div style="padding:40px;text-align:center;color:#B91C1C;font-size:1.1rem;">${msg}</div>`;
}

// ──────────────────────────────────────────────
// Public: KPI Helpers
// ──────────────────────────────────────────────

/**
 * Render a set of 7 KPI cards into a container
 * @param {string} containerId - ID of the .kpi-grid element
 * @param {Object} kpis - { impresiones, clics, ctr, cpc, leads, cpl, inversion } each with {value, previous}
 */
export function renderKpiCards(containerId, kpis) {
  const container = document.getElementById(containerId);
  if (!container) return;

  const metrics = [
    { key: 'impresiones', label: 'IMPRESIONES', format: formatCount, invert: false },
    { key: 'clics', label: 'CLICS', format: formatCount, invert: false },
    { key: 'ctr', label: 'CTR', format: v => formatPct(v), invert: false },
    { key: 'cpc', label: 'CPC', format: formatCurrency, invert: true },
    { key: 'leads', label: 'LEADS', format: formatCount, invert: false },
    { key: 'cpl', label: 'CPL', format: formatCurrency, invert: true },
    { key: 'conversion_rate', label: 'TASA CONV.', format: v => formatPct(v), invert: false, computed: true },
    { key: 'inversion', label: 'INVERSIÓN', format: formatCurrency, invert: false }
  ];

  // Compute conversion_rate from leads/clics if not already present
  const kpisWithCR = { ...kpis };
  if (!kpisWithCR.conversion_rate && kpisWithCR.leads && kpisWithCR.clics) {
    const currClics = kpisWithCR.clics.value;
    const prevClics = kpisWithCR.clics.previous;
    kpisWithCR.conversion_rate = {
      value: currClics > 0 ? (kpisWithCR.leads.value / currClics) * 100 : 0,
      previous: prevClics > 0 ? (kpisWithCR.leads.previous / prevClics) * 100 : 0
    };
  }

  container.innerHTML = metrics.map(m => {
    const kpi = kpisWithCR[m.key];
    if (!kpi) return '';
    const value = m.format(kpi.value);
    const delta = deltaHtml(kpi.value, kpi.previous, m.invert);
    return `
      <div class="kpi-card">
        <div class="kpi-label">${m.label}</div>
        <div class="kpi-value">${value}</div>
        <div class="kpi-delta">${delta}</div>
      </div>`;
  }).join('');
}

// ──────────────────────────────────────────────
// Public: Table Rendering
// ──────────────────────────────────────────────

/**
 * Render a sortable data table
 * @param {string} containerId
 * @param {Array} columns - [{key, label, unit, align?, defaultDesc?, render?}]
 * @param {Array} rows
 */
export function renderSortableTable(containerId, columns, rows, options = {}) {
  const container = document.getElementById(containerId);
  if (!container || !rows) return;

  let sortCol = options.defaultSortCol ?? null;
  let sortAsc = options.defaultSortAsc ?? true;

  // Resolve string key to index
  if (typeof sortCol === 'string') {
    sortCol = columns.findIndex(c => c.key === sortCol);
    if (sortCol === -1) sortCol = null;
  }

  function formatCell(value, unit) {
    if (value == null) return '\u2014';
    switch (unit) {
      case 'eur': return formatCurrency(value);
      case 'pct': return formatPct(value);
      case 'count': return formatCount(value);
      default: return String(value);
    }
  }

  function render() {
    const sorted = [...rows];
    if (sortCol !== null) {
      const col = columns[sortCol];
      sorted.sort((a, b) => {
        const va = a[col.key];
        const vb = b[col.key];
        if (va == null && vb == null) return 0;
        if (va == null) return 1;
        if (vb == null) return -1;
        if (typeof va === 'string') return sortAsc ? va.localeCompare(vb) : vb.localeCompare(va);
        return sortAsc ? va - vb : vb - va;
      });
    }

    let html = '<table class="data-table"><thead><tr>';
    columns.forEach((col, i) => {
      const arrow = sortCol === i ? (sortAsc ? ' &#9650;' : ' &#9660;') : '';
      const align = col.align ? `text-align:${col.align}` : '';
      html += `<th data-col="${i}" style="cursor:pointer;${align}">${col.label}${arrow}</th>`;
    });
    html += '</tr></thead><tbody>';

    sorted.forEach(row => {
      html += '<tr>';
      columns.forEach(col => {
        const align = col.align ? `text-align:${col.align}` : '';
        if (col.render) {
          html += `<td style="${align}">${col.render(row)}</td>`;
        } else {
          html += `<td style="${align}">${formatCell(row[col.key], col.unit)}</td>`;
        }
      });
      html += '</tr>';
    });
    html += '</tbody></table>';

    container.innerHTML = html;

    // Attach sort handlers
    container.querySelectorAll('th[data-col]').forEach(th => {
      th.addEventListener('click', () => {
        const idx = parseInt(th.dataset.col);
        if (sortCol === idx) { sortAsc = !sortAsc; }
        else { sortCol = idx; sortAsc = columns[idx].defaultDesc ? false : true; }
        render();
      });
    });
  }

  render();
}

/**
 * Render a historical table with inline bars for specific columns
 * @param {string} containerId
 * @param {Array} history - monthly data
 * @param {Object} options - { barColumns: ['leads', 'cpl'], barColors: { leads: '#2D5016', cpl: '#D97706' } }
 */
export function renderHistoricalTable(containerId, history, options = {}) {
  const container = document.getElementById(containerId);
  if (!container || !history || !history.length) return;

  const barColumns = options.barColumns || [];
  const barColors = options.barColors || {};

  // Compute max values for bar columns
  const maxVals = {};
  barColumns.forEach(key => {
    maxVals[key] = Math.max(...history.map(h => h[key] || 0), 1);
  });

  const columns = [
    { key: 'month', label: 'Mes', format: formatMonthLabel },
    { key: 'impresiones', label: 'Impresiones', format: formatCount },
    { key: 'clics', label: 'Clics', format: formatCount },
    { key: 'ctr', label: 'CTR', format: v => formatPct(v) },
    { key: 'cpc', label: 'CPC', format: formatCurrency },
    { key: 'leads', label: 'Leads', format: formatCount },
    { key: 'cpl', label: 'CPL', format: formatCurrency },
    { key: 'conversion_rate', label: 'Tasa Conv.', format: v => formatPct(v), computed: row => row.clics > 0 ? (row.leads / row.clics) * 100 : 0 },
    { key: 'inversion', label: 'Inversión', format: formatCurrency }
  ];

  let html = '<table class="data-table"><thead><tr>';
  columns.forEach(col => {
    const align = col.key !== 'month' ? 'text-align:right' : '';
    html += `<th style="${align}">${col.label}</th>`;
  });
  html += '</tr></thead><tbody>';

  // Chronological order: oldest first, newest last
  const sorted = [...history];

  sorted.forEach(row => {
    html += '<tr>';
    columns.forEach(col => {
      const val = col.computed ? col.computed(row) : row[col.key];
      const formatted = col.format ? col.format(val) : (val ?? '\u2014');
      const align = col.key !== 'month' ? 'text-align:right' : '';

      if (barColumns.includes(col.key) && val != null) {
        const pct = Math.min((val / maxVals[col.key]) * 100, 100);
        const color = barColors[col.key] || '#2D5016';
        html += `<td style="${align}">
          <div class="inline-bar-wrapper">
            <span class="inline-bar" style="width:${pct.toFixed(0)}%;background:${color}"></span>
            <span class="inline-value">${formatted}</span>
          </div>
        </td>`;
      } else {
        html += `<td style="${align}">${formatted}</td>`;
      }
    });
    html += '</tr>';
  });

  html += '</tbody></table>';
  container.innerHTML = html;
}

// ──────────────────────────────────────────────
// Public: Merge monthly histories from multiple clients
// ──────────────────────────────────────────────

/**
 * Merge one or more monthly_history arrays into a single combined array.
 * Additive fields (impresiones, clics, leads, inversion) are summed.
 * Ratio fields (ctr, cpc, cpl) are recomputed from totals.
 * Missing months in some sources are handled (only summed where data exists).
 *
 * @param {...Array} histories - one or more monthly_history arrays
 * @returns {Array} merged history, sorted chronologically
 */
export function mergeMonthlyHistories(...histories) {
  // Normalize month keys to YYYY-MM (some sources use YYYY-MM-DD, others YYYY-MM)
  const normalizeMonth = m => typeof m === 'string' ? m.substring(0, 7) : String(m);

  const merged = {};
  histories.forEach(hist => {
    if (!Array.isArray(hist)) return;
    hist.forEach(row => {
      const m = normalizeMonth(row.month);
      if (!merged[m]) {
        merged[m] = { month: m, impresiones: 0, clics: 0, leads: 0, inversion: 0 };
      }
      merged[m].impresiones += row.impresiones || 0;
      merged[m].clics += row.clics || 0;
      merged[m].leads += row.leads || 0;
      merged[m].inversion += row.inversion || 0;
    });
  });
  // Recompute ratios from totals
  return Object.values(merged)
    .map(row => ({
      ...row,
      ctr: row.impresiones > 0 ? (row.clics / row.impresiones) * 100 : 0,
      cpc: row.clics > 0 ? row.inversion / row.clics : 0,
      cpl: row.leads > 0 ? row.inversion / row.leads : 0
    }))
    .sort((a, b) => a.month.localeCompare(b.month));
}

/**
 * Load and merge data for multiple clients for the Agrupado dashboard.
 * @param {Array<string>} clientPaths - e.g. ['../data/trees', '../data/harmonices']
 * @param {string} month - YYYY-MM
 * @returns {Promise<Object>} merged data with google_ads.monthly_history, meta_ads.monthly_history, and combined.monthly_history
 */
export async function loadMergedData(clientPaths, month) {
  const datas = await Promise.all(clientPaths.map(p => fetchJson(`${p}/${month}.json`)));
  const valid = datas.filter(d => d != null);
  if (valid.length === 0) return null;

  const googleHistories = valid.map(d => d.google_ads?.monthly_history || []);
  const metaHistories = valid.map(d => d.meta_ads?.monthly_history || []);

  return {
    google_ads: { monthly_history: mergeMonthlyHistories(...googleHistories) },
    meta_ads: { monthly_history: mergeMonthlyHistories(...metaHistories) },
    combined: { monthly_history: mergeMonthlyHistories(...googleHistories, ...metaHistories) }
  };
}

/**
 * Public helper to fetch the latest manifest from multiple clients
 * and return the union of available months.
 */
export async function loadMergedManifest(clientPaths) {
  const manifests = await Promise.all(clientPaths.map(p => fetchJson(`${p}/manifest.json`)));
  const valid = manifests.filter(m => m != null);
  if (valid.length === 0) return null;

  const allMonths = new Set();
  valid.forEach(m => (m.months || []).forEach(x => allMonths.add(x)));
  const months = Array.from(allMonths).sort();
  const latestMonth = months[months.length - 1];
  return { months, latestMonth };
}

// Re-export utils for convenience
export { formatCurrency, formatPct, formatCount, formatCompact, calcDelta, deltaHtml, formatMonthLabel, formatMonthFull };
