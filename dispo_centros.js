// ═══════════════════════════════════════════════════════════════
// DISPO_CENTROS — Módulo por Establecimiento (solo Centros)
// Archivo: DISPO_CENTROS.XLSX
// v2 — gráficos en reporte, click filtra tabla, export XLSX, descarga original
// ═══════════════════════════════════════════════════════════════

// ── Estado ─────────────────────────────────────────────────────
let centrosData       = [];
let centrosEstab      = [];
let centrosSelected   = '';
let centrosSobMeses   = 6;
let centrosCharts     = {};
let centrosFileId     = '';   // ID del archivo en Drive (para descarga original)

// ── Constantes ─────────────────────────────────────────────────
const CND_ORDER = ['DESABASTECIDO','SUBSTOCK','NORMOSTOCK','SOBRESTOCK','SIN ROTACION'];
const CND_COLORS = {
  DESABASTECIDO : { bg:'rgba(244,63,94,0.85)',  border:'#f43f5e' },
  SUBSTOCK      : { bg:'rgba(249,115,22,0.85)', border:'#f97316' },
  NORMOSTOCK    : { bg:'rgba(34,197,94,0.85)',  border:'#22c55e' },
  SOBRESTOCK    : { bg:'rgba(234,179,8,0.85)',  border:'#eab308' },
  'SIN ROTACION': { bg:'rgba(139,92,246,0.85)', border:'#8b5cf6' },
};
const CND_PILL = {
  DESABASTECIDO:'red', SUBSTOCK:'orange',
  NORMOSTOCK:'green',  SOBRESTOCK:'yellow', 'SIN ROTACION':'violet'
};
const CND_IND_COLOR = {
  DESABASTECIDO:'#e11d48', SUBSTOCK:'#ea580c',
  NORMOSTOCK:'#16a34a', SOBRESTOCK:'#ca8a04', 'SIN ROTACION':'#7c3aed'
};

// ═══════════════════════════════════════════════════════════════
// PARSE
// ═══════════════════════════════════════════════════════════════
function parseCentrosBuffer(buf) {
  const wb = XLSX.read(buf, { type:'array', cellDates:true });
  const sheetName = wb.SheetNames.find(n => /DISP.*CENTROS|SINC/i.test(n)) || wb.SheetNames[0];
  const ws  = wb.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });

  let headerRow = -1;
  for (let i = 0; i < Math.min(30, raw.length); i++) {
    const cols = (raw[i] || []).filter(Boolean).map(c => String(c).toLowerCase());
    if (cols.some(c => c.includes('establecimiento') || c.includes('indicador'))) {
      headerRow = i; break;
    }
  }
  if (headerRow === -1) throw new Error('No se encontró encabezado en DISPO_CENTROS');

  const headers = (raw[headerRow] || []).map(h => String(h || '').trim());
  const colIdx  = {};
  headers.forEach((h, i) => { colIdx[h] = i; });

  const toNum = v => {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    const n = parseFloat(String(v).replace(/,/g, '.'));
    return isNaN(n) ? 0 : n;
  };

  const rows = [];
  for (let i = headerRow + 1; i < raw.length; i++) {
    const r   = raw[i] || [];
    const est = String(r[colIdx['establecimiento']] || '').trim();
    if (!est) continue;
    rows.push({
      codigo_pre        : String(r[colIdx['codigo_pre']]          || '').trim(),
      establecimiento   : est,
      RED               : String(r[colIdx['RED']]                 || '').trim(),
      cod_sismed        : String(r[colIdx['cod_sismed']]          || '').trim(),
      descripcion       : String(r[colIdx['descripcion']]         || '').trim(),
      TIPO              : String(r[colIdx['TIPO']]                || '').trim().toUpperCase(),
      estrategic        : String(r[colIdx['estrategic']]          || '').trim(),
      TIPO_ESTRATEGIA   : String(r[colIdx['TIPO DE ESTRATEGIA']]  || '').trim(),
      SUMINISTRO        : String(r[colIdx['SUMINISTRO']]          || '').trim(),
      STOCK_CENTROS     : toNum(r[colIdx['STOCK_CENTROS']]),
      CPMA_CENTROS      : toNum(r[colIdx['CPMA_CENTROS']]),
      DISP_CENTROS      : toNum(r[colIdx['DISP_CENTROS']]),
      INDICADOR_CENTROS : String(r[colIdx['INDICADOR_CENTROS']]   || '').trim().toUpperCase(),
      FECHA_REPORTE     : r[colIdx['FECHA_REPORTE']] instanceof Date
        ? r[colIdx['FECHA_REPORTE']].toISOString().split('T')[0]
        : String(r[colIdx['FECHA_REPORTE']] || ''),
    });
  }
  return rows;
}

// ═══════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════
function getCentrosRows(estab) {
  return estab ? centrosData.filter(d => d.establecimiento === estab) : centrosData;
}
function cntByInd(rows) {
  const c = {};
  CND_ORDER.forEach(k => { c[k] = 0; });
  rows.forEach(r => { const k = r.INDICADOR_CENTROS; c[k] = (c[k] || 0) + 1; });
  return c;
}
function sobByMeses(rows, umbral) {
  const labels  = [`6–${umbral}m`, `${umbral}–12m`, `12–24m`, `>24m`];
  const buckets = [0, 0, 0, 0];
  rows.filter(r => r.INDICADOR_CENTROS === 'SOBRESTOCK').forEach(r => {
    const m = r.DISP_CENTROS;
    if      (m < umbral) buckets[0]++;
    else if (m < 12)     buckets[1]++;
    else if (m < 24)     buckets[2]++;
    else                 buckets[3]++;
  });
  return { labels, buckets };
}
function top20Sob(rows, umbral) {
  return rows
    .filter(r => r.INDICADOR_CENTROS === 'SOBRESTOCK' && r.DISP_CENTROS >= umbral)
    .sort((a, b) => b.DISP_CENTROS - a.DISP_CENTROS)
    .slice(0, 20);
}

// ── Captura canvas como imagen base64 ─────────────────────────
function canvasToImg(id, w, h) {
  const c = document.getElementById(id);
  if (!c) return '';
  try {
    // Crear canvas offline con fondo blanco para imprimir
    const off = document.createElement('canvas');
    off.width  = w || c.width  || 600;
    off.height = h || c.height || 300;
    const ctx = off.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, off.width, off.height);
    ctx.drawImage(c, 0, 0, off.width, off.height);
    return off.toDataURL('image/png');
  } catch(e) { return ''; }
}

// ═══════════════════════════════════════════════════════════════
// RENDER PRINCIPAL
// ═══════════════════════════════════════════════════════════════
function renderCentrosDashboard(data) {
  centrosData  = data;
  centrosEstab = [...new Set(data.map(d => d.establecimiento))].sort();
  centrosSelected = '';

  // Guardar fileId para descarga original
  const fileObj = (typeof driveFiles !== 'undefined') &&
    driveFiles.find(f => /^DISPO_CENTROS\.XLS/i.test(f.name));
  centrosFileId = fileObj?.id || '';

  const fecha = data[0]?.FECHA_REPORTE || '—';
  document.getElementById('tb-meta').textContent =
    `DISPO_CENTROS · ${data.length.toLocaleString()} registros · Reporte: ${fecha}`;

  document.getElementById('dash-content').innerHTML = buildCentrosHTML();
  injectCentrosStyles();
  renderCentrosView();
}

// ─── HTML SHELL ────────────────────────────────────────────────
function buildCentrosHTML() {
  const opts = centrosEstab.map(e =>
    `<option value="${e}"${e === centrosSelected ? ' selected' : ''}>${e}</option>`
  ).join('');

  return `
  <div class="main" id="centros-main">

    <!-- CONTROLES -->
    <div class="detalle-controls-bar">
      <div class="dc-group">
        <label class="dc-label">📍 Establecimiento</label>
        <select class="inp dc-select" id="cc-estab" onchange="onCentrosEstabChange(this.value)">
          <option value="">— TODOS —</option>
          ${opts}
        </select>
      </div>
      <div class="dc-group">
        <label class="dc-label">⚠️ Umbral sobrestock</label>
        <div style="display:flex;align-items:center;gap:8px">
          <input type="range" id="cc-umbral" min="6" max="36" step="1" value="6"
            oninput="onCentrosUmbralChange(+this.value)"
            style="width:110px;accent-color:var(--accent)">
          <span class="dc-umbral-val" id="cc-umbral-val">≥ 6 meses</span>
        </div>
      </div>
      <div style="margin-left:auto;display:flex;gap:8px;flex-wrap:wrap">
        <button class="btn-sm" onclick="downloadCentrosOriginal()" title="Descargar el archivo XLSX original desde Drive">⬇ Archivo Original</button>
        <button class="btn-sm" onclick="exportCentrosXLSX()">📊 Export Excel</button>
        <button class="btn-sm" onclick="printCentrosReport()">🖨 Imprimir Reporte</button>
      </div>
    </div>

    <!-- FILTRO ACTIVO BADGE -->
    <div id="cc-filter-badge" style="display:none">
      <div style="display:flex;align-items:center;gap:10px;background:rgba(0,61,165,.12);border:1px solid rgba(0,61,165,.3);border-radius:10px;padding:8px 14px;">
        <span style="font-size:12px;color:var(--accent2)">🔍 Filtro activo:</span>
        <span id="cc-filter-label" style="font-size:12px;color:var(--text);font-weight:600"></span>
        <button class="btn-sm" style="margin-left:auto" onclick="clearCentrosChartFilter()">✕ Limpiar filtro</button>
      </div>
    </div>

    <!-- KPIs -->
    <div>
      <div class="sec-hdr"><span class="sec-title">Resumen de Disponibilidad — Centros</span><div class="sec-line"></div></div>
      <div class="kpi-grid" id="ck-grid"></div>
    </div>

    <!-- GRÁFICOS FILA 1 -->
    <div class="g2">
      <div class="card">
        <div class="card-hdr">
          <div>
            <div class="card-title">Indicador Centros — Distribución %</div>
            <div class="card-sub">Clic en segmento para filtrar tabla ↓</div>
          </div>
        </div>
        <div class="chart-h260"><canvas id="cc-ind"></canvas></div>
      </div>
      <div class="card">
        <div class="card-hdr">
          <div>
            <div class="card-title">Tipo de Producto por Indicador</div>
            <div class="card-sub">Clic en barra para filtrar por tipo e indicador ↓</div>
          </div>
        </div>
        <div class="chart-h260"><canvas id="cc-tipo"></canvas></div>
      </div>
    </div>

    <!-- GRÁFICOS FILA 2 -->
    <div class="g2">
      <div class="card">
        <div class="card-hdr">
          <div>
            <div class="card-title">Sobrestock — Distribución por Meses</div>
            <div class="card-sub" id="cc-sob-sub">Agrupado según umbral seleccionado</div>
          </div>
        </div>
        <div class="chart-h260"><canvas id="cc-sob"></canvas></div>
      </div>
      <div class="card">
        <div class="card-hdr">
          <div>
            <div class="card-title">Cobertura — Distribución por Rangos</div>
            <div class="card-sub">Clic en barra para filtrar por rango de cobertura ↓</div>
          </div>
        </div>
        <div class="chart-h260"><canvas id="cc-cobertura"></canvas></div>
      </div>
    </div>

    <!-- COMPARATIVO REDES (solo TODOS) -->
    <div class="card" id="cc-redes-card" style="display:none">
      <div class="card-hdr">
        <div>
          <div class="card-title">Comparativo por Red de Salud</div>
          <div class="card-sub">% Desabastecidos, Substock y Sobrestock por red · Clic para filtrar ↓</div>
        </div>
      </div>
      <div class="chart-h320"><canvas id="cc-redes"></canvas></div>
    </div>

    <!-- COMPARATIVO ESTABLECIMIENTOS (solo TODOS) -->
    <div class="card" id="cc-estabs-card" style="display:none">
      <div class="card-hdr">
        <div>
          <div class="card-title">Establecimientos — % Desabastecidos</div>
          <div class="card-sub">Top 20 · Clic en barra para filtrar ese establecimiento ↓</div>
        </div>
      </div>
      <div class="chart-h360"><canvas id="cc-estabs-desab"></canvas></div>
    </div>

    <!-- TOP SOBRESTOCK -->
    <div class="card">
      <div class="card-hdr">
        <div>
          <div class="card-title" id="cc-topsob-title">Top 20 Mayor Sobrestock (≥ umbral)</div>
          <div class="card-sub" id="cc-topsob-sub">Productos con mayor cobertura · Clic para buscar producto ↓</div>
        </div>
      </div>
      <div class="chart-h360"><canvas id="cc-top-sob"></canvas></div>
    </div>

    <!-- TABLA -->
    <div>
      <div class="sec-hdr"><span class="sec-title">Catálogo — Centros</span><div class="sec-line"></div></div>
      <div class="tbl-card">
        <div class="tbl-bar">
          <span class="tbl-bar-title">Productos</span>
          <input class="inp" type="text" id="ct-search" placeholder="🔍 Buscar producto o código…"
            style="width:220px" oninput="filterCentrosTable()">
          <select class="inp" id="ct-ind" onchange="filterCentrosTable()">
            <option value="">Todos los estados</option>
            ${CND_ORDER.map(i => `<option>${i}</option>`).join('')}
          </select>
          <select class="inp" id="ct-tipo" onchange="filterCentrosTable()">
            <option value="">Todos los tipos</option>
            <option value="M">Medicamentos</option>
            <option value="I">Insumos</option>
          </select>
          <span class="tbl-count" id="ct-count">0 productos</span>
          <button class="btn-sm" onclick="exportCentrosXLSX()" title="Exportar filtrado actual a Excel">📊 Excel filtrado</button>
        </div>
        <div class="tbl-wrap">
          <table>
            <thead>
              <tr>
                <th onclick="sortCentrosTable('cod_sismed')">Código</th>
                <th onclick="sortCentrosTable('descripcion')">Descripción</th>
                <th onclick="sortCentrosTable('TIPO')">Tipo</th>
                <th onclick="sortCentrosTable('establecimiento')" id="ct-estab-col">Establecimiento</th>
                <th onclick="sortCentrosTable('RED')">Red</th>
                <th onclick="sortCentrosTable('estrategic')">Estratégico</th>
                <th onclick="sortCentrosTable('STOCK_CENTROS')" style="text-align:right">Stock</th>
                <th onclick="sortCentrosTable('CPMA_CENTROS')" style="text-align:right">CPMA</th>
                <th onclick="sortCentrosTable('DISP_CENTROS')" style="text-align:right">Disp. (m)</th>
                <th onclick="sortCentrosTable('INDICADOR_CENTROS')">Indicador</th>
              </tr>
            </thead>
            <tbody id="ct-body"></tbody>
          </table>
        </div>
        <div class="tbl-pager">
          <button class="btn-sm" onclick="prevCentrosPage()">← Ant</button>
          <span id="ct-pager">Pág. 1 / 1</span>
          <button class="btn-sm" onclick="nextCentrosPage()">Sig →</button>
        </div>
      </div>
    </div>

    <!-- PRINT AREA (oculto) -->
    <div id="centros-print-area" style="display:none"></div>
  </div>`;
}

// ─── ESTILOS ──────────────────────────────────────────────────
function injectCentrosStyles() {
  if (document.getElementById('detalle-styles') || document.getElementById('centros-styles')) return;
  const s = document.createElement('style');
  s.id = 'centros-styles';
  s.textContent = `
.detalle-controls-bar{display:flex;align-items:flex-end;flex-wrap:wrap;gap:16px;background:var(--surf);border:1px solid var(--border);border-radius:16px;padding:16px 20px;}
.dc-group{display:flex;flex-direction:column;gap:5px;}
.dc-label{font-size:10px;color:var(--text);text-transform:uppercase;letter-spacing:.8px;}
.dc-select{min-width:220px;}
.dc-umbral-val{font-family:'Space Mono',monospace;font-size:11px;color:var(--accent);white-space:nowrap;}
#cc-umbral{accent-color:var(--accent);}
@media print{
  body *{visibility:hidden;}
  #centros-print-area,#centros-print-area *{visibility:visible;}
  #centros-print-area{position:fixed;top:0;left:0;width:100%;display:block!important;background:#fff;color:#000;padding:20px;font-family:Arial,sans-serif;font-size:11px;}
  .pr-title{font-size:18px;font-weight:bold;margin-bottom:4px;}
  .pr-sub{font-size:11px;color:#555;margin-bottom:16px;border-bottom:2px solid #ddd;padding-bottom:8px;}
  .pr-kpi-row{display:flex;gap:10px;margin:14px 0;flex-wrap:wrap;}
  .pr-kpi{border:1px solid #ddd;border-radius:6px;padding:10px 14px;min-width:110px;text-align:center;}
  .pr-kpi-lbl{font-size:9px;color:#888;text-transform:uppercase;margin-bottom:4px;}
  .pr-kpi-val{font-size:20px;font-weight:bold;}
  .pr-kpi-val.red{color:#e11d48;} .pr-kpi-val.ora{color:#ea580c;}
  .pr-kpi-val.yel{color:#ca8a04;} .pr-kpi-val.grn{color:#16a34a;}
  .pr-kpi-val.vio{color:#7c3aed;}
  .pr-section{margin-top:18px;page-break-inside:avoid;}
  .pr-section-title{font-size:13px;font-weight:bold;border-left:4px solid #4f46e5;padding-left:8px;margin-bottom:10px;}
  .pr-charts-row{display:flex;gap:12px;margin:12px 0;}
  .pr-chart-box{flex:1;border:1px solid #e5e7eb;border-radius:6px;padding:10px;}
  .pr-chart-img{width:100%;height:auto;border-radius:4px;}
  .pr-chart-title{font-size:11px;font-weight:bold;margin-bottom:6px;color:#374151;}
  .pr-stats-table{width:100%;border-collapse:collapse;font-size:10px;}
  .pr-stats-table th{background:#f3f4f6;padding:5px 8px;text-align:left;border:1px solid #e5e7eb;font-size:9px;}
  .pr-stats-table td{padding:4px 8px;border:1px solid #e5e7eb;}
  .pr-stats-table tr:nth-child(even) td{background:#f9fafb;}
  .pr-bar-row{display:flex;align-items:center;gap:6px;margin:3px 0;font-size:10px;}
  .pr-bar-fill{height:14px;border-radius:3px;min-width:2px;}
  .pr-table{width:100%;border-collapse:collapse;font-size:9px;margin-top:8px;}
  .pr-table th{background:#1e2032;color:#fff;padding:5px 7px;text-align:left;border:1px solid #374151;}
  .pr-table td{padding:4px 7px;border:1px solid #e5e7eb;}
  .pr-table tr:nth-child(even) td{background:#f9fafb;}
  .pr-footer{margin-top:16px;font-size:9px;color:#9ca3af;border-top:1px solid #e5e7eb;padding-top:6px;}
  .pr-page-break{page-break-before:always;}
}`;
  document.head.appendChild(s);
}

// ═══════════════════════════════════════════════════════════════
// EVENTOS
// ═══════════════════════════════════════════════════════════════
function onCentrosEstabChange(val) {
  centrosSelected = val;
  clearCentrosChartFilter(true);
  renderCentrosView();
}
function onCentrosUmbralChange(v) {
  centrosSobMeses = v;
  document.getElementById('cc-umbral-val').textContent = `≥ ${v} meses`;
  renderCentrosView();
}

// ── Filtro desde gráficos ──────────────────────────────────────
let centrosChartFilter = null; // { field, value, label }

function applyCentrosChartFilter(field, value, label) {
  centrosChartFilter = { field, value, label };
  const badge = document.getElementById('cc-filter-badge');
  const lbl   = document.getElementById('cc-filter-label');
  if (badge) badge.style.display = 'flex';
  if (lbl)   lbl.textContent = label;
  filterCentrosTable();
  // Scroll suave a la tabla
  const tbl = document.querySelector('#centros-main .tbl-card');
  if (tbl) tbl.scrollIntoView({ behavior:'smooth', block:'start' });
}

function clearCentrosChartFilter(silent) {
  centrosChartFilter = null;
  const badge = document.getElementById('cc-filter-badge');
  if (badge) badge.style.display = 'none';
  // Limpiar también selects manuales
  const indEl = document.getElementById('ct-ind');
  const tipEl = document.getElementById('ct-tipo');
  if (indEl) indEl.value = '';
  if (tipEl) tipEl.value = '';
  if (!silent) filterCentrosTable();
}

// ═══════════════════════════════════════════════════════════════
// RENDER VIEW
// ═══════════════════════════════════════════════════════════════
function renderCentrosView() {
  const rows = getCentrosRows(centrosSelected);
  Object.values(centrosCharts).forEach(c => { try { c.destroy(); } catch(e){} });
  centrosCharts = {};
  renderCentrosKPIs(rows);
  renderCentrosCharts(rows);
  filterCentrosTable();
  const redesCard  = document.getElementById('cc-redes-card');
  const estabsCard = document.getElementById('cc-estabs-card');
  if (redesCard)  redesCard.style.display  = centrosSelected ? 'none' : 'block';
  if (estabsCard) estabsCard.style.display = centrosSelected ? 'none' : 'block';
  if (!centrosSelected) {
    renderCentrosRedesChart();
    renderCentrosEstabsChart();
  }
}

// ─── KPIs ──────────────────────────────────────────────────────
function renderCentrosKPIs(rows) {
  const total = rows.length;
  const c     = cntByInd(rows);
  const sob   = rows.filter(r => r.INDICADOR_CENTROS === 'SOBRESTOCK' && r.DISP_CENTROS >= centrosSobMeses).length;
  const pct   = v => total > 0 ? (v / total * 100).toFixed(1) + '%' : '0%';
  const kpis  = [
    { icon:'🧾', lbl:'Total Productos',          val: total.toLocaleString(),  sub:'registros',              cls:'c-cyan'   },
    { icon:'🔴', lbl:'Desabastecidos',            val: c.DESABASTECIDO||0,      sub: pct(c.DESABASTECIDO),    cls:'c-red'    },
    { icon:'🟠', lbl:'Substock',                  val: c.SUBSTOCK||0,           sub: pct(c.SUBSTOCK),         cls:'c-orange' },
    { icon:'🟢', lbl:'Normostock',                val: c.NORMOSTOCK||0,         sub: pct(c.NORMOSTOCK),       cls:'c-green'  },
    { icon:'🟡', lbl:`Sobrestock ≥${centrosSobMeses}m`, val: sob,              sub: pct(sob),                cls:'c-yellow' },
    { icon:'🟡', lbl:'Sobrestock Total',          val: c.SOBRESTOCK||0,         sub: pct(c.SOBRESTOCK),       cls:'c-yellow' },
    { icon:'🟣', lbl:'Sin Rotación',              val: c['SIN ROTACION']||0,    sub: pct(c['SIN ROTACION']),  cls:'c-violet' }
  ];
  document.getElementById('ck-grid').innerHTML = kpis.map(k => `
    <div class="kpi ${k.cls}">
      <div class="kpi-icon">${k.icon}</div>
      <div class="kpi-lbl">${k.lbl}</div>
      <div class="kpi-val">${k.val}</div>
      <div class="kpi-sub">${k.sub}</div>
    </div>`).join('');
}

// ─── CHARTS ────────────────────────────────────────────────────
function renderCentrosCharts(rows) {
  const c     = cntByInd(rows);
  const total = rows.length;

  // ── 1. Doughnut indicador ────────────────────────────────────
  const dLabels = CND_ORDER.filter(k => (c[k]||0) > 0);
  const dVals   = dLabels.map(k => c[k]);
  const ctx1 = document.getElementById('cc-ind');
  if (ctx1) {
    centrosCharts['cc-ind'] = new Chart(ctx1, {
      type: 'doughnut',
      data: {
        labels: dLabels.map((l,i) => `${l} (${total>0?(dVals[i]/total*100).toFixed(1):0}%)`),
        datasets: [{
          data: dVals,
          backgroundColor: dLabels.map(l => CND_COLORS[l]?.bg  || 'rgba(0,0,0,0.1)'),
          borderColor:     dLabels.map(l => CND_COLORS[l]?.border || '#4a5568'),
          borderWidth:2, hoverOffset:8
        }]
      },
      options: {
        responsive:true, maintainAspectRatio:false, cutout:'60%',
        plugins: {
          legend: { position:'right', labels:{ color:'#4a5568', font:{ size:10 }, padding:10 } },
          tooltip: { callbacks: { label: ctx => ` ${ctx.raw.toLocaleString()} (${total>0?(ctx.raw/total*100).toFixed(1):0}%)` } }
        },
        onClick: (e, elems) => {
          if (!elems.length) return;
          const lbl = dLabels[elems[0].index];
          applyCentrosChartFilter('INDICADOR_CENTROS', lbl, `Indicador = ${lbl}`);
        }
      }
    });
  }

  // ── 2. Barras apiladas Tipo x Indicador ──────────────────────
  const byTipo = tipo => CND_ORDER.map(k => rows.filter(r => r.TIPO===tipo && r.INDICADOR_CENTROS===k).length);
  const ctx2 = document.getElementById('cc-tipo');
  if (ctx2) {
    centrosCharts['cc-tipo'] = new Chart(ctx2, {
      type:'bar',
      data: {
        labels: CND_ORDER,
        datasets: [
          { label:'Medicamentos', data: byTipo('M'), backgroundColor:'rgba(59,130,246,0.8)',  borderRadius:4 },
          { label:'Insumos',      data: byTipo('I'), backgroundColor:'rgba(139,92,246,0.8)', borderRadius:4 },
        ]
      },
      options: {
        responsive:true, maintainAspectRatio:false,
        plugins: { legend:{ labels:{ color:'#4a5568', font:{ size:11 } } } },
        scales: {
          x:{ stacked:true, ticks:{ color:'#4a5568', font:{ size:10 } }, grid:{ color:'rgba(0,0,0,0.04)' } },
          y:{ stacked:true, ticks:{ color:'#4a5568', font:{ size:11 } }, grid:{ color:'rgba(0,0,0,0.06)' }, beginAtZero:true }
        },
        onClick: (e, elems) => {
          if (!elems.length) return;
          const ind  = CND_ORDER[elems[0].index];
          const tipo = elems[0].datasetIndex === 0 ? 'M' : 'I';
          const tipoLabel = tipo==='M' ? 'Medicamentos' : 'Insumos';
          // Setear ambos selects y filtrar
          const indEl = document.getElementById('ct-ind');
          const tipEl = document.getElementById('ct-tipo');
          if (indEl) indEl.value = ind;
          if (tipEl) tipEl.value = tipo;
          applyCentrosChartFilter('_multi', ind+'|'+tipo, `${tipoLabel} · ${ind}`);
        }
      }
    });
  }

  // ── 3. Sobrestock por meses ───────────────────────────────────
  const sob     = sobByMeses(rows, centrosSobMeses);
  const totSob  = sob.buckets.reduce((a,b)=>a+b,0);
  document.getElementById('cc-sob-sub').textContent =
    `Umbral: ≥ ${centrosSobMeses} meses · Total sobrestock: ${c.SOBRESTOCK||0} productos`;
  const ctx3 = document.getElementById('cc-sob');
  if (ctx3) {
    // Rangos de meses para filtrado al hacer clic
    const sobRanges = [
      { min:6,           max:centrosSobMeses },
      { min:centrosSobMeses, max:12 },
      { min:12,          max:24 },
      { min:24,          max:Infinity },
    ];
    centrosCharts['cc-sob'] = new Chart(ctx3, {
      type:'bar',
      data: {
        labels: sob.labels,
        datasets: [{ label:'Productos', data: sob.buckets,
          backgroundColor:'rgba(234,179,8,0.8)', borderColor:'#eab308',
          borderWidth:2, borderRadius:6 }]
      },
      options: {
        responsive:true, maintainAspectRatio:false,
        plugins: {
          legend:{ display:false },
          tooltip:{ callbacks:{ label: ctx => ` ${ctx.raw} (${totSob>0?(ctx.raw/totSob*100).toFixed(1):0}% del sobrestock)` } }
        },
        scales: {
          x:{ ticks:{ color:'#1a2745', font:{ size:11 } }, grid:{ color:'rgba(0,0,0,0.04)' } },
          y:{ ticks:{ color:'#1a2745', font:{ size:11 } }, grid:{ color:'rgba(0,0,0,0.06)' }, beginAtZero:true }
        },
        onClick: (e, elems) => {
          if (!elems.length) return;
          const idx = elems[0].index;
          const rng = sobRanges[idx];
          applyCentrosChartFilter('_sob_rango', JSON.stringify(rng),
            `Sobrestock · ${sob.labels[idx]}`);
        }
      }
    });
  }

  // ── 4. Cobertura rangos ───────────────────────────────────────
  const rangos  = ['0m (sin stock)','<1m','1–3m','3–6m','6–12m','12–24m','>24m'];
  const rMin    = [0,   0,  1,  3,  6, 12, 24];
  const rMax    = [0,   1,  3,  6, 12, 24, Infinity];
  const rCounts = rangos.map((_,i) => rows.filter(r => {
    const m = r.DISP_CENTROS;
    if (i===0) return m===0;
    return m >= rMin[i] && m < rMax[i];
  }).length);
  const rColors = ['rgba(244,63,94,0.8)','rgba(249,115,22,0.8)','rgba(234,179,8,0.8)',
    'rgba(34,197,94,0.8)','rgba(34,197,94,0.65)','rgba(6,182,212,0.8)','rgba(139,92,246,0.8)'];
  const ctx4 = document.getElementById('cc-cobertura');
  if (ctx4) {
    centrosCharts['cc-cobertura'] = new Chart(ctx4, {
      type:'bar',
      data: {
        labels: rangos,
        datasets: [{ label:'Productos', data: rCounts, backgroundColor: rColors, borderRadius:5 }]
      },
      options: {
        responsive:true, maintainAspectRatio:false,
        plugins: {
          legend:{ display:false },
          tooltip:{ callbacks:{ label: ctx => ` ${ctx.raw} (${total>0?(ctx.raw/total*100).toFixed(1):0}%)` } }
        },
        scales: {
          x:{ ticks:{ color:'#1a2745', font:{ size:10 } }, grid:{ color:'rgba(0,0,0,0.04)' } },
          y:{ ticks:{ color:'#1a2745', font:{ size:11 } }, grid:{ color:'rgba(0,0,0,0.06)' }, beginAtZero:true }
        },
        onClick: (e, elems) => {
          if (!elems.length) return;
          const idx = elems[0].index;
          applyCentrosChartFilter('_cob_rango', JSON.stringify({ min:rMin[idx], max:rMax[idx], exact0: idx===0 }),
            `Cobertura: ${rangos[idx]}`);
        }
      }
    });
  }

  // ── 5. Top sobrestock ─────────────────────────────────────────
  const topSob = top20Sob(rows, centrosSobMeses);
  document.getElementById('cc-topsob-title').textContent =
    `Top 20 Mayor Sobrestock ≥ ${centrosSobMeses} meses`;
  document.getElementById('cc-topsob-sub').textContent =
    `${topSob.length} productos · Clic para buscar en tabla ↓`;
  const ctx5 = document.getElementById('cc-top-sob');
  if (ctx5) {
    if (topSob.length) {
      const bgs = topSob.map(r =>
        r.DISP_CENTROS>=24 ? 'rgba(244,63,94,0.8)' : r.DISP_CENTROS>=12 ? 'rgba(249,115,22,0.8)' : 'rgba(234,179,8,0.8)');
      centrosCharts['cc-top-sob'] = new Chart(ctx5, {
        type:'bar',
        data: {
          labels: topSob.map(r => r.descripcion.length>32 ? r.descripcion.slice(0,30)+'…' : r.descripcion),
          datasets: [{ label:'Meses cobertura', data: topSob.map(r => +r.DISP_CENTROS.toFixed(1)),
            backgroundColor: bgs, borderRadius:5 }]
        },
        options: {
          indexAxis:'y', responsive:true, maintainAspectRatio:false,
          plugins: {
            legend:{ display:false },
            tooltip:{ callbacks:{
              label: ctx => ` ${ctx.raw} meses`,
              afterLabel: ctx => `Estab: ${topSob[ctx.dataIndex].establecimiento}\nStock: ${topSob[ctx.dataIndex].STOCK_CENTROS.toLocaleString()}`
            }}
          },
          scales: {
            x:{ ticks:{ color:'#1a2745', font:{ size:10 } }, grid:{ color:'rgba(0,0,0,0.06)' },
               title:{ display:true, text:'Meses de cobertura', color:'#1a2745', font:{ size:10 } } },
            y:{ ticks:{ color:'#0d1b2a', font:{ size:10 } }, grid:{ color:'rgba(0,0,0,0.03)' } }
          },
          onClick: (e, elems) => {
            if (!elems.length) return;
            const prod = topSob[elems[0].index];
            const el = document.getElementById('ct-search');
            if (el) { el.value = prod.cod_sismed; filterCentrosTable(); }
            applyCentrosChartFilter('_search', prod.cod_sismed, `Producto: ${prod.descripcion.slice(0,40)}`);
          }
        }
      });
    } else {
      ctx5.parentElement.innerHTML =
        '<div style="display:flex;align-items:center;justify-content:center;height:360px;color:var(--muted);font-size:13px">No hay productos en sobrestock con ese umbral</div>';
    }
  }
}

// ─── COMPARATIVO REDES ─────────────────────────────────────────
function renderCentrosRedesChart() {
  const ctx = document.getElementById('cc-redes');
  if (!ctx) return;
  if (centrosCharts['cc-redes']) centrosCharts['cc-redes'].destroy();
  const redes = [...new Set(centrosData.map(d => d.RED))].filter(Boolean);
  const mkPct = (red, ind) => {
    const sub = centrosData.filter(d => d.RED===red);
    return sub.length>0 ? +(sub.filter(d => d.INDICADOR_CENTROS===ind).length/sub.length*100).toFixed(1) : 0;
  };
  centrosCharts['cc-redes'] = new Chart(ctx, {
    type:'bar',
    data: {
      labels: redes,
      datasets: [
        { label:'% Desab.',     data: redes.map(r=>mkPct(r,'DESABASTECIDO')), backgroundColor:'rgba(244,63,94,0.8)',  borderRadius:4 },
        { label:'% Substock',   data: redes.map(r=>mkPct(r,'SUBSTOCK')),      backgroundColor:'rgba(249,115,22,0.8)', borderRadius:4 },
        { label:'% Sobrestock', data: redes.map(r=>mkPct(r,'SOBRESTOCK')),    backgroundColor:'rgba(234,179,8,0.8)',  borderRadius:4 },
        { label:'% Normostock', data: redes.map(r=>mkPct(r,'NORMOSTOCK')),    backgroundColor:'rgba(34,197,94,0.8)',  borderRadius:4 },
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins: {
        legend:{ labels:{ color:'#1a2745', font:{ size:11 } } },
        tooltip:{ callbacks:{ label: c => ` ${c.raw}%` } }
      },
      scales: {
        x:{ ticks:{ color:'#1a2745', font:{ size:11 } }, grid:{ color:'rgba(0,0,0,0.04)' } },
        y:{ ticks:{ color:'#1a2745', font:{ size:11 }, callback: v=>v+'%' },
           grid:{ color:'rgba(0,0,0,0.06)' }, beginAtZero:true, max:100 }
      },
      onClick: (e, elems) => {
        if (!elems.length) return;
        const red = redes[elems[0].index];
        applyCentrosChartFilter('RED', red, `Red: ${red}`);
      }
    }
  });
}

// ─── TOP ESTABLECIMIENTOS ──────────────────────────────────────
function renderCentrosEstabsChart() {
  const ctx = document.getElementById('cc-estabs-desab');
  if (!ctx) return;
  if (centrosCharts['cc-estabs']) centrosCharts['cc-estabs'].destroy();
  const estabs = centrosEstab.map(e => {
    const sub = centrosData.filter(d => d.establecimiento===e);
    const desab = sub.filter(d => d.INDICADOR_CENTROS==='DESABASTECIDO').length;
    return { e, pct: sub.length>0 ? +(desab/sub.length*100).toFixed(1) : 0 };
  }).sort((a,b)=>b.pct-a.pct).slice(0,20);

  centrosCharts['cc-estabs'] = new Chart(ctx, {
    type:'bar',
    data: {
      labels: estabs.map(x => x.e.length>30 ? x.e.slice(0,28)+'…' : x.e),
      datasets: [{ label:'% Desabastecidos', data: estabs.map(x=>x.pct),
        backgroundColor: estabs.map(x => x.pct>30?'rgba(244,63,94,0.85)':x.pct>15?'rgba(249,115,22,0.8)':'rgba(234,179,8,0.8)'),
        borderRadius:5 }]
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins: {
        legend:{ display:false },
        tooltip:{ callbacks:{ label: c => ` ${c.raw}% desabastecidos` } }
      },
      scales: {
        x:{ ticks:{ color:'#1a2745', font:{ size:10 }, callback: v=>v+'%' },
           grid:{ color:'rgba(0,0,0,0.06)' }, max:100 },
        y:{ ticks:{ color:'#0d1b2a', font:{ size:10 } }, grid:{ color:'rgba(0,0,0,0.03)' } }
      },
      onClick: (e, elems) => {
        if (!elems.length) return;
        const estab = estabs[elems[0].index].e;
        // Cambiar selector de establecimiento
        const sel = document.getElementById('cc-estab');
        if (sel) { sel.value = estab; onCentrosEstabChange(estab); }
      }
    }
  });
}

// ═══════════════════════════════════════════════════════════════
// TABLA — filtro combinado (selects + chartFilter)
// ═══════════════════════════════════════════════════════════════
let ctFiltered = [], ctPage = 1, ctSortCol = 'DISP_CENTROS', ctSortDir = -1;
const CT_PAGE  = 50;

function filterCentrosTable() {
  const rows = getCentrosRows(centrosSelected);
  const s    = (document.getElementById('ct-search')?.value || '').toLowerCase();
  const ind  = document.getElementById('ct-ind')?.value  || '';
  const tipo = document.getElementById('ct-tipo')?.value || '';
  const cf   = centrosChartFilter;

  ctFiltered = rows.filter(r => {
    // Filtros manuales (selects)
    if (ind  && r.INDICADOR_CENTROS !== ind)  return false;
    if (tipo && r.TIPO !== tipo)              return false;
    if (s && !r.descripcion.toLowerCase().includes(s) && !r.cod_sismed.includes(s)) return false;
    // Filtro desde gráfico
    if (cf) {
      if (cf.field === 'INDICADOR_CENTROS' && r.INDICADOR_CENTROS !== cf.value) return false;
      if (cf.field === 'RED'               && r.RED               !== cf.value) return false;
      if (cf.field === '_multi') {
        const [cInd, cTipo] = cf.value.split('|');
        if (r.INDICADOR_CENTROS !== cInd || r.TIPO !== cTipo) return false;
      }
      if (cf.field === '_sob_rango') {
        const rng = JSON.parse(cf.value);
        if (r.INDICADOR_CENTROS !== 'SOBRESTOCK') return false;
        if (r.DISP_CENTROS < rng.min || r.DISP_CENTROS >= rng.max) return false;
      }
      if (cf.field === '_cob_rango') {
        const rng = JSON.parse(cf.value);
        if (rng.exact0 && r.DISP_CENTROS !== 0) return false;
        if (!rng.exact0 && (r.DISP_CENTROS < rng.min || r.DISP_CENTROS >= rng.max)) return false;
      }
      if (cf.field === '_search' && r.cod_sismed !== cf.value) return false;
    }
    return true;
  });

  ctFiltered.sort((a,b) => {
    const av = a[ctSortCol], bv = b[ctSortCol];
    if (typeof av==='number') return (av-bv)*ctSortDir;
    return String(av).localeCompare(String(bv))*ctSortDir;
  });
  ctPage = 1;
  renderCentrosTable();
}

function sortCentrosTable(col) {
  if (ctSortCol===col) ctSortDir*=-1; else { ctSortCol=col; ctSortDir=-1; }
  filterCentrosTable();
}

function renderCentrosTable() {
  const total  = ctFiltered.length;
  const totalP = Math.ceil(total/CT_PAGE)||1;
  const slice  = ctFiltered.slice((ctPage-1)*CT_PAGE, ctPage*CT_PAGE);

  document.getElementById('ct-count').textContent = total.toLocaleString()+' productos';
  document.getElementById('ct-pager').textContent = `Pág. ${ctPage} / ${totalP}`;

  const mBar = v => {
    const fill = Math.min(v/12*100,100);
    const col  = v===0?'var(--red)':v<1?'var(--orange)':v<3?'var(--yellow)':v<6?'var(--green)':'var(--yellow)';
    return `<div class="prog-row" style="min-width:80px">
      <div class="prog-bar"><div class="prog-fill" style="width:${fill}%;background:${col}"></div></div>
      <span class="prog-val">${v.toFixed(1)}</span></div>`;
  };

  document.getElementById('ct-body').innerHTML = !slice.length
    ? `<tr><td colspan="10" style="text-align:center;padding:36px;color:var(--muted)">Sin resultados</td></tr>`
    : slice.map(r => `
    <tr>
      <td class="mono">${r.cod_sismed}</td>
      <td style="max-width:240px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${r.descripcion}">${r.descripcion}</td>
      <td><span class="pill ${r.TIPO==='M'?'pill-blue':'pill-violet'}">${r.TIPO==='M'?'Med.':'Ins.'}</span></td>
      <td style="font-size:11px;color:var(--muted2);max-width:160px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${r.establecimiento}</td>
      <td style="font-size:10px;color:var(--muted)">${r.RED||'—'}</td>
      <td style="font-size:10px;color:var(--muted2)">${r.estrategic||'—'}</td>
      <td class="mono" style="text-align:right">${r.STOCK_CENTROS.toLocaleString()}</td>
      <td class="mono" style="text-align:right">${r.CPMA_CENTROS.toFixed(1)}</td>
      <td>${mBar(r.DISP_CENTROS)}</td>
      <td><span class="pill pill-${CND_PILL[r.INDICADOR_CENTROS]||'gray'}">${r.INDICADOR_CENTROS||'—'}</span></td>
    </tr>`).join('');
}

function prevCentrosPage() { if (ctPage>1) { ctPage--; renderCentrosTable(); } }
function nextCentrosPage() {
  if (ctPage < Math.ceil(ctFiltered.length/CT_PAGE)) { ctPage++; renderCentrosTable(); }
}

// ═══════════════════════════════════════════════════════════════
// EXPORT EXCEL (SheetJS con formato)
// ═══════════════════════════════════════════════════════════════
function exportCentrosXLSX() {
  if (!ctFiltered.length) { alert('No hay datos filtrados para exportar.'); return; }

  const wb = XLSX.utils.book_new();

  // ── Hoja 1: Detalle filtrado ──────────────────────────────────
  const cols = ['cod_sismed','descripcion','TIPO','establecimiento','RED',
    'estrategic','TIPO_ESTRATEGIA','SUMINISTRO',
    'STOCK_CENTROS','CPMA_CENTROS','DISP_CENTROS','INDICADOR_CENTROS','FECHA_REPORTE'];
  const headers = ['Código','Descripción','Tipo','Establecimiento','Red',
    'Estratégico','Tipo Estrategia','Suministro',
    'Stock Centros','CPMA','Disp. Meses','Indicador','Fecha Reporte'];

  const wsData = [headers, ...ctFiltered.map(r => cols.map(c => r[c] ?? ''))];
  const ws1    = XLSX.utils.aoa_to_sheet(wsData);

  // Anchos de columna
  ws1['!cols'] = [8,42,6,28,20,12,14,10,12,10,12,14,12].map(w => ({ wch:w }));
  XLSX.utils.book_append_sheet(wb, ws1, 'Detalle');

  // ── Hoja 2: Resumen estadístico ───────────────────────────────
  const rows = ctFiltered;
  const total = rows.length;
  const c     = cntByInd(rows);
  const sob   = rows.filter(r => r.INDICADOR_CENTROS==='SOBRESTOCK' && r.DISP_CENTROS>=centrosSobMeses).length;
  const pct   = v => total>0 ? (v/total*100).toFixed(2)+'%' : '0%';

  const estabLabel = centrosSelected || 'TODOS LOS ESTABLECIMIENTOS';
  const filterLabel = centrosChartFilter ? centrosChartFilter.label : 'Sin filtro adicional';

  const sumData = [
    ['REPORTE DE DISPONIBILIDAD — CENTROS'],
    ['Establecimiento:', estabLabel],
    ['Filtro activo:', filterLabel],
    ['Fecha generación:', new Date().toLocaleString('es-PE')],
    ['Umbral sobrestock:', `≥ ${centrosSobMeses} meses`],
    [],
    ['RESUMEN POR INDICADOR', '', '', ''],
    ['Indicador', 'Cantidad', 'Porcentaje', '% acumulado'],
  ];
  let acum = 0;
  CND_ORDER.forEach(k => {
    const v = c[k]||0;
    acum += v;
    sumData.push([k, v, pct(v), total>0?(acum/total*100).toFixed(2)+'%':'0%']);
  });
  sumData.push(['TOTAL', total, '100%', '100%']);
  sumData.push([]);
  sumData.push(['SOBRESTOCK DETALLADO', '', '', '']);
  sumData.push([`Sobrestock ≥ ${centrosSobMeses} meses`, sob, pct(sob), '']);
  const sob2 = sobByMeses(rows, centrosSobMeses);
  const totS = sob2.buckets.reduce((a,b)=>a+b,0);
  sob2.labels.forEach((lbl,i) => {
    sumData.push([`  ${lbl}`, sob2.buckets[i], totS>0?(sob2.buckets[i]/totS*100).toFixed(2)+'%':'0%', '']);
  });
  sumData.push([]);
  sumData.push(['DISTRIBUCIÓN POR TIPO', '', '', '']);
  sumData.push(['Tipo', 'Cantidad', '%', '']);
  const med = rows.filter(r=>r.TIPO==='M').length;
  const ins = rows.filter(r=>r.TIPO==='I').length;
  sumData.push(['Medicamentos', med, pct(med), '']);
  sumData.push(['Insumos',      ins, pct(ins), '']);

  const ws2 = XLSX.utils.aoa_to_sheet(sumData);
  ws2['!cols'] = [30,12,12,14].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws2, 'Resumen');

  // ── Hoja 3: Top sobrestock ────────────────────────────────────
  const topSobList = top20Sob(rows, centrosSobMeses);
  if (topSobList.length) {
    const sobHdr = ['#','Código','Descripción','Establecimiento','Red','Stock','CPMA','Disp. (meses)','Indicador'];
    const sobWs  = XLSX.utils.aoa_to_sheet([
      sobHdr,
      ...topSobList.map((r,i) => [i+1, r.cod_sismed, r.descripcion, r.establecimiento,
        r.RED, r.STOCK_CENTROS, +r.CPMA_CENTROS.toFixed(1), +r.DISP_CENTROS.toFixed(1), r.INDICADOR_CENTROS])
    ]);
    sobWs['!cols'] = [4,8,42,28,20,10,10,12,14].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, sobWs, 'Top Sobrestock');
  }

  // ── Hoja 4: Desabastecidos ────────────────────────────────────
  const desabList = rows.filter(r=>r.INDICADOR_CENTROS==='DESABASTECIDO')
    .sort((a,b)=>b.CPMA_CENTROS-a.CPMA_CENTROS);
  if (desabList.length) {
    const dHdr = ['#','Código','Descripción','Establecimiento','Red','Stock','CPMA','Disp. (meses)'];
    const dWs  = XLSX.utils.aoa_to_sheet([
      dHdr,
      ...desabList.map((r,i) => [i+1, r.cod_sismed, r.descripcion, r.establecimiento,
        r.RED, r.STOCK_CENTROS, +r.CPMA_CENTROS.toFixed(1), +r.DISP_CENTROS.toFixed(1)])
    ]);
    dWs['!cols'] = [4,8,42,28,20,10,10,12].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, dWs, 'Desabastecidos');
  }

  const slug = centrosSelected ? `_${centrosSelected.replace(/\s+/g,'_').slice(0,20)}` : '_TODOS';
  const fname = `DISPO_CENTROS${slug}_${new Date().toISOString().split('T')[0]}.xlsx`;
  XLSX.writeFile(wb, fname);
}

// ═══════════════════════════════════════════════════════════════
// DESCARGA DEL ARCHIVO ORIGINAL DESDE DRIVE
// ═══════════════════════════════════════════════════════════════
async function downloadCentrosOriginal() {
  if (!centrosFileId) { alert('No se encontró el ID del archivo en Drive.'); return; }
  const btn = [...document.querySelectorAll('.btn-sm')].find(b => b.textContent.includes('Archivo Original'));
  if (btn) { btn.textContent = '⏳ Descargando…'; btn.disabled = true; }
  try {
    const url = `https://www.googleapis.com/drive/v3/files/${centrosFileId}?alt=media`;
    const res = await authFetch(url);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const blob = await res.blob();
    const a    = document.createElement('a');
    a.href     = URL.createObjectURL(blob);
    a.download = 'DISPO_CENTROS.xlsx';
    a.click();
    URL.revokeObjectURL(a.href);
  } catch(e) {
    alert('Error al descargar el archivo: ' + e.message);
  } finally {
    if (btn) { btn.textContent = '⬇ Archivo Original'; btn.disabled = false; }
  }
}

// ═══════════════════════════════════════════════════════════════
// PRINT REPORT — con imágenes de gráficos
// ═══════════════════════════════════════════════════════════════
function printCentrosReport() {
  const rows  = getCentrosRows(centrosSelected);
  const total = rows.length;
  const c     = cntByInd(rows);
  const sob   = rows.filter(r => r.INDICADOR_CENTROS==='SOBRESTOCK' && r.DISP_CENTROS>=centrosSobMeses).length;
  const pct   = v => total>0 ? (v/total*100).toFixed(1)+'%' : '0%';

  const topSobList = top20Sob(rows, centrosSobMeses).slice(0,15);
  const desabList  = rows.filter(r=>r.INDICADOR_CENTROS==='DESABASTECIDO')
    .sort((a,b)=>b.CPMA_CENTROS-a.CPMA_CENTROS).slice(0,15);

  const fecha      = new Date().toLocaleDateString('es-PE',{day:'2-digit',month:'long',year:'numeric'});
  const repFecha   = rows[0]?.FECHA_REPORTE || '—';
  const estabLabel = centrosSelected || 'TODOS LOS ESTABLECIMIENTOS';
  const filterLabel = centrosChartFilter ? ` · Filtro: ${centrosChartFilter.label}` : '';

  // Capturar gráficos como imágenes
  const imgInd      = canvasToImg('cc-ind',      560, 280);
  const imgTipo     = canvasToImg('cc-tipo',     560, 280);
  const imgSob      = canvasToImg('cc-sob',      560, 280);
  const imgCob      = canvasToImg('cc-cobertura',560, 280);
  const imgTopSob   = canvasToImg('cc-top-sob',  860, 400);
  const imgRedes    = canvasToImg('cc-redes',    860, 340);
  const imgEstabs   = canvasToImg('cc-estabs-desab', 860, 380);

  // Barra visual para stats (HTML)
  const maxVal = Math.max(...CND_ORDER.map(k => c[k]||0));
  const statBar = (k) => {
    const v    = c[k]||0;
    const fill = maxVal>0 ? Math.round(v/maxVal*100) : 0;
    return `<div class="pr-bar-row">
      <div style="width:110px;font-size:9px">${k}</div>
      <div class="pr-bar-fill" style="width:${fill}%;background:${CND_IND_COLOR[k]};flex:unset;min-width:${fill?2:0}px"></div>
      <span style="margin-left:6px;font-weight:bold">${v}</span>
      <span style="margin-left:4px;color:#888">(${pct(v)})</span>
    </div>`;
  };

  const sob2  = sobByMeses(rows, centrosSobMeses);
  const totS  = sob2.buckets.reduce((a,b)=>a+b,0);
  const maxS  = Math.max(...sob2.buckets);
  const sobBar = (v,lbl) => {
    const fill = maxS>0 ? Math.round(v/maxS*100) : 0;
    return `<div class="pr-bar-row">
      <div style="width:70px;font-size:9px">${lbl}</div>
      <div class="pr-bar-fill" style="width:${fill}%;background:#ca8a04;flex:unset;min-width:${fill?2:0}px"></div>
      <span style="margin-left:6px;font-weight:bold">${v}</span>
      <span style="margin-left:4px;color:#888">(${totS>0?(v/totS*100).toFixed(1):0}%)</span>
    </div>`;
  };

  document.getElementById('centros-print-area').innerHTML = `

    <!-- ══ PÁGINA 1: RESUMEN + GRÁFICOS ════════════════════════ -->
    <div class="pr-title">DIRESA CALLAO — DISPOALM · Disponibilidad Centros</div>
    <div class="pr-sub">
      Establecimiento: <strong>${estabLabel}</strong> &nbsp;|&nbsp;
      Fecha reporte: <strong>${repFecha}</strong> &nbsp;|&nbsp;
      Generado: ${fecha} &nbsp;|&nbsp;
      Umbral sobrestock: ≥ ${centrosSobMeses} meses${filterLabel}
    </div>

    <!-- KPIs -->
    <div class="pr-kpi-row">
      <div class="pr-kpi"><div class="pr-kpi-lbl">Total Productos</div><div class="pr-kpi-val">${total.toLocaleString()}</div></div>
      <div class="pr-kpi"><div class="pr-kpi-lbl">Desabastecidos</div><div class="pr-kpi-val red">${c.DESABASTECIDO||0}<br><small>${pct(c.DESABASTECIDO)}</small></div></div>
      <div class="pr-kpi"><div class="pr-kpi-lbl">Substock</div><div class="pr-kpi-val ora">${c.SUBSTOCK||0}<br><small>${pct(c.SUBSTOCK)}</small></div></div>
      <div class="pr-kpi"><div class="pr-kpi-lbl">Normostock</div><div class="pr-kpi-val grn">${c.NORMOSTOCK||0}<br><small>${pct(c.NORMOSTOCK)}</small></div></div>
      <div class="pr-kpi"><div class="pr-kpi-lbl">Sobrestock Total</div><div class="pr-kpi-val yel">${c.SOBRESTOCK||0}<br><small>${pct(c.SOBRESTOCK)}</small></div></div>
      <div class="pr-kpi"><div class="pr-kpi-lbl">Sobrestock ≥${centrosSobMeses}m</div><div class="pr-kpi-val yel">${sob}<br><small>${pct(sob)}</small></div></div>
      <div class="pr-kpi"><div class="pr-kpi-lbl">Sin Rotación</div><div class="pr-kpi-val vio">${c['SIN ROTACION']||0}<br><small>${pct(c['SIN ROTACION'])}</small></div></div>
    </div>

    <!-- SECCIÓN 1: Distribución e indicadores -->
    <div class="pr-section">
      <div class="pr-section-title">Distribución por Indicador</div>
      <div class="pr-charts-row">
        <div class="pr-chart-box" style="flex:1">
          <div class="pr-chart-title">Gráfico — Indicador Centros (%)</div>
          ${imgInd ? `<img class="pr-chart-img" src="${imgInd}">` : '<p style="color:#888;font-size:10px">Gráfico no disponible</p>'}
        </div>
        <div class="pr-chart-box" style="flex:1">
          <div class="pr-chart-title">Estadísticas detalladas — Indicador Centros</div>
          ${CND_ORDER.map(k => statBar(k)).join('')}
          <div style="margin-top:10px;border-top:1px solid #eee;padding-top:6px;font-size:9px;color:#555">
            <strong>Total analizado:</strong> ${total.toLocaleString()} productos
          </div>
        </div>
      </div>
    </div>

    <!-- SECCIÓN 2: Tipo de Producto -->
    <div class="pr-section">
      <div class="pr-section-title">Composición por Tipo de Producto</div>
      <div class="pr-charts-row">
        <div class="pr-chart-box" style="flex:2">
          <div class="pr-chart-title">Gráfico — Medicamentos vs Insumos por Indicador</div>
          ${imgTipo ? `<img class="pr-chart-img" src="${imgTipo}">` : ''}
        </div>
        <div class="pr-chart-box" style="flex:1">
          <div class="pr-chart-title">Detalle por Tipo</div>
          <table class="pr-stats-table">
            <thead><tr><th>Indicador</th><th>Med.</th><th>Ins.</th><th>Total</th></tr></thead>
            <tbody>
              ${CND_ORDER.map(k => {
                const med = rows.filter(r=>r.TIPO==='M'&&r.INDICADOR_CENTROS===k).length;
                const ins = rows.filter(r=>r.TIPO==='I'&&r.INDICADOR_CENTROS===k).length;
                return `<tr><td>${k}</td><td style="text-align:right">${med}</td><td style="text-align:right">${ins}</td><td style="text-align:right;font-weight:bold">${med+ins}</td></tr>`;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- SECCIÓN 3: Sobrestock y Cobertura -->
    <div class="pr-section pr-page-break">
      <div class="pr-section-title">Análisis de Sobrestock y Cobertura</div>
      <div class="pr-charts-row">
        <div class="pr-chart-box" style="flex:1">
          <div class="pr-chart-title">Sobrestock por Rangos de Meses (umbral ≥ ${centrosSobMeses}m)</div>
          ${imgSob ? `<img class="pr-chart-img" src="${imgSob}">` : ''}
          <div style="margin-top:8px">
            ${sob2.labels.map((lbl,i) => sobBar(sob2.buckets[i], lbl)).join('')}
          </div>
        </div>
        <div class="pr-chart-box" style="flex:1">
          <div class="pr-chart-title">Distribución General por Cobertura (meses)</div>
          ${imgCob ? `<img class="pr-chart-img" src="${imgCob}">` : ''}
        </div>
      </div>
    </div>

    <!-- SECCIÓN 4: Comparativo por Red (si aplica) -->
    ${imgRedes ? `
    <div class="pr-section">
      <div class="pr-section-title">Comparativo por Red de Salud</div>
      <div class="pr-charts-row">
        <div class="pr-chart-box" style="flex:1">
          <div class="pr-chart-title">% por Indicador según Red</div>
          <img class="pr-chart-img" src="${imgRedes}">
        </div>
      </div>
    </div>` : ''}

    <!-- SECCIÓN 5: Top Establecimientos desabastecidos (si aplica) -->
    ${imgEstabs ? `
    <div class="pr-section">
      <div class="pr-section-title">Top Establecimientos — % Desabastecidos</div>
      <div class="pr-chart-box">
        <img class="pr-chart-img" src="${imgEstabs}">
      </div>
    </div>` : ''}

    <!-- SECCIÓN 6: Top 20 Sobrestock -->
    <div class="pr-section pr-page-break">
      <div class="pr-section-title">Top 20 Mayor Sobrestock — Gráfico</div>
      <div class="pr-chart-box">
        ${imgTopSob ? `<img class="pr-chart-img" src="${imgTopSob}">` : '<p style="color:#888;font-size:10px">Sin datos</p>'}
      </div>
    </div>

    <!-- SECCIÓN 7: Tablas de rankings -->
    <div class="pr-section">
      <div class="pr-section-title">🟡 Top 15 Mayor Sobrestock ≥ ${centrosSobMeses} meses</div>
      <table class="pr-table">
        <thead><tr><th>#</th><th>Código</th><th>Descripción</th><th>Establecimiento</th><th>Red</th><th>Stock</th><th>CPMA</th><th>Disp.(m)</th><th>Indicador</th></tr></thead>
        <tbody>${topSobList.map((r,i)=>`
          <tr>
            <td>${i+1}</td><td>${r.cod_sismed}</td><td>${r.descripcion}</td>
            <td>${r.establecimiento}</td><td>${r.RED||'—'}</td>
            <td style="text-align:right">${r.STOCK_CENTROS.toLocaleString()}</td>
            <td style="text-align:right">${r.CPMA_CENTROS.toFixed(1)}</td>
            <td style="text-align:right;font-weight:bold;color:#ca8a04">${r.DISP_CENTROS.toFixed(1)}</td>
            <td>${r.INDICADOR_CENTROS}</td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>

    <div class="pr-section">
      <div class="pr-section-title">🔴 Top 15 Desabastecidos (Mayor CPMA)</div>
      <table class="pr-table">
        <thead><tr><th>#</th><th>Código</th><th>Descripción</th><th>Establecimiento</th><th>Red</th><th>Stock</th><th>CPMA</th><th>Disp.(m)</th></tr></thead>
        <tbody>${desabList.map((r,i)=>`
          <tr>
            <td>${i+1}</td><td>${r.cod_sismed}</td><td>${r.descripcion}</td>
            <td>${r.establecimiento}</td><td>${r.RED||'—'}</td>
            <td style="text-align:right;color:#e11d48">${r.STOCK_CENTROS.toLocaleString()}</td>
            <td style="text-align:right">${r.CPMA_CENTROS.toFixed(1)}</td>
            <td style="text-align:right;font-weight:bold;color:#e11d48">${r.DISP_CENTROS.toFixed(1)}</td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>

    <!-- REPORTE DEL FILTRADO ACTUAL (si hay filtro) -->
    ${centrosChartFilter && ctFiltered.length ? `
    <div class="pr-section pr-page-break">
      <div class="pr-section-title">📋 Reporte Detallado — ${centrosChartFilter.label} (${ctFiltered.length} productos)</div>
      <table class="pr-table">
        <thead><tr><th>#</th><th>Código</th><th>Descripción</th><th>Establecimiento</th><th>Red</th><th>Stock</th><th>CPMA</th><th>Disp.(m)</th><th>Indicador</th></tr></thead>
        <tbody>${ctFiltered.slice(0,60).map((r,i)=>`
          <tr>
            <td>${i+1}</td><td>${r.cod_sismed}</td><td>${r.descripcion}</td>
            <td>${r.establecimiento}</td><td>${r.RED||'—'}</td>
            <td style="text-align:right">${r.STOCK_CENTROS.toLocaleString()}</td>
            <td style="text-align:right">${r.CPMA_CENTROS.toFixed(1)}</td>
            <td style="text-align:right">${r.DISP_CENTROS.toFixed(1)}</td>
            <td>${r.INDICADOR_CENTROS}</td>
          </tr>`).join('')}
        </tbody>
      </table>
      ${ctFiltered.length>60 ? `<p style="font-size:9px;color:#888;margin-top:4px">Mostrando 60 de ${ctFiltered.length} registros. Usa "Export Excel" para el listado completo.</p>` : ''}
    </div>` : ''}

    <div class="pr-footer">
      DISPOALM Dashboard · DIRESA Callao · DEMID · Generado: ${new Date().toLocaleString('es-PE')} · ${total.toLocaleString()} registros procesados.
    </div>`;

  setTimeout(() => window.print(), 300);
}
