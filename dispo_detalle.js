// ═══════════════════════════════════════════════════════════════
// DISPO_CENTROS_ALM_DETALLE — Módulo por Establecimiento
// Archivo específico: DISPO_CENTROS_ALM_DETALLE.XLSX
// ═══════════════════════════════════════════════════════════════

const DETALLE_FILE_NAME = 'DISPO_CENTROS_ALM_DETALLE';
const DETALLE_FILE_ID   = ''; // Se resuelve dinámicamente desde driveFiles

// ── Estado del módulo ──────────────────────────────────────────
let detalleData      = [];   // Todos los registros parseados
let detalleEstab     = [];   // Lista de establecimientos únicos
let detalleSelected  = '';   // Establecimiento activo
let detalleSobMeses  = 6;    // Umbral sobrestock (meses)
let detalleCharts    = {};   // Chart.js instances
let detalleView      = 'centros'; // 'centros' | 'alm' | 'ambos'

// ── Indicadores ────────────────────────────────────────────────
const IND_ORDER = ['DESABASTECIDO','SUBSTOCK','NORMOSTOCK','SOBRESTOCK','SIN ROTACION'];
const IND_COLORS = {
  DESABASTECIDO : { bg: 'rgba(244,63,94,0.85)',   border: '#f43f5e' },
  SUBSTOCK      : { bg: 'rgba(249,115,22,0.85)',  border: '#f97316' },
  NORMOSTOCK    : { bg: 'rgba(34,197,94,0.85)',   border: '#22c55e' },
  SOBRESTOCK    : { bg: 'rgba(234,179,8,0.85)',   border: '#eab308' },
  'SIN ROTACION': { bg: 'rgba(139,92,246,0.85)',  border: '#8b5cf6' },
};
const IND_PILL_MAP = {
  DESABASTECIDO:'red', SUBSTOCK:'orange',
  NORMOSTOCK:'green',  SOBRESTOCK:'yellow', 'SIN ROTACION':'violet'
};

// ═══════════════════════════════════════════════════════════════
// PARSE — Lee DISPO_CENTROS_ALM_DETALLE desde un ArrayBuffer
// ═══════════════════════════════════════════════════════════════
function parseDetalleBuffer(buf) {
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  const sheetName = wb.SheetNames.find(n => /DISP.*CENTROS|SINC/i.test(n)) || wb.SheetNames[0];
  const ws  = wb.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  // Encontrar fila de encabezado
  let headerRow = -1;
  for (let i = 0; i < Math.min(30, raw.length); i++) {
    const cols = (raw[i] || []).filter(Boolean).map(c => String(c).toLowerCase());
    if (cols.some(c => c.includes('establecimiento') || c.includes('indicador'))) {
      headerRow = i; break;
    }
  }
  if (headerRow === -1) throw new Error('No se encontró encabezado en DISPO_CENTROS_ALM_DETALLE');

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
    const r = raw[i] || [];
    const est = String(r[colIdx['establecimiento']] || '').trim();
    if (!est) continue;

    const obj = {
      codigo_pre       : String(r[colIdx['codigo_pre']]        || '').trim(),
      establecimiento  : est,
      RED              : String(r[colIdx['RED']]               || '').trim(),
      cod_sismed       : String(r[colIdx['cod_sismed']]        || '').trim(),
      descripcion      : String(r[colIdx['descripcion']]       || '').trim(),
      TIPO             : String(r[colIdx['TIPO']]              || '').trim().toUpperCase(),
      estrategic       : String(r[colIdx['estrategic']]        || '').trim(),
      SUMINISTRO       : String(r[colIdx['SUMINISTRO']]        || '').trim(),
      // Centros
      STOCK_CENTROS    : toNum(r[colIdx['STOCK_CENTROS']]),
      CPMA_CENTROS     : toNum(r[colIdx['CPMA_CENTROS']]),
      DISP_CENTROS     : toNum(r[colIdx['DISP_CENTROS']]),
      INDICADOR_CENTROS: String(r[colIdx['INDICADOR_CENTROS']] || '').trim().toUpperCase(),
      // Almacén
      STOCK_ALM        : toNum(r[colIdx['STOCK_ALM']]),
      CMPA_ALM         : toNum(r[colIdx['CMPA_ALM']]),
      DIP_ALM          : toNum(r[colIdx['DIP_ALM']]),
      INDICADOR_ALM    : String(r[colIdx['INDICADOR_ALM']]     || '').trim().toUpperCase(),
      // Fecha
      FECHA_REPORTE    : r[colIdx['FECHA_REPORTE']] instanceof Date
        ? r[colIdx['FECHA_REPORTE']].toISOString().split('T')[0]
        : String(r[colIdx['FECHA_REPORTE']] || ''),
    };
    rows.push(obj);
  }
  return rows;
}

// ═══════════════════════════════════════════════════════════════
// CARGAR ARCHIVO DESDE DRIVE
// ═══════════════════════════════════════════════════════════════
async function loadDetalleFile() {
  const file = driveFiles.find(f =>
    f.name.toUpperCase().includes(DETALLE_FILE_NAME) ||
    f.name.toUpperCase().includes('DISPO_CENTROS_ALM_DETALLE')
  );
  if (!file) return null;

  if (fileDataCache['__detalle__']) return fileDataCache['__detalle__'];

  const url = `https://www.googleapis.com/drive/v3/files/${file.id}?alt=media`;
  const res  = await authFetch(url);
  if (!res.ok) throw new Error(`HTTP ${res.status} al descargar DISPO_CENTROS_ALM_DETALLE`);
  const buf  = await res.arrayBuffer();
  const data = parseDetalleBuffer(buf);

  fileDataCache['__detalle__'] = data;
  return data;
}

// ═══════════════════════════════════════════════════════════════
// DATOS FILTRADOS POR ESTABLECIMIENTO
// ═══════════════════════════════════════════════════════════════
function getDetalleForEstab(estab) {
  if (!estab) return detalleData;
  return detalleData.filter(d => d.establecimiento === estab);
}

function countByIndicador(rows, field) {
  const counts = {};
  IND_ORDER.forEach(k => { counts[k] = 0; });
  rows.forEach(r => {
    const k = r[field];
    if (counts[k] !== undefined) counts[k]++;
    else counts[k] = (counts[k] || 0) + 1;
  });
  return counts;
}

function getSobrestockByMeses(rows, field, umbral) {
  // Distribución sobrestock: < umbral, umbral-12, 12-24, >24
  const labels = [`6–${umbral}m`, `${umbral}–12m`, `12–24m`, `>24m`];
  const buckets = [0, 0, 0, 0];
  rows.filter(r => r[field] === 'SOBRESTOCK').forEach(r => {
    const meses = field === 'INDICADOR_CENTROS' ? r.DISP_CENTROS : r.DIP_ALM;
    if (meses < umbral)       buckets[0]++;
    else if (meses < 12)      buckets[1]++;
    else if (meses < 24)      buckets[2]++;
    else                      buckets[3]++;
  });
  return { labels, buckets };
}

function getTop20Sobrestock(rows, field, mesesField, umbral) {
  return rows
    .filter(r => r[field] === 'SOBRESTOCK' && r[mesesField] >= umbral)
    .sort((a, b) => b[mesesField] - a[mesesField])
    .slice(0, 20);
}

// ═══════════════════════════════════════════════════════════════
// RENDER — PANEL PRINCIPAL DE ESTABLECIMIENTOS
// ═══════════════════════════════════════════════════════════════
let detalleFileId = '';

function renderDetalleDashboard(data) {
  detalleData  = data;
  detalleEstab = [...new Set(data.map(d => d.establecimiento))].sort();
  detalleSelected = '';

  const fileObj = (typeof driveFiles !== 'undefined') &&
    driveFiles.find(f => /DISPO_CENTROS_ALM_DETALLE/i.test(f.name));
  detalleFileId = fileObj?.id || '';

  const fecha = data[0]?.FECHA_REPORTE || '—';
  document.getElementById('tb-meta').textContent =
    `DISPO_CENTROS_ALM_DETALLE · ${data.length.toLocaleString()} registros · Reporte: ${fecha}`;

  document.getElementById('dash-content').innerHTML = buildDetalleHTML();
  bindDetalleEvents();
  renderDetalleView();
}

function buildDetalleHTML() {
  const estabOptions = detalleEstab.map(e =>
    `<option value="${e}"${e === detalleSelected ? ' selected' : ''}>${e}</option>`
  ).join('');

  return `
  <div class="main" id="detalle-main">
    <!-- CONTROLES -->
    <div class="detalle-controls-bar">
      <div class="dc-group">
        <label class="dc-label">📍 Establecimiento</label>
        <select class="inp dc-select" id="dc-estab" onchange="onEstabChange(this.value)">
          <option value="">— TODOS —</option>
          ${estabOptions}
        </select>
      </div>
      <div class="dc-group">
        <label class="dc-label">🔭 Vista</label>
        <div class="dc-segmented" id="dc-view">
          <button class="seg-btn active" data-v="centros" onclick="onViewChange('centros',this)">Centros</button>
          <button class="seg-btn" data-v="alm" onclick="onViewChange('alm',this)">Almacén</button>
          <button class="seg-btn" data-v="ambos" onclick="onViewChange('ambos',this)">Ambos</button>
        </div>
      </div>
      <div class="dc-group">
        <label class="dc-label">⚠️ Umbral sobrestock</label>
        <div style="display:flex;align-items:center;gap:8px">
          <input type="range" id="dc-umbral" min="6" max="36" step="1" value="6"
            oninput="onUmbralChange(+this.value)" style="width:100px;accent-color:var(--accent)">
          <span class="dc-umbral-val" id="dc-umbral-val">≥ 6 meses</span>
        </div>
      </div>
      <div style="margin-left:auto;display:flex;gap:8px;flex-wrap:wrap">
        <button class="btn-sm" onclick="downloadDetalleOriginal()" title="Descargar archivo XLSX original desde Drive">⬇ Archivo Original</button>
        <button class="btn-sm" onclick="exportDetalleXLSX()">📊 Export Excel</button>
        <button class="btn-sm" onclick="printDetalleReport()">🖨 Imprimir Reporte</button>
      </div>
    </div>

    <!-- FILTRO ACTIVO BADGE -->
    <div id="dd-filter-badge" style="display:none">
      <div style="display:flex;align-items:center;gap:10px;background:rgba(108,99,255,.12);border:1px solid rgba(108,99,255,.3);border-radius:10px;padding:8px 14px;">
        <span style="font-size:12px;color:var(--accent2)">🔍 Filtro activo:</span>
        <span id="dd-filter-label" style="font-size:12px;color:var(--text);font-weight:600"></span>
        <button class="btn-sm" style="margin-left:auto" onclick="clearDetalleChartFilter()">✕ Limpiar filtro</button>
      </div>
    </div>

    <!-- KPIs -->
    <div>
      <div class="sec-hdr"><span class="sec-title">Resumen de Disponibilidad</span><div class="sec-line"></div></div>
      <div class="kpi-grid" id="dk-grid"></div>
    </div>

    <!-- GRÁFICOS INTERACTIVOS -->
    <div class="g2" id="d-charts-row1">
      <div class="card">
        <div class="card-hdr">
          <div>
            <div class="card-title" id="dct1-title">Indicador Centros — Distribución %</div>
            <div class="card-sub" id="dct1-sub">Haz clic en un segmento para filtrar la tabla</div>
          </div>
        </div>
        <div class="chart-h260"><canvas id="dc-ind-centros"></canvas></div>
      </div>
      <div class="card">
        <div class="card-hdr">
          <div>
            <div class="card-title" id="dct2-title">Indicador Almacén — Distribución %</div>
            <div class="card-sub" id="dct2-sub">Distribución por indicador de almacén</div>
          </div>
        </div>
        <div class="chart-h260"><canvas id="dc-ind-alm"></canvas></div>
      </div>
    </div>

    <!-- SOBRESTOCK ANALYSIS -->
    <div class="g2" id="d-charts-row2">
      <div class="card">
        <div class="card-hdr">
          <div>
            <div class="card-title">Sobrestock Centros — Distribución por Meses</div>
            <div class="card-sub" id="dcsob-sub">Agrupado según umbral seleccionado</div>
          </div>
        </div>
        <div class="chart-h260"><canvas id="dc-sob-centros"></canvas></div>
      </div>
      <div class="card">
        <div class="card-hdr">
          <div>
            <div class="card-title">Sobrestock Almacén — Distribución por Meses</div>
            <div class="card-sub">Productos con exceso de stock en almacén central</div>
          </div>
        </div>
        <div class="chart-h260"><canvas id="dc-sob-alm"></canvas></div>
      </div>
    </div>

    <!-- COMPARATIVO ENTRE REDES (solo modo TODOS) -->
    <div class="card" id="d-redes-card" style="display:none">
      <div class="card-hdr">
        <div>
          <div class="card-title">Comparativo por Red de Salud</div>
          <div class="card-sub">% Desabastecidos y Sobrestock por red</div>
        </div>
      </div>
      <div class="chart-h320"><canvas id="dc-redes"></canvas></div>
    </div>

    <!-- TOP SOBRESTOCK -->
    <div class="card">
      <div class="card-hdr">
        <div>
          <div class="card-title" id="d-topsob-title">Top 20 Mayor Sobrestock (≥ umbral)</div>
          <div class="card-sub" id="d-topsob-sub">Productos con más meses de cobertura en sobrestock</div>
        </div>
      </div>
      <div class="chart-h360"><canvas id="dc-top-sob"></canvas></div>
    </div>

    <!-- TABLA DETALLE -->
    <div>
      <div class="sec-hdr"><span class="sec-title">Catálogo Detallado</span><div class="sec-line"></div></div>
      <div class="tbl-card">
        <div class="tbl-bar">
          <span class="tbl-bar-title">Productos por Establecimiento</span>
          <input class="inp" type="text" id="dt-search" placeholder="🔍 Buscar producto o código…"
            style="width:220px" oninput="filterDetalleTable()">
          <select class="inp" id="dt-ind-c" onchange="filterDetalleTable()">
            <option value="">Centros: Todos</option>
            ${IND_ORDER.map(i=>`<option>${i}</option>`).join('')}
          </select>
          <select class="inp" id="dt-ind-a" onchange="filterDetalleTable()">
            <option value="">Almacén: Todos</option>
            ${IND_ORDER.map(i=>`<option>${i}</option>`).join('')}
          </select>
          <select class="inp" id="dt-tipo" onchange="filterDetalleTable()">
            <option value="">Todos los tipos</option>
            <option value="M">Medicamentos</option>
            <option value="I">Insumos</option>
          </select>
          <span class="tbl-count" id="dt-count">0 productos</span>
          <button class="btn-sm" onclick="exportDetalleXLSX()" title="Exportar filtrado actual a Excel">📊 Excel filtrado</button>
        </div>
        <div class="tbl-wrap">
          <table>
            <thead>
              <tr>
                <th onclick="sortDetalleTable('cod_sismed')">Código</th>
                <th onclick="sortDetalleTable('descripcion')">Descripción</th>
                <th onclick="sortDetalleTable('TIPO')">Tipo</th>
                <th onclick="sortDetalleTable('establecimiento')" id="dt-estab-col">Establecimiento</th>
                <th onclick="sortDetalleTable('RED')">Red</th>
                <th onclick="sortDetalleTable('STOCK_CENTROS')" style="text-align:right">Stock Centro</th>
                <th onclick="sortDetalleTable('DISP_CENTROS')" style="text-align:right">Disp.C (m)</th>
                <th onclick="sortDetalleTable('INDICADOR_CENTROS')">Ind. Centro</th>
                <th onclick="sortDetalleTable('STOCK_ALM')" style="text-align:right">Stock Alm</th>
                <th onclick="sortDetalleTable('DIP_ALM')" style="text-align:right">Disp.A (m)</th>
                <th onclick="sortDetalleTable('INDICADOR_ALM')">Ind. Almacén</th>
              </tr>
            </thead>
            <tbody id="dt-body"></tbody>
          </table>
        </div>
        <div class="tbl-pager">
          <button class="btn-sm" onclick="prevDetallePage()">← Ant</button>
          <span id="dt-pager">Pág. 1 / 1</span>
          <button class="btn-sm" onclick="nextDetallePage()">Sig →</button>
        </div>
      </div>
    </div>

    <!-- PRINT AREA (oculto, visible solo al imprimir) -->
    <div id="detalle-print-area" style="display:none"></div>
  </div>`;
}

// ── Estilos extra para el módulo ─────────────────────────────
function injectDetalleStyles() {
  if (document.getElementById('detalle-styles')) return;
  const s = document.createElement('style');
  s.id = 'detalle-styles';
  s.textContent = `
.detalle-controls-bar{
  display:flex;align-items:flex-end;flex-wrap:wrap;gap:16px;
  background:var(--surf);border:1px solid var(--border);
  border-radius:16px;padding:16px 20px;
}
.dc-group{display:flex;flex-direction:column;gap:5px;}
.dc-label{font-size:10px;color:var(--text);text-transform:uppercase;letter-spacing:.8px;}
.dc-select{min-width:220px;}
.dc-segmented{display:flex;border:1px solid var(--border);border-radius:8px;overflow:hidden;}
.seg-btn{
  padding:7px 16px;background:transparent;border:none;color:var(--muted2);
  font-size:12px;font-family:'DM Sans',sans-serif;cursor:pointer;transition:all .15s;
}
.seg-btn:not(:last-child){border-right:1px solid var(--border);}
.seg-btn.active{background:var(--accent);color:#fff;}
.dc-umbral-val{font-family:'Space Mono',monospace;font-size:11px;color:var(--accent);white-space:nowrap;}
#dc-umbral{accent-color:var(--accent);}

/* Print styles */
@media print {
  body * { visibility: hidden; }
  #detalle-print-area, #detalle-print-area * { visibility: visible; }
  #detalle-print-area {
    position: fixed; top: 0; left: 0; width: 100%;
    display: block !important;
    background: #fff; color: #000;
    padding: 20px; font-family: Arial, sans-serif;
  }
  .print-header { margin-bottom: 20px; }
  .print-title { font-size: 18px; font-weight: bold; margin-bottom: 4px; }
  .print-sub { font-size: 12px; color: #666; }
  .print-kpi-row { display: flex; gap: 12px; margin: 16px 0; flex-wrap: wrap; }
  .print-kpi { border: 1px solid #ddd; border-radius: 8px; padding: 12px 16px; min-width: 130px; }
  .print-kpi-lbl { font-size: 10px; color: #888; text-transform: uppercase; margin-bottom: 4px; }
  .print-kpi-val { font-size: 22px; font-weight: bold; }
  .print-kpi-val.red { color: #e11d48; }
  .print-kpi-val.orange { color: #ea580c; }
  .print-kpi-val.yellow { color: #ca8a04; }
  .print-kpi-val.green { color: #16a34a; }
  .print-charts-row { display: flex; gap: 16px; margin: 16px 0; page-break-inside: avoid; }
  .print-chart-box { flex: 1; border: 1px solid #eee; border-radius: 8px; padding: 12px; }
  .print-chart-title { font-size: 12px; font-weight: bold; margin-bottom: 8px; }
  .print-bar-row { display: flex; align-items: center; gap: 8px; margin: 3px 0; font-size: 10px; }
  .print-bar { height: 14px; border-radius: 3px; }
  .print-table { width: 100%; border-collapse: collapse; margin-top: 16px; font-size: 10px; page-break-inside: auto; }
  .print-table th { background: #f3f4f6; padding: 6px 8px; text-align: left; border: 1px solid #e5e7eb; }
  .print-table td { padding: 5px 8px; border: 1px solid #e5e7eb; }
  .print-table tr:nth-child(even) td { background: #f9fafb; }
  .print-footer { margin-top: 20px; font-size: 10px; color: #999; border-top: 1px solid #eee; padding-top: 8px; }
  .print-sob-section { margin-top: 20px; page-break-inside: avoid; }
}`;
  document.head.appendChild(s);
}

// ═══════════════════════════════════════════════════════════════
// BIND EVENTS
// ═══════════════════════════════════════════════════════════════
function bindDetalleEvents() {
  injectDetalleStyles();
}

function onEstabChange(val) {
  detalleSelected = val;
  renderDetalleView();
}

function onViewChange(v, btn) {
  detalleView = v;
  document.querySelectorAll('.seg-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  renderDetalleView();
}

function onUmbralChange(v) {
  detalleSobMeses = v;
  document.getElementById('dc-umbral-val').textContent = `≥ ${v} meses`;
  renderDetalleView();
}

// ═══════════════════════════════════════════════════════════════
// RENDER VIEW
// ═══════════════════════════════════════════════════════════════
function renderDetalleView() {
  const rows = getDetalleForEstab(detalleSelected);

  // Destroy charts
  Object.values(detalleCharts).forEach(c => { try { c.destroy(); } catch(e){} });
  detalleCharts = {};

  renderDetalleKPIs(rows);
  renderDetalleCharts(rows);
  filterDetalleTable();

  // Comparativo redes solo cuando es TODOS
  const redasCard = document.getElementById('d-redes-card');
  if (redasCard) redasCard.style.display = detalleSelected ? 'none' : 'block';
  if (!detalleSelected) renderRedesChart();
}

// ─── KPIs ──────────────────────────────────────────────────────
function renderDetalleKPIs(rows) {
  const total = rows.length;
  const cC = countByIndicador(rows, 'INDICADOR_CENTROS');
  const cA = countByIndicador(rows, 'INDICADOR_ALM');
  const pct = (v) => total > 0 ? (v / total * 100).toFixed(1) + '%' : '0%';
  const sobC = rows.filter(r => r.INDICADOR_CENTROS === 'SOBRESTOCK' && r.DISP_CENTROS >= detalleSobMeses).length;
  const sobA = rows.filter(r => r.INDICADOR_ALM    === 'SOBRESTOCK' && r.DIP_ALM >= detalleSobMeses).length;

  const kpis = [
    { icon:'🧾', lbl:'Total Productos',   val: total.toLocaleString(),          sub:'registros',              cls:'c-cyan'   },
    { icon:'🔴', lbl:'Desab. Centros',    val: cC.DESABASTECIDO||0,             sub: pct(cC.DESABASTECIDO),   cls:'c-red'    },
    { icon:'🔴', lbl:'Desab. Almacén',    val: cA.DESABASTECIDO||0,             sub: pct(cA.DESABASTECIDO),   cls:'c-red'    },
    { icon:'🟡', lbl:`Sobrestock C ≥${detalleSobMeses}m`,  val: sobC,           sub: pct(sobC),               cls:'c-yellow' },
    { icon:'🟡', lbl:`Sobrestock A ≥${detalleSobMeses}m`,  val: sobA,           sub: pct(sobA),               cls:'c-yellow' },
    { icon:'🟢', lbl:'Normostock C',      val: cC.NORMOSTOCK||0,                sub: pct(cC.NORMOSTOCK),      cls:'c-green'  },
    { icon:'🟠', lbl:'Substock C',        val: cC.SUBSTOCK||0,                  sub: pct(cC.SUBSTOCK),        cls:'c-orange' },
    { icon:'🟣', lbl:'Sin Rotación C',    val: cC['SIN ROTACION']||0,           sub: pct(cC['SIN ROTACION']), cls:'c-violet' },
  ];

  document.getElementById('dk-grid').innerHTML = kpis.map(k => `
    <div class="kpi ${k.cls}">
      <div class="kpi-icon">${k.icon}</div>
      <div class="kpi-lbl">${k.lbl}</div>
      <div class="kpi-val">${k.val}</div>
      <div class="kpi-sub">${k.sub}</div>
    </div>`).join('');
}

// ─── CHARTS ────────────────────────────────────────────────────
const CJ_DEFAULTS = {
  plugins: { legend: { labels: { color: '#1a2745', font: { size: 11 } } } },
  scales: {}
};

function mkDoughnut(id, labels, values, title) {
  const ctx = document.getElementById(id);
  if (!ctx) return;
  const colors = labels.map(l => IND_COLORS[l]?.bg || 'rgba(136,137,166,0.7)');
  const borders = labels.map(l => IND_COLORS[l]?.border || '#8889a6');
  const total = values.reduce((a, b) => a + b, 0);

  detalleCharts[id] = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: labels.map((l, i) => `${l} (${total > 0 ? (values[i]/total*100).toFixed(1) : 0}%)`),
      datasets: [{ data: values, backgroundColor: colors, borderColor: borders, borderWidth: 2, hoverOffset: 8 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '60%',
      plugins: {
        legend: { position: 'right', labels: { color: '#1a2745', font: { size: 10 }, padding: 10 } },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const pct = total > 0 ? (ctx.raw / total * 100).toFixed(1) : 0;
              return ` ${ctx.raw.toLocaleString()} productos (${pct}%)`;
            }
          }
        }
      },
      onClick: (e, elems) => {
        if (!elems.length) return;
        const label = labels[elems[0].index];
        applyIndicadorFilter('INDICADOR_CENTROS', label);
      }
    }
  });
}

function mkDoughnutAlm(id, labels, values) {
  const ctx = document.getElementById(id);
  if (!ctx) return;
  const colors = labels.map(l => IND_COLORS[l]?.bg || 'rgba(136,137,166,0.7)');
  const borders = labels.map(l => IND_COLORS[l]?.border || '#8889a6');
  const total = values.reduce((a, b) => a + b, 0);

  detalleCharts[id] = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: labels.map((l, i) => `${l} (${total > 0 ? (values[i]/total*100).toFixed(1) : 0}%)`),
      datasets: [{ data: values, backgroundColor: colors, borderColor: borders, borderWidth: 2, hoverOffset: 8 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '60%',
      plugins: {
        legend: { position: 'right', labels: { color: '#1a2745', font: { size: 10 }, padding: 10 } },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const pct = total > 0 ? (ctx.raw / total * 100).toFixed(1) : 0;
              return ` ${ctx.raw.toLocaleString()} productos (${pct}%)`;
            }
          }
        }
      },
      onClick: (e, elems) => {
        if (!elems.length) return;
        const label = labels[elems[0].index];
        applyIndicadorFilter('INDICADOR_ALM', label);
      }
    }
  });
}

function mkBarSobrestock(id, labels, buckets, color) {
  const ctx = document.getElementById(id);
  if (!ctx) return;
  const total = buckets.reduce((a, b) => a + b, 0);

  detalleCharts[id] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Productos',
        data: buckets,
        backgroundColor: color,
        borderColor: color.replace('0.85', '1'),
        borderWidth: 2, borderRadius: 6
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const pct = total > 0 ? (ctx.raw / total * 100).toFixed(1) : 0;
              return ` ${ctx.raw} productos (${pct}% del sobrestock)`;
            }
          }
        }
      },
      scales: {
        x: { ticks: { color: '#8889a6', font: { size: 11 } }, grid: { color: 'rgba(255,255,255,0.04)' } },
        y: { ticks: { color: '#8889a6', font: { size: 11 } }, grid: { color: 'rgba(255,255,255,0.06)' }, beginAtZero: true }
      }
    }
  });
}

function mkTopSobBar(id, items, mesesField, label) {
  const ctx = document.getElementById(id);
  if (!ctx) return;
  const labels = items.map(r => r.descripcion.length > 30 ? r.descripcion.slice(0,28)+'…' : r.descripcion);
  const vals   = items.map(r => +(r[mesesField]).toFixed(1));
  const bgs    = vals.map(v => v >= 24 ? 'rgba(244,63,94,0.8)' : v >= 12 ? 'rgba(249,115,22,0.8)' : 'rgba(234,179,8,0.8)');

  detalleCharts[id] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Meses de cobertura',
        data: vals,
        backgroundColor: bgs,
        borderRadius: 5
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            afterLabel: (ctx) => {
              const r = items[ctx.dataIndex];
              return `${r.establecimiento}\nStock: ${r.STOCK_CENTROS?.toLocaleString() || r.STOCK_ALM?.toLocaleString()}`;
            }
          }
        }
      },
      scales: {
        x: {
          ticks: { color: '#1a2745', font: { size: 10 } },
          grid: { color: 'rgba(0,0,0,0.06)' },
          title: { display: true, text: 'Meses de cobertura', color: '#1a2745', font: { size: 10 } }
        },
        y: { ticks: { color: '#0d1b2a', font: { size: 10 } }, grid: { color: 'rgba(0,0,0,0.03)' } }
      }
    }
  });
}

function renderDetalleCharts(rows) {
  const cC  = countByIndicador(rows, 'INDICADOR_CENTROS');
  const cA  = countByIndicador(rows, 'INDICADOR_ALM');
  const cLbls = IND_ORDER.filter(k => (cC[k] || 0) > 0);
  const aLbls = IND_ORDER.filter(k => (cA[k] || 0) > 0);

  // Doughnuts indicador
  mkDoughnut('dc-ind-centros', cLbls, cLbls.map(k => cC[k] || 0));
  mkDoughnutAlm('dc-ind-alm', aLbls, aLbls.map(k => cA[k] || 0));

  // Sobrestock por meses
  const sobC = getSobrestockByMeses(rows, 'INDICADOR_CENTROS', detalleSobMeses);
  const sobA = getSobrestockByMeses(rows, 'INDICADOR_ALM',     detalleSobMeses);
  document.getElementById('dcsob-sub').textContent =
    `Umbral actual: ≥ ${detalleSobMeses} meses (arrastra el slider para cambiar)`;
  mkBarSobrestock('dc-sob-centros', sobC.labels, sobC.buckets, 'rgba(234,179,8,0.8)');
  mkBarSobrestock('dc-sob-alm',     sobA.labels, sobA.buckets, 'rgba(249,115,22,0.8)');

  // Top sobrestock
  const mField = detalleView === 'alm' ? 'INDICADOR_ALM' : 'INDICADOR_CENTROS';
  const mMeses = detalleView === 'alm' ? 'DIP_ALM' : 'DISP_CENTROS';
  const topSob = getTop20Sobrestock(rows, mField, mMeses, detalleSobMeses);
  document.getElementById('d-topsob-title').textContent =
    `Top 20 Sobrestock ≥ ${detalleSobMeses} meses (${detalleView === 'alm' ? 'Almacén' : 'Centros'})`;
  document.getElementById('d-topsob-sub').textContent =
    `${topSob.length} productos encontrados con criterio actual`;
  if (topSob.length > 0) {
    mkTopSobBar('dc-top-sob', topSob, mMeses, 'Meses');
  } else {
    const ctx = document.getElementById('dc-top-sob');
    if (ctx) ctx.parentElement.innerHTML =
      '<div style="display:flex;align-items:center;justify-content:center;height:360px;color:var(--muted);font-size:13px">No hay productos en sobrestock con ese umbral</div>';
  }
}

function renderRedesChart() {
  const ctx = document.getElementById('dc-redes');
  if (!ctx) return;
  if (detalleCharts['dc-redes']) { detalleCharts['dc-redes'].destroy(); }

  const redes = [...new Set(detalleData.map(d => d.RED))].filter(Boolean);
  const desabCentros = redes.map(r => {
    const sub = detalleData.filter(d => d.RED === r);
    return sub.length > 0 ? +(sub.filter(d => d.INDICADOR_CENTROS === 'DESABASTECIDO').length / sub.length * 100).toFixed(1) : 0;
  });
  const sobCentros = redes.map(r => {
    const sub = detalleData.filter(d => d.RED === r);
    return sub.length > 0 ? +(sub.filter(d => d.INDICADOR_CENTROS === 'SOBRESTOCK').length / sub.length * 100).toFixed(1) : 0;
  });
  const desabAlm = redes.map(r => {
    const sub = detalleData.filter(d => d.RED === r);
    return sub.length > 0 ? +(sub.filter(d => d.INDICADOR_ALM === 'DESABASTECIDO').length / sub.length * 100).toFixed(1) : 0;
  });

  detalleCharts['dc-redes'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: redes,
      datasets: [
        { label: '% Desab. Centros',  data: desabCentros, backgroundColor: 'rgba(244,63,94,0.8)',  borderRadius: 4 },
        { label: '% Sobrestock C',    data: sobCentros,   backgroundColor: 'rgba(234,179,8,0.8)',   borderRadius: 4 },
        { label: '% Desab. Almacén',  data: desabAlm,     backgroundColor: 'rgba(249,115,22,0.8)', borderRadius: 4 },
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#1a2745', font: { size: 11 } } },
        tooltip: { callbacks: { label: c => ` ${c.raw}%` } }
      },
      scales: {
        x: { ticks: { color: '#1a2745', font: { size: 11 } }, grid: { color: 'rgba(0,0,0,0.04)' } },
        y: {
          ticks: { color: '#1a2745', font: { size: 11 }, callback: v => v + '%' },
          grid: { color: 'rgba(0,0,0,0.06)' }, beginAtZero: true, max: 100
        }
      }
    }
  });
}


// ═══════════════════════════════════════════════════════════════
// TABLA DETALLE
// ═══════════════════════════════════════════════════════════════
let dtFiltered = [], dtPage = 1, dtSortCol = 'DISP_CENTROS', dtSortDir = -1;
const DT_PAGE = 50;

function filterDetalleTable() {
  const rows  = getDetalleForEstab(detalleSelected);
  const s     = (document.getElementById('dt-search')?.value || '').toLowerCase();
  const indC  = document.getElementById('dt-ind-c')?.value  || '';
  const indA  = document.getElementById('dt-ind-a')?.value  || '';
  const tipo  = document.getElementById('dt-tipo')?.value   || '';

  dtFiltered = rows.filter(r => {
    if (indC && r.INDICADOR_CENTROS !== indC) return false;
    if (indA && r.INDICADOR_ALM     !== indA) return false;
    if (tipo && r.TIPO !== tipo) return false;
    if (s && !r.descripcion.toLowerCase().includes(s) && !r.cod_sismed.includes(s)) return false;
    return true;
  });

  dtFiltered.sort((a, b) => {
    const av = a[dtSortCol], bv = b[dtSortCol];
    if (typeof av === 'number') return (av - bv) * dtSortDir;
    return String(av).localeCompare(String(bv)) * dtSortDir;
  });
  dtPage = 1;
  renderDetalleTable();
}

function sortDetalleTable(col) {
  if (dtSortCol === col) dtSortDir *= -1; else { dtSortCol = col; dtSortDir = -1; }
  filterDetalleTable();
}

function renderDetalleTable() {
  const total  = dtFiltered.length;
  const totalP = Math.ceil(total / DT_PAGE) || 1;
  const slice  = dtFiltered.slice((dtPage - 1) * DT_PAGE, dtPage * DT_PAGE);
  const showEstab = !detalleSelected;

  document.getElementById('dt-count').textContent = total.toLocaleString() + ' productos';
  document.getElementById('dt-pager').textContent = `Pág. ${dtPage} / ${totalP}`;

  const mesesBar = (v, field) => {
    const fill = Math.min(v / 12 * 100, 100);
    const col  = v === 0 ? 'var(--red)' : v < 1 ? 'var(--orange)' : v < 3 ? 'var(--yellow)' : v < 6 ? 'var(--green)' : 'var(--yellow)';
    return `<div class="prog-row" style="min-width:80px">
      <div class="prog-bar"><div class="prog-fill" style="width:${fill}%;background:${col}"></div></div>
      <span class="prog-val">${v.toFixed(1)}</span></div>`;
  };

  document.getElementById('dt-body').innerHTML = !slice.length
    ? `<tr><td colspan="11" style="text-align:center;padding:36px;color:var(--muted)">Sin resultados</td></tr>`
    : slice.map(r => `
    <tr>
      <td class="mono">${r.cod_sismed}</td>
      <td style="max-width:240px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${r.descripcion}">${r.descripcion}</td>
      <td><span class="pill ${r.TIPO === 'M' ? 'pill-blue' : 'pill-violet'}">${r.TIPO === 'M' ? 'Med.' : 'Ins.'}</span></td>
      <td style="font-size:11px;color:var(--muted2);max-width:160px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${r.establecimiento}</td>
      <td style="font-size:10px;color:var(--muted)">${r.RED || '—'}</td>
      <td class="mono" style="text-align:right">${r.STOCK_CENTROS.toLocaleString()}</td>
      <td>${mesesBar(r.DISP_CENTROS, 'DISP_CENTROS')}</td>
      <td><span class="pill pill-${IND_PILL_MAP[r.INDICADOR_CENTROS]||'gray'}">${r.INDICADOR_CENTROS||'—'}</span></td>
      <td class="mono" style="text-align:right">${r.STOCK_ALM.toLocaleString()}</td>
      <td>${mesesBar(Math.max(0, r.DIP_ALM), 'DIP_ALM')}</td>
      <td><span class="pill pill-${IND_PILL_MAP[r.INDICADOR_ALM]||'gray'}">${r.INDICADOR_ALM||'—'}</span></td>
    </tr>`).join('');
}

function prevDetallePage() { if (dtPage > 1) { dtPage--; renderDetalleTable(); } }
function nextDetallePage() {
  const tp = Math.ceil(dtFiltered.length / DT_PAGE);
  if (dtPage < tp) { dtPage++; renderDetalleTable(); }
}

// ═══════════════════════════════════════════════════════════════
// CSV EXPORT
// ═══════════════════════════════════════════════════════════════
function exportDetalleCSV() {
  const cols = ['cod_sismed','descripcion','TIPO','establecimiento','RED',
    'STOCK_CENTROS','CPMA_CENTROS','DISP_CENTROS','INDICADOR_CENTROS',
    'STOCK_ALM','CMPA_ALM','DIP_ALM','INDICADOR_ALM','FECHA_REPORTE'];
  const rows = [cols.join(','), ...dtFiltered.map(r =>
    cols.map(c => `"${String(r[c] ?? '').replace(/"/g, '""')}"`).join(',')
  )];
  const blob = new Blob([rows.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  const estabSlug = detalleSelected ? `_${detalleSelected.replace(/\s+/g,'_').slice(0,20)}` : '_TODOS';
  a.download = `DISPO_DETALLE${estabSlug}_${new Date().toISOString().split('T')[0]}.csv`;
  a.click();
}

// ═══════════════════════════════════════════════════════════════
// PRINT REPORT
// ═══════════════════════════════════════════════════════════════
function printDetalleReport() {
  const rows = getDetalleForEstab(detalleSelected);
  const total = rows.length;
  const cC = countByIndicador(rows, 'INDICADOR_CENTROS');
  const cA = countByIndicador(rows, 'INDICADOR_ALM');
  const sobC = rows.filter(r => r.INDICADOR_CENTROS === 'SOBRESTOCK' && r.DISP_CENTROS >= detalleSobMeses).length;
  const sobA = rows.filter(r => r.INDICADOR_ALM    === 'SOBRESTOCK' && r.DIP_ALM >= detalleSobMeses).length;
  const pct  = v => total > 0 ? (v / total * 100).toFixed(1) + '%' : '0%';

  const topSobC = getTop20Sobrestock(rows, 'INDICADOR_CENTROS', 'DISP_CENTROS', detalleSobMeses).slice(0, 15);
  const desabC  = rows.filter(r => r.INDICADOR_CENTROS === 'DESABASTECIDO')
    .sort((a, b) => b.CPMA_CENTROS - a.CPMA_CENTROS).slice(0, 15);

  const fecha = new Date().toLocaleDateString('es-PE', {day:'2-digit',month:'long',year:'numeric'});
  const reportFecha = rows[0]?.FECHA_REPORTE || '—';
  const estabLabel  = detalleSelected || 'TODOS LOS ESTABLECIMIENTOS';

  // Render bar row for print
  const printBar = (v, max) => {
    const pctB = Math.min(v / max * 100, 100);
    const col  = v === 0 ? '#e11d48' : v < 6 ? '#ca8a04' : '#16a34a';
    return `<div style="display:flex;align-items:center;gap:6px;width:100%">
      <div style="flex:1;height:10px;background:#e5e7eb;border-radius:3px;overflow:hidden">
        <div style="width:${pctB}%;height:100%;background:${col};border-radius:3px"></div>
      </div>
      <span style="font-family:monospace;font-size:9px;min-width:28px">${v.toFixed(1)}</span>
    </div>`;
  };

  const indColor = { DESABASTECIDO:'#e11d48', SUBSTOCK:'#ea580c', NORMOSTOCK:'#16a34a', SOBRESTOCK:'#ca8a04', 'SIN ROTACION':'#7c3aed' };
  const indPrintRow = (label, val, field) =>
    `<div class="print-bar-row">
      <div style="width:12px;height:12px;background:${indColor[label]||'#999'};border-radius:2px"></div>
      <span style="flex:1">${label}</span>
      <span style="font-weight:bold">${val}</span>
      <span style="color:#888;min-width:42px;text-align:right">${pct(val)}</span>
    </div>`;

  document.getElementById('detalle-print-area').innerHTML = `
    <div class="print-header">
      <div class="print-title">DIRESA CALLAO — DISPOALM · Reporte de Disponibilidad</div>
      <div class="print-sub">
        Establecimiento: <strong>${estabLabel}</strong> &nbsp;|&nbsp;
        Fecha reporte: <strong>${reportFecha}</strong> &nbsp;|&nbsp;
        Generado: ${fecha} &nbsp;|&nbsp;
        Umbral sobrestock: ≥ ${detalleSobMeses} meses
      </div>
    </div>

    <!-- KPIs -->
    <div class="print-kpi-row">
      <div class="print-kpi"><div class="print-kpi-lbl">Total Productos</div><div class="print-kpi-val">${total.toLocaleString()}</div></div>
      <div class="print-kpi"><div class="print-kpi-lbl">Desab. Centros</div><div class="print-kpi-val red">${cC.DESABASTECIDO||0} <small style="font-size:12px">(${pct(cC.DESABASTECIDO)})</small></div></div>
      <div class="print-kpi"><div class="print-kpi-lbl">Desab. Almacén</div><div class="print-kpi-val red">${cA.DESABASTECIDO||0} <small style="font-size:12px">(${pct(cA.DESABASTECIDO)})</small></div></div>
      <div class="print-kpi"><div class="print-kpi-lbl">Sobrestock C ≥${detalleSobMeses}m</div><div class="print-kpi-val yellow">${sobC} <small style="font-size:12px">(${pct(sobC)})</small></div></div>
      <div class="print-kpi"><div class="print-kpi-lbl">Sobrestock A ≥${detalleSobMeses}m</div><div class="print-kpi-val yellow">${sobA} <small style="font-size:12px">(${pct(sobA)})</small></div></div>
      <div class="print-kpi"><div class="print-kpi-lbl">Normostock C</div><div class="print-kpi-val green">${cC.NORMOSTOCK||0} <small style="font-size:12px">(${pct(cC.NORMOSTOCK)})</small></div></div>
    </div>

    <!-- Distribución -->
    <div class="print-charts-row">
      <div class="print-chart-box">
        <div class="print-chart-title">Distribución — Indicador Centros</div>
        ${IND_ORDER.map(k => indPrintRow(k, cC[k]||0, 'INDICADOR_CENTROS')).join('')}
      </div>
      <div class="print-chart-box">
        <div class="print-chart-title">Distribución — Indicador Almacén</div>
        ${IND_ORDER.map(k => indPrintRow(k, cA[k]||0, 'INDICADOR_ALM')).join('')}
      </div>
    </div>

    <!-- Top sobrestock -->
    <div class="print-sob-section">
      <div class="print-chart-title">🟡 Top 15 Mayor Sobrestock Centros ≥ ${detalleSobMeses} meses</div>
      <table class="print-table">
        <thead><tr><th>#</th><th>Código</th><th>Descripción</th><th>Establecimiento</th><th>Stock</th><th>CPMA</th><th>Disp. (m)</th><th>Indicador</th></tr></thead>
        <tbody>
          ${topSobC.map((r, i) => `
          <tr>
            <td>${i+1}</td>
            <td>${r.cod_sismed}</td>
            <td>${r.descripcion}</td>
            <td>${r.establecimiento}</td>
            <td style="text-align:right">${r.STOCK_CENTROS.toLocaleString()}</td>
            <td style="text-align:right">${r.CPMA_CENTROS.toFixed(1)}</td>
            <td style="text-align:right;font-weight:bold;color:#ca8a04">${r.DISP_CENTROS.toFixed(1)}</td>
            <td>${r.INDICADOR_CENTROS}</td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>

    <!-- Desabastecidos -->
    <div class="print-sob-section">
      <div class="print-chart-title">🔴 Top 15 Desabastecidos Centros (Mayor Consumo)</div>
      <table class="print-table">
        <thead><tr><th>#</th><th>Código</th><th>Descripción</th><th>Establecimiento</th><th>Stock</th><th>CPMA</th><th>Disp. (m)</th><th>Indicador</th></tr></thead>
        <tbody>
          ${desabC.map((r, i) => `
          <tr>
            <td>${i+1}</td>
            <td>${r.cod_sismed}</td>
            <td>${r.descripcion}</td>
            <td>${r.establecimiento}</td>
            <td style="text-align:right;color:#e11d48">${r.STOCK_CENTROS.toLocaleString()}</td>
            <td style="text-align:right">${r.CPMA_CENTROS.toFixed(1)}</td>
            <td style="text-align:right;font-weight:bold;color:#e11d48">${r.DISP_CENTROS.toFixed(1)}</td>
            <td style="color:#e11d48;font-weight:bold">${r.INDICADOR_CENTROS}</td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>

    <div class="print-footer">
      DISPOALM Dashboard · DIRESA Callao · DEMID · Generado: ${new Date().toLocaleString('es-PE')} · ${total.toLocaleString()} registros procesados.
    </div>`;

  setTimeout(() => window.print(), 200);
}

// ═══════════════════════════════════════════════════════════════
// CHART FILTER — Detalle
// ═══════════════════════════════════════════════════════════════
let detalleChartFilter = null;

function applyDetalleChartFilter(field, value, label) {
  detalleChartFilter = { field, value, label };
  const badge = document.getElementById('dd-filter-badge');
  const lbl   = document.getElementById('dd-filter-label');
  if (badge) badge.style.display = 'flex';
  if (lbl)   lbl.textContent = label;
  filterDetalleTable();
  const tbl = document.querySelector('#detalle-main .tbl-card');
  if (tbl) tbl.scrollIntoView({ behavior:'smooth', block:'start' });
}

function clearDetalleChartFilter() {
  detalleChartFilter = null;
  const badge = document.getElementById('dd-filter-badge');
  if (badge) badge.style.display = 'none';
  const indC = document.getElementById('dt-ind-c');
  const indA = document.getElementById('dt-ind-a');
  const tip  = document.getElementById('dt-tipo');
  if (indC) indC.value = '';
  if (indA) indA.value = '';
  if (tip)  tip.value  = '';
  filterDetalleTable();
}

// ── Patch filterDetalleTable to respect detalleChartFilter ────
const _origFilterDetalle = filterDetalleTable;
filterDetalleTable = function() {
  const rows  = getDetalleForEstab(detalleSelected);
  const s     = (document.getElementById('dt-search')?.value || '').toLowerCase();
  const indC  = document.getElementById('dt-ind-c')?.value  || '';
  const indA  = document.getElementById('dt-ind-a')?.value  || '';
  const tipo  = document.getElementById('dt-tipo')?.value   || '';
  const cf    = detalleChartFilter;

  dtFiltered = rows.filter(r => {
    if (indC && r.INDICADOR_CENTROS !== indC) return false;
    if (indA && r.INDICADOR_ALM     !== indA) return false;
    if (tipo && r.TIPO !== tipo) return false;
    if (s && !r.descripcion.toLowerCase().includes(s) && !r.cod_sismed.includes(s)) return false;
    if (cf) {
      if (cf.field === 'INDICADOR_CENTROS' && r.INDICADOR_CENTROS !== cf.value) return false;
      if (cf.field === 'INDICADOR_ALM'     && r.INDICADOR_ALM     !== cf.value) return false;
      if (cf.field === 'RED'               && r.RED               !== cf.value) return false;
      if (cf.field === '_search'           && r.cod_sismed         !== cf.value) return false;
    }
    return true;
  });

  dtFiltered.sort((a, b) => {
    const av = a[dtSortCol], bv = b[dtSortCol];
    if (typeof av === 'number') return (av - bv) * dtSortDir;
    return String(av).localeCompare(String(bv)) * dtSortDir;
  });
  dtPage = 1;
  renderDetalleTable();
};

// ── Patch doughnut clicks to use applyDetalleChartFilter ──────
// (Los mkDoughnut y mkDoughnutAlm originales ya usan applyIndicadorFilter;
//  redirigimos esa función al nuevo sistema)
function applyIndicadorFilter(field, value) {
  const label = field === 'INDICADOR_CENTROS'
    ? `Indicador Centros = ${value}`
    : `Indicador Almacén = ${value}`;
  applyDetalleChartFilter(field, value, label);
}

// ═══════════════════════════════════════════════════════════════
// EXPORT EXCEL — Detalle (4 hojas)
// ═══════════════════════════════════════════════════════════════
function exportDetalleXLSX() {
  if (!dtFiltered.length) { alert('No hay datos filtrados para exportar.'); return; }

  const wb = XLSX.utils.book_new();

  // Hoja 1: Detalle filtrado
  const cols = ['cod_sismed','descripcion','TIPO','establecimiento','RED',
    'STOCK_CENTROS','CPMA_CENTROS','DISP_CENTROS','INDICADOR_CENTROS',
    'STOCK_ALM','CMPA_ALM','DIP_ALM','INDICADOR_ALM','FECHA_REPORTE'];
  const hdrs = ['Código','Descripción','Tipo','Establecimiento','Red',
    'Stock Centro','CPMA Centro','Disp.C (m)','Indicador Centro',
    'Stock Almacén','CMPA Almacén','Disp.A (m)','Indicador Almacén','Fecha Reporte'];
  const ws1 = XLSX.utils.aoa_to_sheet([hdrs, ...dtFiltered.map(r => cols.map(c => r[c] ?? ''))]);
  ws1['!cols'] = [8,42,6,28,20,12,10,10,14,12,10,10,14,12].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws1, 'Detalle');

  // Hoja 2: Resumen
  const rows  = getDetalleForEstab(detalleSelected);
  const total = rows.length;
  const cC    = countByIndicador(rows, 'INDICADOR_CENTROS');
  const cA    = countByIndicador(rows, 'INDICADOR_ALM');
  const pct   = v => total>0 ? (v/total*100).toFixed(2)+'%' : '0%';
  const estabLabel   = detalleSelected || 'TODOS';
  const filterLabel  = detalleChartFilter ? detalleChartFilter.label : 'Sin filtro';

  const sumData = [
    ['REPORTE DISPO_CENTROS_ALM_DETALLE'],
    ['Establecimiento:', estabLabel],
    ['Filtro:', filterLabel],
    ['Fecha:', new Date().toLocaleString('es-PE')],
    ['Umbral sobrestock:', `≥ ${detalleSobMeses} meses`],
    [],
    ['INDICADOR','CENTROS Cant.','CENTROS %','ALMACÉN Cant.','ALMACÉN %'],
    ...IND_ORDER.map(k => [k, cC[k]||0, pct(cC[k]), cA[k]||0, pct(cA[k])]),
    ['TOTAL', total, '100%', total, '100%'],
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(sumData);
  ws2['!cols'] = [22,14,12,14,12].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws2, 'Resumen');

  // Hoja 3: Top sobrestock centros
  const topC = getTop20Sobrestock(rows, 'INDICADOR_CENTROS', 'DISP_CENTROS', detalleSobMeses);
  if (topC.length) {
    const hdr = ['#','Código','Descripción','Establecimiento','Stock C','CPMA C','Disp.C (m)','Ind. C'];
    const ws3 = XLSX.utils.aoa_to_sheet([hdr, ...topC.map((r,i)=>[i+1,r.cod_sismed,r.descripcion,r.establecimiento,r.STOCK_CENTROS,+r.CPMA_CENTROS.toFixed(1),+r.DISP_CENTROS.toFixed(1),r.INDICADOR_CENTROS])]);
    ws3['!cols'] = [4,8,42,28,12,10,12,14].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, ws3, 'Top Sobrestock C');
  }

  // Hoja 4: Desabastecidos centros
  const desab = rows.filter(r=>r.INDICADOR_CENTROS==='DESABASTECIDO').sort((a,b)=>b.CPMA_CENTROS-a.CPMA_CENTROS);
  if (desab.length) {
    const hdr = ['#','Código','Descripción','Establecimiento','Stock C','CPMA C','Disp.C (m)'];
    const ws4 = XLSX.utils.aoa_to_sheet([hdr, ...desab.map((r,i)=>[i+1,r.cod_sismed,r.descripcion,r.establecimiento,r.STOCK_CENTROS,+r.CPMA_CENTROS.toFixed(1),+r.DISP_CENTROS.toFixed(1)])]);
    ws4['!cols'] = [4,8,42,28,12,10,12].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, ws4, 'Desabastecidos');
  }

  const slug = detalleSelected ? `_${detalleSelected.replace(/\s+/g,'_').slice(0,20)}` : '_TODOS';
  XLSX.writeFile(wb, `DISPO_DETALLE${slug}_${new Date().toISOString().split('T')[0]}.xlsx`);
}

// ═══════════════════════════════════════════════════════════════
// DESCARGA ORIGINAL DESDE DRIVE — Detalle
// ═══════════════════════════════════════════════════════════════
async function downloadDetalleOriginal() {
  if (!detalleFileId) { alert('No se encontró el ID del archivo en Drive.'); return; }
  const btn = [...document.querySelectorAll('.btn-sm')].find(b => b.textContent.includes('Archivo Original'));
  if (btn) { btn.textContent = '⏳ Descargando…'; btn.disabled = true; }
  try {
    const url = `https://www.googleapis.com/drive/v3/files/${detalleFileId}?alt=media`;
    const res = await authFetch(url);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const blob = await res.blob();
    const a    = document.createElement('a');
    a.href     = URL.createObjectURL(blob);
    a.download = 'DISPO_CENTROS_ALM_DETALLE.xlsx';
    a.click();
    URL.revokeObjectURL(a.href);
  } catch(e) {
    alert('Error al descargar: ' + e.message);
  } finally {
    if (btn) { btn.textContent = '⬇ Archivo Original'; btn.disabled = false; }
  }
}
