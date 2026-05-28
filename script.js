// ═══════════════════════════════════════════════════════════════
// CONFIG
// ═══════════════════════════════════════════════════════════════
const FOLDER_ID = '1hAjQoKcETPn-s1kOEa4mhCsmVTQr3VC1';
const FOLDER_NAME = 'CARPETA_ROTACION_2024';
const EXCLUDED_FILES = new Set([]);
// Archivos con vistas especiales
const DETALLE_FILE_PATTERN = /DISPO_CENTROS_ALM_DETALLE/i;
const CENTROS_FILE_PATTERN_JS = /^DISPO_CENTROS\.XLS[XM]?$/i;

// ── Service Account credentials (embedded) ─────────────────────
const SA_CREDS = {
  client_email: "lector-sheets-api@enduring-signal-372414.iam.gserviceaccount.com",
  private_key_id: "f7c5d4030438ca5e96cf1150eac289519580d407",
  private_key: `-----BEGIN PRIVATE KEY-----
MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQDJyK82oRocEhEN
iOCKSzPKoKrkLChKFkEqgqTeN3zIPWPdX/E1Wy4MCQMUPVtNeC5kM1dpVX/FqWFh
ZFVmIFRTXukCt14iaPmvcG+/0bRLl0VWJKehr47SWLVGk03Sc1Mr4qZmmnMD8awV
B78gmjHXPIZgVm9K7lBtfA5DUtdNv0dy7EhmyME+YB2S2uPWYhATuiljZzFagDYe
eFHdtq9HD6CfCJZ+6Cv+KQmnSmMhKOUwZQsN7SLZKbebQmwMx37d9NrSsd4H3m9w
gIrD6zvaA5wo4gszxkqVvr+Xa5nJCGZXdTiyx/RmsI1Po3oyW4yjwDdNcPJQ/0Fm
RkwkAws3AgMBAAECggEAFEsRi/3xwUbz3heY9W+Qut0jsqTaPHE2ho7zA8Iy03SR
GY9vk70iyWUSoTN2fjkK+VrMfLRzkC5zs7bCe+3ebTALYRGN7wn9qm41XuY5Kq0A
9SPkk/cy9uai/IEN/49Hmw9FSuZXkNr6Qncv+phnUM0uenvW34NOJWSnjo/DRXc8
dgEagDuRaUSUWOnc/0IDo2AZJesNOykZrqOdzSbRlRDRWGAj9Oni9h6s9H9TUIuI
FjJsCenYx+dRQ56NcohI/m0mD/zfoz20QtA4TA+eRUdobW/jRDpEYynzMR+P6MUW
W7p5IBu8NVtJZ7UKxTMqhOJpXgYvftngGb37Bn1JiQKBgQDt2YflF6CR5K6v+uoX
FN84DY2v/l/hk9w1WnyTl55URi9WdMVsLhr+iOiORD7/s/OBVNAXLx5F39EowGKQ
Df7p1AEXVEwN23k7XBcgAzehKuvuat1t8gQj2upxE5iYY55vgi5w28A/N0xkIADq
3+rAy9L1lkjeCHCApVnsYJIvgwKBgQDZLpjDvOskF8fQvCvDgc/KpdHghDXCbQWd
u0V8IwlfjWMDFC6RFh5573WTErXyIgTGkBgl2HFOta5tuefbpHM6eSXjYHQy5wMm
T1p/canrhAfN3qHQj6cEOO60ZEGVhbDYACnPn9e6HUWuGYRmm4CxbgKv49dFJM9E
+5HiSg0TPQKBgDgXB3wYEU62bhekBVZCHs3aLGKnizboJpMSbrRjmfvnvTwXN4Nm
cAk8ghMitvSYHEGQna4J94qk2G9SxyuPkWDVFRjKUVEnQ7Si5/UtrQ290HueCe/m
leFau9TEuUSeoXtMaOVGfVSrFsG55l19RqAGsQ4nTbHGrbk4xTcAIxk/AoGBAI4u
pFJY5vQtIWJ1Ho2S+Y9ouZsr0uojZ67YwGfBpExLCswkNUxE19TqQOX1NOp8/PVl
tnloVt25saUiNTmlUs6wK27NYaVFedll0BBpQF3whQ9xNiyNL0FjaFru28TgzmLK
0JSdx6SPr2OITxbBrSH7rISsO+2HaK53qA0mf6LdAoGBAOio6L3VNrEbj95vjxpW
YMwqlM12FLdvV81ZkNYzwNRNfirQitEUX+pNCazNfPTrbpvTMvqO10NVTwKVNDBx
COj7fhwav4qSbMZGx/vqy7z12xUSznxUd/yLQY5UcGNgdXWkFdGgMquaFwE7r/Ht
dLiJAAl/TY1qk8Bbcdi7JsHn
-----END PRIVATE KEY-----`
};

let accessToken = null;
let tokenExpiry  = 0;

let driveFiles   = [];
let fileDataCache= {};
let activeFileId = null;
let charts = {};
let allData = [], filteredData = [];
let sortCol = 'CMPA', sortDir = -1;
let page = 1;
let activeView = { profile: 'rotacion', showConsolidadoVfCols: false, showAlmCols: false };
const PAGE = 50;

// ── JWT / Service-Account auth ─────────────────────────────────
// Encode helpers (no external libs needed)
function b64url(str) {
  return btoa(str).replace(/\+/g,'-').replace(/\//g,'_').replace(/=/g,'');
}
function b64urlFromBytes(bytes) {
  let bin = '';
  bytes.forEach(b => bin += String.fromCharCode(b));
  return b64url(bin);
}

async function importPrivateKey(pem) {
  const pemBody = pem
    .replace(/-----BEGIN PRIVATE KEY-----|-----END PRIVATE KEY-----|\n/g,'');
  const der = Uint8Array.from(atob(pemBody), c => c.charCodeAt(0));
  return crypto.subtle.importKey(
    'pkcs8', der.buffer,
    { name:'RSASSA-PKCS1-v1_5', hash:'SHA-256' },
    false, ['sign']
  );
}

async function getAccessToken() {
  const now = Math.floor(Date.now() / 1000);
  if (accessToken && now < tokenExpiry - 60) return accessToken;

  const header  = b64url(JSON.stringify({ alg:'RS256', typ:'JWT' }));
  const payload = b64url(JSON.stringify({
    iss  : SA_CREDS.client_email,
    scope: 'https://www.googleapis.com/auth/drive.readonly',
    aud  : 'https://oauth2.googleapis.com/token',
    iat  : now,
    exp  : now + 3600
  }));

  const sigInput = `${header}.${payload}`;
  const key = await importPrivateKey(SA_CREDS.private_key);
  const sigBytes = new Uint8Array(
    await crypto.subtle.sign(
      'RSASSA-PKCS1-v1_5', key,
      new TextEncoder().encode(sigInput)
    )
  );
  const jwt = `${sigInput}.${b64urlFromBytes(sigBytes)}`;

  const resp = await fetch('https://oauth2.googleapis.com/token', {
    method : 'POST',
    headers: { 'Content-Type':'application/x-www-form-urlencoded' },
    body   : `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`
  });
  const json = await resp.json();
  if (!json.access_token) throw new Error('Auth fallida: ' + JSON.stringify(json));

  accessToken = json.access_token;
  tokenExpiry  = now + (json.expires_in || 3600);
  return accessToken;
}

// ── Drive API helpers ──────────────────────────────────────────
async function authFetch(url) {
  const token = await getAccessToken();
  return fetch(url, { headers: { Authorization: `Bearer ${token}` } });
}

async function fetchDriveFiles() {
  const q = encodeURIComponent(
    `'${FOLDER_ID}' in parents and trashed=false and ` +
    `(mimeType='application/vnd.ms-excel' or ` +
    `mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or ` +
    `name contains '.xlsm' or name contains '.xlsx' or name contains '.xls')`
  );
  const url = `https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name,mimeType,modifiedTime,size)&orderBy=name`;
  const res  = await authFetch(url);
  const json = await res.json();
  if (json.error) throw new Error(json.error.message);
  return (json.files || []).filter(f => !EXCLUDED_FILES.has(String(f.name || '').toUpperCase()));
}

function goBack() {
  Object.values(charts).forEach(c => { try{c.destroy();}catch(e){} });
  charts = {};
  fileDataCache = {};
  driveFiles = [];
  activeFileId = null;
  accessToken = null;
  tokenExpiry  = 0;
  document.getElementById('dash-content').innerHTML =
    `<div class="loading-overlay"><div class="loader"></div><div class="loading-text">Reconectando…</div></div>`;
  initApp();
}

async function refreshFiles() {
  const btn = document.getElementById('refresh-btn');
  btn.classList.add('spinning');
  try {
    driveFiles = await fetchDriveFiles();
    fileDataCache = {};
    buildTabs();
    if (activeFileId && driveFiles.find(f => f.id === activeFileId)) {
      selectFile(activeFileId);
    } else if (driveFiles.length > 0) {
      selectFile(driveFiles[0].id);
    }
  } catch(e) { alert('Error al actualizar: ' + e.message); }
  finally { btn.classList.remove('spinning'); }
}

// ═══════════════════════════════════════════════════════════════
// TABS
// ═══════════════════════════════════════════════════════════════
/*function buildTabs() {
  const nav = document.getElementById('file-nav');
  // Keep refresh button
  const refreshBtn = nav.querySelector('.nav-refresh');
  nav.innerHTML = '';
  nav.appendChild(refreshBtn);

  const dotColors = ['#6c63ff','#22c55e','#f97316','#3b82f6','#f43f5e','#eab308','#06b6d4','#8b5cf6'];

  driveFiles.forEach((f, i) => {
    const tab = document.createElement('div');
    tab.className = 'file-tab' + (f.id === activeFileId ? ' active' : '');
    tab.dataset.id = f.id;
    const color = dotColors[i % dotColors.length];
    const shortName = f.name.replace(/\.(xlsm?|xlsx?)$/i,'').slice(0,28);
    const date = f.modifiedTime ? new Date(f.modifiedTime).toLocaleDateString('es-PE',{day:'2-digit',month:'short'}) : '';
    tab.innerHTML = `
      <div class="tab-dot" style="background:${color}"></div>
      <span>${shortName}</span>
      <span class="tab-badge">${date}</span>`;
    tab.onclick = () => selectFile(f.id);
    nav.insertBefore(tab, refreshBtn);
  });
}*/

function buildTabs() {
  const nav = document.getElementById('file-nav');
  // Mantener el botón de actualizar en su lugar
  const refreshBtn = nav.querySelector('.nav-refresh');
  nav.innerHTML = '';
  if (refreshBtn) nav.appendChild(refreshBtn);

  // Paleta de colores original para los puntitos de las pestañas
  const dotColors = ['#6c63ff', '#22c55e', '#f97316', '#3b82f6', '#f43f5e', '#eab308', '#06b6d4', '#8b5cf6'];

  // 1. Filtrar estrictamente para dejar solo los dos archivos que deseas mostrar
  const filteredFiles = driveFiles.filter(f => {
    const isDetalle = DETALLE_FILE_PATTERN.test(f.name);
    const isCentros = CENTROS_FILE_PATTERN_JS.test(f.name);
    return isDetalle || isCentros;
  });

  // 2. Renderizar las pestañas con el diseño original de la app
  filteredFiles.forEach((f, i) => {
    const tab = document.createElement('div');
    // Aplica la clase 'active' si este archivo es el seleccionado actualmente
    tab.className = 'file-tab' + (f.id === activeFileId ? ' active' : '');
    tab.dataset.id = f.id;
    
    // Asignar un color del array de diseño
    const color = dotColors[i % dotColors.length];
    
    // Limpiar el nombre para que no muestre el ".xlsx" o ".xlsm"
    const shortName = f.name.replace(/\.(xlsm?|xlsx?)$/i, '').replace(/_/g, ' ').slice(0, 28);
    
    // Formatear la fecha corta para el badge de la derecha
    const date = f.modifiedTime ? new Date(f.modifiedTime).toLocaleDateString('es-PE', { day: '2-digit', month: 'short' }) : '';
    
    // Inyectar la estructura visual idéntica a la original
    tab.innerHTML = `
      <div class="tab-dot" style="background:${color}"></div>
      <span>${shortName}</span>
      <span class="tab-badge">${date}</span>`;
    
    // Asignar el evento de clic
    tab.onclick = () => selectFile(f.id);
    
    // Insertar la pestaña justo antes del botón de actualizar
    if (refreshBtn) {
      nav.insertBefore(tab, refreshBtn);
    } else {
      nav.appendChild(tab);
    }
  });
}

// ═══════════════════════════════════════════════════════════════
// SELECT + LOAD FILE
// ═══════════════════════════════════════════════════════════════
async function selectFile(id) {
  activeFileId = id;

  // Update active tab
  document.querySelectorAll('.file-tab').forEach(t => {
    t.classList.toggle('active', t.dataset.id === id);
  });

  const file = driveFiles.find(f => f.id === id);
  document.getElementById('tb-meta').textContent = file ? file.name : '—';

  // Show loader
  document.getElementById('dash-content').innerHTML = `
    <div class="loading-overlay">
      <div class="loader"></div>
      <div class="loading-text">Cargando ${file?.name || ''}…</div>
    </div>`;

  // Destroy old charts
  Object.values(charts).forEach(c => { try{c.destroy();}catch(e){} });
  charts = {};

  try {
    let data;
    if (fileDataCache[id]) {
      data = fileDataCache[id];
    } else {
      data = await downloadAndParse(id);
      fileDataCache[id] = data;
    }

    if (!data || data.length === 0) {
      document.getElementById('dash-content').innerHTML =
        '<div class="loading-overlay"><p style="color:var(--muted)">No se encontraron datos en este archivo.</p></div>';
      return;
    }

    // ── Routing por tipo de archivo ──────────────────────────────
    const fname = file?.name || '';
    if (DETALLE_FILE_PATTERN.test(fname)) {
      renderDetalleDashboard(data.__detalleRaw || data);
    } else if (CENTROS_FILE_PATTERN_JS.test(fname)) {
      renderCentrosDashboard(data.__centrosRaw || data);
    } else {
      renderDashboard(data, file);
    }
  } catch(e) {
    document.getElementById('dash-content').innerHTML =
      `<div class="loading-overlay"><p style="color:var(--red)">Error al cargar: ${e.message}</p></div>`;
    console.error(e);
  }
}

async function downloadAndParse(fileId) {
  // Download via Drive with Bearer token (Service Account)
  const url = `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`;
  const res = await authFetch(url);
  if (!res.ok) throw new Error(`HTTP ${res.status} al descargar el archivo`);
  const buf = await res.arrayBuffer();

  // ── Routing especial por tipo de archivo ──────────────────────
  const file = driveFiles.find(f => f.id === fileId);
  const fname = file?.name || '';
  if (DETALLE_FILE_PATTERN.test(fname)) {
    const rows = parseDetalleBuffer(buf);
    return Object.assign(rows, { __detalleRaw: rows });
  }
  if (CENTROS_FILE_PATTERN_JS.test(fname)) {
    const rows = parseCentrosBuffer(buf);
    return Object.assign(rows, { __centrosRaw: rows });
  }
  // ─────────────────────────────────────────────────────────────

  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  // Find the right sheet
  const sheetName = wb.SheetNames.find(n =>
    /DISPALM|DISPO|DISP/i.test(n)
  ) || wb.SheetNames[0];

  const ws = wb.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  const norm = s => String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase().replace(/[^A-Z0-9]/g, '');

  const HEADER_ALIASES = {
    MEDCOD: [
      'MEDCOD','CODMED','MEDCODIGO','CODIGO','CODIGOITEM','CODIGOPRODUCTO',
      'CODIGODEPRODUCTO','CODIGOINSUMO','ITEM','CODPRODUCTO','CODSISMED'
    ],
    NOMBRE_PRODUCTO: ['NOMBREPRODUCTO','PRODUCTO','DESCRIPCION','DESCRIPCIONPRODUCTO','NOMBRE'],
    STOCK_ACTUAL: ['STOCKACTUAL','STOCK','SALDO','EXISTENCIA','STOCKDISPONIBLE'],
    STOCK_CENTROS: ['STOCKCENTROS'],
    STOCK_ALM: ['STOCKALM'],
    STOCK_TOTAL: ['STOCKTOTAL'],
    CMPA: ['CMPA','CONSUMOPROMEDIO','CONSUMOPROMEDIOMENSUAL','CONSUMOMENSUAL','CPM'],
    CPMA: ['CPMA'],
    CPMA_CENTROS: ['CPMACENTROS'],
    CMPA_ALM: ['CMPAALM'],
    MES_DISPONIBLES: ['MESDISPONIBLES','MESESDISPONIBLES','COBERTURA','MESESCOBERTURA','COBERTURAMESES'],
    DISP_CENTROS: ['DISPCENTROS'],
    DIP_ALM: ['DIPALM'],
    DISPO: ['DISPO'],
    PRECIO: ['PRECIO'],
    ESTRATEGIC: ['ESTRATEGIC','ESTRATEGICO','TIPODEESTRATEGIA'],
    TIPO: ['TIPO'],
    forma_farm: ['FORMAFARM','FORMAFARMACEUTICA','FORMA'],
    INDICADOR: ['INDICADOR','ESTADOSTOCK','ESTADO'],
    INDICADOR_CENTROS: ['INDICADORCENTROS'],
    INDICADOR_ALM: ['INDICADORALM'],
    INDICADOR_NEW: ['INDICADORNEW'],
    NIVEL_ROTACION: ['NIVELROTACION','NIVELDEROTACION','ROTACION'],
    ESTADO_CONSUMO: ['ESTADOCONSUMO','CONSUMOESTADO'],
    INTERPRETACION: ['INTERPRETACION','OBSERVACION','COMENTARIO'],
    FECHA_VENCIMIENTO_CENTRO: ['FECHAVENCIMIENTOCENTRO'],
    FECHA_VENCIMIENTO_ALMACEN: ['FECHAVENCIMIENTOALMACEN'],
    FECHA_REPORTE: ['FECHAREPORTE','FECHA']
  };

  const aliasToCanonical = {};
  Object.entries(HEADER_ALIASES).forEach(([canonical, aliases]) => {
    aliases.forEach(a => { aliasToCanonical[norm(a)] = canonical; });
  });

  const resolveHeader = (txt) => {
    const k = norm(txt);
    if (!k) return '';
    if (aliasToCanonical[k]) return aliasToCanonical[k];

    if (/MEDCOD|CODIGO?MED|CODIGOPROD|CODIGOITEM|^CODIGO$|^COD$|ITEMCODE|CODPRODUCTO|CODIGOMEDICAMENTO/.test(k)) return 'MEDCOD';
    if (/NOMBREPRODUCTO|DESCRIPCIONPRODUCTO|DESCRIPCION|PRODUCTO|NOMBREMED|NOMBREINSUMO/.test(k)) return 'NOMBRE_PRODUCTO';
    if (/STOCKACTUAL|STOCK|SALDO|EXISTENCIA|DISPONIBLE/.test(k)) return 'STOCK_ACTUAL';
    if (/CMPA|CONSUMOPROMEDIO|CONSUMOMENSUAL|CPM/.test(k)) return 'CMPA';
    if (/MES(ES)?DISPONIBLES|COBERTURA(MESES)?/.test(k)) return 'MES_DISPONIBLES';
    if (/^PRECIO$|COSTO|PUNITARIO/.test(k)) return 'PRECIO';
    if (/ESTRATEGIC|ESTRATEGICO|TIPODEESTRATEGIA/.test(k)) return 'ESTRATEGIC';
    if (/^TIPO$|TIPOPRODUCTO|CLASE/.test(k)) return 'TIPO';
    if (/FORMAFARM|FORMAFARMACEUTICA|PRESENTACION|FORMA/.test(k)) return 'forma_farm';
    if (/INDICADOR|ESTADOSTOCK|ESTADO/.test(k)) return 'INDICADOR';
    if (/NIVELROTACION|ROTACION/.test(k)) return 'NIVEL_ROTACION';
    if (/ESTADOCONSUMO|CONSUMOESTADO/.test(k)) return 'ESTADO_CONSUMO';
    if (/INTERPRETACION|OBSERVACION|COMENTARIO/.test(k)) return 'INTERPRETACION';
    if (/FECHAVENCIMIENTOCENTRO/.test(k)) return 'FECHA_VENCIMIENTO_CENTRO';
    if (/FECHAVENCIMIENTOALMACEN/.test(k)) return 'FECHA_VENCIMIENTO_ALMACEN';
    if (/FECHAREPORTE|FECHACORTE|FECHAACTUALIZACION|^FECHA$/.test(k)) return 'FECHA_REPORTE';

    return txt;
  };

  // Find best header row by alias score (tolerante a archivos con encabezados distintos)
  let headerRow = -1;
  let bestScore = -1;
  for (let i = 0; i < Math.min(80, raw.length); i++) {
    const row = raw[i] || [];
    const found = new Set();
    row.forEach(c => {
      const key = resolveHeader(c);
      if (key) found.add(key);
    });
    const score = found.size + (found.has('MEDCOD') ? 2 : 0) + (found.has('NOMBRE_PRODUCTO') ? 1 : 0);
    if (score > bestScore) {
      bestScore = score;
      headerRow = i;
    }
  }
  if (headerRow === -1 || bestScore < 2) {
    throw new Error('No se reconocieron encabezados válidos en el archivo');
  }

  const headers = raw[headerRow].map(h => {
    const txt = h ? String(h).trim() : '';
    return resolveHeader(txt);
  });

  const toNum = v => {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
    const s = String(v).trim().replace(/\s+/g, '');
    if (!s) return 0;
    const n = Number(
      s.replace(/\.(?=\d{3}(?:\D|$))/g, '').replace(',', '.')
    );
    return Number.isFinite(n) ? n : 0;
  };

  const toDateText = v => {
    if (v == null || v === '') return '';
    if (v instanceof Date && !Number.isNaN(v.getTime())) return v.toISOString().split('T')[0];
    const txt = String(v).trim();
    if (!txt) return '';
    const d = new Date(txt);
    if (!Number.isNaN(d.getTime())) return d.toISOString().split('T')[0];
    return txt;
  };

  const rows = [];
  for (let i = headerRow + 1; i < raw.length; i++) {
    const row = raw[i];
    if (!row || row.every(c => c == null || String(c).trim() === '')) continue;
    const obj = {};
    headers.forEach((h, j) => {
      if (!h) return;
      const val = row[j] !== undefined && row[j] !== null ? row[j] : '';
      if (obj[h] === undefined || obj[h] === '') obj[h] = val;
    });

    if (!obj.MEDCOD) {
      const codeEntry = Object.entries(obj).find(([k, v]) => /COD|ITEM/i.test(norm(k)) && v !== '' && v != null);
      if (codeEntry) obj.MEDCOD = codeEntry[1];
    }
    if (!obj.NOMBRE_PRODUCTO) {
      const nameEntry = Object.entries(obj).find(([k, v]) => /NOMBRE|DESCRIP|PRODUCT/i.test(norm(k)) && v !== '' && v != null);
      if (nameEntry) obj.NOMBRE_PRODUCTO = nameEntry[1];
    }

    const hasCoreData = Boolean(obj.MEDCOD || obj.NOMBRE_PRODUCTO);
    if (!hasCoreData) continue;

    // Normalize
    obj.MEDCOD         = String(obj.MEDCOD || '').replace(/\.0+$/,'').trim();
    if (!obj.MEDCOD) obj.MEDCOD = String(i - headerRow).padStart(5,'0');
    else if (/^\d+$/.test(obj.MEDCOD)) obj.MEDCOD = obj.MEDCOD.padStart(5,'0');
    obj.STOCK_ACTUAL   = toNum(
      obj.STOCK_ACTUAL ?? obj.STOCK_TOTAL ?? obj.STOCK_ALM ?? obj.STOCK_CENTROS
    );
    obj.STOCK_CENTROS  = toNum(obj.STOCK_CENTROS);
    obj.STOCK_ALM      = toNum(obj.STOCK_ALM);
    obj.STOCK_TOTAL    = toNum(obj.STOCK_TOTAL);
    obj.CMPA           = toNum(
      obj.CMPA ?? obj.CMPA_ALM ?? obj.CPMA_CENTROS ?? obj.CPMA
    );
    obj.CPMA_CENTROS   = toNum(obj.CPMA_CENTROS);
    obj.CMPA_ALM       = toNum(obj.CMPA_ALM);
    obj.CPMA           = toNum(obj.CPMA);
    obj.MES_DISPONIBLES= toNum(
      obj.MES_DISPONIBLES ?? obj.DIP_ALM ?? obj.DISP_CENTROS ?? obj.DISPO
    );
    obj.DISP_CENTROS   = toNum(obj.DISP_CENTROS);
    obj.DIP_ALM        = toNum(obj.DIP_ALM);
    obj.DISPO          = toNum(obj.DISPO);
    obj.NOMBRE_PRODUCTO= String(obj.NOMBRE_PRODUCTO     || '');
    obj.TIPO           = String(obj.TIPO                || '').toUpperCase().trim();
    obj.PRECIO         = toNum(obj.PRECIO);
    obj.ESTRATEGIC     = String(obj.ESTRATEGIC          || '').trim();
    obj.forma_farm     = String(obj.forma_farm          || '');
    obj.INDICADOR_CENTROS = String(obj.INDICADOR_CENTROS || '').toUpperCase().trim();
    obj.INDICADOR_ALM     = String(obj.INDICADOR_ALM     || '').toUpperCase().trim();
    obj.INDICADOR_NEW     = String(obj.INDICADOR_NEW     || '').toUpperCase().trim();
    obj.INDICADOR      = String(
      obj.INDICADOR ?? obj.INDICADOR_ALM ?? obj.INDICADOR_CENTROS ?? obj.INDICADOR_NEW ?? ''
    ).toUpperCase().trim();
    const hasIndicadorNew = Boolean(obj.INDICADOR_NEW);
    if (hasIndicadorNew) obj.INDICADOR = obj.INDICADOR_NEW;
    obj.NIVEL_ROTACION = String(obj.NIVEL_ROTACION      || '').toUpperCase().trim();
    obj.ESTADO_CONSUMO = String(obj.ESTADO_CONSUMO      || '').toUpperCase().trim();
    obj.INTERPRETACION = String(obj.INTERPRETACION      || '');

    if (!obj.INDICADOR) {
      const m = obj.MES_DISPONIBLES;
      obj.INDICADOR = m === 0 ? 'DESABASTECIDO' : m < 1 ? 'SUBSTOCK' : m <= 3 ? 'NORMOSTOCK' : 'SOBRESTOCK';
    }
    if (!obj.NIVEL_ROTACION) {
      const c = obj.CMPA;
      obj.NIVEL_ROTACION = c <= 0 ? 'SIN ROTACION' : c >= 1000 ? 'ALTA ROTACION' : c >= 100 ? 'MEDIA ROTACION' : 'BAJA ROTACION';
    }
    if (!obj.ESTADO_CONSUMO) {
      const c = obj.CMPA;
      obj.ESTADO_CONSUMO = c <= 0 ? 'SIN DATOS' : c >= 1000 ? 'CONSUMO MUY RECIENTE' : c >= 100 ? 'CONSUMO RECIENTE' : 'CONSUMO ANTIGUO';
    }
    if (!obj.forma_farm) obj.forma_farm = '—';
    if (!obj.INTERPRETACION || (hasIndicadorNew && /^Indicador\s+/i.test(obj.INTERPRETACION))) {
      obj.INTERPRETACION = `Indicador ${obj.INDICADOR} con ${obj.MES_DISPONIBLES.toFixed(2)} meses de cobertura.`;
    }
    obj.FECHA_VENCIMIENTO_CENTRO = toDateText(obj.FECHA_VENCIMIENTO_CENTRO);
    obj.FECHA_VENCIMIENTO_ALMACEN = toDateText(obj.FECHA_VENCIMIENTO_ALMACEN);
    if (obj.FECHA_REPORTE instanceof Date) {
      obj.FECHA_REPORTE = obj.FECHA_REPORTE.toISOString().split('T')[0];
    } else { obj.FECHA_REPORTE = String(obj.FECHA_REPORTE || ''); }
    rows.push(obj);
  }
  return rows;
}

// ═══════════════════════════════════════════════════════════════
// RENDER DASHBOARD
// ═══════════════════════════════════════════════════════════════

function renderDashboard(data, file) {
  allData = data;
  filteredData = [...data];
  page = 1; sortCol = 'CMPA'; sortDir = -1;
  activeView.showConsolidadoVfCols = /CONSOLIDADO[_\s-]*VF/i.test(file?.name || '') ||
    data.some(d => d.FECHA_VENCIMIENTO_CENTRO || d.FECHA_VENCIMIENTO_ALMACEN || d.INDICADOR_NEW);
  activeView.showAlmCols = /CENTROS[_\s-]*ALM/i.test(file?.name || '') ||
    data.some(d => d.INDICADOR_ALM || d.INDICADOR_CENTROS || d.DIP_ALM || d.DISP_CENTROS);

  if (activeView.showConsolidadoVfCols) activeView.profile = 'consolidado_vf';
  else if (activeView.showAlmCols) activeView.profile = 'centros_alm';
  else activeView.profile = 'rotacion';

  const fecha = data[0]?.FECHA_REPORTE || file?.modifiedTime?.split('T')[0] || '';
  document.getElementById('tb-meta').textContent = `${file?.name || ''} · ${data.length.toLocaleString()} productos · Reporte: ${fecha}`;
  document.getElementById('footer-txt').textContent = `DISPOALM Dashboard · Carpeta: ${FOLDER_NAME} · Archivo: ${file?.name || ''} · Fecha reporte: ${fecha} · ${data.length.toLocaleString()} productos · Datos procesados localmente.`;

  // Build HTML
  document.getElementById('dash-content').innerHTML = buildDashHTML();

  // Render all
  renderKPIs(data);
  renderCharts(data);
  renderAlerts(data);
  filterTable();
}

function buildDashHTML() {
  const isAlm = activeView.profile === 'centros_alm';
  const isVf = activeView.profile === 'consolidado_vf';
  const titles = isVf ? {
    c1t:'Estado de Disponibilidad', c1s:'Según INDICADOR_NEW',
    c2t:'Stock Centros vs Almacén', c2s:'Comparativo de stock físico',
    c3t:'Estrategia', c3s:'Distribución por campo strategic',
    c4t:'Cobertura (DISPO)', c4s:'Distribución de DISPO',
    c5t:'Vencimientos', c5s:'Riesgo por fechas de vencimiento',
    a1t:'🔴 Alertas de Vencimiento/Stock', a1s:'Productos desabastecidos o con vencimiento próximo',
    topT:'📈 Top 12 Mayor CPMA', topS:'Consumo promedio mensual',
    tmt:'Medicamentos (M)', tms:'Indicador New por tipo',
    tit:'Insumos (I)', tis:'Indicador New por tipo'
  } : isAlm ? {
    c1t:'Indicador Almacén', c1s:'Distribución por INDICADOR_ALM',
    c2t:'Indicador Centros', c2s:'Distribución por INDICADOR_CENTROS',
    c3t:'Tipo de Producto', c3s:'Medicamentos vs Insumos',
    c4t:'Cobertura Almacén (DIP_ALM)', c4s:'Distribución por meses',
    c5t:'Cobertura Centros (DISP_CENTROS)', c5s:'Distribución por meses',
    a1t:'🔴 Alertas ALM/Centros', a1s:'Desabastecidos y substock críticos',
    topT:'📈 Top 12 Mayor CMPA_ALM', topS:'Consumo promedio mensual de almacén',
    tmt:'Medicamentos (M)', tms:'Indicador ALM por tipo',
    tit:'Insumos (I)', tis:'Indicador ALM por tipo'
  } : {
    c1t:'Estado de Stock', c1s:'Distribución por indicador de disponibilidad',
    c2t:'Nivel de Rotación', c2s:'Velocidad de consumo de productos',
    c3t:'Forma Farmacéutica', c3s:'Top 8 presentaciones',
    c4t:'Cobertura de Stock', c4s:'Distribución por meses disponibles',
    c5t:'Estado de Consumo', c5s:'Frecuencia de despacho',
    a1t:'🔴 Alertas Críticas', a1s:'Desabastecidos con alta rotación — acción inmediata',
    topT:'📈 Top 12 Mayor Consumo (CMPA)', topS:'Unidades/mes promedio',
    tmt:'Medicamentos (M)', tms:'Estado por tipo Medicamento',
    tit:'Insumos (I)', tis:'Estado por tipo Insumo'
  };

  return `
  <div class="main">
    <!-- KPIs -->
    <div>
      <div class="sec-hdr"><span class="sec-title">Indicadores Clave</span><div class="sec-line"></div></div>
      <div class="kpi-grid" id="kpi-grid"></div>
    </div>

    <!-- Charts row 1 -->
    <div class="g2">
      <div class="card">
        <div class="card-hdr"><div><div class="card-title">${titles.c1t}</div><div class="card-sub">${titles.c1s}</div></div></div>
        <div class="chart-h200"><canvas id="cIndicador"></canvas></div>
      </div>
      <div class="card">
        <div class="card-hdr"><div><div class="card-title">${titles.c2t}</div><div class="card-sub">${titles.c2s}</div></div></div>
        <div class="chart-h200"><canvas id="cRotacion"></canvas></div>
      </div>
    </div>

    <!-- Charts row 2 -->
    <div class="g3">
      <div class="card">
        <div class="card-hdr"><div><div class="card-title">${titles.c3t}</div><div class="card-sub">${titles.c3s}</div></div></div>
        <div class="chart-h200"><canvas id="cForma"></canvas></div>
      </div>
      <div class="card">
        <div class="card-hdr"><div><div class="card-title">${titles.c4t}</div><div class="card-sub">${titles.c4s}</div></div></div>
        <div class="chart-h200"><canvas id="cCobertura"></canvas></div>
      </div>
      <div class="card">
        <div class="card-hdr"><div><div class="card-title">${titles.c5t}</div><div class="card-sub">${titles.c5s}</div></div></div>
        <div class="chart-h200"><canvas id="cConsumo"></canvas></div>
      </div>
    </div>

    <!-- Alerts + Top consumo -->
    <div class="g2">
      <div class="card">
        <div class="card-hdr">
          <div><div class="card-title">${titles.a1t}</div><div class="card-sub">${titles.a1s}</div></div>
        </div>
        <div class="alert-scroll" id="alerts-list"></div>
      </div>
      <div class="card">
        <div class="card-hdr"><div><div class="card-title">${titles.topT}</div><div class="card-sub">${titles.topS}</div></div></div>
        <div class="chart-h360"><canvas id="cTopConsumo"></canvas></div>
      </div>
    </div>

    <!-- M vs I -->
    <div class="g2">
      <div class="card">
        <div class="card-hdr"><div><div class="card-title">${titles.tmt}</div><div class="card-sub">${titles.tms}</div></div></div>
        <div class="chart-h200"><canvas id="cTipoM"></canvas></div>
      </div>
      <div class="card">
        <div class="card-hdr"><div><div class="card-title">${titles.tit}</div><div class="card-sub">${titles.tis}</div></div></div>
        <div class="chart-h200"><canvas id="cTipoI"></canvas></div>
      </div>
    </div>

    <!-- Table -->
    <div>
      <div class="sec-hdr"><span class="sec-title">Catálogo Completo</span><div class="sec-line"></div></div>
      <div class="tbl-card">
        <div class="tbl-bar">
          <span class="tbl-bar-title">Tabla de Productos</span>
          <input class="inp" type="text" id="t-search" placeholder="🔍 Buscar nombre o código…" style="width:220px" oninput="filterTable()">
          <select class="inp" id="t-ind" onchange="filterTable()">
            <option value="">Todos los estados</option>
            <option>DESABASTECIDO</option><option>SUBSTOCK</option>
            <option>NORMOSTOCK</option><option>SOBRESTOCK</option>
            <option>SIN ROTACION</option>
          </select>
          ${!isAlm && !isVf ? `<select class="inp" id="t-rot" onchange="filterTable()">
            <option value="">Toda rotación</option>
            <option>ALTA ROTACION</option><option>MEDIA ROTACION</option>
            <option>BAJA ROTACION</option><option>SIN ROTACION</option>
          </select>` : ''}
          <select class="inp" id="t-tipo" onchange="filterTable()">
            <option value="">M + I</option>
            <option value="M">Medicamento</option>
            <option value="I">Insumo</option>
          </select>
          <span class="tbl-count" id="t-count"></span>
        </div>
        <div class="tbl-wrap">
          <table>
            <thead><tr>
              <th onclick="sortTbl('MEDCOD')">Código</th>
              <th onclick="sortTbl('NOMBRE_PRODUCTO')">Producto</th>
              <th onclick="sortTbl('TIPO')">Tipo</th>
              ${isAlm || isVf ? `<th onclick="sortTbl('PRECIO')" style="text-align:right">Precio</th>` : ''}
              ${isVf ? `<th onclick="sortTbl('ESTRATEGIC')">Estratégic</th>` : ''}
              ${isAlm ? `<th onclick="sortTbl('STOCK_CENTROS')" style="text-align:right">Stock Centros</th><th onclick="sortTbl('CPMA_CENTROS')" style="text-align:right">CPMA Centros</th><th onclick="sortTbl('DISP_CENTROS')">Disp Centros</th><th onclick="sortTbl('INDICADOR_CENTROS')">Indicador Centros</th><th onclick="sortTbl('STOCK_ALM')" style="text-align:right">Stock Alm</th><th onclick="sortTbl('CMPA_ALM')" style="text-align:right">CMPA Alm</th><th onclick="sortTbl('DIP_ALM')">Dip Alm</th><th onclick="sortTbl('INDICADOR_ALM')">Indicador Alm</th>` : ''}
              ${isVf ? `<th onclick="sortTbl('STOCK_CENTROS')" style="text-align:right">Stock Centros</th><th onclick="sortTbl('STOCK_ALM')" style="text-align:right">Stock Alm</th><th onclick="sortTbl('STOCK_TOTAL')" style="text-align:right">Stock Total</th><th onclick="sortTbl('CPMA')" style="text-align:right">CPMA</th><th onclick="sortTbl('DISPO')">Dispo</th>` : ''}
              ${!isVf ? `<th onclick="sortTbl('forma_farm')">Forma</th>` : ''}
              ${!isVf ? `<th onclick="sortTbl('STOCK_ACTUAL')" style="text-align:right">Stock</th>` : ''}
              ${!isVf ? `<th onclick="sortTbl('CMPA')" style="text-align:right">CMPA</th>` : ''}
              <th onclick="sortTbl('MES_DISPONIBLES')">Meses</th>
              <th onclick="sortTbl('${isVf ? 'INDICADOR_NEW' : 'INDICADOR'}')">Indicador</th>
              ${!isAlm && !isVf ? `<th onclick="sortTbl('NIVEL_ROTACION')">Rotación</th><th>Consumo</th>` : ''}
              ${isVf ? `<th onclick="sortTbl('FECHA_VENCIMIENTO_CENTRO')">Venc. Centro</th><th onclick="sortTbl('FECHA_VENCIMIENTO_ALMACEN')">Venc. Almacén</th>` : ''}
              <th>Interpretación</th>
            </tr></thead>
            <tbody id="t-body"></tbody>
          </table>
        </div>
        <div class="tbl-pager">
          <button class="btn-sm" onclick="prevPage()">← Anterior</button>
          <span id="t-pager"></span>
          <button class="btn-sm" onclick="nextPage()">Siguiente →</button>
          <button class="btn-sm" onclick="exportCSV()" style="margin-left:auto">⬇ Exportar CSV</button>
        </div>
      </div>
    </div>
  </div>`;
}

// ═══════════════════════════════════════════════════════════════
// KPIs
// ═══════════════════════════════════════════════════════════════
function renderKPIs(data) {
  const isAlm = activeView.profile === 'centros_alm';
  const isVf = activeView.profile === 'consolidado_vf';
  const n = k => data.filter(d => d.INDICADOR === k).length;
  const total = data.length;
  const desab = n('DESABASTECIDO'), sub = n('SUBSTOCK'),
        normo = n('NORMOSTOCK'),   sobre = n('SOBRESTOCK'),
        sinR  = n('SIN ROTACION');
  const crit  = data.filter(d => d.INDICADOR==='DESABASTECIDO' && d.NIVEL_ROTACION==='ALTA ROTACION').length;
  const altaR = data.filter(d => d.NIVEL_ROTACION==='ALTA ROTACION').length;
  const p = (a,b) => b ? (a/b*100).toFixed(1)+'%' : '0%';

  const toDate = s => {
    if (!s) return null;
    const d = new Date(s);
    return Number.isNaN(d.getTime()) ? null : d;
  };

  let kpis;
  if (isVf) {
    const stockC = data.reduce((a,d)=>a+d.STOCK_CENTROS,0);
    const stockA = data.reduce((a,d)=>a+d.STOCK_ALM,0);
    const stockT = data.reduce((a,d)=>a+d.STOCK_TOTAL,0);
    const today = new Date();
    today.setHours(0,0,0,0);
    const near90 = data.filter(d => {
      const c = toDate(d.FECHA_VENCIMIENTO_CENTRO);
      const a = toDate(d.FECHA_VENCIMIENTO_ALMACEN);
      const nearest = [c,a].filter(Boolean).sort((x,y)=>x-y)[0];
      if (!nearest) return false;
      const diff = Math.ceil((nearest - today) / 86400000);
      return diff >= 0 && diff <= 90;
    }).length;
    const iNew = k => data.filter(d => (d.INDICADOR_NEW || d.INDICADOR) === k).length;
    kpis = [
      {lbl:'Total', val:total, sub:'productos', icon:'📦', c:'c-cyan'},
      {lbl:'Stock Centros', val:Math.round(stockC), sub:'unidades', icon:'🏥', c:'c-blue'},
      {lbl:'Stock Alm', val:Math.round(stockA), sub:'unidades', icon:'🏬', c:'c-violet'},
      {lbl:'Stock Total', val:Math.round(stockT), sub:'unidades', icon:'📊', c:'c-green'},
      {lbl:'Desabast.', val:iNew('DESABASTECIDO'), sub:p(iNew('DESABASTECIDO'),total), icon:'🚨', c:'c-red'},
      {lbl:'Substock', val:iNew('SUBSTOCK'), sub:p(iNew('SUBSTOCK'),total), icon:'⚠️', c:'c-orange'},
      {lbl:'Sobrestock', val:iNew('SOBRESTOCK'), sub:p(iNew('SOBRESTOCK'),total), icon:'✅', c:'c-yellow'},
      {lbl:'Vence ≤90d', val:near90, sub:'centro/almacén', icon:'⏰', c:'c-red'}
    ];
  } else if (isAlm) {
    const nAlm = k => data.filter(d => (d.INDICADOR_ALM || d.INDICADOR) === k).length;
    const nCen = k => data.filter(d => (d.INDICADOR_CENTROS || d.INDICADOR) === k).length;
    const critCov = data.filter(d => d.DISP_CENTROS < 1 || d.DIP_ALM < 1).length;
    const okCov = data.filter(d => d.DISP_CENTROS >= 3 && d.DIP_ALM >= 3).length;
    const stockC = data.reduce((a,d)=>a+d.STOCK_CENTROS,0);
    const stockA = data.reduce((a,d)=>a+d.STOCK_ALM,0);
    kpis = [
      {lbl:'Total', val:total, sub:'productos', icon:'📦', c:'c-cyan'},
      {lbl:'Stock Centros', val:Math.round(stockC), sub:'unidades', icon:'🏥', c:'c-blue'},
      {lbl:'Stock Alm', val:Math.round(stockA), sub:'unidades', icon:'🏬', c:'c-violet'},
      {lbl:'Desab Centros', val:nCen('DESABASTECIDO'), sub:p(nCen('DESABASTECIDO'),total), icon:'🚨', c:'c-red'},
      {lbl:'Desab Alm', val:nAlm('DESABASTECIDO'), sub:p(nAlm('DESABASTECIDO'),total), icon:'🔥', c:'c-red'},
      {lbl:'Substock Alm', val:nAlm('SUBSTOCK'), sub:p(nAlm('SUBSTOCK'),total), icon:'⚠️', c:'c-orange'},
      {lbl:'Cobertura Crítica', val:critCov, sub:`${p(critCov,total)} (<1 mes en centros o alm)`, icon:'⛔', c:'c-red'},
      {lbl:'Cobertura Adecuada', val:okCov, sub:`${p(okCov,total)} (≥3 meses en ambos)`, icon:'✅', c:'c-green'}
    ];
  } else {
    kpis = [
      {lbl:'Total',       val:total,  sub:'productos',         icon:'📦', c:'c-cyan'},
      {lbl:'Desabast.',   val:desab,  sub:p(desab,total),      icon:'🚨', c:'c-red'},
      {lbl:'Substock',    val:sub,    sub:p(sub,total),         icon:'⚠️', c:'c-orange'},
      {lbl:'Normostock',  val:normo,  sub:p(normo,total),       icon:'✅', c:'c-green'},
      {lbl:'Sobrestock',  val:sobre,  sub:p(sobre,total),       icon:'📊', c:'c-yellow'},
      {lbl:'Sin Rotac.',  val:sinR,   sub:p(sinR,total),        icon:'🔄', c:'c-violet'},
      {lbl:'Críticos',    val:crit,   sub:'desab+alta rot.',    icon:'🔥', c:'c-red'},
      {lbl:'Alta Rotac.', val:altaR,  sub:p(altaR,total),       icon:'⚡', c:'c-blue'},
    ];
  }

  document.getElementById('kpi-grid').innerHTML = kpis.map(k=>`
    <div class="kpi ${k.c}">
      <div class="kpi-icon">${k.icon}</div>
      <div class="kpi-lbl">${k.lbl}</div>
      <div class="kpi-val">${k.val.toLocaleString()}</div>
      <div class="kpi-sub">${k.sub}</div>
    </div>`).join('');
}

// ═══════════════════════════════════════════════════════════════
// CHARTS
// ═══════════════════════════════════════════════════════════════
const PAL = {
  DESABASTECIDO:'#f43f5e', SUBSTOCK:'#f97316',
  NORMOSTOCK:'#22c55e',    SOBRESTOCK:'#eab308',
  'SIN ROTACION':'#8b5cf6',
  'ALTA ROTACION':'#3b82f6','MEDIA ROTACION':'#22c55e',
  'BAJA ROTACION':'#eab308',
};

const baseOpts = (legend=true) => ({
  responsive:true, maintainAspectRatio:false,
  plugins:{
    legend:{display:legend,position:'bottom',labels:{
      color:'#5a5b72',font:{family:'DM Sans',size:11},padding:12,boxWidth:10
    }},
    tooltip:{
      backgroundColor:'#0f1018',borderColor:'rgba(255,255,255,.1)',borderWidth:1,
      titleColor:'#ecedf5',bodyColor:'#8889a6',padding:11,
      titleFont:{family:'Space Mono',size:11},bodyFont:{family:'DM Sans',size:12}
    }
  }
});

function cnt(data, field) {
  return data.reduce((m,d)=>{const k=d[field]||'N/A';m[k]=(m[k]||0)+1;return m;},{});
}

function renderCharts(data) {
  Object.values(charts).forEach(c=>{try{c.destroy();}catch(e){}});
  charts={};
  const isAlm = activeView.profile === 'centros_alm';
  const isVf = activeView.profile === 'consolidado_vf';

  const IND_ORD = ['DESABASTECIDO','SUBSTOCK','NORMOSTOCK','SOBRESTOCK','SIN ROTACION'];
  const ROT_ORD = ['ALTA ROTACION','MEDIA ROTACION','BAJA ROTACION','SIN ROTACION'];

  const covBins = vals => {
    const bins=[
      {lo:0,hi:0.01,lbl:'0 meses'},
      {lo:0.01,hi:1,lbl:'< 1 mes'},
      {lo:1,hi:3,lbl:'1–3 m'},
      {lo:3,hi:6,lbl:'3–6 m'},
      {lo:6,hi:12,lbl:'6–12 m'},
      {lo:12,hi:9999,lbl:'> 12 m'},
    ];
    return {
      labels: bins.map(b=>b.lbl),
      data: bins.map(b=>vals.filter(v=>v>=b.lo&&v<b.hi).length)
    };
  };

  if (isVf) {
    const ic = data.reduce((m,d)=>{const k=d.INDICADOR_NEW||d.INDICADOR||'N/A';m[k]=(m[k]||0)+1;return m;},{});
    charts.ind = new Chart(document.getElementById('cIndicador'),{
      type:'doughnut',
      data:{labels:IND_ORD,datasets:[{data:IND_ORD.map(k=>ic[k]||0),backgroundColor:IND_ORD.map(k=>PAL[k]),borderWidth:2,borderColor:'#0f1018',hoverOffset:5}]},
      options:{...baseOpts(),cutout:'62%'}
    });

    const stockC = data.reduce((a,d)=>a+d.STOCK_CENTROS,0);
    const stockA = data.reduce((a,d)=>a+d.STOCK_ALM,0);
    charts.rot = new Chart(document.getElementById('cRotacion'),{
      type:'doughnut',
      data:{labels:['Centros','Almacén'],datasets:[{data:[stockC,stockA],backgroundColor:['#3b82f6','#8b5cf6'],borderWidth:2,borderColor:'#0f1018',hoverOffset:5}]},
      options:{...baseOpts(),cutout:'62%'}
    });

    const ec = cnt(data,'ESTRATEGIC');
    const eTop = Object.entries(ec).sort((a,b)=>b[1]-a[1]).slice(0,8);
    charts.forma = new Chart(document.getElementById('cForma'),{
      type:'bar',
      data:{labels:eTop.map(([k])=>k||'N/A'),datasets:[{data:eTop.map(([,v])=>v),backgroundColor:'rgba(108,99,255,.7)',borderColor:'#6c63ff',borderWidth:1,borderRadius:4}]},
      options:{...baseOpts(false),indexAxis:'y',scales:{x:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#5a5b72',font:{size:9}}},y:{grid:{display:false},ticks:{color:'#ecedf5',font:{size:9}}}}}
    });

    const dispo = covBins(data.map(d=>d.DISPO));
    charts.cob = new Chart(document.getElementById('cCobertura'),{
      type:'bar',
      data:{labels:dispo.labels,datasets:[{data:dispo.data,backgroundColor:['#f43f5e','#f97316','#eab308','#22c55e','#3b82f6','#06b6d4'],borderRadius:5}]},
      options:{...baseOpts(false),scales:{x:{grid:{display:false},ticks:{color:'#ecedf5',font:{size:10}}},y:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#5a5b72',font:{size:9}}}}}
    });

    const toDate = s => { const d = new Date(s); return Number.isNaN(d.getTime()) ? null : d; };
    const today = new Date(); today.setHours(0,0,0,0);
    const vb = {VENCIDO:0,'<= 90 días':0,'91–180 días':0,'> 180 días':0,'SIN FECHA':0};
    data.forEach(d => {
      const c = toDate(d.FECHA_VENCIMIENTO_CENTRO);
      const a = toDate(d.FECHA_VENCIMIENTO_ALMACEN);
      const nearest = [c,a].filter(Boolean).sort((x,y)=>x-y)[0];
      if (!nearest) { vb['SIN FECHA']++; return; }
      const diff = Math.ceil((nearest - today)/86400000);
      if (diff < 0) vb['VENCIDO']++;
      else if (diff <= 90) vb['<= 90 días']++;
      else if (diff <= 180) vb['91–180 días']++;
      else vb['> 180 días']++;
    });
    const vOrder = ['VENCIDO','<= 90 días','91–180 días','> 180 días','SIN FECHA'];
    charts.cons = new Chart(document.getElementById('cConsumo'),{
      type:'doughnut',
      data:{labels:vOrder,datasets:[{data:vOrder.map(k=>vb[k]||0),backgroundColor:['#f43f5e','#f97316','#eab308','#22c55e','#5a5b72'],borderWidth:2,borderColor:'#0f1018',hoverOffset:5}]},
      options:{...baseOpts(),cutout:'62%'}
    });

    const top12=[...data].sort((a,b)=>b.CPMA-a.CPMA).slice(0,12);
    const trunc=s=>s.length>38?s.slice(0,36)+'…':s;
    charts.top = new Chart(document.getElementById('cTopConsumo'),{
      type:'bar',
      data:{labels:top12.map(d=>trunc(d.NOMBRE_PRODUCTO)),datasets:[{data:top12.map(d=>Math.round(d.CPMA)),backgroundColor:top12.map(d=>PAL[d.INDICADOR_NEW||d.INDICADOR]||'#6c63ff'),borderRadius:4}]},
      options:{...baseOpts(false),indexAxis:'y',scales:{x:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#5a5b72',font:{size:9}}},y:{grid:{display:false},ticks:{color:'#ecedf5',font:{size:9}}}}}
    });

    ['M','I'].forEach(tipo => {
      const tData = data.filter(d=>d.TIPO===tipo);
      const tc = tData.reduce((m,d)=>{const k=d.INDICADOR_NEW||d.INDICADOR||'N/A';m[k]=(m[k]||0)+1;return m;},{});
      charts['tipo'+tipo] = new Chart(document.getElementById('cTipo'+tipo),{
        type:'bar',
        data:{labels:IND_ORD,datasets:[{data:IND_ORD.map(k=>tc[k]||0),backgroundColor:IND_ORD.map(k=>PAL[k]),borderRadius:5}]},
        options:{...baseOpts(false),scales:{x:{grid:{display:false},ticks:{color:'#ecedf5',font:{size:9}}},y:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#5a5b72',font:{size:9}}}}}
      });
    });
    return;
  }

  if (isAlm) {
    const ia = data.reduce((m,d)=>{const k=d.INDICADOR_ALM||d.INDICADOR||'N/A';m[k]=(m[k]||0)+1;return m;},{});
    charts.ind = new Chart(document.getElementById('cIndicador'),{
      type:'doughnut',
      data:{labels:IND_ORD,datasets:[{data:IND_ORD.map(k=>ia[k]||0),backgroundColor:IND_ORD.map(k=>PAL[k]),borderWidth:2,borderColor:'#0f1018',hoverOffset:5}]},
      options:{...baseOpts(),cutout:'62%'}
    });

    const ic = data.reduce((m,d)=>{const k=d.INDICADOR_CENTROS||d.INDICADOR||'N/A';m[k]=(m[k]||0)+1;return m;},{});
    charts.rot = new Chart(document.getElementById('cRotacion'),{
      type:'doughnut',
      data:{labels:IND_ORD,datasets:[{data:IND_ORD.map(k=>ic[k]||0),backgroundColor:IND_ORD.map(k=>PAL[k]),borderWidth:2,borderColor:'#0f1018',hoverOffset:5}]},
      options:{...baseOpts(),cutout:'62%'}
    });

    const tcTipo = cnt(data,'TIPO');
    charts.forma = new Chart(document.getElementById('cForma'),{
      type:'bar',
      data:{labels:Object.keys(tcTipo),datasets:[{data:Object.values(tcTipo),backgroundColor:'rgba(108,99,255,.7)',borderColor:'#6c63ff',borderWidth:1,borderRadius:4}]},
      options:{...baseOpts(false),indexAxis:'y',scales:{x:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#5a5b72',font:{size:9}}},y:{grid:{display:false},ticks:{color:'#ecedf5',font:{size:9}}}}}
    });

    const dipAlm = covBins(data.map(d=>d.DIP_ALM));
    charts.cob = new Chart(document.getElementById('cCobertura'),{
      type:'bar',
      data:{labels:dipAlm.labels,datasets:[{data:dipAlm.data,backgroundColor:['#f43f5e','#f97316','#eab308','#22c55e','#3b82f6','#06b6d4'],borderRadius:5}]},
      options:{...baseOpts(false),scales:{x:{grid:{display:false},ticks:{color:'#ecedf5',font:{size:10}}},y:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#5a5b72',font:{size:9}}}}}
    });

    const dispC = covBins(data.map(d=>d.DISP_CENTROS));
    charts.cons = new Chart(document.getElementById('cConsumo'),{
      type:'bar',
      data:{labels:dispC.labels,datasets:[{data:dispC.data,backgroundColor:['#22c55e','#3b82f6','#eab308','#8b5cf6','#f97316','#5a5b72'],borderRadius:5}]},
      options:{...baseOpts(false),scales:{x:{grid:{display:false},ticks:{color:'#ecedf5',font:{size:10}}},y:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#5a5b72',font:{size:9}}}}}
    });

    const top12=[...data].sort((a,b)=>b.CMPA_ALM-a.CMPA_ALM).slice(0,12);
    const trunc=s=>s.length>38?s.slice(0,36)+'…':s;
    charts.top = new Chart(document.getElementById('cTopConsumo'),{
      type:'bar',
      data:{labels:top12.map(d=>trunc(d.NOMBRE_PRODUCTO)),datasets:[{data:top12.map(d=>Math.round(d.CMPA_ALM)),backgroundColor:top12.map(d=>PAL[d.INDICADOR_ALM||d.INDICADOR]||'#6c63ff'),borderRadius:4}]},
      options:{...baseOpts(false),indexAxis:'y',scales:{x:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#5a5b72',font:{size:9}}},y:{grid:{display:false},ticks:{color:'#ecedf5',font:{size:9}}}}}
    });

    ['M','I'].forEach(tipo => {
      const tData = data.filter(d=>d.TIPO===tipo);
      const ta = tData.reduce((m,d)=>{const k=d.INDICADOR_ALM||d.INDICADOR||'N/A';m[k]=(m[k]||0)+1;return m;},{});
      charts['tipo'+tipo] = new Chart(document.getElementById('cTipo'+tipo),{
        type:'bar',
        data:{labels:IND_ORD,datasets:[{data:IND_ORD.map(k=>ta[k]||0),backgroundColor:IND_ORD.map(k=>PAL[k]),borderRadius:5}]},
        options:{...baseOpts(false),scales:{x:{grid:{display:false},ticks:{color:'#ecedf5',font:{size:9}}},y:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#5a5b72',font:{size:9}}}}}
      });
    });
    return;
  }

  // 1. Indicador donut
  const ic = cnt(data,'INDICADOR');
  charts.ind = new Chart(document.getElementById('cIndicador'),{
    type:'doughnut',
    data:{labels:IND_ORD,datasets:[{
      data:IND_ORD.map(k=>ic[k]||0),
      backgroundColor:IND_ORD.map(k=>PAL[k]),
      borderWidth:2,borderColor:'#0f1018',hoverOffset:5
    }]},
    options:{...baseOpts(),cutout:'62%'}
  });

  // 2. Rotación donut
  const rc = cnt(data,'NIVEL_ROTACION');
  charts.rot = new Chart(document.getElementById('cRotacion'),{
    type:'doughnut',
    data:{labels:ROT_ORD,datasets:[{
      data:ROT_ORD.map(k=>rc[k]||0),
      backgroundColor:['#3b82f6','#22c55e','#eab308','#8b5cf6'],
      borderWidth:2,borderColor:'#0f1018',hoverOffset:5
    }]},
    options:{...baseOpts(),cutout:'62%'}
  });

  // 3. Forma farmacéutica
  const fc = cnt(data,'forma_farm');
  const fTop = Object.entries(fc).sort((a,b)=>b[1]-a[1]).slice(0,8);
  charts.forma = new Chart(document.getElementById('cForma'),{
    type:'bar',
    data:{labels:fTop.map(([k])=>k),datasets:[{
      data:fTop.map(([,v])=>v),
      backgroundColor:'rgba(108,99,255,.7)',borderColor:'#6c63ff',
      borderWidth:1,borderRadius:4
    }]},
    options:{...baseOpts(false),indexAxis:'y',
      scales:{
        x:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#5a5b72',font:{size:9}}},
        y:{grid:{display:false},ticks:{color:'#ecedf5',font:{size:9}}}
      }
    }
  });

  // 4. Cobertura meses
  const cobBins=[
    {lo:0,hi:0.01,lbl:'0 meses'},
    {lo:0.01,hi:1,lbl:'< 1 mes'},
    {lo:1,hi:3,lbl:'1–3 m'},
    {lo:3,hi:6,lbl:'3–6 m'},
    {lo:6,hi:12,lbl:'6–12 m'},
    {lo:12,hi:9999,lbl:'> 12 m'},
  ];
  charts.cob = new Chart(document.getElementById('cCobertura'),{
    type:'bar',
    data:{labels:cobBins.map(b=>b.lbl),datasets:[{
      data:cobBins.map(b=>data.filter(d=>d.MES_DISPONIBLES>=b.lo&&d.MES_DISPONIBLES<b.hi).length),
      backgroundColor:['#f43f5e','#f97316','#eab308','#22c55e','#3b82f6','#06b6d4'],
      borderRadius:5
    }]},
    options:{...baseOpts(false),
      scales:{
        x:{grid:{display:false},ticks:{color:'#ecedf5',font:{size:10}}},
        y:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#5a5b72',font:{size:9}}}
      }
    }
  });

  // 5. Estado consumo
  const cc = cnt(data,'ESTADO_CONSUMO');
  const CO_ORD=['CONSUMO MUY RECIENTE','CONSUMO RECIENTE','CONSUMO ANTIGUO','SIN DATOS'];
  charts.cons = new Chart(document.getElementById('cConsumo'),{
    type:'doughnut',
    data:{labels:CO_ORD,datasets:[{
      data:CO_ORD.map(k=>cc[k]||0),
      backgroundColor:['#22c55e','#3b82f6','#eab308','#5a5b72'],
      borderWidth:2,borderColor:'#0f1018',hoverOffset:5
    }]},
    options:{...baseOpts(),cutout:'62%'}
  });

  // 6. Top 12 consumo
  const top12=[...data].sort((a,b)=>b.CMPA-a.CMPA).slice(0,12);
  const trunc=s=>s.length>38?s.slice(0,36)+'…':s;
  charts.top = new Chart(document.getElementById('cTopConsumo'),{
    type:'bar',
    data:{labels:top12.map(d=>trunc(d.NOMBRE_PRODUCTO)),datasets:[{
      data:top12.map(d=>Math.round(d.CMPA)),
      backgroundColor:top12.map(d=>PAL[d.INDICADOR]||'#6c63ff'),
      borderRadius:4
    }]},
    options:{...baseOpts(false),indexAxis:'y',
      scales:{
        x:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#5a5b72',font:{size:9}}},
        y:{grid:{display:false},ticks:{color:'#ecedf5',font:{size:9}}}
      }
    }
  });

  // 7+8. Tipo M vs I
  ['M','I'].forEach(tipo => {
    const tData = data.filter(d=>d.TIPO===tipo);
    const tc = cnt(tData,'INDICADOR');
    charts['tipo'+tipo] = new Chart(document.getElementById('cTipo'+tipo),{
      type:'bar',
      data:{labels:IND_ORD,datasets:[{
        data:IND_ORD.map(k=>tc[k]||0),
        backgroundColor:IND_ORD.map(k=>PAL[k]),borderRadius:5
      }]},
      options:{...baseOpts(false),
        scales:{
          x:{grid:{display:false},ticks:{color:'#ecedf5',font:{size:9}}},
          y:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#5a5b72',font:{size:9}}}
        }
      }
    });
  });
}

// ═══════════════════════════════════════════════════════════════
// ALERTS
// ═══════════════════════════════════════════════════════════════
function renderAlerts(data) {
  const isAlm = activeView.profile === 'centros_alm';
  const isVf = activeView.profile === 'consolidado_vf';

  let crit;
  if (isVf) {
    const toDate = s => { const d = new Date(s); return Number.isNaN(d.getTime()) ? null : d; };
    const today = new Date(); today.setHours(0,0,0,0);
    crit = data.filter(d => {
      const ind = d.INDICADOR_NEW || d.INDICADOR;
      if (ind === 'DESABASTECIDO' || ind === 'SUBSTOCK') return true;
      const c = toDate(d.FECHA_VENCIMIENTO_CENTRO);
      const a = toDate(d.FECHA_VENCIMIENTO_ALMACEN);
      const nearest = [c,a].filter(Boolean).sort((x,y)=>x-y)[0];
      if (!nearest) return false;
      const diff = Math.ceil((nearest - today)/86400000);
      return diff >= 0 && diff <= 90;
    }).sort((a,b)=>b.CPMA-a.CPMA).slice(0,25);
  } else if (isAlm) {
    crit = data.filter(d => {
      const ia = d.INDICADOR_ALM || d.INDICADOR;
      return ia === 'DESABASTECIDO' || ia === 'SUBSTOCK' || (d.DIP_ALM < 1 && d.CMPA_ALM > 0);
    }).sort((a,b)=>b.CMPA_ALM-a.CMPA_ALM).slice(0,25);
  } else {
    crit = data
      .filter(d=>d.INDICADOR==='DESABASTECIDO'&&d.NIVEL_ROTACION==='ALTA ROTACION')
      .sort((a,b)=>b.CMPA-a.CMPA).slice(0,25);
  }

  const el = document.getElementById('alerts-list');
  if(!el) return;
  if(!crit.length){el.innerHTML='<p style="text-align:center;color:var(--muted);padding:20px;font-size:12px">Sin alertas críticas</p>';return;}
  el.innerHTML = crit.map(d=>`
    <div class="a-item crit">
      <div class="a-dot"></div>
      <div style="flex:1;min-width:0">
        <div class="a-name">${d.NOMBRE_PRODUCTO}</div>
        <div class="a-meta">Cód: ${d.MEDCOD} · ${isVf ? `CPMA: ${Math.round(d.CPMA).toLocaleString()} u/mes · Dispo: ${d.DISPO.toFixed(2)} · Venc.C: ${d.FECHA_VENCIMIENTO_CENTRO || '—'}` : isAlm ? `CMPA_ALM: ${Math.round(d.CMPA_ALM).toLocaleString()} u/mes · DIP_ALM: ${d.DIP_ALM.toFixed(2)} · IND_ALM: ${d.INDICADOR_ALM || '—'}` : `CMPA: ${Math.round(d.CMPA).toLocaleString()} u/mes · ${d.forma_farm}`}</div>
        <div class="a-meta" style="color:rgba(244,63,94,.65);margin-top:2px;font-style:italic">${d.INTERPRETACION}</div>
      </div>
      <span class="a-tag">${isVf ? (d.INDICADOR_NEW || d.INDICADOR || 'Revisar') : isAlm ? (d.INDICADOR_ALM || 'Revisar') : 'Sin stock'}</span>
    </div>`).join('');
}

// ═══════════════════════════════════════════════════════════════
// TABLE
// ═══════════════════════════════════════════════════════════════
const IND_PILL = {
  DESABASTECIDO:'red', SUBSTOCK:'orange',
  NORMOSTOCK:'green',  SOBRESTOCK:'yellow', 'SIN ROTACION':'violet'
};
const ROT_PILL = {
  'ALTA ROTACION':'blue','MEDIA ROTACION':'green',
  'BAJA ROTACION':'yellow','SIN ROTACION':'gray'
};

function filterTable() {
  const isVf = activeView.profile === 'consolidado_vf';
  const s = (document.getElementById('t-search')?.value||'').toLowerCase();
  const fi = document.getElementById('t-ind')?.value||'';
  const fr = document.getElementById('t-rot')?.value||'';
  const ft = document.getElementById('t-tipo')?.value||'';
  filteredData = allData.filter(d=>{
    const indicator = isVf ? (d.INDICADOR_NEW || d.INDICADOR) : d.INDICADOR;
    if(fi && indicator!==fi) return false;
    if(fr && d.NIVEL_ROTACION!==fr) return false;
    if(ft && d.TIPO!==ft) return false;
    if(s && !d.NOMBRE_PRODUCTO.toLowerCase().includes(s) && !d.MEDCOD.includes(s)) return false;
    return true;
  });
  filteredData.sort((a,b)=>{
    const av=a[sortCol], bv=b[sortCol];
    if(typeof av==='number') return (av-bv)*sortDir;
    return String(av).localeCompare(String(bv))*sortDir;
  });
  page=1; renderTable();
}

function sortTbl(col) {
  if(sortCol===col) sortDir*=-1; else{sortCol=col;sortDir=-1;}
  filterTable();
}

function renderTable() {
  const isAlm = activeView.profile === 'centros_alm';
  const isVf = activeView.profile === 'consolidado_vf';

  const total = filteredData.length;
  const totalP = Math.ceil(total/PAGE)||1;
  const slice = filteredData.slice((page-1)*PAGE, page*PAGE);

  const cnt = document.getElementById('t-count');
  const pager = document.getElementById('t-pager');
  const tbody = document.getElementById('t-body');
  if(!cnt||!pager||!tbody) return;

  cnt.textContent = total.toLocaleString()+' productos';
  pager.textContent = `Pág. ${page} / ${totalP}`;

  if(!slice.length){
    const cols = isAlm ? 18 : isVf ? 15 : 11;
    tbody.innerHTML=`<tr><td colspan="${cols}" style="text-align:center;padding:36px;color:var(--muted)">Sin resultados</td></tr>`;
    return;
  }

  const mesesBar = v => {
    const fill = Math.min(v/12*100,100);
    const col = v===0?'var(--red)':v<1?'var(--orange)':v<3?'var(--yellow)':'var(--green)';
    return `<div class="prog-row">
      <div class="prog-bar"><div class="prog-fill" style="width:${fill}%;background:${col}"></div></div>
      <span class="prog-val">${v.toFixed(1)}</span></div>`;
  };

  tbody.innerHTML = slice.map(d=>`
    <tr>
      <td class="mono">${d.MEDCOD}</td>
      <td style="max-width:270px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${d.NOMBRE_PRODUCTO}">${d.NOMBRE_PRODUCTO}</td>
      <td><span class="pill ${d.TIPO==='M'?'pill-blue':'pill-violet'}">${d.TIPO==='M'?'Med.':'Ins.'}</span></td>
      ${isAlm || isVf ? `<td class="mono" style="text-align:right">${d.PRECIO ? d.PRECIO.toLocaleString() : '0'}</td>` : ''}
      ${isVf ? `<td style="font-size:11px;color:var(--muted2)">${d.ESTRATEGIC || '—'}</td>` : ''}
      ${isAlm ? `<td class="mono" style="text-align:right">${d.STOCK_CENTROS.toLocaleString()}</td><td class="mono" style="text-align:right">${Math.round(d.CPMA_CENTROS).toLocaleString()}</td><td style="font-size:11px;color:var(--muted2)">${d.DISP_CENTROS.toFixed(2)}</td><td><span class="pill pill-${IND_PILL[d.INDICADOR_CENTROS]||'gray'}">${d.INDICADOR_CENTROS || '—'}</span></td><td class="mono" style="text-align:right">${d.STOCK_ALM.toLocaleString()}</td><td class="mono" style="text-align:right">${Math.round(d.CMPA_ALM).toLocaleString()}</td><td style="font-size:11px;color:var(--muted2)">${d.DIP_ALM.toFixed(2)}</td><td><span class="pill pill-${IND_PILL[d.INDICADOR_ALM]||'gray'}">${d.INDICADOR_ALM || '—'}</span></td>` : ''}
      ${isVf ? `<td class="mono" style="text-align:right">${d.STOCK_CENTROS.toLocaleString()}</td><td class="mono" style="text-align:right">${d.STOCK_ALM.toLocaleString()}</td><td class="mono" style="text-align:right">${d.STOCK_TOTAL.toLocaleString()}</td><td class="mono" style="text-align:right">${Math.round(d.CPMA).toLocaleString()}</td><td style="font-size:11px;color:var(--muted2)">${d.DISPO.toFixed(2)}</td>` : ''}
      ${!isVf ? `<td style="font-size:11px;color:var(--muted2)">${d.forma_farm}</td>` : ''}
      ${!isVf ? `<td class="mono" style="text-align:right">${d.STOCK_ACTUAL.toLocaleString()}</td>` : ''}
      ${!isVf ? `<td class="mono" style="text-align:right">${Math.round(d.CMPA).toLocaleString()}</td>` : ''}
      <td style="min-width:120px">${mesesBar(d.MES_DISPONIBLES)}</td>
      <td><span class="pill pill-${IND_PILL[isVf ? (d.INDICADOR_NEW || d.INDICADOR) : d.INDICADOR]||'gray'}">${isVf ? (d.INDICADOR_NEW || d.INDICADOR || '—') : d.INDICADOR}</span></td>
      ${!isAlm && !isVf ? `<td><span class="pill pill-${ROT_PILL[d.NIVEL_ROTACION]||'gray'}">${d.NIVEL_ROTACION}</span></td><td style="font-size:10px;color:var(--muted2)">${d.ESTADO_CONSUMO}</td>` : ''}
      ${isVf ? `<td class="mono" style="font-size:10px">${d.FECHA_VENCIMIENTO_CENTRO || '—'}</td><td class="mono" style="font-size:10px">${d.FECHA_VENCIMIENTO_ALMACEN || '—'}</td>` : ''}
      <td style="font-size:10px;color:var(--muted);max-width:200px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${d.INTERPRETACION}">${d.INTERPRETACION}</td>
    </tr>`).join('');
}

function prevPage(){if(page>1){page--;renderTable();}}
function nextPage(){const tp=Math.ceil(filteredData.length/PAGE);if(page<tp){page++;renderTable();}}

// ═══════════════════════════════════════════════════════════════
// CSV EXPORT
// ═══════════════════════════════════════════════════════════════
function exportCSV() {
  const isAlm = activeView.profile === 'centros_alm';
  const isVf = activeView.profile === 'consolidado_vf';

  const cols=[
    'MEDCOD','NOMBRE_PRODUCTO','TIPO',
    ...(isAlm || isVf ? ['PRECIO'] : []),
    ...(isVf ? ['ESTRATEGIC'] : []),
    ...(isAlm ? ['STOCK_CENTROS','CPMA_CENTROS','DISP_CENTROS','INDICADOR_CENTROS','STOCK_ALM','CMPA_ALM','DIP_ALM','INDICADOR_ALM'] : []),
    ...(isVf ? ['STOCK_CENTROS','STOCK_ALM','STOCK_TOTAL','CPMA','DISPO'] : []),
    'forma_farm','STOCK_ACTUAL','CMPA','MES_DISPONIBLES',
    ...(isVf ? ['INDICADOR_NEW'] : ['INDICADOR']),
    'NIVEL_ROTACION','ESTADO_CONSUMO',
    ...(isVf ? ['FECHA_VENCIMIENTO_CENTRO','FECHA_VENCIMIENTO_ALMACEN'] : []),
    'INTERPRETACION','FECHA_REPORTE'
  ];
  const rows=[cols.join(','),...filteredData.map(d=>cols.map(c=>`"${String(d[c]||'').replace(/"/g,'""')}"`).join(','))];
  const blob=new Blob([rows.join('\n')],{type:'text/csv;charset=utf-8;'});
  const a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download=`DISPOALM_export_${new Date().toISOString().split('T')[0]}.csv`;
  a.click();
}

// ═══════════════════════════════════════════════════════════════
// UTILS
// ═══════════════════════════════════════════════════════════════
function showError(msg){const e=document.getElementById('error-msg');e.textContent=msg;e.style.display='block';}
function hideError(){document.getElementById('error-msg').style.display='none';}



// ═══════════════════════════════════════════════════════════════
//  INIT — Service Account → Google Drive
// ═══════════════════════════════════════════════════════════════
async function initApp() {
  const content = document.getElementById('dash-content');
  content.innerHTML = `
    <div class="loading-overlay">
      <div class="loader"></div>
      <div class="loading-text">Autenticando con Google Drive…</div>
    </div>`;

  try {
    // 1. Get OAuth2 token via Service Account JWT
    await getAccessToken();

    // 2. List Excel files in the Drive folder
    content.querySelector('.loading-text').textContent = 'Buscando archivos en Drive…';
    driveFiles = await fetchDriveFiles();

    if (!driveFiles.length) {
      content.innerHTML = `<div class="loading-overlay"><p style="color:var(--muted)">No se encontraron archivos Excel en la carpeta <b>${FOLDER_NAME}</b>.<br><small style="color:var(--muted2)">Verifica que la carpeta esté compartida con la cuenta de servicio.</small></p></div>`;
      return;
    }

    // 3. Build tabs and load the first file
    buildTabs();
    selectFile(driveFiles[0].id);

  } catch(err) {
    console.error(err);
    content.innerHTML = `
      <div class="loading-overlay">
        <p style="color:var(--red);font-weight:600;margin-bottom:8px">Error al conectar con Google Drive</p>
        <p style="color:var(--muted2);font-size:12px;max-width:480px;text-align:center">${err.message}</p>
        <button class="btn-sm" onclick="initApp()" style="margin-top:16px">↺ Reintentar</button>
      </div>`;
  }
}

window.onload = initApp;