/* ============================================================
   Dashboard Pareto 80/20 — v2 (fechas robustas, string YMD)
   ============================================================ */

const state = {
  rawRows: [],
  charts: { pareto: null, marca: null, evolucion: null },
  lastSKU: [],
  lastMetric: 'valor',
  lastUmbral: 80
};

/* ---------- Formato ---------- */
const fmtMoney  = (n) => '$' + (isFinite(n) ? n : 0).toLocaleString('es-CO', { maximumFractionDigits: 0 });
const fmtMoney2 = (n) => '$' + (isFinite(n) ? n : 0).toLocaleString('es-CO', { maximumFractionDigits: 2 });
const fmtNum    = (n) => (isFinite(n) ? n : 0).toLocaleString('es-CO');
const escapeHtml = (s) => String(s ?? '')
  .replace(/&/g, '&amp;').replace(/</g, '&lt;')
  .replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/* ---------- Fechas: siempre 'YYYY-MM-DD' ---------- */
const pad2 = (n) => String(n).padStart(2, '0');

function dateObjToYMD(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return null;
  // Si es exactamente medianoche UTC, probablemente vino de un serial Excel
  const isUtcMidnight = d.getUTCHours() === 0 && d.getUTCMinutes() === 0
    && d.getUTCSeconds() === 0 && d.getUTCMilliseconds() === 0;
  if (isUtcMidnight) {
    return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth()+1)}-${pad2(d.getUTCDate())}`;
  }
  return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`;
}

function serialToYMD(serial) {
  if (!isFinite(serial) || serial <= 0) return null;
  const ms = Date.UTC(1899, 11, 30) + Math.floor(serial) * 86400000;
  const d = new Date(ms);
  return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth()+1)}-${pad2(d.getUTCDate())}`;
}

function stringToYMD(v) {
  if (typeof v !== 'string') return null;
  const s = v.trim();
  if (!s) return null;
  // YYYY-MM-DD (posible con hora)
  let m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) return `${m[1]}-${pad2(+m[2])}-${pad2(+m[3])}`;
  // DD/MM/YYYY o MM/DD/YYYY (si a>12 es día, si b>12 es mes)
  m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})/);
  if (m) {
    let y = +m[3]; if (y < 100) y += y < 50 ? 2000 : 1900;
    let a = +m[1], b = +m[2], day, mon;
    if (a > 12)       { day = a; mon = b; }
    else if (b > 12)  { day = b; mon = a; }
    else              { day = a; mon = b; }
    return `${y}-${pad2(mon)}-${pad2(day)}`;
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return dateObjToYMD(d);
  return null;
}

function toYMD(v) {
  if (v instanceof Date)   return dateObjToYMD(v);
  if (typeof v === 'number') return serialToYMD(v);
  if (typeof v === 'string') return stringToYMD(v);
  return null;
}

/* ---------- Encabezados ---------- */
const FIELD_ALIASES = {
  mes:          ['MES'],
  oc:           ['OC'],
  negocio:      ['NEGOCIO'],
  fechaOC:      ['FECHA OC', 'FECHA'],
  pedido:       ['# PEDIDO', 'PEDIDO'],
  factura:      ['FACTURA'],
  facturaElec:  ['# FACTURA ELECT.', '# FACTURA ELECT'],
  descripcion:  ['DESCRIPCION', 'DESCRIPCIÓN'],
  cant:         ['CANT.', 'CANT', 'CANTIDAD'],
  costoOC:      ['COSTO OC', 'COSTO'],
  ean:          ['EAN'],
  capris:       ['CAPRIS', 'SKU'],
  ref:          ['REF'],
  marcas:       ['MARCAS', 'MARCA', 'COLUMNA1', 'COLUMNA 1', 'BRAND'],
  departamento: ['DEPARTAMENTO', 'DEPTO', 'DEPARTAMENT']
};

const normHeader = (h) => String(h ?? '').trim().toUpperCase().replace(/\s+/g, ' ');

function findHeaderRow(aoa) {
  for (let i = 0; i < Math.min(aoa.length, 25); i++) {
    const row = (aoa[i] || []).map(normHeader);
    if (row.includes('MES') && row.includes('OC') && row.some(c => c.startsWith('FECHA'))) return i;
  }
  return 0;
}

function buildColumnMap(headers) {
  const map = {};
  const normalized = headers.map(normHeader);
  for (const [field, aliases] of Object.entries(FIELD_ALIASES)) {
    for (const alias of aliases) {
      const idx = normalized.indexOf(normHeader(alias));
      if (idx !== -1) { map[field] = idx; break; }
    }
  }
  return map;
}

/* ---------- Parseo de filas ---------- */
function parseRows(aoa, headerRowIdx) {
  const headers = aoa[headerRowIdx] || [];
  const map = buildColumnMap(headers);
  const rows = [];
  for (let i = headerRowIdx + 1; i < aoa.length; i++) {
    const r = aoa[i] || [];
    const get = (idx) => (idx === undefined ? '' : r[idx]);
    const ref = String(get(map.ref) ?? '').trim();
    const capris = String(get(map.capris) ?? '').trim();
    const descripcion = String(get(map.descripcion) ?? '').trim();
    if (!ref && !capris && !descripcion) continue;
    if (/^TOTAL$/i.test(descripcion) && !ref && !capris) continue;

    const fechaYMD = toYMD(get(map.fechaOC));
    if (!fechaYMD) continue;

    rows.push({
      mes: String(get(map.mes) ?? '').trim(),
      oc: String(get(map.oc) ?? '').trim(),
      negocio: String(get(map.negocio) ?? '').trim(),
      fechaYMD,
      pedido: String(get(map.pedido) ?? '').trim(),
      factura: String(get(map.factura) ?? '').trim(),
      descripcion,
      cant: Number(get(map.cant)) || 0,
      costoOC: Number(get(map.costoOC)) || 0,
      capris,
      ref,
      marca: String(get(map.marcas) ?? '').trim() || 'SIN MARCA',
      departamento: String(get(map.departamento) ?? '').trim()
    });
  }
  return { rows, map };
}

/* ---------- Carga ---------- */
async function loadFile(file) {
  const buf = await file.arrayBuffer();
  // NO usar cellDates:true → evita ambigüedad UTC/local
  const wb = XLSX.read(buf, { type: 'array' });
  const sheetName = wb.SheetNames.find(n => /CONSOLIDADO/i.test(n)) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
  const headerRowIdx = findHeaderRow(aoa);
  const { rows } = parseRows(aoa, headerRowIdx);
  if (!rows.length) {
    alert('No se encontraron filas válidas en la hoja "' + sheetName + '".');
    return;
  }
  state.rawRows = rows;
  console.log('✔ Filas cargadas:', rows.length);
  const rango = [...new Set(rows.map(r => r.fechaYMD))].sort();
  console.log('   Rango fechas:', rango[0], '→', rango[rango.length-1],
              '| Días únicos en archivo:', rango.length);
  initFilters();
  applyFilters();
}

/* ---------- Filtros ---------- */
function initFilters() {
  const marcas = [...new Set(state.rawRows.map(r => r.marca).filter(Boolean))].sort();
  const negocios = [...new Set(state.rawRows.map(r => r.negocio).filter(Boolean))].sort();

  const selM = document.getElementById('filtroMarca');
  const selN = document.getElementById('filtroNegocio');
  selM.innerHTML = '<option value="">Todas</option>' +
    marcas.map(m => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join('');
  selN.innerHTML = '<option value="">Todos</option>' +
    negocios.map(n => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join('');

  // Rango de fechas inicial = min → max (strings 'YYYY-MM-DD')
  const fechas = state.rawRows.map(r => r.fechaYMD).filter(Boolean).sort();
  if (fechas.length) {
    document.getElementById('fechaDesde').value = fechas[0];
    document.getElementById('fechaHasta').value = fechas[fechas.length - 1];
  }
}

function getFiltered() {
  // Comparación LEXICOGRÁFICA de 'YYYY-MM-DD' === orden cronológico
  const desde   = document.getElementById('fechaDesde').value; // '' o 'YYYY-MM-DD'
  const hasta   = document.getElementById('fechaHasta').value;
  const marca   = document.getElementById('filtroMarca').value;
  const negocio = document.getElementById('filtroNegocio').value;

  return state.rawRows.filter(r => {
    if (desde   && r.fechaYMD < desde) return false;
    if (hasta   && r.fechaYMD > hasta) return false;
    if (marca   && r.marca !== marca) return false;
    if (negocio && r.negocio !== negocio) return false;
    return true;
  });
}

/* ---------- Agrupación por SKU ---------- */
function groupBySKU(rows) {
  const map = new Map();
  rows.forEach(r => {
    const key = r.ref || r.capris || r.descripcion || '(sin ref)';
    if (!map.has(key)) {
      map.set(key, {
        ref: r.ref, capris: r.capris, descripcion: r.descripcion,
        marca: r.marca, unidades: 0, valor: 0
      });
    }
    const g = map.get(key);
    g.unidades += r.cant;
    g.valor    += r.costoOC;
  });
  return [...map.values()];
}

/* ---------- Refresco general ---------- */
function applyFilters() {
  if (!state.rawRows.length) return;

  const metric = document.getElementById('filtroMetrica').value;
  const umbral = Number(document.getElementById('umbral').value) || 80;
  const topN   = Number(document.getElementById('topN').value) || 20;
  document.getElementById('lblUmbral').textContent = umbral;

  const filtered = getFiltered();
  const fechasFil = [...new Set(filtered.map(r => r.fechaYMD))].sort();
  console.log(
    '🔎 Filas filtradas:', filtered.length,
    '| Rango efectivo:', fechasFil[0], '→', fechasFil[fechasFil.length-1],
    '| Días únicos:', fechasFil.length
  );

  const skus = groupBySKU(filtered);
  const metricValue = (s) => metric === 'valor' ? s.valor : s.unidades;

  skus.sort((a, b) => metricValue(b) - metricValue(a));

  const totalMetric   = skus.reduce((s, x) => s + metricValue(x), 0);
  const totalUnidades = filtered.reduce((s, r) => s + r.cant, 0);
  const totalValor    = filtered.reduce((s, r) => s + r.costoOC, 0);

  /* --- Pareto --- */
  let cum = 0, skusUmbral = 0, valorConcentrado = 0;
  const skuConAcum = skus.map((s) => {
    const v = metricValue(s);
    const share = totalMetric > 0 ? (v / totalMetric) * 100 : 0;
    const acumAnterior = cum;
    cum += share;
    if (acumAnterior < umbral - 1e-9) {
      skusUmbral += 1;
      valorConcentrado += v;
    }
    return { ...s, share, acumulado: cum };
  });

  const totalMetricConc = metric === 'valor' ? totalValor : totalUnidades;
  const pctConcentrado  = totalMetricConc > 0 ? (valorConcentrado / totalMetricConc) * 100 : 0;

  /* --- Días y OCs --- */
  const numDias = fechasFil.length;                 // días con al menos 1 fila

  const ocsVistas = new Set();
  const fechasOC  = new Set();
  filtered.forEach(r => {
    const oc = String(r.oc ?? '').trim();
    if (!oc) return;
    if (String(r.pedido  ?? '').toUpperCase() === 'ANULADO') return;
    if (String(r.factura ?? '').toUpperCase() === 'ANULADO') return;
    ocsVistas.add(oc);
    if (r.fechaYMD) fechasOC.add(r.fechaYMD);
  });
  const numOCs    = ocsVistas.size;
  const diasConOC = fechasOC.size;

  /* --- KPIs existentes --- */
  document.getElementById('kpiTotalSKUs').textContent      = fmtNum(skuConAcum.length);
  document.getElementById('kpiTotalUnidades').textContent  = fmtNum(totalUnidades);
  document.getElementById('kpiValorTotal').textContent     = fmtMoney(totalValor);
  document.getElementById('kpiSKUsUmbral').textContent     = fmtNum(skusUmbral);
  document.getElementById('kpiPctConcentrado').textContent = pctConcentrado.toFixed(1) + '%';
  document.getElementById('kpiPromedioSKU').textContent    = fmtMoney(skuConAcum.length ? totalValor / skuConAcum.length : 0);
  document.getElementById('kpiUnidadesDia').textContent    = fmtNum(numDias ? totalUnidades / numDias : 0);
  document.getElementById('kpiValorDia').textContent       = fmtMoney(numDias ? totalValor / numDias : 0);

  /* --- NUEVOS KPIs --- */
  // Promedio Venta Diaria = Σ valor / días con ventas
  const promVentaDia = numDias > 0 ? totalValor / numDias : 0;
  document.getElementById('kpiPromVentaDia').textContent = fmtMoney2(promVentaDia);

  // Promedio OC Recibidas Diarias = OCs únicas válidas / días con OC
  const promOCDia = diasConOC > 0 ? numOCs / diasConOC : 0;
  document.getElementById('kpiPromOCDia').textContent = promOCDia.toFixed(2);

  /* --- Gráficas y tabla --- */
  renderPareto(skuConAcum.slice(0, topN), metric, umbral);
  renderMarca(filtered, metric);
  renderEvolucion(filtered, metric);
  renderTable(skuConAcum, metric, umbral);

  state.lastSKU    = skuConAcum;
  state.lastMetric = metric;
  state.lastUmbral = umbral;
}

/* ---------- Gráficas ---------- */
function destroyChart(key) {
  if (state.charts[key]) { state.charts[key].destroy(); state.charts[key] = null; }
}

function umbralLinePlugin(umbral) {
  return {
    id: 'umbralLine',
    afterDraw(chart) {
      const { ctx, chartArea, scales } = chart;
      if (!scales.y1) return;
      const y = scales.y1.getPixelForValue(umbral);
      ctx.save();
      ctx.beginPath();
      ctx.setLineDash([6, 4]);
      ctx.strokeStyle = 'rgba(220,38,38,0.8)';
      ctx.lineWidth = 1.6;
      ctx.moveTo(chartArea.left, y);
      ctx.lineTo(chartArea.right, y);
      ctx.stroke();
      ctx.restore();
    }
  };
}

function renderPareto(data, metric, umbral) {
  const labels      = data.map(d => d.ref || d.capris || '(sin ref)');
  const values      = data.map(d => metric === 'valor' ? d.valor : d.unidades);
  const acumulados  = data.map(d => d.acumulado);
  const metricLabel = metric === 'valor' ? 'Valor ($)' : 'Unidades';

  destroyChart('pareto');
  const ctx = document.getElementById('chartPareto').getContext('2d');
  state.charts.pareto = new Chart(ctx, {
    data: {
      labels,
      datasets: [
        {
          type: 'bar', label: metricLabel, data: values, yAxisID: 'y',
          backgroundColor: 'rgba(37,99,235,0.75)',
          borderColor: 'rgba(37,99,235,1)',
          borderWidth: 1, order: 2
        },
        {
          type: 'line', label: '% Acumulado', data: acumulados, yAxisID: 'y1',
          borderColor: '#dc2626', backgroundColor: 'rgba(220,38,38,0.1)',
          tension: 0.25, pointRadius: 3, fill: false, order: 1
        }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { position: 'top' },
        datalabels: { display: false },
        tooltip: {
          callbacks: {
            label: (c) => c.dataset.yAxisID === 'y1'
              ? `Acum: ${c.parsed.y.toFixed(2)}%`
              : (metric === 'valor' ? `Valor: ${fmtMoney(c.parsed.y)}` : `Unidades: ${fmtNum(c.parsed.y)}`)
          }
        }
      },
      scales: {
        x: { type: 'category', ticks: { autoSkip: true, maxRotation: 60, minRotation: 45, font: { size: 10 } } },
        y: { position: 'left', beginAtZero: true, title: { display: true, text: metricLabel } },
        y1: {
          position: 'right', beginAtZero: true, max: 100,
          grid: { drawOnChartArea: false },
          title: { display: true, text: '% Acumulado' },
          ticks: { callback: v => v + '%' }
        }
      }
    },
    plugins: [umbralLinePlugin(umbral)]
  });
}

function renderMarca(rows, metric) {
  const map = new Map();
  rows.forEach(r => {
    const m = r.marca || 'SIN MARCA';
    if (!map.has(m)) map.set(m, { valor: 0, unidades: 0 });
    const g = map.get(m);
    g.valor    += r.costoOC;
    g.unidades += r.cant;
  });
  const arr = [...map.entries()]
    .map(([k, v]) => ({ marca: k, valor: v.valor, unidades: v.unidades }))
    .sort((a, b) => metric === 'valor' ? b.valor - a.valor : b.unidades - a.unidades)
    .slice(0, 15);

  destroyChart('marca');
  const ctx = document.getElementById('chartMarca').getContext('2d');
  state.charts.marca = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: arr.map(x => x.marca),
      datasets: [{
        label: metric === 'valor' ? 'Valor ($)' : 'Unidades',
        data: arr.map(x => metric === 'valor' ? x.valor : x.unidades),
        backgroundColor: 'rgba(5,150,105,0.8)',
        borderColor: 'rgba(5,150,105,1)',
        borderWidth: 1
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        datalabels: { display: false },
        tooltip: { callbacks: { label: (c) => metric === 'valor' ? fmtMoney(c.parsed.y) : fmtNum(c.parsed.y) } }
      },
      scales: { y: { beginAtZero: true } }
    }
  });
}

function renderEvolucion(rows, metric) {
  // Agrupa por r.fechaYMD (string 'YYYY-MM-DD')
  const map = new Map();
  rows.forEach(r => {
    const k = r.fechaYMD;
    if (!k) return;
    if (!map.has(k)) map.set(k, { valor: 0, unidades: 0 });
    const g = map.get(k);
    g.valor    += r.costoOC;
    g.unidades += r.cant;
  });
  const keys   = [...map.keys()].sort();
  const values = keys.map(k => metric === 'valor' ? map.get(k).valor : map.get(k).unidades);

  destroyChart('evolucion');
  const ctx = document.getElementById('chartEvolucion').getContext('2d');
  state.charts.evolucion = new Chart(ctx, {
    type: 'line',
    data: {
      labels: keys,
      datasets: [{
        label: metric === 'valor' ? 'Valor diario ($)' : 'Unidades diarias',
        data: values,
        borderColor: '#8b5cf6',
        backgroundColor: 'rgba(139,92,246,0.15)',
        fill: true, tension: 0.3, pointRadius: 2
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        datalabels: { display: false },
        tooltip: { callbacks: { label: (c) => metric === 'valor' ? fmtMoney(c.parsed.y) : fmtNum(c.parsed.y) } }
      },
      scales: {
        x: { ticks: { autoSkip: true, maxRotation: 45, font: { size: 10 } } },
        y: { beginAtZero: true }
      }
    }
  });
}

/* ---------- Tabla ---------- */
function renderTable(skuConAcum, metric, umbral) {
  const tbody = document.querySelector('#tablaSKUs tbody');
  if (!skuConAcum.length) {
    tbody.innerHTML = '<tr><td colspan="9" class="empty">Sin datos para los filtros seleccionados.</td></tr>';
    return;
  }
  tbody.innerHTML = skuConAcum.map(s => {
    const inPareto = s.share > 0 && (s.acumulado - s.share) < (umbral - 1e-9);
    const badge = inPareto
      ? '<span class="badge badge-yes">✓</span>'
      : '<span class="badge badge-no">—</span>';
    return `<tr>
      <td>${escapeHtml(s.ref)}</td>
      <td>${escapeHtml(s.capris)}</td>
      <td title="${escapeHtml(s.descripcion)}">${escapeHtml(s.descripcion)}</td>
      <td>${escapeHtml(s.marca)}</td>
      <td class="num">${fmtNum(s.unidades)}</td>
      <td class="num">${fmtMoney(s.valor)}</td>
      <td class="num">${s.share.toFixed(2)}%</td>
      <td class="num">${s.acumulado.toFixed(2)}%</td>
      <td class="center">${badge}</td>
    </tr>`;
  }).join('');
}

/* ---------- Exportar CSV ---------- */
function exportCSV() {
  if (!state.lastSKU.length) { alert('No hay datos para exportar.'); return; }
  const headers = ['REF (SKU)','CAPRIS','Descripcion','Marca','Unidades','Valor','Participacion','% Acumulado','Umbral ' + state.lastUmbral + '%'];
  const lines = [headers.join(',')];
  state.lastSKU.forEach(s => {
    const inPareto = s.share > 0 && (s.acumulado - s.share) < (state.lastUmbral - 1e-9);
    const row = [s.ref, s.capris, s.descripcion, s.marca,
                 s.unidades, s.valor.toFixed(2),
                 s.share.toFixed(2), s.acumulado.toFixed(2),
                 inPareto ? 'SI' : 'NO'];
    lines.push(row.map(v => `"${String(v ?? '').replace(/"/g,'""')}"`).join(','));
  });
  const blob = new Blob(['\uFEFF' + lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `pareto_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

/* ---------- Plantilla Excel ---------- */
function downloadTemplate() {
  const headers = [
    'MES','OC','NEGOCIO','FECHA OC','# PEDIDO','FACTURA','# FACTURA ELECT.',
    'DESCRIPCION','CANT.','COSTO OC','EAN','CAPRIS','REF','MARCAS','DEPARTAMENTO'
  ];
  const sample = [
    'ENERO','15071115','MKP','2026-01-01','GX-10631','GX-10648','GX-10645',
    'MP CINTURON PORTAHERRAMIENTA 12 BOLSILLO MAX 10 Kg', 1, 32287,
    '6942210221315','789418','THT16P20325','TOTAL','HERRAMIENTAS'
  ];
  const ws = XLSX.utils.aoa_to_sheet([headers, sample]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'CONSOLIDADO 2026');
  XLSX.writeFile(wb, 'plantilla_pareto.xlsx');
}

/* ---------- Eventos ---------- */
function bindEvents() {
  document.getElementById('fileInput').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try { await loadFile(file); }
    catch (err) { console.error(err); alert('Error al leer el archivo: ' + err.message); }
  });

  ['fechaDesde','fechaHasta','filtroMarca','filtroNegocio','filtroMetrica','topN']
    .forEach(id => document.getElementById(id).addEventListener('change', applyFilters));

  const umbral = document.getElementById('umbral');
  umbral.addEventListener('input', () => {
    document.getElementById('lblUmbral').textContent = umbral.value;
  });
  umbral.addEventListener('change', applyFilters);

  document.getElementById('btnExport').addEventListener('click', exportCSV);
  document.getElementById('btnTemplate').addEventListener('click', downloadTemplate);
}

/* ---------- Init ---------- */
document.addEventListener('DOMContentLoaded', () => {
  if (window.Chart) {
    Chart.defaults.font.family = 'system-ui, -apple-system, "Segoe UI", Roboto, sans-serif';
    Chart.defaults.color = '#374151';
  }
  if (window.ChartDataLabels && window.Chart) {
    Chart.register(window.ChartDataLabels);
    Chart.defaults.plugins.datalabels.display = false;
  }
  bindEvents();
});