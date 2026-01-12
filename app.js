/* PowerBI Lite | GitHub Pages | XLSX + Chart.js
   Hojas:
   - DATA_KPIS: anio, mes_num, sucursal, vendedor, kpi, valor
   - DATA_DIM : anio, mes_num, sucursal, dimension, clave, nombre, ventas_con_iva, venta_sin_iva, costo, utilidad, margen
*/

let KPIS = [];
let DIM = [];
let chartMain = null;
let chartDim = null;

const $ = (id) => document.getElementById(id);

const els = {
  // nav
  navExec: $("navExec"),
  navAccum: $("navAccum"),
  navDim: $("navDim"),
  pageTitle: $("pageTitle"),
  modePill: $("modePill"),

  // file actions
  fileInput: $("fileInput"),
  btnDemo: $("btnDemo"),
  btnReset: $("btnReset"),
  dataStatus: $("dataStatus"),

  // slicers
  yearSelect: $("yearSelect"),
  monthSelect: $("monthSelect"),
  branchSelect: $("branchSelect"),
  sellerSelect: $("sellerSelect"),
  kpiSelect: $("kpiSelect"),
  viewMode: $("viewMode"),

  // accum controls
  accControls: $("accControls"),
  yearASelect: $("yearASelect"),
  yearBSelect: $("yearBSelect"),
  accStartMonth: $("accStartMonth"),
  accEndMonth: $("accEndMonth"),
  accChart: $("accChart"),

  chipYTD: $("chipYTD"),
  chipEneJun: $("chipEneJun"),
  chipJulNov: $("chipJulNov"),
  chipFull: $("chipFull"),

  // KPI cards
  kpiLabelA: $("kpiLabelA"),
  kpiLabelB: $("kpiLabelB"),
  kpiValueA: $("kpiValueA"),
  kpiValueB: $("kpiValueB"),
  kpiDiffAbs: $("kpiDiffAbs"),
  kpiDiffPct: $("kpiDiffPct"),
  badgeA: $("badgeA"),
  badgeB: $("badgeB"),
  badgeDiff: $("badgeDiff"),
  badgePct: $("badgePct"),

  // main chart/table
  chartTitle: $("chartTitle"),
  chartMeta: $("chartMeta"),
  chartMain: $("chartMain"),
  tableTitle: $("tableTitle"),
  tableMeta: $("tableMeta"),
  warnings: $("warnings"),
  btnExportTable: $("btnExportTable"),
  mainTable: $("mainTable"),
  mainTbody: $("mainTable").querySelector("tbody"),
  thGroup: $("thGroup"),

  // dim page
  dimPage: $("dimPage"),
  dimSelect: $("dimSelect"),
  dimMetric: $("dimMetric"),
  topN: $("topN"),
  chartDim: $("chartDim"),
  dimTable: $("dimTable"),
  dimTbody: $("dimTable").querySelector("tbody"),
  dimMeta: $("dimMeta"),
  dimTableMeta: $("dimTableMeta"),
  dimWarnings: $("dimWarnings"),
  btnExportDim: $("btnExportDim"),

  // tooltip
  kpiTooltip: $("kpiTooltip"),
};

// =================== CONFIG “EJECUTIVO” ===================
// Semáforo por KPI (simple y editable):
// - good: umbral para verde (>= para KPIs que “más es mejor”)
// - warn: umbral para amarillo
// - direction: "up" (más es mejor) o "down" (menos es mejor)
// - unit: "money"|"percent"|"int"|"number" para formato forzado (si no, se infiere por nombre)
const KPI_RULES = {
  "Margen de Utilidad": { direction:"up", unit:"percent", good: 0.35, warn: 0.25 },
  "Ticket Promedio (CON IVA)": { direction:"up", unit:"money", good: 550, warn: 450 },
  "Transacciones": { direction:"up", unit:"int", good: 1200, warn: 900 },
  // Ejemplo para “menos es mejor”:
  // "Devoluciones": { direction:"down", unit:"int", good: 10, warn: 25 },
};

// ===== Tooltips tipo Power BI (Definición + Fórmula) =====
// Si un KPI no está aquí, mostrará un tooltip genérico.
const KPI_TOOLTIPS = {
  "Ventas Totales (CON IVA)": {
    definicion: "Importe total vendido incluyendo IVA para el filtro seleccionado.",
    formula: "SUM(valor) donde kpi = 'Ventas Totales (CON IVA)'",
    interpretacion: "Sirve para medir el tamaño de venta (ingreso bruto con IVA)."
  },
  "Ventas Contado (CON IVA)": {
    definicion: "Importe vendido en contado (incluye IVA).",
    formula: "SUM(valor) donde kpi = 'Ventas Contado (CON IVA)'",
    interpretacion: "Mide flujo de caja inmediato."
  },
  "Ventas Crédito (CON IVA)": {
    definicion: "Importe vendido a crédito (incluye IVA).",
    formula: "SUM(valor) donde kpi = 'Ventas Crédito (CON IVA)'",
    interpretacion: "Mide colocación a crédito y necesidad de cobranza."
  },
  "Utilidad Total (SIN IVA)": {
    definicion: "Utilidad generada (sin IVA) para el filtro seleccionado.",
    formula: "SUM(valor) donde kpi = 'Utilidad Total (SIN IVA)'",
    interpretacion: "Indica ganancia operativa antes de gastos (según tu definición de utilidad)."
  },
  "Margen de Utilidad": {
    definicion: "Porcentaje de utilidad sobre ventas (según definición del archivo).",
    formula: "Margen = Utilidad / Ventas (idealmente SIN IVA)  (según cómo lo calcule el origen)",
    interpretacion: "Más alto = mejor rentabilidad. Revisa cambios por descuentos/costos."
  },
  "Transacciones": {
    definicion: "Número de tickets/operaciones (transacciones) en el periodo.",
    formula: "SUM(valor) donde kpi = 'Transacciones'",
    interpretacion: "Mide volumen de actividad; junto con Ticket Promedio explica crecimiento."
  },
  "Ticket Promedio (CON IVA)": {
    definicion: "Promedio de venta por transacción (incluye IVA).",
    formula: "Ticket = Ventas Totales (CON IVA) / Transacciones (según origen)",
    interpretacion: "Sube por mix, precios o más piezas por ticket."
  }
};

const MONTH_NAMES = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];

// =================== helpers ===================
function normalizeKey(k){
  if(!k) return "";
  return String(k).trim().toLowerCase()
    .replace(/\s+/g, "_")
    .replace(/[^\wáéíóúüñ]/gi, "");
}

function toNumber(v){
  if(v === null || v === undefined) return NaN;
  if(typeof v === "number") return v;
  const s = String(v).trim();
  if(!s) return NaN;

  const hasComma = s.includes(",");
  const hasDot = s.includes(".");
  let cleaned = s;

  if(hasComma && hasDot) cleaned = s.replace(/,/g, "");
  else if(hasComma && !hasDot) cleaned = s.replace(/\./g, "").replace(",", ".");
  else cleaned = s.replace(/,/g, "");

  const n = Number(cleaned);
  return Number.isFinite(n) ? n : NaN;
}

function fmtMoney(n){
  if(!Number.isFinite(n)) return "—";
  return new Intl.NumberFormat("es-MX", { style:"currency", currency:"MXN", maximumFractionDigits:2 }).format(n);
}
function fmtInt(n){
  if(!Number.isFinite(n)) return "—";
  return new Intl.NumberFormat("es-MX", { maximumFractionDigits:0 }).format(n);
}
function fmtNum(n, d=2){
  if(!Number.isFinite(n)) return "—";
  return new Intl.NumberFormat("es-MX", { maximumFractionDigits:d }).format(n);
}
function fmtPct(r){
  if(!Number.isFinite(r)) return "—";
  return new Intl.NumberFormat("es-MX", { style:"percent", maximumFractionDigits:2 }).format(r);
}

function uniqueSorted(arr){
  return [...new Set(arr)]
    .filter(v => v !== null && v !== undefined && String(v).trim() !== "")
    .sort((a,b)=>{
      const na = Number(a), nb = Number(b);
      if(Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
      return String(a).localeCompare(String(b), "es");
    });
}

function setSelectOptions(selectEl, values, includeAll=true, allLabel="Todos"){
  const prev = selectEl.value;
  selectEl.innerHTML = "";

  if(includeAll){
    const opt = document.createElement("option");
    opt.value = "__ALL__";
    opt.textContent = allLabel;
    selectEl.appendChild(opt);
  }
  for(const v of values){
    const opt = document.createElement("option");
    opt.value = String(v);
    opt.textContent = String(v);
    selectEl.appendChild(opt);
  }

  if(prev && [...selectEl.options].some(o=>o.value===prev)){
    selectEl.value = prev;
  }else{
    selectEl.value = includeAll ? "__ALL__" : (values[0] ?? "");
  }
}

function aggSum(values){
  if(!values.length) return NaN;
  let s=0;
  for(const v of values) s += v;
  return s;
}

function aggStats(values){
  if(!values.length) return { sum:NaN, avg:NaN, max:NaN, min:NaN };
  let sum=0, max=-Infinity, min=Infinity;
  for(const v of values){
    sum += v;
    if(v > max) max = v;
    if(v < min) min = v;
  }
  return { sum, avg: sum/values.length, max, min };
}

function groupBy(rows, key){
  const m = new Map();
  for(const r of rows){
    const k = String(r[key] ?? "");
    if(!m.has(k)) m.set(k, []);
    m.get(k).push(r);
  }
  return m;
}

function escapeHtml(s){
  return String(s)
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}

function monthsList(){ return Array.from({length:12}, (_,i)=> i+1); }
function clampMonth(m){
  const n = Number(m);
  if(!Number.isFinite(n)) return 1;
  return Math.min(12, Math.max(1, Math.trunc(n)));
}

function computeDelta(a, b){
  if(!Number.isFinite(a) || !Number.isFinite(b)) return { abs:NaN, pct:NaN };
  const abs = a - b;
  const pct = (b===0) ? NaN : (abs/b);
  return { abs, pct };
}

function inferUnitFromKpiName(kpi){
  const k = (kpi||"").toLowerCase();
  if(k.includes("margen") || k.includes("%") || k.includes("porcentaje")) return "percent";
  if(k.includes("venta") || k.includes("utilidad") || k.includes("costo") || k.includes("ticket")) return "money";
  if(k.includes("transac") || k.includes("operac") || k.includes("clientes") || k.includes("sku")) return "int";
  return "number";
}

function formatByUnit(v, unit){
  if(!Number.isFinite(v)) return "—";
  if(unit === "money") return fmtMoney(v);
  if(unit === "int") return fmtInt(v);
  if(unit === "percent") {
    // acepta 0.35 o 35
    if(v >= 0 && v <= 1.5) return fmtPct(v);
    return fmtPct(v/100);
  }
  return fmtNum(v);
}

function getKpiRule(kpi){
  const rule = KPI_RULES[kpi];
  if(rule) return rule;
  return { direction:"up", unit: inferUnitFromKpiName(kpi), good: null, warn: null };
}

function setBadge(el, text, cls){
  el.classList.remove("good","warn","bad");
  el.textContent = text;
  if(cls) el.classList.add(cls);
}

function classifySemaforo(value, kpi){
  const rule = getKpiRule(kpi);
  if(rule.good === null || rule.warn === null || !Number.isFinite(value)) return null;

  const dir = rule.direction;
  if(dir === "up"){
    if(value >= rule.good) return "good";
    if(value >= rule.warn) return "warn";
    return "bad";
  }else{
    if(value <= rule.good) return "good";
    if(value <= rule.warn) return "warn";
    return "bad";
  }
}

// =================== Tooltip functions ===================
function getKpiTooltipContent(kpi){
  const t = KPI_TOOLTIPS[kpi];
  const fallback = {
    definicion: "KPI de tu archivo. Si quieres tooltip específico, agrégalo en KPI_TOOLTIPS.",
    formula: "SUM(valor) filtrado por año/mes/sucursal/vendedor y el KPI seleccionado.",
    interpretacion: "Úsalo para lectura ejecutiva; compara vs base y revisa tendencia."
  };

  const data = t || fallback;

  return `
    <div class="ttTitle">${escapeHtml(kpi || "KPI")}</div>

    <div class="ttRow">
      <div class="ttLabel">Definición</div>
      <div class="ttText">${escapeHtml(data.definicion)}</div>
    </div>

    <div class="ttRow">
      <div class="ttLabel">Fórmula</div>
      <div class="ttText">${escapeHtml(data.formula)}</div>
    </div>

    <div class="ttRow">
      <div class="ttLabel">Interpretación</div>
      <div class="ttText">${escapeHtml(data.interpretacion)}</div>
    </div>

    <div class="ttFoot">Fuente: hoja <b>DATA_KPIS</b> | Se aplica filtro global y/o acumulado.</div>
  `;
}

function positionTooltipNearElement(ttEl, anchorEl){
  const r = anchorEl.getBoundingClientRect();
  const margin = 12;

  // derecha si cabe, si no izquierda
  const preferredLeft = r.right + margin;
  const maxLeft = window.innerWidth - ttEl.offsetWidth - margin;
  let left = Math.min(preferredLeft, maxLeft);
  if(left < margin) left = margin;

  // centrado vertical relativo a la tarjeta
  let top = r.top + (r.height/2) - (ttEl.offsetHeight/2);

  // clamp
  const maxTop = window.innerHeight - ttEl.offsetHeight - margin;
  top = Math.max(margin, Math.min(top, maxTop));

  ttEl.style.left = `${Math.round(left)}px`;
  ttEl.style.top = `${Math.round(top)}px`;
}

function showKpiTooltip(anchorEl){
  if(!els.kpiTooltip) return;
  const kpi = els.kpiSelect?.value || "";
  els.kpiTooltip.innerHTML = getKpiTooltipContent(kpi);
  els.kpiTooltip.classList.remove("hidden");

  requestAnimationFrame(()=>{
    positionTooltipNearElement(els.kpiTooltip, anchorEl);
  });
}

function hideKpiTooltip(){
  if(!els.kpiTooltip) return;
  els.kpiTooltip.classList.add("hidden");
}

// =================== Excel read ===================
function readFile(file){
  return new Promise((resolve, reject)=>{
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("No se pudo leer el archivo."));
    reader.onload = (e) => resolve(e.target.result);
    reader.readAsArrayBuffer(file);
  });
}

function sheetToJson(wb, sheetName){
  const ws = wb.Sheets[sheetName];
  if(!ws) return null;
  return XLSX.utils.sheet_to_json(ws, { defval:"" });
}

function normalizeKPIS(rows){
  const req = ["anio","mes_num","sucursal","vendedor","kpi","valor"];
  const out = [];
  for(const r of rows){
    const o = {};
    for(const [k,v] of Object.entries(r)) o[normalizeKey(k)] = v;
    if(req.some(c => !(c in o))) continue;

    const anio = toNumber(o.anio);
    const mes = toNumber(o.mes_num);
    const valor = toNumber(o.valor);
    if(!Number.isFinite(anio) || !Number.isFinite(mes) || !Number.isFinite(valor)) continue;

    out.push({
      anio: Math.trunc(anio),
      mes_num: clampMonth(mes),
      sucursal: String(o.sucursal ?? "").trim() || "SIN_SUCURSAL",
      vendedor: String(o.vendedor ?? "").trim() || "SIN_VENDEDOR",
      kpi: String(o.kpi ?? "").trim() || "SIN_KPI",
      valor
    });
  }
  return out;
}

function normalizeDIM(rows){
  const req = ["anio","mes_num","sucursal","dimension","clave","nombre","ventas_con_iva","venta_sin_iva","costo","utilidad","margen"];
  const out = [];
  for(const r of rows){
    const o = {};
    for(const [k,v] of Object.entries(r)) o[normalizeKey(k)] = v;
    if(req.some(c => !(c in o))) continue;

    const anio = toNumber(o.anio);
    const mes = toNumber(o.mes_num);
    if(!Number.isFinite(anio) || !Number.isFinite(mes)) continue;

    out.push({
      anio: Math.trunc(anio),
      mes_num: clampMonth(mes),
      sucursal: String(o.sucursal ?? "").trim() || "SIN_SUCURSAL",
      dimension: String(o.dimension ?? "").trim().toUpperCase() || "SIN_DIM",
      clave: String(o.clave ?? "").trim(),
      nombre: String(o.nombre ?? "").trim(),
      ventas_con_iva: toNumber(o.ventas_con_iva),
      venta_sin_iva: toNumber(o.venta_sin_iva),
      costo: toNumber(o.costo),
      utilidad: toNumber(o.utilidad),
      margen: toNumber(o.margen),
    });
  }
  return out;
}

async function loadExcel(file){
  const buf = await readFile(file);
  const wb = XLSX.read(buf, { type:"array" });

  const names = wb.SheetNames.map(n=>String(n));
  const find = (t)=> names.find(n => n.toLowerCase() === t.toLowerCase()) || null;

  const shK = find("DATA_KPIS");
  const shD = find("DATA_DIM");

  if(!shK && !shD) throw new Error("No encontré hojas DATA_KPIS ni DATA_DIM.");

  KPIS = shK ? normalizeKPIS(sheetToJson(wb, shK) || []) : [];
  DIM  = shD ? normalizeDIM(sheetToJson(wb, shD) || []) : [];

  els.dataStatus.textContent = `Archivo: ${file.name} | KPIs: ${KPIS.length} filas | DIM: ${DIM.length} filas`;
  setupSelectors();
  renderAll();
}

// =================== filtering ===================
function getGlobalFilters(){
  const y = els.yearSelect.value;
  const m = els.monthSelect.value;
  const b = els.branchSelect.value;
  const s = els.sellerSelect.value;
  return {
    anio: (y==="__ALL__") ? null : Number(y),
    mes_num: (m==="__ALL__") ? null : Number(m),
    sucursal: (b==="__ALL__") ? null : b,
    vendedor: (s==="__ALL__") ? null : s
  };
}

function filterKPIS(f, overrides={}){
  const ff = { ...f, ...overrides };
  const kpi = els.kpiSelect.value;

  return KPIS.filter(r=>{
    if(r.kpi !== kpi) return false;

    if(ff.anio !== null && r.anio !== ff.anio) return false;
    if(ff.mes_num !== null && r.mes_num !== ff.mes_num) return false;

    // anti doble conteo
    if(String(r.sucursal).toUpperCase() === "TODAS") return false;

    if(ff.sucursal !== null){
      const bs = String(ff.sucursal).toUpperCase();
      if(bs !== "TODAS" && r.sucursal !== ff.sucursal) return false;
    }
    if(ff.vendedor !== null && r.vendedor !== ff.vendedor) return false;

    return true;
  });
}

function filterDIM(f, overrides={}){
  const ff = { ...f, ...overrides };
  return DIM.filter(r=>{
    if(ff.anio !== null && r.anio !== ff.anio) return false;
    if(ff.mes_num !== null && r.mes_num !== ff.mes_num) return false;

    if(String(r.sucursal).toUpperCase() === "TODAS") return false;

    if(ff.sucursal !== null){
      const bs = String(ff.sucursal).toUpperCase();
      if(bs !== "TODAS" && r.sucursal !== ff.sucursal) return false;
    }
    return true;
  });
}

// =================== selectors ===================
function setupSelectors(){
  const src = KPIS.length ? KPIS : DIM;
  if(!src.length){
    setSelectOptions(els.yearSelect, [], true, "Todos");
    setSelectOptions(els.monthSelect, [], true, "Todos");
    setSelectOptions(els.branchSelect, ["TODAS"], false);
    setSelectOptions(els.sellerSelect, ["TODOS"], true, "Todos");
    setSelectOptions(els.kpiSelect, [], false);
    return;
  }

  const years = uniqueSorted(src.map(r=>r.anio));
  const monthsPresent = uniqueSorted(src.map(r=>r.mes_num));
  const branches = uniqueSorted(src.map(r=>r.sucursal));
  const sellers = uniqueSorted(KPIS.map(r=>r.vendedor));
  const kpis = uniqueSorted(KPIS.map(r=>r.kpi));

  setSelectOptions(els.yearSelect, years, true, "Todos");
  setSelectOptions(els.monthSelect, monthsPresent.length ? monthsPresent : monthsList(), true, "Todos");

  setSelectOptions(els.branchSelect, ["TODAS", ...branches], false);
  setSelectOptions(els.sellerSelect, sellers.length ? sellers : ["TODOS"], true, "Todos");

  setSelectOptions(els.kpiSelect, kpis, false);

  // defaults
  const lastYear = years[years.length-1];
  const prevYear = years.length >= 2 ? years[years.length-2] : lastYear;
  const lastMonth = monthsPresent.length ? monthsPresent[monthsPresent.length-1] : 12;

  els.yearSelect.value = String(lastYear);
  els.monthSelect.value = String(lastMonth);
  els.branchSelect.value = "TODAS";
  els.sellerSelect.value = "__ALL__";

  setSelectOptions(els.yearASelect, years, false);
  setSelectOptions(els.yearBSelect, years, false);
  els.yearASelect.value = String(lastYear);
  els.yearBSelect.value = String(prevYear);

  setSelectOptions(els.accStartMonth, monthsList(), false);
  setSelectOptions(els.accEndMonth, monthsList(), false);
  els.accStartMonth.value = "1";
  els.accEndMonth.value = String(lastMonth);

  els.viewMode.value = "period";
  els.accChart.value = "bars";
  toggleAccumUI();
}

// =================== pages/nav ===================
function setActiveNav(which){
  [els.navExec, els.navAccum, els.navDim].forEach(b=>b.classList.remove("active"));

  if(which === "exec"){
    els.navExec.classList.add("active");
    els.pageTitle.textContent = "Executive";
    els.dimPage.classList.add("hidden");
  }else if(which === "accum"){
    els.navAccum.classList.add("active");
    els.pageTitle.textContent = "Acumulados";
    els.dimPage.classList.add("hidden");
  }else{
    els.navDim.classList.add("active");
    els.pageTitle.textContent = "Dimensiones";
    els.dimPage.classList.remove("hidden");
  }
  renderAll();
}

function toggleAccumUI(){
  const isAccum = els.viewMode.value === "accum";
  els.accControls.style.display = isAccum ? "grid" : "none";
  els.modePill.textContent = isAccum ? "Acumulado" : "Periodo";
}

// =================== computations ===================
function sumRangeForYear({anio, lo, hi, f}){
  const rows = filterKPIS(f, { anio, mes_num:null }).filter(r => r.mes_num >= lo && r.mes_num <= hi);
  return { rows, sum: aggSum(rows.map(r=>r.valor)), n: rows.length };
}

function monthlySumsForYear({anio, f}){
  const rows = filterKPIS(f, { anio, mes_num:null });
  const byM = groupBy(rows, "mes_num");
  return monthsList().map(m => aggSum((byM.get(String(m))||[]).map(r=>r.valor)));
}

function cumulativeFromMonthly(monthly, lo, hi){
  let acc = 0;
  return monthsList().map(m=>{
    const v = monthly[m-1];
    if(m < lo || m > hi) return null;
    acc += (Number.isFinite(v) ? v : 0);
    return acc;
  });
}

function resolvePeriodContext(){
  const f = getGlobalFilters();
  const kpi = els.kpiSelect.value;
  const rule = getKpiRule(kpi);
  const unit = rule.unit || inferUnitFromKpiName(kpi);

  const hasYear = f.anio !== null;
  const hasMonth = f.mes_num !== null;

  let cur = NaN, base = NaN, meta = "";

  if(hasYear && hasMonth){
    const rowsCur = filterKPIS(f);
    cur = aggSum(rowsCur.map(r=>r.valor));
    const rowsBase = filterKPIS(f, { anio: f.anio - 1, mes_num: f.mes_num });
    base = aggSum(rowsBase.map(r=>r.valor));
    meta = `${f.anio} ${MONTH_NAMES[f.mes_num-1]} vs ${f.anio-1} ${MONTH_NAMES[f.mes_num-1]}`;
  }else if(hasYear && !hasMonth){
    const rowsCur = filterKPIS(f, { anio: f.anio, mes_num:null });
    cur = aggSum(rowsCur.map(r=>r.valor));
    const rowsBase = filterKPIS(f, { anio: f.anio - 1, mes_num:null });
    base = aggSum(rowsBase.map(r=>r.valor));
    meta = `Año ${f.anio} vs Año ${f.anio-1} (suma anual)`;
  }else{
    const yA = Number(els.yearASelect.value);
    const yB = Number(els.yearBSelect.value);
    const rowsCur = filterKPIS({ ...f, anio:yA }, { mes_num:null });
    cur = aggSum(rowsCur.map(r=>r.valor));
    const rowsBase = filterKPIS({ ...f, anio:yB }, { mes_num:null });
    base = aggSum(rowsBase.map(r=>r.valor));
    meta = `Año ${yA} vs ${yB} (suma anual)`;
  }

  const d = computeDelta(cur, base);
  return { f, kpi, unit, cur, base, d, meta };
}

function resolveAccumContext(){
  const f = getGlobalFilters();
  const kpi = els.kpiSelect.value;
  const rule = getKpiRule(kpi);
  const unit = rule.unit || inferUnitFromKpiName(kpi);

  const yA = Number(els.yearASelect.value);
  const yB = Number(els.yearBSelect.value);
  const lo = clampMonth(els.accStartMonth.value);
  const hi = clampMonth(els.accEndMonth.value);
  const a = sumRangeForYear({ anio:yA, lo:Math.min(lo,hi), hi:Math.max(lo,hi), f });
  const b = sumRangeForYear({ anio:yB, lo:Math.min(lo,hi), hi:Math.max(lo,hi), f });
  const d = computeDelta(a.sum, b.sum);

  const meta = `${yA} vs ${yB} | ${MONTH_NAMES[Math.min(lo,hi)-1]}–${MONTH_NAMES[Math.max(lo,hi)-1]}`;
  return { f, kpi, unit, yA, yB, lo:Math.min(lo,hi), hi:Math.max(lo,hi), a, b, d, meta };
}

// =================== render KPI strip ===================
function renderKpiStrip(mode){
  const kpi = els.kpiSelect.value;

  if(mode === "accum"){
    const c = resolveAccumContext();

    els.kpiLabelA.textContent = "Acum Año A";
    els.kpiLabelB.textContent = "Acum Año B";
    els.kpiValueA.textContent = formatByUnit(c.a.sum, c.unit);
    els.kpiValueB.textContent = formatByUnit(c.b.sum, c.unit);
    els.kpiDiffAbs.textContent = formatByUnit(c.d.abs, c.unit);
    els.kpiDiffPct.textContent = fmtPct(c.d.pct);

    const sem = classifySemaforo(c.a.sum, kpi);
    setBadge(els.badgeA, sem ? (sem==="good"?"VERDE":sem==="warn"?"AMARILLO":"ROJO") : "—", sem);

    setBadge(els.badgeB, "Base", null);
    setBadge(els.badgeDiff, Number.isFinite(c.d.abs) ? (c.d.abs>=0?"▲":"▼") : "—", null);
    setBadge(els.badgePct, Number.isFinite(c.d.pct) ? (c.d.pct>=0?"▲":"▼") : "—", null);

    return;
  }

  const c = resolvePeriodContext();
  els.kpiLabelA.textContent = "Valor";
  els.kpiLabelB.textContent = "Comparativo";
  els.kpiValueA.textContent = formatByUnit(c.cur, c.unit);
  els.kpiValueB.textContent = formatByUnit(c.base, c.unit);
  els.kpiDiffAbs.textContent = formatByUnit(c.d.abs, c.unit);
  els.kpiDiffPct.textContent = fmtPct(c.d.pct);

  const sem = classifySemaforo(c.cur, kpi);
  setBadge(els.badgeA, sem ? (sem==="good"?"VERDE":sem==="warn"?"AMARILLO":"ROJO") : "—", sem);
  setBadge(els.badgeB, "Base", null);
  setBadge(els.badgeDiff, Number.isFinite(c.d.abs) ? (c.d.abs>=0?"▲":"▼") : "—", null);
  setBadge(els.badgePct, Number.isFinite(c.d.pct) ? (c.d.pct>=0?"▲":"▼") : "—", null);
}

// =================== render main chart + table ===================
function renderMain(){
  const view = els.viewMode.value;
  toggleAccumUI();

  if(!KPIS.length){
    els.chartTitle.textContent = "Tendencia";
    els.chartMeta.textContent = "Sin DATA_KPIS";
    els.tableTitle.textContent = "Detalle";
    els.tableMeta.textContent = "—";
    els.warnings.textContent = "No hay datos KPI.";
    els.mainTbody.innerHTML = "";
    if(chartMain){ chartMain.destroy(); chartMain = null; }
    renderKpiStrip(view);
    return;
  }

  if(view === "accum"){
    const c = resolveAccumContext();
    renderKpiStrip("accum");

    els.chartTitle.textContent = `Acumulado ${c.kpi}`;
    els.chartMeta.textContent = c.meta;

    const chartMode = els.accChart.value;
    if(chartMode === "cumline"){
      const labels = monthsList().map(m=>MONTH_NAMES[m-1]);
      const aMonthly = monthlySumsForYear({ anio:c.yA, f:c.f });
      const bMonthly = monthlySumsForYear({ anio:c.yB, f:c.f });
      const aCum = cumulativeFromMonthly(aMonthly, c.lo, c.hi);
      const bCum = cumulativeFromMonthly(bMonthly, c.lo, c.hi);

      drawChart("line", labels, [
        { label: `${c.yA} acumulado`, data: aCum },
        { label: `${c.yB} acumulado`, data: bCum },
      ], c.unit);
    }else{
      drawChart("bar",
        [`${c.yA} (${MONTH_NAMES[c.lo-1]}–${MONTH_NAMES[c.hi-1]})`, `${c.yB} (${MONTH_NAMES[c.lo-1]}–${MONTH_NAMES[c.hi-1]})`],
        [{ label:`Acumulado (${c.kpi})`, data:[c.a.sum, c.b.sum] }],
        c.unit
      );
    }

    const branchUpper = String(c.f.sucursal ?? "TODAS").toUpperCase();
    const groupKey = (branchUpper === "TODAS") ? "sucursal" : "vendedor";
    els.thGroup.textContent = (groupKey === "sucursal") ? "Sucursal" : "Vendedor";
    els.tableTitle.textContent = "Detalle (Año A)";
    els.tableMeta.textContent = `Agrupado por ${els.thGroup.textContent} | Filas: ${fmtInt(c.a.n)}`;

    renderStatsTable(c.a.rows, groupKey, c.unit);
    els.warnings.textContent = (c.a.n===0 && c.b.n===0) ? "Sin filas para el rango." : "";
    return;
  }

  // Period mode
  const c = resolvePeriodContext();
  renderKpiStrip("period");

  els.chartTitle.textContent = `Tendencia ${c.kpi}`;
  els.chartMeta.textContent = c.meta;

  const years = uniqueSorted(KPIS.map(r=>r.anio));
  const year = (c.f.anio !== null) ? c.f.anio : Number(years[years.length-1] ?? NaN);
  const labels = monthsList().map(m=>MONTH_NAMES[m-1]);

  const curMonthly = monthlySumsForYear({ anio: year, f: c.f });
  const baseMonthly = monthlySumsForYear({ anio: year-1, f: c.f });

  drawChart("line", labels, [
    { label:`${year}`, data: curMonthly },
    { label:`${year-1}`, data: baseMonthly },
  ], c.unit);

  let rowsScope = [];
  if(c.f.anio !== null && c.f.mes_num !== null){
    rowsScope = filterKPIS(c.f);
  }else if(c.f.anio !== null){
    rowsScope = filterKPIS(c.f, { anio:c.f.anio, mes_num:null });
  }else{
    rowsScope = filterKPIS(c.f, { anio:year, mes_num:null });
  }

  const branchUpper = String(c.f.sucursal ?? "TODAS").toUpperCase();
  const groupKey = (branchUpper === "TODAS") ? "sucursal" : "vendedor";
  els.thGroup.textContent = (groupKey === "sucursal") ? "Sucursal" : "Vendedor";
  els.tableTitle.textContent = "Detalle";
  els.tableMeta.textContent = `Agrupado por ${els.thGroup.textContent} | Filas: ${fmtInt(rowsScope.length)}`;

  renderStatsTable(rowsScope, groupKey, c.unit);
  els.warnings.textContent = rowsScope.length ? "" : "Sin filas para el filtro actual.";
}

function drawChart(type, labels, datasets, unit){
  const options = {
    responsive:true,
    maintainAspectRatio:false,
    spanGaps:true,
    plugins:{
      legend:{ labels:{ color:"#cbd5e1" } },
      tooltip:{
        callbacks:{
          label:(ctx)=>{
            const v = ctx.parsed.y;
            return `${ctx.dataset.label}: ${formatByUnit(v, unit)}`;
          }
        }
      }
    },
    scales:{
      x:{ ticks:{ color:"#94a3b8" }, grid:{ color:"rgba(31,42,68,.15)" } },
      y:{ ticks:{ color:"#94a3b8" }, grid:{ color:"rgba(31,42,68,.55)" } }
    }
  };

  if(chartMain){ chartMain.destroy(); chartMain = null; }
  chartMain = new Chart(els.chartMain, {
    type,
    data:{ labels, datasets },
    options
  });
}

function renderStatsTable(rows, groupKey, unit){
  const grouped = groupBy(rows, groupKey);
  const items = [];
  for(const [key, arr] of grouped.entries()){
    const vals = arr.map(r=>r.valor).filter(Number.isFinite);
    const st = aggStats(vals);
    items.push({ group:key, n:arr.length, ...st });
  }
  items.sort((a,b)=> (b.sum - a.sum));

  els.mainTbody.innerHTML = "";
  for(const it of items){
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(it.group)}</td>
      <td>${fmtInt(it.n)}</td>
      <td>${formatByUnit(it.sum, unit)}</td>
      <td>${formatByUnit(it.avg, unit)}</td>
      <td>${formatByUnit(it.max, unit)}</td>
      <td>${formatByUnit(it.min, unit)}</td>
    `;
    els.mainTbody.appendChild(tr);
  }
}

// =================== Dimensiones ===================
function renderDim(){
  if(els.dimPage.classList.contains("hidden")) return;

  if(!DIM.length){
    els.dimMeta.textContent = "Sin DATA_DIM";
    els.dimWarnings.textContent = "No hay datos Dimensiones.";
    els.dimTbody.innerHTML = "";
    if(chartDim){ chartDim.destroy(); chartDim = null; }
    return;
  }

  const f = getGlobalFilters();
  const dim = els.dimSelect.value;
  const metric = els.dimMetric.value;
  const topN = Number(els.topN.value);

  const years = uniqueSorted(DIM.map(r=>r.anio));
  const months = uniqueSorted(DIM.map(r=>r.mes_num));
  const anio = (f.anio !== null) ? f.anio : Number(years[years.length-1] ?? NaN);
  const mes = (f.mes_num !== null) ? f.mes_num : Number(months[months.length-1] ?? NaN);

  const rows = filterDIM(f, { anio, mes_num: mes }).filter(r => r.dimension === String(dim).toUpperCase());

  const items = rows
    .map(r=>({
      clave:r.clave,
      nombre:r.nombre,
      ventas:r.ventas_con_iva,
      utilidad:r.utilidad,
      margen:r.margen,
      metricValue:r[metric]
    }))
    .filter(it => Number.isFinite(it.metricValue))
    .sort((a,b)=> (b.metricValue - a.metricValue))
    .slice(0, topN);

  els.dimMeta.textContent = `${dim} | ${metric} | ${anio}-${mes} | Sucursal: ${f.sucursal ?? "TODAS"}`;
  els.dimWarnings.textContent = items.length ? "" : "Sin filas para este filtro/dimensión.";

  els.dimTbody.innerHTML = "";
  for(const it of items){
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(it.clave)}</td>
      <td>${escapeHtml(it.nombre)}</td>
      <td>${fmtMoney(it.ventas)}</td>
      <td>${fmtMoney(it.utilidad)}</td>
      <td>${fmtPct(it.margen >= 0 && it.margen <= 1.5 ? it.margen : it.margen/100)}</td>
    `;
    els.dimTbody.appendChild(tr);
  }
  els.dimTableMeta.textContent = `Top ${items.length} | Filas fuente: ${fmtInt(rows.length)}`;

  const labels = items.map(it => it.nombre || it.clave).reverse();
  const data = items.map(it => it.metricValue).reverse();
  const options = {
    responsive:true,
    maintainAspectRatio:false,
    indexAxis:"y",
    plugins:{ legend:{ labels:{ color:"#cbd5e1" } } },
    scales:{
      x:{ ticks:{ color:"#94a3b8" }, grid:{ color:"rgba(31,42,68,.55)" } },
      y:{ ticks:{ color:"#94a3b8" }, grid:{ color:"rgba(31,42,68,.15)" } }
    }
  };

  if(chartDim){ chartDim.destroy(); chartDim=null; }
  chartDim = new Chart(els.chartDim, {
    type:"bar",
    data:{ labels, datasets:[{ label:`${metric} (Top ${items.length})`, data }] },
    options
  });
}

// =================== export CSV ===================
function exportTableCSV(tableEl, filename){
  const headers = [...tableEl.querySelectorAll("thead th")].map(th=>th.textContent.trim());
  const bodyRows = [...tableEl.querySelectorAll("tbody tr")].map(tr =>
    [...tr.querySelectorAll("td")].map(td=>td.textContent.trim())
  );
  if(bodyRows.length === 0){ alert("No hay datos para exportar."); return; }

  const rows = [headers, ...bodyRows];
  const csv = rows.map(r => r.map(cell=>{
    const c = String(cell).replaceAll('"','""');
    return `"${c}"`;
  }).join(",")).join("\n");

  const blob = new Blob([csv], { type:"text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// =================== presets acumulado ===================
function setAccumRange(lo, hi){
  els.viewMode.value = "accum";
  toggleAccumUI();
  els.accStartMonth.value = String(clampMonth(lo));
  els.accEndMonth.value = String(clampMonth(hi));
  renderAll();
}

// =================== render all ===================
function renderAll(){
  toggleAccumUI();
  renderMain();
  renderDim();
}

// =================== demo ===================
function makeDemo(){
  const suc = ["GENERAL","EXPRESS","ADELITAS","SAN AGUST","H ILUSTRES"];
  const kpis = [
    "Ventas Totales (CON IVA)","Ventas Contado (CON IVA)","Ventas Crédito (CON IVA)",
    "Transacciones","Ticket Promedio (CON IVA)","Utilidad Total (SIN IVA)","Margen de Utilidad"
  ];
  const years = [2024, 2025];
  const kRows = [];

  for(const y of years){
    for(let m=1; m<=12; m++){
      for(const s of suc){
        for(const k of kpis){
          let base = 100;
          if(k.includes("Ventas")) base = 500000;
          if(k.includes("Utilidad")) base = 90000;
          if(k.includes("Transacciones")) base = 1200;
          if(k.includes("Ticket")) base = 500;
          if(k.includes("Margen")) base = 0.35;

          const season = (m>=10 && m<=12) ? 1.12 : (m<=2 ? 0.92 : 1.0);
          const mult = (Math.random()*0.30 + 0.85) * season * (y===2025 ? 1.10 : 1.0) * (s==="GENERAL" ? 1.25 : 1.0);

          const val = k.includes("Margen") ? Math.min(0.65, Math.max(0.15, base*mult)) : Math.round(base*mult*100)/100;

          kRows.push({ anio:y, mes_num:m, sucursal:s, vendedor:"TODOS", kpi:k, valor:val });
        }
      }
    }
  }

  const dimRows = [];
  const marcas = Array.from({length:50}, (_,i)=>({ clave:`M${i+1}`, nombre:`Marca ${i+1}` }));
  const cats = Array.from({length:40}, (_,i)=>({ clave:`C${i+1}`, nombre:`Categoria ${i+1}` }));

  for(const y of years){
    for(let m=1; m<=12; m++){
      for(const s of suc){
        const list = (Math.random() > 0.5) ? marcas : cats;
        const dim = (list === marcas) ? "MARCA" : "CATEGORIA";
        for(const it of list){
          const ventas = Math.round((Math.random()*120000 + 5000) * (y===2025?1.08:1.0));
          const costo = Math.round(ventas * (Math.random()*0.70 + 0.20));
          const utilidad = ventas - costo;
          const margen = ventas === 0 ? 0 : (utilidad / ventas);

          dimRows.push({
            anio:y, mes_num:m, sucursal:s, dimension:dim,
            clave:it.clave, nombre:it.nombre,
            ventas_con_iva:ventas, venta_sin_iva: Math.round(ventas/1.16),
            costo, utilidad, margen
          });
        }
      }
    }
  }

  KPIS = kRows;
  DIM = dimRows;
  els.dataStatus.textContent = `DEMO | KPIs: ${KPIS.length} | DIM: ${DIM.length}`;
  setupSelectors();
  renderAll();
}

// =================== events ===================
function bindEvents(){
  // nav
  els.navExec.addEventListener("click", ()=> setActiveNav("exec"));
  els.navAccum.addEventListener("click", ()=>{
    els.viewMode.value = "accum";
    toggleAccumUI();
    setActiveNav("accum");
  });
  els.navDim.addEventListener("click", ()=> setActiveNav("dim"));

  // file
  els.fileInput.addEventListener("change", async (e)=>{
    const file = e.target.files?.[0];
    if(!file) return;
    try{ await loadExcel(file); }
    catch(err){ alert(err?.message || "Error al cargar archivo."); console.error(err); }
    els.fileInput.value = "";
  });

  // demo / reset
  els.btnDemo.addEventListener("click", makeDemo);
  els.btnReset.addEventListener("click", ()=>{
    KPIS = []; DIM = [];
    els.dataStatus.textContent = "Sin datos cargados";
    if(chartMain){ chartMain.destroy(); chartMain=null; }
    if(chartDim){ chartDim.destroy(); chartDim=null; }
    els.mainTbody.innerHTML = "";
    els.dimTbody.innerHTML = "";
    setActiveNav("exec");
    setupSelectors();
    renderAll();
  });

  // slicers render
  [els.yearSelect, els.monthSelect, els.branchSelect, els.sellerSelect, els.kpiSelect, els.viewMode,
   els.yearASelect, els.yearBSelect, els.accStartMonth, els.accEndMonth, els.accChart].forEach(el=>{
    el.addEventListener("change", renderAll);
  });

  els.viewMode.addEventListener("change", ()=>{
    toggleAccumUI();
    renderAll();
  });

  // presets
  els.chipYTD.addEventListener("click", ()=>{
    const m = (els.monthSelect.value==="__ALL__") ? 12 : Number(els.monthSelect.value);
    setAccumRange(1, clampMonth(m));
  });
  els.chipEneJun.addEventListener("click", ()=> setAccumRange(1,6));
  els.chipJulNov.addEventListener("click", ()=> setAccumRange(7,11));
  els.chipFull.addEventListener("click", ()=> setAccumRange(1,12));

  // export
  els.btnExportTable.addEventListener("click", ()=> exportTableCSV(els.mainTable, "detalle_kpi.csv"));
  els.btnExportDim.addEventListener("click", ()=> exportTableCSV(els.dimTable, "ranking_dim.csv"));

  // dim slicers
  [els.dimSelect, els.dimMetric, els.topN].forEach(el=> el.addEventListener("change", renderDim));

  // ===== Tooltips PowerBI en tarjetas KPI (hover) =====
  const kpiCards = document.querySelectorAll(".kpiStrip .kpiCard");
  kpiCards.forEach(card=>{
    card.addEventListener("mouseenter", ()=> showKpiTooltip(card));
    card.addEventListener("mouseleave", hideKpiTooltip);
  });

  // Reposicionar/ocultar tooltip si se hace scroll o resize
  window.addEventListener("scroll", hideKpiTooltip, { passive:true });
  window.addEventListener("resize", hideKpiTooltip);
}

// init
bindEvents();
setActiveNav("exec");
setupSelectors();
toggleAccumUI();
renderAll();
