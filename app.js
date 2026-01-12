/* Estructura final:
  - Barra fija global:
    Año A, Año B, Modo (Periodo/Acumulado),
    Mes (Periodo) o DE/A (Acumulado),
    Sucursal, Vendedor

  - Tabs: Sucursal, Empleado, Marcas, Categorías

  Excel:
  - DATA_KPIS: anio, mes_num, sucursal, vendedor, kpi, valor
  - DATA_DIM : anio, mes_num, sucursal, dimension(MARCA|CATEGORIA), nombre, ventas_con_iva, utilidad, margen, etc.
*/

let KPIS = [];
let DIM = [];

let chartSucursal = null;
let chartMarcas = null;
let chartCategorias = null;

const $ = (id) => document.getElementById(id);

const els = {
  // file
  fileInput: $("fileInput"),
  btnReset: $("btnReset"),
  dataStatus: $("dataStatus"),

  // global filters
  yearA: $("yearA"),
  yearB: $("yearB"),
  mode: $("mode"),
  month: $("month"),
  accStart: $("accStart"),
  accEnd: $("accEnd"),
  branch: $("branch"),
  seller: $("seller"),

  periodMonthWrap: $("periodMonthWrap"),
  accStartWrap: $("accStartWrap"),
  accEndWrap: $("accEndWrap"),

  // tabs
  tabSucursal: $("tabSucursal"),
  tabEmpleado: $("tabEmpleado"),
  tabMarcas: $("tabMarcas"),
  tabCategorias: $("tabCategorias"),

  // pages
  pageSucursal: $("pageSucursal"),
  pageEmpleado: $("pageEmpleado"),
  pageMarcas: $("pageMarcas"),
  pageCategorias: $("pageCategorias"),

  // Sucursal page
  kpiSucursal: $("kpiSucursal"),
  metaSucursal: $("metaSucursal"),
  sucA: $("sucA"),
  sucB: $("sucB"),
  sucDiffAbs: $("sucDiffAbs"),
  sucDiffPct: $("sucDiffPct"),
  chartSucursal: $("chartSucursal"),
  chartTitleSucursal: $("chartTitleSucursal"),
  chartMetaSucursal: $("chartMetaSucursal"),
  tableSucursal: $("tableSucursal"),
  tableSucursalBody: $("tableSucursal").querySelector("tbody"),
  warnSucursal: $("warnSucursal"),

  // Empleado page
  kpiEmpleado: $("kpiEmpleado"),
  empSearch: $("empSearch"),
  metaEmpleado: $("metaEmpleado"),
  empTableMeta: $("empTableMeta"),
  tableEmpleado: $("tableEmpleado"),
  tableEmpleadoBody: $("tableEmpleado").querySelector("tbody"),
  warnEmpleado: $("warnEmpleado"),

  // Marcas page
  metricMarca: $("metricMarca"),
  topMarca: $("topMarca"),
  metaMarcas: $("metaMarcas"),
  chartMetaMarcas: $("chartMetaMarcas"),
  chartMarcas: $("chartMarcas"),
  tableMarcas: $("tableMarcas"),
  tableMarcasBody: $("tableMarcas").querySelector("tbody"),
  warnMarcas: $("warnMarcas"),

  // Categorías page
  metricCat: $("metricCat"),
  topCat: $("topCat"),
  metaCategorias: $("metaCategorias"),
  chartMetaCategorias: $("chartMetaCategorias"),
  chartCategorias: $("chartCategorias"),
  tableCategorias: $("tableCategorias"),
  tableCategoriasBody: $("tableCategorias").querySelector("tbody"),
  warnCategorias: $("warnCategorias"),
};

const MONTH_NAMES = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];

// ---------- helpers ----------
function normalizeKey(k){
  return String(k||"").trim().toLowerCase().replace(/\s+/g,"_").replace(/[^\wáéíóúüñ]/gi,"");
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
function fmtPct(r){
  if(!Number.isFinite(r)) return "—";
  // margen viene 0.35
  return new Intl.NumberFormat("es-MX", { style:"percent", maximumFractionDigits:2 }).format(r);
}
function fmtNum(n, d=2){
  if(!Number.isFinite(n)) return "—";
  return new Intl.NumberFormat("es-MX", { maximumFractionDigits:d }).format(n);
}

function clampMonth(m){
  const n = Number(m);
  if(!Number.isFinite(n)) return 1;
  return Math.min(12, Math.max(1, Math.trunc(n)));
}

function monthsList(){ return Array.from({length:12}, (_,i)=> i+1); }

function uniqueSorted(arr){
  return [...new Set(arr)]
    .filter(v => v !== null && v !== undefined && String(v).trim() !== "")
    .sort((a,b)=>{
      const na = Number(a), nb = Number(b);
      if(Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
      return String(a).localeCompare(String(b), "es");
    });
}

function setSelectOptions(el, values, includeAll=false, allLabel="TODOS"){
  const prev = el.value;
  el.innerHTML = "";

  if(includeAll){
    const opt = document.createElement("option");
    opt.value = "__ALL__";
    opt.textContent = allLabel;
    el.appendChild(opt);
  }

  for(const v of values){
    const opt = document.createElement("option");
    opt.value = String(v);
    opt.textContent = String(v);
    el.appendChild(opt);
  }

  if(prev && [...el.options].some(o=>o.value===prev)) el.value = prev;
  else el.value = (includeAll ? "__ALL__" : (values[0] ?? ""));
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

function sum(arr){
  let s=0;
  for(const v of arr) s += v;
  return s;
}

function computeDelta(a, b){
  if(!Number.isFinite(a) || !Number.isFinite(b)) return { abs:NaN, pct:NaN };
  const abs = a - b;
  const pct = (b===0) ? NaN : abs / b;
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
  if(unit === "percent") return fmtPct(v); // margen decimal 0.35
  return fmtNum(v);
}

// ---------- read Excel ----------
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
  // mínimo para marcas/categorías
  const req = ["anio","mes_num","sucursal","dimension","nombre","ventas_con_iva","utilidad","margen"];
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
      dimension: String(o.dimension ?? "").trim().toUpperCase(),
      nombre: String(o.nombre ?? "").trim() || "SIN_NOMBRE",
      ventas_con_iva: toNumber(o.ventas_con_iva),
      utilidad: toNumber(o.utilidad),
      margen: toNumber(o.margen), // decimal
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

  els.dataStatus.textContent = `Archivo: ${file.name} | KPIs: ${KPIS.length} | DIM: ${DIM.length}`;
  setupSelectors();
  renderAll();
}

// ---------- selectors ----------
function setupSelectors(){
  const src = KPIS.length ? KPIS : DIM;
  if(!src.length){
    // limpiar selects
    setSelectOptions(els.yearA, [], false);
    setSelectOptions(els.yearB, [], false);
    setSelectOptions(els.month, monthsList().map(m=>MONTH_NAMES[m-1]), false);
    setSelectOptions(els.accStart, monthsList(), false);
    setSelectOptions(els.accEnd, monthsList(), false);
    setSelectOptions(els.branch, ["TODAS"], false);
    setSelectOptions(els.seller, ["TODOS"], false);
    setSelectOptions(els.kpiSucursal, [], false);
    setSelectOptions(els.kpiEmpleado, [], false);
    return;
  }

  // years/months from KPIS if available
  const years = uniqueSorted((KPIS.length ? KPIS : src).map(r=>r.anio));
  const monthsPresent = uniqueSorted((KPIS.length ? KPIS : src).map(r=>r.mes_num));
  const branches = uniqueSorted((KPIS.length ? KPIS : src).map(r=>r.sucursal));
  const sellers = uniqueSorted(KPIS.map(r=>r.vendedor));
  const kpis = uniqueSorted(KPIS.map(r=>r.kpi));

  setSelectOptions(els.yearA, years, false);
  setSelectOptions(els.yearB, years, false);

  // month shown as number 1-12 (más claro para cálculo)
  setSelectOptions(els.month, monthsList(), false);
  setSelectOptions(els.accStart, monthsList(), false);
  setSelectOptions(els.accEnd, monthsList(), false);

  setSelectOptions(els.branch, ["TODAS", ...branches], false);
  setSelectOptions(els.seller, ["TODOS", ...sellers], false);

  setSelectOptions(els.kpiSucursal, kpis, false);
  setSelectOptions(els.kpiEmpleado, kpis, false);

  // defaults
  const lastYear = years[years.length-1];
  const prevYear = years.length >= 2 ? years[years.length-2] : lastYear;
  const lastMonth = monthsPresent.length ? monthsPresent[monthsPresent.length-1] : 12;

  els.yearA.value = String(lastYear);
  els.yearB.value = String(prevYear);
  els.mode.value = "period";
  els.month.value = String(lastMonth);
  els.accStart.value = "1";
  els.accEnd.value = String(lastMonth);
  els.branch.value = "TODAS";
  els.seller.value = "TODOS";

  syncModeUI();
}

// ---------- mode ui ----------
function syncModeUI(){
  const isAccum = els.mode.value === "accum";
  els.periodMonthWrap.style.display = isAccum ? "none" : "block";
  els.accStartWrap.style.display = isAccum ? "block" : "none";
  els.accEndWrap.style.display = isAccum ? "block" : "none";
}

// ---------- global filters getter ----------
function getGlobal(){
  return {
    yearA: Number(els.yearA.value),
    yearB: Number(els.yearB.value),
    mode: els.mode.value, // period | accum
    month: clampMonth(els.month.value),
    accStart: clampMonth(els.accStart.value),
    accEnd: clampMonth(els.accEnd.value),
    branch: els.branch.value, // TODAS or name
    seller: els.seller.value, // TODOS or name
  };
}

// ---------- KPI filters ----------
function rowsForKpi({anio, kpi, branch, seller, month, mode, accStart, accEnd}){
  let lo = accStart, hi = accEnd;
  if(lo > hi){ const t=lo; lo=hi; hi=t; }

  return KPIS.filter(r=>{
    if(r.anio !== anio) return false;
    if(r.kpi !== kpi) return false;

    // filters
    if(branch !== "TODAS" && r.sucursal !== branch) return false;
    if(seller !== "TODOS" && r.vendedor !== seller) return false;

    if(mode === "period"){
      if(r.mes_num !== month) return false;
    }else{
      if(r.mes_num < lo || r.mes_num > hi) return false;
    }
    return true;
  });
}

function sumValue(rows){
  return sum(rows.map(r=>r.valor).filter(Number.isFinite));
}

// ---------- charts ----------
function destroyChart(ch){ if(ch){ ch.destroy(); } }

function drawLineChart(canvasEl, labels, datasets, unit){
  const options = {
    responsive:true,
    maintainAspectRatio:false,
    spanGaps:true,
    plugins:{
      legend:{ labels:{ color:"#cbd5e1" } },
      tooltip:{
        callbacks:{
          label:(ctx)=> `${ctx.dataset.label}: ${formatByUnit(ctx.parsed.y, unit)}`
        }
      }
    },
    scales:{
      x:{ ticks:{ color:"#94a3b8" }, grid:{ color:"rgba(31,42,68,.15)" } },
      y:{ ticks:{ color:"#94a3b8" }, grid:{ color:"rgba(31,42,68,.55)" } }
    }
  };

  return new Chart(canvasEl, { type:"line", data:{ labels, datasets }, options });
}

function drawBarChart(canvasEl, labels, datasets, unit){
  const options = {
    responsive:true,
    maintainAspectRatio:false,
    plugins:{
      legend:{ labels:{ color:"#cbd5e1" } },
      tooltip:{
        callbacks:{
          label:(ctx)=> `${ctx.dataset.label}: ${formatByUnit(ctx.parsed.y, unit)}`
        }
      }
    },
    scales:{
      x:{ ticks:{ color:"#94a3b8" }, grid:{ color:"rgba(31,42,68,.15)" } },
      y:{ ticks:{ color:"#94a3b8" }, grid:{ color:"rgba(31,42,68,.55)" } }
    }
  };

  return new Chart(canvasEl, { type:"bar", data:{ labels, datasets }, options });
}

// ---------- TAB NAV ----------
function setTab(active){
  // buttons
  [els.tabSucursal, els.tabEmpleado, els.tabMarcas, els.tabCategorias].forEach(b=>b.classList.remove("active"));
  // pages
  [els.pageSucursal, els.pageEmpleado, els.pageMarcas, els.pageCategorias].forEach(p=>p.classList.add("hidden"));

  if(active === "sucursal"){
    els.tabSucursal.classList.add("active");
    els.pageSucursal.classList.remove("hidden");
  }else if(active === "empleado"){
    els.tabEmpleado.classList.add("active");
    els.pageEmpleado.classList.remove("hidden");
  }else if(active === "marcas"){
    els.tabMarcas.classList.add("active");
    els.pageMarcas.classList.remove("hidden");
  }else{
    els.tabCategorias.classList.add("active");
    els.pageCategorias.classList.remove("hidden");
  }
  renderAll();
}

// ---------- Render: Sucursal ----------
function renderSucursal(){
  if(els.pageSucursal.classList.contains("hidden")) return;

  if(!KPIS.length){
    els.warnSucursal.textContent = "Sin DATA_KPIS.";
    els.tableSucursalBody.innerHTML = "";
    destroyChart(chartSucursal); chartSucursal=null;
    return;
  }

  const g = getGlobal();
  const kpi = els.kpiSucursal.value;
  const unit = inferUnitFromKpiName(kpi);

  // total A/B
  const rowsA = rowsForKpi({ anio:g.yearA, kpi, branch:g.branch, seller:g.seller, month:g.month, mode:g.mode, accStart:g.accStart, accEnd:g.accEnd });
  const rowsB = rowsForKpi({ anio:g.yearB, kpi, branch:g.branch, seller:g.seller, month:g.month, mode:g.mode, accStart:g.accStart, accEnd:g.accEnd });

  const sumA = sumValue(rowsA);
  const sumB = sumValue(rowsB);
  const d = computeDelta(sumA, sumB);

  els.sucA.textContent = formatByUnit(sumA, unit);
  els.sucB.textContent = formatByUnit(sumB, unit);
  els.sucDiffAbs.textContent = formatByUnit(d.abs, unit);
  els.sucDiffPct.textContent = fmtPct(d.pct);

  // meta text
  const rangeTxt = (g.mode==="period")
    ? `${MONTH_NAMES[g.month-1]}`
    : `${MONTH_NAMES[Math.min(g.accStart,g.accEnd)-1]}–${MONTH_NAMES[Math.max(g.accStart,g.accEnd)-1]}`;

  els.metaSucursal.textContent = `KPI: ${kpi} | ${rangeTxt} | Sucursal: ${g.branch} | Vendedor: ${g.seller} | A: ${g.yearA} vs B: ${g.yearB}`;

  // chart
  destroyChart(chartSucursal); chartSucursal=null;

  if(g.mode === "period"){
    // en periodo, gráfica: tendencia mensual A vs B (mismos filtros branch/seller, y kpi)
    const labels = monthsList().map(m=>MONTH_NAMES[m-1]);

    const monthlySum = (anio)=>{
      const rows = KPIS.filter(r=>{
        if(r.anio !== anio) return false;
        if(r.kpi !== kpi) return false;
        if(g.branch !== "TODAS" && r.sucursal !== g.branch) return false;
        if(g.seller !== "TODOS" && r.vendedor !== g.seller) return false;
        return true;
      });
      const byM = groupBy(rows, "mes_num");
      return monthsList().map(m => sumValue(byM.get(String(m)) || []));
    };

    const aData = monthlySum(g.yearA);
    const bData = monthlySum(g.yearB);

    els.chartTitleSucursal.textContent = "Tendencia mensual";
    els.chartMetaSucursal.textContent = `${g.yearA} vs ${g.yearB}`;

    chartSucursal = drawLineChart(els.chartSucursal, labels, [
      { label: String(g.yearA), data: aData },
      { label: String(g.yearB), data: bData },
    ], unit);

  }else{
    // acumulado: barras A vs B del rango
    let lo=g.accStart, hi=g.accEnd;
    if(lo>hi){ const t=lo; lo=hi; hi=t; }

    els.chartTitleSucursal.textContent = "Acumulado (rango DE/A)";
    els.chartMetaSucursal.textContent = `${MONTH_NAMES[lo-1]}–${MONTH_NAMES[hi-1]}`;

    chartSucursal = drawBarChart(els.chartSucursal,
      [`${g.yearA}`, `${g.yearB}`],
      [{ label: kpi, data: [sumA, sumB] }],
      unit
    );
  }

  // table by sucursal (si branch=TODAS, mostramos ranking por sucursal; si se eligió una sucursal, mostramos por vendedor)
  const groupKey = (g.branch === "TODAS") ? "sucursal" : "vendedor";
  const labelA = g.yearA;
  const labelB = g.yearB;

  const buildGroupSums = (anio)=>{
    // mismo KPI, mismo modo/rango, y filtros globales (excepto el grouping)
    let lo=g.accStart, hi=g.accEnd;
    if(lo>hi){ const t=lo; lo=hi; hi=t; }

    const rows = KPIS.filter(r=>{
      if(r.anio !== anio) return false;
      if(r.kpi !== kpi) return false;

      // si branch=TODAS, dejamos todas; si no, filtramos
      if(g.branch !== "TODAS" && r.sucursal !== g.branch) return false;

      // vendedor global: si estás en sucursal y eliges vendedor, se respeta
      if(g.seller !== "TODOS" && r.vendedor !== g.seller) return false;

      if(g.mode==="period"){
        if(r.mes_num !== g.month) return false;
      }else{
        if(r.mes_num < lo || r.mes_num > hi) return false;
      }
      return true;
    });

    const m = groupBy(rows, groupKey);
    const out = [];
    for(const [key, arr] of m.entries()){
      out.push({ key, value: sumValue(arr) });
    }
    return out;
  };

  const aGroups = buildGroupSums(g.yearA);
  const bGroups = buildGroupSums(g.yearB);
  const allKeys = uniqueSorted([...aGroups.map(x=>x.key), ...bGroups.map(x=>x.key)]);

  const mapA = new Map(aGroups.map(x=>[x.key, x.value]));
  const mapB = new Map(bGroups.map(x=>[x.key, x.value]));

  const rowsTable = allKeys.map(k=>{
    const va = mapA.get(k);
    const vb = mapB.get(k);
    const dd = computeDelta(va, vb);
    return { k, va, vb, dabs: dd.abs, dpct: dd.pct };
  }).sort((x,y)=> (y.va||0) - (x.va||0));

  els.tableSucursalBody.innerHTML = "";
  for(const r of rowsTable){
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${r.k}</td>
      <td>${formatByUnit(r.va, unit)}</td>
      <td>${formatByUnit(r.vb, unit)}</td>
      <td>${formatByUnit(r.dabs, unit)}</td>
      <td>${fmtPct(r.dpct)}</td>
    `;
    els.tableSucursalBody.appendChild(tr);
  }

  els.warnSucursal.textContent = rowsTable.length ? "" : "Sin filas para este filtro.";
}

// ---------- Render: Empleado ----------
function renderEmpleado(){
  if(els.pageEmpleado.classList.contains("hidden")) return;

  if(!KPIS.length){
    els.warnEmpleado.textContent = "Sin DATA_KPIS.";
    els.tableEmpleadoBody.innerHTML = "";
    return;
  }

  const g = getGlobal();
  const kpi = els.kpiEmpleado.value;
  const unit = inferUnitFromKpiName(kpi);

  const rangeTxt = (g.mode==="period")
    ? `${MONTH_NAMES[g.month-1]}`
    : `${MONTH_NAMES[Math.min(g.accStart,g.accEnd)-1]}–${MONTH_NAMES[Math.max(g.accStart,g.accEnd)-1]}`;

  els.metaEmpleado.textContent = `KPI: ${kpi} | ${rangeTxt} | Sucursal: ${g.branch} | A: ${g.yearA} vs B: ${g.yearB}`;

  // Empleado = TODOS (tabla completa), pero respeta filtro sucursal
  let lo=g.accStart, hi=g.accEnd;
  if(lo>hi){ const t=lo; lo=hi; hi=t; }

  const baseFilter = (anio)=> KPIS.filter(r=>{
    if(r.anio !== anio) return false;
    if(r.kpi !== kpi) return false;
    if(g.branch !== "TODAS" && r.sucursal !== g.branch) return false;
    // en tabla de empleados, NO filtramos por vendedor global (porque sería auto-filtrarte)
    if(g.mode==="period"){
      if(r.mes_num !== g.month) return false;
    }else{
      if(r.mes_num < lo || r.mes_num > hi) return false;
    }
    return true;
  });

  const rowsA = baseFilter(g.yearA);
  const rowsB = baseFilter(g.yearB);

  // group by (vendedor + sucursal) para evitar mezcla si hay mismo nombre en varias sucursales
  const keyFn = (r)=> `${r.vendedor}|||${r.sucursal}`;

  const mapSum = (rows)=>{
    const m = new Map();
    for(const r of rows){
      const key = keyFn(r);
      m.set(key, (m.get(key)||0) + (Number.isFinite(r.valor)?r.valor:0));
    }
    return m;
  };

  const aMap = mapSum(rowsA);
  const bMap = mapSum(rowsB);

  const allKeys = uniqueSorted([...aMap.keys(), ...bMap.keys()]);
  const tableRows = allKeys.map(k=>{
    const [vend, suc] = k.split("|||");
    const va = aMap.get(k);
    const vb = bMap.get(k);
    const d = computeDelta(va, vb);
    return { vend, suc, va, vb, dabs:d.abs, dpct:d.pct };
  }).sort((x,y)=> (y.va||0) - (x.va||0));

  // search filter
  const q = (els.empSearch.value || "").trim().toLowerCase();
  const filtered = q
    ? tableRows.filter(r => r.vend.toLowerCase().includes(q))
    : tableRows;

  els.empTableMeta.textContent = `Empleados: ${fmtInt(filtered.length)} (de ${fmtInt(tableRows.length)})`;

  els.tableEmpleadoBody.innerHTML = "";
  for(const r of filtered){
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${r.vend}</td>
      <td>${r.suc}</td>
      <td>${formatByUnit(r.va, unit)}</td>
      <td>${formatByUnit(r.vb, unit)}</td>
      <td>${formatByUnit(r.dabs, unit)}</td>
      <td>${fmtPct(r.dpct)}</td>
    `;
    els.tableEmpleadoBody.appendChild(tr);
  }

  els.warnEmpleado.textContent = filtered.length ? "" : "Sin filas para este filtro / búsqueda.";
}

// ---------- Render: Dimensiones (Marcas/Categorías) ----------
function renderDimPage(which){
  // which: "MARCA" | "CATEGORIA"
  if(which === "MARCA" && els.pageMarcas.classList.contains("hidden")) return;
  if(which === "CATEGORIA" && els.pageCategorias.classList.contains("hidden")) return;

  if(!DIM.length){
    if(which==="MARCA"){
      els.warnMarcas.textContent = "Sin DATA_DIM.";
      els.tableMarcasBody.innerHTML = "";
      destroyChart(chartMarcas); chartMarcas=null;
    }else{
      els.warnCategorias.textContent = "Sin DATA_DIM.";
      els.tableCategoriasBody.innerHTML = "";
      destroyChart(chartCategorias); chartCategorias=null;
    }
    return;
  }

  const g = getGlobal();
  // en dims, normalmente usamos el año A como “principal” (si quieres comparar A vs B en dims, lo hacemos después)
  const anio = g.yearA;

  let lo=g.accStart, hi=g.accEnd;
  if(lo>hi){ const t=lo; lo=hi; hi=t; }

  // en dims, usamos mismo modo/rango global:
  const rows = DIM.filter(r=>{
    if(r.anio !== anio) return false;
    if(r.dimension !== which) return false;
    if(g.branch !== "TODAS" && r.sucursal !== g.branch) return false;
    if(g.mode === "period"){
      if(r.mes_num !== g.month) return false;
    }else{
      if(r.mes_num < lo || r.mes_num > hi) return false;
    }
    return true;
  });

  const metric = (which==="MARCA") ? els.metricMarca.value : els.metricCat.value;
  const topN = Number((which==="MARCA") ? els.topMarca.value : els.topCat.value);

  // aggregate by nombre
  const m = new Map();
  for(const r of rows){
    const key = r.nombre;
    if(!m.has(key)){
      m.set(key, { nombre:key, ventas:0, utilidad:0, margenSum:0, margenCount:0 });
    }
    const it = m.get(key);
    if(Number.isFinite(r.ventas_con_iva)) it.ventas += r.ventas_con_iva;
    if(Number.isFinite(r.utilidad)) it.utilidad += r.utilidad;
    if(Number.isFinite(r.margen)){
      // margen: promedio simple (luego afinamos a ponderado por ventas si quieres)
      it.margenSum += r.margen;
      it.margenCount += 1;
    }
  }

  const items = [...m.values()].map(it=>{
    const margen = it.margenCount ? (it.margenSum / it.margenCount) : NaN;
    return { ...it, margen };
  });

  const metricValue = (it)=>{
    if(metric === "ventas_con_iva") return it.ventas;
    if(metric === "utilidad") return it.utilidad;
    if(metric === "margen") return it.margen;
    return it.ventas;
  };

  const sorted = items
    .filter(it => Number.isFinite(metricValue(it)))
    .sort((a,b)=> (metricValue(b) - metricValue(a)))
    .slice(0, topN);

  const rangeTxt = (g.mode==="period")
    ? `${MONTH_NAMES[g.month-1]}`
    : `${MONTH_NAMES[lo-1]}–${MONTH_NAMES[hi-1]}`;

  if(which==="MARCA"){
    els.metaMarcas.textContent = `Año: ${anio} | ${rangeTxt} | Sucursal: ${g.branch}`;
    els.chartMetaMarcas.textContent = `Métrica: ${metric} | Top ${sorted.length}`;
  }else{
    els.metaCategorias.textContent = `Año: ${anio} | ${rangeTxt} | Sucursal: ${g.branch}`;
    els.chartMetaCategorias.textContent = `Métrica: ${metric} | Top ${sorted.length}`;
  }

  // table
  const tbody = (which==="MARCA") ? els.tableMarcasBody : els.tableCategoriasBody;
  tbody.innerHTML = "";
  for(const it of sorted){
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${it.nombre}</td>
      <td>${fmtMoney(it.ventas)}</td>
      <td>${fmtMoney(it.utilidad)}</td>
      <td>${fmtPct(it.margen)}</td>
    `;
    tbody.appendChild(tr);
  }

  // chart (bar horizontal)
  const labels = sorted.map(it=>it.nombre).reverse();
  const data = sorted.map(it=>metricValue(it)).reverse();
  const unit = (metric === "margen") ? "percent" : "money";

  const canvas = (which==="MARCA") ? els.chartMarcas : els.chartCategorias;
  if(which==="MARCA"){ destroyChart(chartMarcas); chartMarcas=null; }
  else { destroyChart(chartCategorias); chartCategorias=null; }

  const options = {
    responsive:true,
    maintainAspectRatio:false,
    indexAxis:"y",
    plugins:{
      legend:{ labels:{ color:"#cbd5e1" } },
      tooltip:{ callbacks:{ label:(ctx)=> `${formatByUnit(ctx.parsed.x, unit)}` } }
    },
    scales:{
      x:{ ticks:{ color:"#94a3b8" }, grid:{ color:"rgba(31,42,68,.55)" } },
      y:{ ticks:{ color:"#94a3b8" }, grid:{ color:"rgba(31,42,68,.15)" } }
    }
  };

  const ch = new Chart(canvas, {
    type:"bar",
    data:{ labels, datasets:[{ label: which, data }] },
    options
  });

  if(which==="MARCA") chartMarcas = ch; else chartCategorias = ch;

  const warnEl = (which==="MARCA") ? els.warnMarcas : els.warnCategorias;
  warnEl.textContent = sorted.length ? "" : "Sin filas para este filtro.";
}

// ---------- render all ----------
function renderAll(){
  syncModeUI();
  renderSucursal();
  renderEmpleado();
  renderDimPage("MARCA");
  renderDimPage("CATEGORIA");
}

// ---------- events ----------
function bindEvents(){
  els.fileInput.addEventListener("change", async (e)=>{
    const file = e.target.files?.[0];
    if(!file) return;
    try{ await loadExcel(file); }
    catch(err){ alert(err?.message || "Error al cargar archivo."); console.error(err); }
    els.fileInput.value = "";
  });

  els.btnReset.addEventListener("click", ()=>{
    KPIS=[]; DIM=[];
    els.dataStatus.textContent = "Sin datos cargados";
    destroyChart(chartSucursal); chartSucursal=null;
    destroyChart(chartMarcas); chartMarcas=null;
    destroyChart(chartCategorias); chartCategorias=null;
    setupSelectors();
    renderAll();
  });

  // global filters
  [els.yearA, els.yearB, els.mode, els.month, els.accStart, els.accEnd, els.branch, els.seller].forEach(el=>{
    el.addEventListener("change", renderAll);
  });

  // KPI selects
  els.kpiSucursal.addEventListener("change", renderAll);
  els.kpiEmpleado.addEventListener("change", renderAll);

  // employee search
  els.empSearch.addEventListener("input", renderEmpleado);

  // dim controls
  [els.metricMarca, els.topMarca].forEach(el=> el.addEventListener("change", ()=> renderDimPage("MARCA")));
  [els.metricCat, els.topCat].forEach(el=> el.addEventListener("change", ()=> renderDimPage("CATEGORIA")));

  // tabs
  els.tabSucursal.addEventListener("click", ()=> setTab("sucursal"));
  els.tabEmpleado.addEventListener("click", ()=> setTab("empleado"));
  els.tabMarcas.addEventListener("click", ()=> setTab("marcas"));
  els.tabCategorias.addEventListener("click", ()=> setTab("categorias"));
}

// init
bindEvents();
setupSelectors();
syncModeUI();
setTab("sucursal");
renderAll();
