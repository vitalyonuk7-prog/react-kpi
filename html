/******************************************************
 * KPI Apps Script Backend (Code.gs)
 * Підтримує шити:
 *   [PROJECT] MONTHLY - <Manager>
 *   [PROJECT] WEEKLY  - <Manager>
 * Також "TOTAL" як окрема вкладка (агрегація по бренду).
 ******************************************************/

const CFG = {
  HTML_FILE_NAME: 'kpi_dash',
  TZ: () => Session.getScriptTimeZone(),
};

/** HTML */
function doGet() {
  return HtmlService.createHtmlOutputFromFile(CFG.HTML_FILE_NAME)
    .setTitle('KPI')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ====================== ПУБЛІЧНЕ API ======================= */

function kpiListManagers(project) {
  const map = scanProjectSheets_(project);         // {mgr:{monthly,weekly}}
  const managers = Object.keys(map).sort((a,b)=>a.localeCompare(b));
  // додамо TOTAL першим (як брендова сума)
  if (!managers.includes('TOTAL')) managers.unshift('TOTAL');
  return managers;
}

/** manager === "TOTAL" або "" → агрегуємо всіх менеджерів бренду */
function kpiLoadProject(project, _segIgnored, manager) {
  const proj = String(project||'').trim();
  const mgr  = normalizeManagerName_(String(manager||'').trim()); // "" → "", "TOTAL KPI" → "TOTAL"

  const projMap = scanProjectSheets_(proj); // {mgr:{monthly,weekly}}
  if (!Object.keys(projMap).length) return { metrics: [], months: [], nonZero: {} };

  // Кого беремо
  const mgrList = (!mgr || mgr === 'TOTAL') ? Object.keys(projMap)
                                            : (projMap[mgr] ? [mgr] : []);

  // ===== Акумулятори
  const metricSet = new Set();

  // MONTHLY: ym -> {metric:sum}
  const monthlyByYm = new Map();

  // WEEKLY: ym -> (dateKey -> {date, metrics(sum)})
  /** @type {Map<string, Map<string, {date:Date, metrics:Object}>>} */
  const weeklyByYm = new Map();

  // ===== Читаємо усіх потрібних менеджерів
  mgrList.forEach(m => {
    const pair = projMap[m] || {};
    // MONTHLY
    if (pair.monthly) {
      const { headers, rows } = getTable_(pair.monthly);
      const cDate = headers.indexOf('Date');
      if (cDate < 0) throw new Error(`Аркуш "${pair.monthly.getName()}" не містить колонки "Date" (MONTHLY).`);
      const metrics = headers.filter((h,i)=>i!==cDate);
      metrics.forEach(mm=>metricSet.add(mm));

      rows.forEach(r=>{
        const d = toDate_(r[cDate]);
        const ym = ymKey_(d);
        if (!monthlyByYm.has(ym)) monthlyByYm.set(ym, createZeroMap_(metrics));
        const bucket = monthlyByYm.get(ym);
        metrics.forEach(mn=>{
          bucket[mn] += toNumber_(r[headers.indexOf(mn)]);
        });
      });
    }

    // WEEKLY
    if (pair.weekly) {
      const { headers, rows } = getTable_(pair.weekly);
      const cDate = headers.indexOf('Date');
      if (cDate < 0) throw new Error(`Аркуш "${pair.weekly.getName()}" не містить колонки "Date" (WEEKLY).`);
      const metrics = headers.filter((h,i)=>i!==cDate);
      metrics.forEach(mm=>metricSet.add(mm));

      rows.forEach(r=>{
        const d = toDate_(r[cDate]);
        const ym = ymKey_(d);
        const dKey = Utilities.formatDate(d, CFG.TZ(), 'yyyy-MM-dd');

        if (!weeklyByYm.has(ym)) weeklyByYm.set(ym, new Map());
        const mapByDate = weeklyByYm.get(ym);
        if (!mapByDate.has(dKey)) mapByDate.set(dKey, { date: d, metrics: createZeroMap_(metrics) });

        const cell = mapByDate.get(dKey);
        metrics.forEach(mn=>{
          cell.metrics[mn] += toNumber_(r[headers.indexOf(mn)]);
        });
      });
    }
  });

  // ===== Формуємо відповідь
  const metrics = Array.from(metricSet).sort((a,b)=>a.localeCompare(b));

  // об’єднаний список місяців
  const ymUnion = new Set([
    ...Array.from(monthlyByYm.keys()),
    ...Array.from(weeklyByYm.keys()),
  ]);
  const ymList = Array.from(ymUnion).sort();

  const months = [];
  const nonZero = Object.create(null);
  metrics.forEach(m=>nonZero[m]=false);

  ymList.forEach(ym=>{
    const monthly = monthlyByYm.get(ym) || createZeroMap_(metrics);

    // зібрати й відсортувати тижні (агреговані по даті)
    const arr = weeklyByYm.get(ym)
      ? Array.from(weeklyByYm.get(ym).values()).sort((a,b)=>a.date - b.date)
      : [];

    // сума тижнів
    const sumWeeks = createZeroMap_(metrics);
    arr.forEach(w => addIntoMap_(sumWeeks, w.metrics));

    // weeks у формат фронта
    const weeks = arr.map((w,idx)=>({
      weekNo: idx+1,
      label : Utilities.formatDate(w.date, CFG.TZ(), 'yyyy-MM-dd'),
      weekly: w.metrics
    }));

    // nonZero
    metrics.forEach(k=>{
      if (Number(monthly[k]||0)!==0) nonZero[k] = true;
      if (Number(sumWeeks[k]||0)!==0) nonZero[k] = true;
      weeks.forEach(w=>{ if (Number(w.weekly[k]||0)!==0) nonZero[k]=true; });
    });

    months.push({
      ym,
      label: monthLabel_(ym),
      monthly,
      sumWeeks,
      weeks
    });
  });

  return { metrics, months, nonZero };
}

/* ====================== СКАНЕР ТАБІВ ======================= */

function scanProjectSheets_(project) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rx = new RegExp(
    String.raw`^\[\s*${escapeRegex_(project)}\s*\]\s*-?\s*(MONTHLY|WEEKLY)\s*-\s*(.+)$`,
    'i'
  );

  /** @type {{[mgr:string]: {monthly: GoogleAppsScript.Spreadsheet.Sheet|null, weekly: GoogleAppsScript.Spreadsheet.Sheet|null}}} */
  const map = {};

  ss.getSheets().forEach(sh=>{
    const m = sh.getName().trim().match(rx);
    if (!m) return;
    const kind = (m[1]||'').toUpperCase(); // MONTHLY|WEEKLY
    const raw  = String(m[2]||'').trim();
    const mgr  = normalizeManagerName_(raw);

    if (!map[mgr]) map[mgr] = { monthly:null, weekly:null };
    if (kind==='MONTHLY') map[mgr].monthly = sh;
    if (kind==='WEEKLY')  map[mgr].weekly  = sh;
  });

  return map;
}

/* ====================== УТИЛІТИ ======================= */

function getTable_(sheet){
  const v = sheet.getDataRange().getValues();
  if (!v || v.length<2) return { headers:[], rows:[] };
  const headers = v[0].map(x=>String(x||'').trim());
  const rows = v.slice(1).filter(r => r.some(c => c!=='' && c!=null));
  return { headers, rows };
}

function normalizeManagerName_(s){
  const t = s.replace(/\s+/g,' ').trim();
  if (/^total(\s+kpi)?$/i.test(t)) return 'TOTAL';
  return t;
}

function createZeroMap_(metrics){
  const o = Object.create(null);
  metrics.forEach(m=>o[m]=0);
  return o;
}
function addIntoMap_(dst, src){
  Object.keys(src).forEach(k=> dst[k] = (dst[k]||0) + (Number(src[k]||0)));
}

function toDate_(v){
  if (v instanceof Date) return v;
  const s = String(v||'').trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2])-1, Number(m[3]));
  const d = new Date(s);
  if (!isNaN(d)) return d;
  throw new Error(`Не вдалось розпізнати дату: "${s}"`);
}

function ymKey_(d){
  const y = d.getFullYear();
  const m = ('0'+(d.getMonth()+1)).slice(-2);
  return `${y}-${m}`;
}

function monthLabel_(ym){
  const [y,m] = ym.split('-').map(Number);
  const d = new Date(y, m-1, 1);
  return Utilities.formatDate(d, CFG.TZ(), 'MMM yyyy');
}

function toNumber_(v){
  if (v==='' || v==null) return 0;
  if (typeof v==='number') return isFinite(v)?v:0;
  let s = String(v).trim();
  s = s.replace(/[€\s]/g,'');
  if (s.endsWith('%')){
    const n = Number(s.slice(0,-1).replace(',', '.'));
    return isFinite(n)?n:0;
    }
  s = s.replace(',', '.');
  const n = Number(s);
  return isFinite(n)?n:0;
}

function escapeRegex_(s){
  return String(s).replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
}
