/* ===================================================================
   维修数据分析 — app.js v1.0
   纯前端、本地处理。加载 Excel → 数据概览 → 分析面板
   =================================================================== */
'use strict';

/* ---- 列定义 ---- */
const COLS = [
  {L:'A',cn:'维修单编号',en:'Ticket ID',type:'text'},
  {L:'B',cn:'维修状态',en:'Status',type:'enum'},
  {L:'C',cn:'维修单类型',en:'Ticket Type',type:'enum'},
  {L:'D',cn:'部门名称',en:'Department',type:'text'},
  {L:'E',cn:'报修部门名称',en:'Report Dept',type:'text'},
  {L:'F',cn:'车间名称',en:'Workshop',type:'text'},
  {L:'G',cn:'车间分类',en:'Workshop Cat',type:'enum'},
  {L:'H',cn:'产线',en:'Line',type:'text'},
  {L:'I',cn:'工序',en:'Process',type:'text'},
  {L:'J',cn:'设备分类名称',en:'Equip Cat',type:'text'},
  {L:'K',cn:'二级设备',en:'Sub Equip',type:'text'},
  {L:'L',cn:'设备名称',en:'Equip Name',type:'text'},
  {L:'M',cn:'设备编号',en:'Equip No',type:'text'},
  {L:'N',cn:'资产编码',en:'Asset Code',type:'text'},
  {L:'O',cn:'存放位置',en:'Location',type:'text'},
  {L:'P',cn:'故障等级',en:'Fault Level',type:'enum'},
  {L:'Q',cn:'紧急程度',en:'Urgency',type:'enum'},
  {L:'R',cn:'重复故障关联',en:'Repeat Fault',type:'text'},
  {L:'S',cn:'班次',en:'Shift',type:'enum'},
  {L:'T',cn:'问题描述',en:'Description',type:'text'},
  {L:'U',cn:'申请人',en:'Applicant',type:'text'},
  {L:'V',cn:'接收时间',en:'Received Time',type:'datetime'},
  {L:'W',cn:'故障原因',en:'Fault Cause',type:'text'},
  {L:'X',cn:'维修人',en:'Repair Person',type:'text'},
  {L:'Y',cn:'申请时间',en:'Apply Time',type:'datetime'},
  {L:'Z',cn:'驳回原因',en:'Reject Reason',type:'text'},
  {L:'AA',cn:'班长确认时间',en:'Confirm Time',type:'datetime'},
  {L:'AB',cn:'班长确认人',en:'Confirmer',type:'text'},
  {L:'AC',cn:'故障级别',en:'Fault Grade',type:'text'},
  {L:'AD',cn:'设备维修分类',en:'Repair Cat',type:'enum'},
  {L:'AE',cn:'故障部件名称',en:'Fault Part',type:'text'},
  {L:'AF',cn:'维修开始时间',en:'Repair Start',type:'datetime'},
  {L:'AG',cn:'维修完成时间',en:'Repair End',type:'datetime'},
  {L:'AH',cn:'参与维修人',en:'Participants',type:'text'},
  {L:'AI',cn:'等待时长',en:'Wait Time',type:'num'},
  {L:'AJ',cn:'维修时差',en:'Repair Diff',type:'num'},
  {L:'AK',cn:'维修时长',en:'Repair Duration',type:'num'},
  {L:'AL',cn:'故障原因描述',en:'Cause Desc',type:'text'},
  {L:'AM',cn:'维修内容',en:'Repair Content',type:'text'},
  {L:'AN',cn:'备件价值',en:'Parts Cost',type:'num'},
  {L:'AO',cn:'委外服务商',en:'Outsource',type:'text'},
  {L:'AP',cn:'委外费用',en:'Outsource Cost',type:'num'},
  {L:'AQ',cn:'维修总价值',en:'Total Cost',type:'num'},
  {L:'AR',cn:'是否影响产能',en:'Affects Capacity',type:'enum'},
  {L:'AS',cn:'是否影响质量',en:'Affects Quality',type:'enum'},
  {L:'AT',cn:'产量影响时长/h',en:'Impact Duration',type:'num'},
  {L:'AU',cn:'产能影响数量/kg',en:'Impact Qty',type:'num'},
  {L:'AV',cn:'确认人',en:'Confirmer',type:'text'},
  {L:'AW',cn:'确认时间',en:'Confirm Time',type:'datetime'},
  {L:'AX',cn:'维修质量评分',en:'Quality Score',type:'num'},
  {L:'AY',cn:'现场维修评分',en:'Onsite Score',type:'num'},
  {L:'AZ',cn:'重复维修评分',en:'Repeat Score',type:'num'},
  {L:'BA',cn:'接单结单及时性',en:'Timeliness',type:'num'},
  {L:'BB',cn:'现场维修时长',en:'Onsite Duration',type:'num'},
  {L:'BC',cn:'维修评价描述',en:'Eval Desc',type:'text'},
  {L:'BD',cn:'总分',en:'Total Score',type:'num'},
  {L:'BE',cn:'Gap',en:'Gap',type:'num'},
];

const COL_MAP = {};
COLS.forEach(c => { COL_MAP[c.L] = c; });

/* ---- 工具函数 ---- */
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];

function toast(m, persistent) {
  const t = $('#toast');
  t.innerHTML = '<span class="toast-msg">' + m + '</span>';
  t.classList.remove('toast-result');
  t.classList.add('show');
  clearTimeout(toast._t);
  if (!persistent) toast._t = setTimeout(() => t.classList.remove('show'), 2400);
}

function yieldUI() {
  return new Promise(resolve => requestAnimationFrame(() => setTimeout(resolve, 0)));
}

function toastResult(m, onOpen) {
  const t = $('#toast');
  const openBtn = onOpen ? '<button class="toast-btn toast-btn-open">打开</button>' : '';
  t.innerHTML = '<span class="toast-msg">' + m + '</span>' + openBtn +
    '<button class="toast-btn toast-btn-close">完成</button>';
  t.classList.add('toast-result');
  t.classList.add('show');
  clearTimeout(toast._t);
  if (onOpen) t.querySelector('.toast-btn-open').addEventListener('click', () => { onOpen(); });
  t.querySelector('.toast-btn-close').addEventListener('click', () => { t.classList.remove('show'); t.classList.remove('toast-result'); });
}

function esc(v) {
  if (v == null) return '';
  if (v instanceof Date) return v.toLocaleString('zh-CN');
  return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function colToIdx(L) {
  let r = 0;
  for (const ch of L) r = r * 26 + (ch.charCodeAt(0) - 64);
  return r - 1;
}

function serialToDate(n) {
  const wd = Math.floor(n), frac = n - wd;
  return new Date(1899, 11, 30 + (n < 60 ? wd : wd - 1), 0, 0, 0, Math.round(frac * 86400000));
}

function parseDateVal(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return v;
  if (typeof v === 'number') return serialToDate(v);
  if (typeof v === 'string') {
    const d = new Date(v.replace(/-/g, '/'));
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function fmtDate(d) {
  if (!d) return '';
  const p = n => String(n).padStart(2, '0');
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
}

function fmtDateTime(d) {
  if (!d) return '';
  const p = n => String(n).padStart(2, '0');
  return fmtDate(d) + ' ' + p(d.getHours()) + ':' + p(d.getMinutes());
}

/* ---- 全局状态 ---- */
const S = {
  fileName: '',
  rows: [],       // parsed data rows (array of objects keyed by column letter)
  stats: {},      // computed statistics
  analysisFilter: { months: [], workshop: '' },
  exportFilter: { months: [], workshop: '' },
};

/* ---- 初始化 ---- */
function initLoad() {
  const dz = $('#dropzone'), fi = $('#fileInput');
  $('#pickBtn').addEventListener('click', e => { e.stopPropagation(); fi.click(); });
  dz.addEventListener('click', () => fi.click());
  fi.addEventListener('change', () => { if (fi.files[0]) handleFile(fi.files[0]); });
  ['dragover', 'dragenter'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('drag'); }));
  ['dragleave', 'drop'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('drag'); }));
  dz.addEventListener('drop', e => { if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); });
  $('#todayStamp').textContent = '版本时间：2026.07.20';
}

/* ---- 文件处理 ---- */
async function handleFile(file) {
  var name = file.name || '';
  var ext = name.slice(name.lastIndexOf('.') + 1).toLowerCase();
  if (ext !== 'xlsx' && ext !== 'xlsm' && ext !== 'xls') {
    toast('文件格式不对：请上传 .xlsx / .xlsm / .xls 文件');
    return;
  }
  if (file.size > 50 * 1024 * 1024) {
    toast('文件过大：请上传 50MB 以内的文件');
    return;
  }
  try {
    toast('正在解析…');
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellFormula: true, cellNF: true });
    const sheetName = wb.SheetNames[0];
    if (!sheetName) { toast('未找到工作表：文件内无任何 Sheet'); return; }
    const ws = wb.Sheets[sheetName];
    const ref = ws['!ref'];
    if (!ref) { toast('工作表为空：第一个 Sheet 没有数据'); return; }
    const m = ref.match(/A\d+:(\w+?)(\d+)/);
    if (!m) { toast('表格式不对：无法识别数据范围（' + ref + '）'); return; }
    const maxCol = colToIdx(m[1]) + 1;
    if (maxCol < COLS.length) {
      toast('表格式不对：期望至少 ' + COLS.length + ' 列（A~' + COLS[COLS.length - 1].L + '），实际仅 ' + maxCol + ' 列');
      return;
    }
    const maxRow = parseInt(m[2]);
    if (maxRow < 2) { toast('工作表无数据行：仅有表头，缺少数据'); return; }
    S.fileName = name;
    parseData(ws);
    if (!S.rows.length) { toast('解析后无有效数据行：所有行均为空'); return; }
    computeStats();
    showOverview();
    toast('已载入：共 ' + S.rows.length + ' 条维修记录');
  } catch (err) {
    console.error(err);
    toast('载入失败：' + (err.message || '未知错误'));
  }
}

/* ---- 解析数据 ---- */
function parseData(ws) {
  const ref = ws['!ref'];
  if (!ref) { S.rows = []; return; }
  const m = ref.match(/A\d+:(\w+?)(\d+)/);
  const maxRow = m ? parseInt(m[2]) : 0;
  const maxCol = m ? colToIdx(m[1]) + 1 : COLS.length;

  S.rows = [];
  for (let r = 2; r <= maxRow; r++) {
    const row = {};
    let hasData = false;
    for (let c = 0; c < maxCol; c++) {
      const colLetter = XLSX.utils.encode_col(c);
      const cell = ws[colLetter + r];
      if (cell && cell.v != null && cell.v !== '') {
        row[colLetter] = cell.v;
        hasData = true;
      }
    }
    if (hasData) S.rows.push(row);
  }
}

/* ---- 统计计算 ---- */
function countBy(colLetter) {
  const counts = {};
  S.rows.forEach(row => {
    const v = row[colLetter];
    if (v == null || v === '') return;
    const key = String(v).trim();
    counts[key] = (counts[key] || 0) + 1;
  });
  return counts;
}

function computeStats() {
  S.stats = {};
  // 维修状态
  S.stats.status = countBy('B');
  // 维修单类型
  S.stats.ticketType = countBy('C');
  // 车间分类
  S.stats.workshopCat = countBy('G');
  // 车间名称
  S.stats.workshop = countBy('F');
  // 紧急程度
  S.stats.urgency = countBy('Q');
  // 班次
  S.stats.shift = countBy('S');
  // 故障等级
  S.stats.faultLevel = countBy('P');
  // 设备维修分类
  S.stats.repairCat = countBy('AD');
  // 是否影响产能
  S.stats.affectsCap = countBy('AR');
  // 是否影响质量
  S.stats.affectsQual = countBy('AS');
  // 维修质量评分
  S.stats.qualityScore = countBy('AX');
  // 设备名称（top 20）
  const equipCounts = countBy('L');
  S.stats.equipment = Object.entries(equipCounts).sort((a, b) => b[1] - a[1]).slice(0, 20);
  // 设备数量
  S.stats.equipCount = Object.keys(equipCounts).length;
  // 车间数量
  S.stats.workshopCount = Object.keys(S.stats.workshop).length;
  // 日期范围
  const dates = [];
  S.rows.forEach(row => {
    const d = parseDateVal(row['Y']);
    if (d) dates.push(d);
  });
  if (dates.length) {
    dates.sort((a, b) => a - b);
    S.stats.dateMin = dates[0];
    S.stats.dateMax = dates[dates.length - 1];
  }
}

/* ---- 展示概览 ---- */
function showOverview() {
  $('#stage-load').classList.remove('active');
  $('#stage-overview').classList.add('active');

  // 统计条
  $('#cntTotal').textContent = S.rows.length;
  $('#cntAssets').textContent = S.stats.equipCount;
  $('#cntWorkshops').textContent = S.stats.workshopCount;
  if (S.stats.dateMin && S.stats.dateMax) {
    const range = fmtDate(S.stats.dateMin) + '~' + fmtDate(S.stats.dateMax);
    $('#cntDateRange').textContent = range;
    $('#cntDateRange').style.fontSize = '.8rem';
  }

  // 图表
  renderHBarChart('#chartStatus', S.stats.status, 'alt');
  renderHBarChart('#chartWorkshop', S.stats.workshop);
  renderHBarChart('#chartCategory', S.stats.workshopCat, 'alt2');
  renderHBarChart('#chartUrgency', S.stats.urgency, 'alt3');
  renderHBarChart('#chartShift', S.stats.shift, 'alt');
  renderHBarChart('#chartFaultLevel', S.stats.faultLevel, 'alt2');

  // 数据预览
  renderPreview();

  // 绑定按钮（仅绑定一次）
  if (!_overviewBound) {
    _overviewBound = true;
    $('#backBtn').addEventListener('click', () => {
      $('#stage-overview').classList.remove('active');
      $('#stage-load').classList.add('active');
      S.rows = [];
      S.stats = {};
      const fi = $('#fileInput');
      if (fi) fi.value = '';
    });
    $('#exportBtn').addEventListener('click', openExport);
    $('#exportCancel').addEventListener('click', closeExport);
    $('#exportGo').addEventListener('click', doExport);

    // 可折叠区域
    $$('.collapsible-header').forEach(h => {
      h.addEventListener('click', () => {
        h.classList.toggle('collapsed');
        $('#' + h.dataset.target).classList.toggle('collapsed');
      });
    });
  }

  $('#exportBtn').disabled = false;

  // 分析面板
  initAnalysisPanel();
  initDataFilters();

  // 智能分析建议
  generateSuggestions();
}

/* ---- 智能分析建议 ---- */
const CROSS_WORKSHOP_METHODS = [3, 4, 15, 25, 44, 47, 57];
const MIN_ROWS_BY_DIM = { 1: 30, 2: 50, 3: 30, 4: 30, 5: 30, 6: 30, 7: 30, 8: 50, 9: 50, 10: 30, 11: 30 };

function extractTables(html) {
  const wrap = document.createElement('div');
  wrap.innerHTML = html;
  wrap.querySelectorAll('script').forEach(s => s.remove());
  const tables = wrap.querySelectorAll('table.result-table');
  const tableData = [];
  tables.forEach(table => {
    const headers = [...table.querySelectorAll('thead th')].map(th => th.textContent.trim());
    const rows = [...table.querySelectorAll('tbody tr')].map(tr =>
      [...tr.querySelectorAll('td')].map(td => td.textContent.trim())
    );
    tableData.push({ headers, rows });
  });
  return tableData;
}

async function generateSuggestions() {
  const listEl = $('#suggestionsList');
  const countEl = $('#suggestionCount');
  if (!listEl) return;
  listEl.innerHTML = '<p class="suggestion-loading">正在运行全局分析…</p>';
  countEl.textContent = '';

  let runCount = 0;
  async function runMethod(method, rows) {
    try {
      const html = method.fn(rows);
      const tables = extractTables(html);
      return { name: method.name, tables };
    } catch (e) {
      return { name: method.name, tables: [] };
    }
  }

  // === 阶段 1: 全局运行全部 81 个方法 ===
  const globalResults = {};
  for (let i = 0; i < ANALYSIS_METHODS.length; i++) {
    const method = ANALYSIS_METHODS[i];
    globalResults[method.id] = await runMethod(method, S.rows);
    runCount++;
    if (runCount % 10 === 0) {
      listEl.innerHTML = '<p class="suggestion-loading">全局分析中… (' + runCount + ' 次运行)</p>';
      await yieldUI();
    }
  }

  // === 阶段 2: 按车间运行非跨车间方法 ===
  const workshopGroups = {};
  S.rows.forEach(r => { const ws = r['F']; if (ws) (workshopGroups[ws] || (workshopGroups[ws] = [])).push(r); });
  const workshopNames = Object.keys(workshopGroups).filter(ws => workshopGroups[ws].length >= 30);
  const workshopResults = {};

  for (let wi = 0; wi < workshopNames.length; wi++) {
    const ws = workshopNames[wi];
    const wsRows = workshopGroups[ws];
    workshopResults[ws] = {};
    for (let i = 0; i < ANALYSIS_METHODS.length; i++) {
      const method = ANALYSIS_METHODS[i];
      if (CROSS_WORKSHOP_METHODS.includes(method.id)) continue;
      const minRows = MIN_ROWS_BY_DIM[method.dim] || 30;
      if (wsRows.length < minRows) continue;
      if (method.dim === 3) {
        const months = new Set();
        wsRows.forEach(r => { const d = parseDateVal(r['Y']); if (d) months.add(d.getFullYear() + '-' + (d.getMonth() + 1)); });
        if (months.size < 3) continue;
      }
      if (method.dim === 4) {
        const equips = new Set();
        wsRows.forEach(r => { const e = r['L']; if (e) equips.add(e); });
        if (equips.size < 3) continue;
      }
      workshopResults[ws][method.id] = await runMethod(method, wsRows);
      runCount++;
      if (runCount % 10 === 0) {
        listEl.innerHTML = '<p class="suggestion-loading">车间分析中（' + ws + ' ' + (wi + 1) + '/' + workshopNames.length + '）… (' + runCount + ' 次运行)</p>';
        await yieldUI();
      }
    }
  }

  // === 阶段 3: 三层建议引擎 ===
  listEl.innerHTML = '<p class="suggestion-loading">正在生成建议…</p>';
  await yieldUI();

  const suggestions = buildSuggestions(globalResults, workshopResults, workshopNames);
  _lastSuggestions = suggestions;
  renderSuggestions(suggestions);
}

function parseNum(s) {
  if (s == null) return null;
  const n = parseFloat(String(s).replace(/[%,]/g, ''));
  return isNaN(n) ? null : n;
}

function findTable(tables, headerKeywords) {
  for (const t of tables) {
    if (!t.headers) continue;
    const headerStr = t.headers.join(' ');
    if (headerKeywords.every(kw => headerStr.includes(kw))) return t;
  }
  return null;
}

function findRow(table, colKeyword) {
  if (!table || !table.rows) return null;
  for (const row of table.rows) {
    if (row.some(cell => String(cell).includes(colKeyword))) return row;
  }
  return null;
}

function findCellValue(table, rowKeyword, colIndex) {
  const row = findRow(table, rowKeyword);
  if (!row) return null;
  return parseNum(row[colIndex]);
}

function buildSuggestions(globalResults, workshopResults, workshopNames) {
  const suggestions = [];
  const seen = new Set();
  const nameCache = {};
  function methodName(id) {
    if (id in nameCache) return nameCache[id];
    const m = ANALYSIS_METHODS.find(am => am.id === id);
    const name = m ? (m.name + ' - ' + m.desc) : '';
    nameCache[id] = name;
    return name;
  }
  function add(severity, text, methodIds, scope, group) {
    const key = (scope || 'global') + '|' + text.replace(/\d+(\.\d+)?/g, '#');
    if (seen.has(key)) return;
    seen.add(key);
    const names = methodIds.map(methodName).filter(Boolean);
    suggestions.push({ severity, text, methods: names, scope: scope || 'global', group: group || '' });
  }
  runGlobalRules(globalResults, add);
  for (const ws of workshopNames) runWorkshopRules(ws, workshopResults[ws], globalResults, add);
  runCrossWorkshopComparison(workshopResults, workshopNames, globalResults, add);
  for (const ws of workshopNames) runCrossDiagnosis(ws, workshopResults[ws], globalResults, add);
  const overlapPairs = { cwHighWait: 'highWait', cwHighUrgent: 'highUrgent' };
  const suppressGlobalIfMerged = ['shiftImbalance', 'workloadImbalance', 'causeConcentration', 'highWait', 'highUrgent', 'risingTrend', 'scoreDecline'];
  const diagSuppress = { diag_preventive: ['highUrgent', 'risingTrend'], diag_process: ['highWait'], diag_aging: ['lowMTBF', 'faultConcentration'] };
  const filtered = suggestions.filter(s => {
    if (s.scope !== 'global' || !suppressGlobalIfMerged.includes(s.group)) return true;
    return !suggestions.some(o => o.group === s.group && o.scope !== 'global');
  }).filter(s => {
    if (!s.group || !overlapPairs[s.group]) return true;
    const baseGroup = overlapPairs[s.group];
    return !suggestions.some(o => o.group === baseGroup && o.scope === s.scope);
  }).filter(s => {
    if (!s.group || s.scope === 'global') return true;
    for (const diag in diagSuppress) {
      if (diagSuppress[diag].includes(s.group)) {
        if (suggestions.some(o => o.group === diag && o.scope === s.scope)) return false;
      }
    }
    return true;
  });
  const merged = mergeSuggestions(filtered);
  return merged.sort((a, b) => {
    const order = { high: 0, medium: 1, low: 2 };
    return (order[a.severity] ?? 3) - (order[b.severity] ?? 3);
  });
}

function mergeSuggestions(suggestions) {
  const GROUP_META = {
    lowMTBF: { summary: '设备 MTBF 过短，建议评估更换或大修', extract: (s, text) => { const m = s.text.match(/的 (.+?) 的 MTBF 仅 ([\d.]+) 天/); return m ? s.scope + '(' + m[1] + ' ' + m[2] + '天)' : text; } },
    repeatFault: { summary: '存在重复故障设备，建议制定根因分析计划', extract: (s, text) => { const m = s.text.match(/存在重复故障 (\d+) 次的设备/); return m ? s.scope + '(' + m[1] + '次)' : text; } },
    faultConcentration: { summary: '设备故障高度集中，建议重点维护', extract: (s, text) => { const m = text.match(/前 3 台设备累计 ([\d.]+)%/); return m ? s.scope + '(' + m[1] + '%)' : text; } },
    causeConcentration: { summary: '故障原因集中度高，建议制定根因消除计划', extract: (s, text) => { const m = text.match(/累计占比 ([\d.]+)%/); return m ? s.scope + '(' + m[1] + '%)' : text; } },
    highWait: { summary: '维修等待占比偏高，派单效率偏低', extract: (s, text) => { const m = text.match(/维修等待占比达 ([\d.]+)%/); return m ? s.scope + '(' + m[1] + '%)' : text; } },
    risingTrend: { summary: '最近一月维修量环比上升，建议排查设备老化', extract: (s, text) => { const m = text.match(/环比上升 ([\d.]+)%/); return m ? s.scope + '(+' + m[1] + '%)' : text; } },
    shiftImbalance: { summary: '班次维修量失衡，建议调整班次人员配置', extract: (s, text) => { const m = text.match(/是最低的 ([\d.]+) 倍/); return m ? s.scope + '(' + m[1] + '倍)' : text; } },
    lowScore: { summary: '维修评分偏低，建议制定质量改进计划', extract: (s, text) => { const m = text.match(/(.+?) 的平均维修评分仅 ([\d.]+) 分/); return m ? m[1] + '(' + m[2] + '分)' : text; } },
    highReject: { summary: '驳回率偏高，建议统一维修申请标准并加强培训', extract: (s, text) => { const m = text.match(/(.+?) 驳回率偏高（([\d.]+)%）/); return m ? m[1] + '(' + m[2] + '%)' : text; } },
    highOutsource: { summary: '委外费用占比接近或超过备件费用，建议增加维修技能培训', extract: (s) => s.scope },
    workloadImbalance: { summary: '维修人工作量分配不均，建议优化维修人员任务分配', extract: (s, text) => { const m = text.match(/是最低的 ([\d.]+) 倍/); return m ? s.scope + '(' + m[1] + '倍)' : text; } },
    aFaultHigh: { summary: 'A 级故障占比偏高，建议加强设备风险评估', extract: (s, text) => { const m = text.match(/A 级故障占比达 ([\d.]+)%/); return m ? s.scope + '(' + m[1] + '%)' : text; } },
    highUrgent: { summary: '突发/特急维修占比偏高，建议提高预防性维护频率', extract: (s, text) => { const m = text.match(/占比达 ([\d.]+)%/); return m ? s.scope + '(' + m[1] + '%)' : text; } },
    scoreDecline: { summary: '维修评分呈下降趋势', extract: (s, text) => { const m = text.match(/下降 ([\d.]+)%/); return m ? s.scope + '(-' + m[1] + '%)' : text; } },
    cwHighWait: { summary: '维修等待占比明显高于各车间平均，调度效率落后', extract: (s, text) => { const m = text.match(/维修等待占比 ([\d.]+)%/); return m ? s.scope + '(' + m[1] + '%)' : text; } },
    cwHighDur: { summary: '平均维修时长明显高于各车间平均，维修效率偏低', extract: (s, text) => { const m = text.match(/平均维修时长 ([\d.]+)/); return m ? s.scope + '(' + m[1] + 'h)' : text; } },
    cwHighCost: { summary: '平均维修成本明显高于各车间平均，需关注成本控制', extract: (s, text) => { const m = text.match(/平均维修成本 ([\d.]+)/); return m ? s.scope + '(' + m[1] + ')' : text; } },
    cwLowScore: { summary: '平均维修评分明显低于各车间平均，维修质量需改善', extract: (s, text) => { const m = text.match(/平均维修评分 ([\d.]+)/); return m ? s.scope + '(' + m[1] + '分)' : text; } },
    cwHighUrgent: { summary: '突发/特急维修占比明显高于各车间平均，预防性维护不足', extract: (s, text) => { const m = text.match(/占比 ([\d.]+)%/); return m ? s.scope + '(' + m[1] + '%)' : text; } },
    diag_aging: { summary: '综合诊断：设备老化严重，建议全面评估设备更换计划并优先安排大修', extract: (s) => s.scope },
    diag_process: { summary: '综合诊断：维修流程效率低，建议全面审查维修申请、派单、维修、验收全流程', extract: (s) => s.scope },
    diag_preventive: { summary: '综合诊断：预防性维护严重不足，建议立即加强预防性维护计划并提高巡检频率', extract: (s) => s.scope },
    diag_team: { summary: '综合诊断：维修团队能力不足，建议加强技能培训并调整人员配置', extract: (s) => s.scope },
    diag_quality: { summary: '综合诊断：维修质量差导致重复花费，建议开展根因分析并提升自修能力', extract: (s) => s.scope },
  };
  const grouped = {};
  const ungrouped = [];
  for (const s of suggestions) {
    if (!s.group || s.scope === 'global') { ungrouped.push(s); continue; }
    if (!grouped[s.group]) grouped[s.group] = [];
    grouped[s.group].push(s);
  }
  const result = [...ungrouped];
  for (const g in grouped) {
    const items = grouped[g];
    if (items.length <= 1) { result.push(...items); continue; }
    const meta = GROUP_META[g];
    const sev = items[0].severity;
    if (meta) {
      const detail = items.map(s => {
        const text = s.text.replace(s.scope, '').replace(/^[\s的]+/, '');
        return meta.extract(s, text);
      }).join('、');
      result.push({ severity: sev, text: items.length + ' 个车间的' + meta.summary + '：' + detail, methods: items[0].methods, scope: 'multi', group: g });
    } else if (g.startsWith('diag_general_')) {
      const desc = g.replace('diag_general_', '');
      const detail = items.map(s => s.scope).join('、');
      result.push({ severity: sev, text: items.length + ' 个车间的综合诊断：存在 ' + desc + ' 等多个问题，建议制定综合改进计划——涉及车间：' + detail, methods: items[0].methods, scope: 'multi', group: g });
    } else {
      const detail = items.map(s => s.text.replace(s.scope, '').replace(/^[\s的]+/, '')).join('；');
      result.push({ severity: sev, text: items.length + ' 个车间存在相同问题：' + detail, methods: items[0].methods, scope: 'multi', group: g });
    }
  }
  return result;
}

function getParetoCum(r, id) {
  if (!r[id] || !r[id].tables.length) return null;
  const t = r[id].tables[0];
  if (!t.headers || !t.headers.some(h => h.includes('累计占比')) || t.rows.length < 3) return null;
  return parseNum(t.rows[2][3]);
}
function getUrgentPct(r, id) {
  if (!r[id] || !r[id].tables.length || !r[id].tables[0].rows) return null;
  let sum = 0;
  for (const row of r[id].tables[0].rows) {
    if (String(row[0]).includes('突发') || String(row[0]).includes('特急')) {
      const p = parseNum(row[2]);
      if (p != null) sum += p;
    }
  }
  return sum;
}
function getTrendChange(r, id) {
  if (!r[id] || !r[id].tables.length) return null;
  const t = r[id].tables[0];
  if (!t.rows || t.rows.length < 3) return null;
  const last = parseNum(t.rows[t.rows.length - 1][1]);
  const prev = parseNum(t.rows[t.rows.length - 2][1]);
  if (last == null || prev == null || prev <= 0) return null;
  return (last - prev) / prev * 100;
}
function getFirstRowVal(r, id, colIdx) {
  if (!r[id] || !r[id].tables.length || !r[id].tables[0].rows) return null;
  return parseNum(r[id].tables[0].rows[0][colIdx]);
}
function getRowByLabel(r, id, label, colIdx) {
  if (!r[id] || !r[id].tables.length) return null;
  return findCellValue(r[id].tables[0], label, colIdx);
}
function getShiftImbalance(r, id) {
  if (!r[id] || !r[id].tables.length) return null;
  const t = r[id].tables[0];
  if (!t.rows || t.rows.length < 2) return null;
  const vals = t.rows.map(row => parseNum(row[1])).filter(v => v != null);
  if (vals.length < 2) return null;
  const max = Math.max(...vals), min = Math.min(...vals);
  if (min < 5) return null;
  return max / min;
}
function getWorkloadImbalance(r, id) {
  if (!r[id] || !r[id].tables.length) return null;
  const t = r[id].tables[0];
  if (!t.rows || t.rows.length < 5) return null;
  const top = parseNum(t.rows[0][2]);
  const bot = parseNum(t.rows[Math.min(4, t.rows.length - 1)][2]);
  if (top == null || bot == null || bot < 5) return null;
  return top / bot;
}
function getRepeatFaultCount(r, id) {
  if (!r[id] || !r[id].tables.length || !r[id].tables[0].rows) return null;
  for (const row of r[id].tables[0].rows) {
    const c = parseNum(row[row.length - 1]);
    if (c != null && c >= 5) return c;
  }
  return null;
}
function getLowMTBFName(r, id) {
  if (!r[id] || !r[id].tables.length || !r[id].tables[0].rows) return null;
  for (const row of r[id].tables[0].rows) {
    const mtbf = parseNum(row[3]);
    if (mtbf != null && mtbf < 7) return { name: row[1], mtbf };
  }
  return null;
}
function getAFaultPct(r, id) {
  if (!r[id] || !r[id].tables.length || !r[id].tables[0].rows) return null;
  for (const row of r[id].tables[0].rows) {
    if (String(row[0]).includes('A')) return parseNum(row[2]);
  }
  return null;
}
function getOutsourceRatio(r, id) {
  if (!r[id] || !r[id].tables.length) return null;
  const op = findCellValue(r[id].tables[0], '委外', 1);
  const pp = findCellValue(r[id].tables[0], '备件', 1);
  if (op == null || pp == null) return null;
  return op / pp;
}

function runRules(r, add, prefix) {
  const p = prefix || '';
  const sc = prefix ? prefix : 'global';
  const cum13 = getParetoCum(r, 13);
  if (cum13 != null && cum13 > 70) add('high', p + (prefix ? ' ' : '') + '设备故障高度集中：前 3 台设备累计' + (prefix ? ' ' : '贡献 ') + cum13.toFixed(1) + '%，建议' + (prefix ? '重点维护' : '将这 3 台设备列为重点维护对象') + '。', [13, 16, 29, 33], sc, 'faultConcentration');
  const cum14 = getParetoCum(r, 14);
  if (cum14 != null && cum14 > 60) add('medium', p + '故障原因集中度高：前 3 个故障原因累计占比 ' + cum14.toFixed(1) + '%，建议制定根因消除计划。', [14, 71, 72, 74], sc, 'causeConcentration');
  if (r[44] && r[44].tables.length && r[44].tables[0].rows) {
    for (const row of r[44].tables[0].rows) {
      const rate = parseNum(row[3]);
      if (rate != null && rate > 15) { add('medium', row[0] + ' 驳回率偏高（' + rate.toFixed(1) + '%），建议统一维修申请标准并加强培训。', [44], sc, 'highReject'); }
    }
  }
  const waitPct = getRowByLabel(r, 39, '等待占比', 1);
  if (waitPct != null && waitPct > 30) add('medium', p + '维修等待占比达 ' + waitPct.toFixed(1) + '%，' + (prefix ? '派单效率偏低' : '建议优化调度流程') + '。', [39, 11, 42, 45], sc, 'highWait');
  const trend = getTrendChange(r, 21);
  if (trend != null && trend > 20) add('medium', p + '最近一月维修量环比上升 ' + trend.toFixed(1) + '%，建议排查设备老化。', [21, 22, 27, 69], sc, 'risingTrend');
  const mtbf = getLowMTBFName(r, 29);
  if (mtbf) add('high', p + (prefix ? '的 ' : '') + mtbf.name + ' 的 MTBF 仅 ' + mtbf.mtbf.toFixed(1) + ' 天，建议评估更换或大修。', [29, 31, 32, 38], sc, 'lowMTBF');
  const shift = getShiftImbalance(r, 9);
  if (shift != null && shift > 2.5) add('low', p + '班次维修量失衡，最高是最低的 ' + shift.toFixed(1) + ' 倍，建议调整班次人员配置。', [9, 26, 40], sc, 'shiftImbalance');
  if (r[57] && r[57].tables.length && r[57].tables[0].rows) {
    for (const row of r[57].tables[0].rows) {
      const score = parseNum(row[2]);
      if (score != null && score < 3) { add('medium', row[0] + ' 的平均维修评分仅 ' + score.toFixed(2) + ' 分，建议制定质量改进计划。', [57, 54, 58, 59], sc, 'lowScore'); }
    }
  }
  const outRatio = getOutsourceRatio(r, 48);
  if (outRatio != null && outRatio > 0.8) add('low', p + '委外费用占比接近或超过备件费用，建议增加维修技能培训。', [48, 12, 47], sc, 'highOutsource');
  const repCount = getRepeatFaultCount(r, 32);
  if (repCount != null) add('high', p + '存在重复故障 ' + repCount + ' 次的设备，建议制定根因分析计划。', [32, 34, 35], sc, 'repeatFault');
  const wl = getWorkloadImbalance(r, 17);
  if (wl != null && wl > 4) add('low', p + '维修人工作量分配不均，最高是最低的 ' + wl.toFixed(1) + ' 倍，建议优化维修人员任务分配。', [17, 18, 41], sc, 'workloadImbalance');
  const aPct = getAFaultPct(r, 7);
  if (aPct != null && aPct > 20) add('high', p + 'A 级故障占比达 ' + aPct.toFixed(1) + '%，建议加强设备风险评估。', [7, 27, 53, 60], sc, 'aFaultHigh');
  const up = getUrgentPct(r, 8);
  if (up != null && up > 40) add('medium', p + '突发/特急维修占比达 ' + up.toFixed(1) + '%，建议提高预防性维护频率。', [8, 21, 69, 70], sc, 'highUrgent');
  const scoreTrend = getTrendChange(r, 59);
  if (scoreTrend != null && scoreTrend < -10) add('medium', p + '维修评分呈下降趋势（最近一月下降 ' + Math.abs(scoreTrend).toFixed(1) + '%）。', [59, 28, 54, 58], sc, 'scoreDecline');
}

function runGlobalRules(r, add) { runRules(r, add, ''); }
function runWorkshopRules(ws, wr, gr, add) { runRules(wr, add, ws); }

function runCrossWorkshopComparison(workshopResults, workshopNames, gr, add) {
  const metrics = {};
  for (const ws of workshopNames) {
    const wr = workshopResults[ws];
    const m = {};
    if (wr[39] && wr[39].tables.length) m.waitPct = findCellValue(wr[39].tables[0], '等待占比', 1);
    if (wr[10] && wr[10].tables.length) m.durMean = findCellValue(wr[10].tables[0], '均值', 1);
    if (wr[12] && wr[12].tables.length) m.costMean = findCellValue(wr[12].tables[0], '维修总价值(AQ) 均值', 1);
    if (wr[54] && wr[54].tables.length) m.scoreMean = findCellValue(wr[54].tables[0], '均值', 1);
    const up = getUrgentPct(wr, 8);
    if (up != null) m.urgentPct = up;
    metrics[ws] = m;
  }
  function gMean(key) {
    const vals = workshopNames.map(ws => metrics[ws][key]).filter(v => v != null);
    if (vals.length < 2) return null;
    return vals.reduce((s, v) => s + v, 0) / vals.length;
  }
  const gM = { waitPct: gMean('waitPct'), durMean: gMean('durMean'), costMean: gMean('costMean'), scoreMean: gMean('scoreMean'), urgentPct: gMean('urgentPct') };
  for (const ws of workshopNames) {
    const m = metrics[ws];
    if (m.waitPct != null && gM.waitPct != null && m.waitPct > gM.waitPct * 1.5) {
      add('medium', ws + ' 维修等待占比 ' + m.waitPct.toFixed(1) + '%，是各车间平均的 ' + (m.waitPct / gM.waitPct).toFixed(1) + ' 倍，调度效率明显落后。', [39, 11, 42, 45], ws, 'cwHighWait');
    }
    if (m.durMean != null && gM.durMean != null && m.durMean > gM.durMean * 1.5) {
      add('medium', ws + ' 平均维修时长 ' + m.durMean.toFixed(2) + '，是各车间平均的 ' + (m.durMean / gM.durMean).toFixed(1) + ' 倍，维修效率明显偏低。', [10, 40, 41, 43], ws, 'cwHighDur');
    }
    if (m.costMean != null && gM.costMean != null && m.costMean > gM.costMean * 1.5) {
      add('medium', ws + ' 平均维修成本 ' + m.costMean.toFixed(2) + '，是各车间平均的 ' + (m.costMean / gM.costMean).toFixed(1) + ' 倍，需关注成本控制。', [12, 47, 51, 52], ws, 'cwHighCost');
    }
    if (m.scoreMean != null && gM.scoreMean != null && m.scoreMean < gM.scoreMean * 0.7) {
      add('medium', ws + ' 平均维修评分 ' + m.scoreMean.toFixed(2) + '，明显低于各车间平均 ' + gM.scoreMean.toFixed(2) + '，维修质量需改善。', [54, 57, 58, 59], ws, 'cwLowScore');
    }
    if (m.urgentPct != null && gM.urgentPct != null && m.urgentPct > gM.urgentPct * 1.5) {
      add('low', ws + ' 突发/特急维修占比 ' + m.urgentPct.toFixed(1) + '%，是各车间平均的 ' + (m.urgentPct / gM.urgentPct).toFixed(1) + ' 倍，预防性维护明显不足。', [8, 21, 69, 70], ws, 'cwHighUrgent');
    }
  }
}

function runCrossDiagnosis(ws, wr, gr, add) {
  const issues = [];
  function push(key, methods) { issues.push({ key, methods }); }

  if (getParetoCum(wr, 13) > 70) push('faultConcentration', [13, 16, 29]);
  if (getLowMTBFName(wr, 29)) push('lowMTBF', [29, 31, 32]);
  const costMean = getRowByLabel(wr, 12, '维修总价值(AQ) 均值', 1);
  const gCost = getRowByLabel(gr, 12, '维修总价值(AQ) 均值', 1);
  if (costMean != null && gCost != null && costMean > gCost * 1.5) push('highCost', [12, 47, 51]);
  if (getRowByLabel(wr, 39, '等待占比', 1) > 30) push('highWait', [39, 11, 42]);
  if (getRowByLabel(wr, 54, '均值', 1) != null && getRowByLabel(wr, 54, '均值', 1) < 3) push('lowScore', [54, 58, 59]);
  if (gr[44] && gr[44].tables.length) {
    for (const row of gr[44].tables[0].rows) {
      if (String(row[0]) === ws) {
        if (parseNum(row[3]) > 15) push('highReject', [44]);
        break;
      }
    }
  }
  if (getUrgentPct(wr, 8) > 40) push('highUrgent', [8, 21, 69]);
  if (getTrendChange(wr, 21) > 20) push('risingTrend', [21, 22, 27]);
  if (getRepeatFaultCount(wr, 32) != null) push('repeatFault', [32, 34, 35]);
  const oRatio = getOutsourceRatio(wr, 48);
  if (oRatio != null && oRatio > 0.8) push('highOutsource', [48, 12, 47]);
  const wl = getWorkloadImbalance(wr, 17);
  if (wl != null && wl > 4) push('workloadImbalance', [17, 18, 41]);

  if (issues.length < 2) return;
  const iKeys = issues.map(i => i.key);
  const allM = issues.flatMap(i => i.methods);
  const sev = issues.length >= 3 ? 'high' : 'medium';

  if (iKeys.includes('faultConcentration') && iKeys.includes('lowMTBF') && iKeys.includes('highCost')) {
    add(sev, ws + ' 综合诊断：设备故障集中 + MTBF 过短 + 维修成本偏高，设备老化严重，建议全面评估设备更换计划并优先安排大修。', allM, ws, 'diag_aging'); return;
  }
  if (iKeys.includes('highWait') && iKeys.includes('lowScore') && iKeys.includes('highReject')) {
    add(sev, ws + ' 综合诊断：等待时间长 + 评分低 + 驳回率高，维修流程效率低，建议全面审查维修申请、派单、维修、验收全流程。', allM, ws, 'diag_process'); return;
  }
  if (iKeys.includes('highUrgent') && iKeys.includes('risingTrend') && iKeys.includes('lowMTBF')) {
    add(sev, ws + ' 综合诊断：突发维修多 + 维修量上升 + MTBF 短，预防性维护严重不足，建议立即加强预防性维护计划并提高巡检频率。', allM, ws, 'diag_preventive'); return;
  }
  if (iKeys.includes('workloadImbalance') && iKeys.includes('lowScore') && iKeys.includes('lowMTBF')) {
    add(sev, ws + ' 综合诊断：工作量不均 + 评分低 + MTBF 短，维修团队能力不足，建议加强技能培训并调整人员配置。', allM, ws, 'diag_team'); return;
  }
  if (iKeys.includes('highCost') && iKeys.includes('highOutsource') && iKeys.includes('repeatFault')) {
    add(sev, ws + ' 综合诊断：成本高 + 委外多 + 重复故障，维修质量差导致重复花费，建议开展根因分析并提升自修能力。', allM, ws, 'diag_quality'); return;
  }
  const labels = {
    faultConcentration: '设备故障集中', lowMTBF: 'MTBF 过短', highCost: '维修成本偏高',
    highWait: '等待时间长', lowScore: '维修评分低', highReject: '驳回率高',
    highUrgent: '突发维修多', risingTrend: '维修量上升', repeatFault: '重复故障',
    highOutsource: '委外费用高', workloadImbalance: '工作量不均',
  };
  const desc = issues.map(i => labels[i.key] || i.key).join('、');
  add(sev, ws + ' 综合诊断：存在 ' + desc + ' 等多个问题，建议制定综合改进计划。', allM, ws, 'diag_general_' + desc);
}

const SEVERITY_ICONS = { high: '🔴', medium: '🟡', low: '🟢' };

function renderSuggestions(suggestions) {
  const listEl = $('#suggestionsList');
  const countEl = $('#suggestionCount');
  if (!suggestions.length) {
    listEl.innerHTML = '<p class="suggestion-loading">基于当前数据分析，未发现需要特别关注的异常指标。</p>';
    countEl.textContent = '';
    return;
  }
  countEl.textContent = '（' + suggestions.length + ' 条）';
  const sorted = suggestions.slice().sort((a, b) => {
    const order = { high: 0, medium: 1, low: 2 };
    return (order[a.severity] ?? 3) - (order[b.severity] ?? 3);
  });
  listEl.innerHTML = sorted.map(s =>
    '<div class="suggestion-item">' +
    '<span class="suggestion-severity">' + (SEVERITY_ICONS[s.severity] || '⚪') + '</span>' +
    '<span class="suggestion-text">' + esc(s.text) + '</span>' +
    '<span class="suggestion-info"><span class="suggestion-info-icon">ℹ️</span>' +
    '<span class="suggestion-info-tip">分析方法：' + s.methods.join('、') + '</span></span>' +
    '</div>'
  ).join('');
}

/* ---- 分析面板 ---- */
const DIM_NAMES = ['', '描述性统计', '排名与帕累托', '时间趋势', '设备可靠性', '维修效率', '成本与影响', '评分与质量', '相关性分析', '回归与预测', '文本分析', '大数据分析'];
function scoreBadge(score, label) {
  const v = parseFloat(score);
  const stars = Math.max(0.5, Math.round(v * 2) / 2);
  const full = Math.floor(stars);
  const half = stars % 1 !== 0;
  const empty = 5 - full - (half ? 1 : 0);
  let html = '<span class="star-badge" title="' + (label || '综合评分') + ' ' + score + '">';
  for (let i = 0; i < full; i++) html += '<span class="star-full">★</span>';
  if (half) html += '<span class="star-half">★</span>';
  for (let i = 0; i < empty; i++) html += '<span class="star-empty">★</span>';
  html += '</span>';
  return html;
}

let _selectedMethodId = null;
let _overviewBound = false;
let _analysisBound = false;
let _favorites = new Set(JSON.parse(localStorage.getItem('ma_fav') || '[]'));
let _lastSuggestions = [];

function toggleFav(id) {
  if (_favorites.has(id)) _favorites.delete(id);
  else _favorites.add(id);
  localStorage.setItem('ma_fav', JSON.stringify([..._favorites]));
  renderMethodList();
}

function initAnalysisPanel() {
  renderMethodList();
  if (!_analysisBound) {
    _analysisBound = true;
    $('#methodSearch').addEventListener('input', renderMethodList);
    $('#filterDimension').addEventListener('change', renderMethodList);
    $('#filterScore').addEventListener('change', renderMethodList);
  }
}

function renderMethodList() {
  const q = ($('#methodSearch').value || '').toLowerCase().trim();
  const dim = $('#filterDimension').value;
  const minScore = parseFloat($('#filterScore').value) || 0;
  let methods = ANALYSIS_METHODS.filter(m => {
    if (dim && m.dim !== parseInt(dim)) return false;
    if (minScore && parseFloat(m.score) < minScore) return false;
    if (q && !m.name.toLowerCase().includes(q) && !m.desc.toLowerCase().includes(q)) return false;
    return true;
  }).sort((a, b) => {
    const af = _favorites.has(a.id) ? 1 : 0;
    const bf = _favorites.has(b.id) ? 1 : 0;
    if (af !== bf) return bf - af;
    return parseFloat(b.score) - parseFloat(a.score);
  });
  const list = $('#methodList');
  if (!methods.length) {
    list.innerHTML = '<p class="no-method">无匹配方法</p>';
    return;
  }
  list.innerHTML = methods.map(m =>
    '<div class="method-card' + (m.id === _selectedMethodId ? ' selected' : '') + '" data-mid="' + m.id + '">' +
    '<div class="mc-header">' +
    '<span class="mc-dim">' + DIM_NAMES[m.dim] + '</span>' +
    '<span class="mc-fav-row">' +
    '<span class="mc-fav' + (_favorites.has(m.id) ? ' active' : '') + '" data-fav="' + m.id + '">♥</span>' +
    scoreBadge(m.score) +
    '</span>' +
    '</div>' +
    '<div class="mc-name">' + esc(m.name) + '</div>' +
    '<div class="mc-fields">字段: ' + esc(fmtFields(m.fields)) + '</div>' +
    '</div>'
  ).join('');
  $$('.method-card', list).forEach(card => {
    card.addEventListener('click', e => {
      if (e.target.classList.contains('mc-fav')) { e.stopPropagation(); toggleFav(parseInt(e.target.dataset.fav)); return; }
      selectMethod(parseInt(card.dataset.mid));
    });
  });
}

function fmtFields(fields) {
  return fields.split(',').map(L => {
    const t = L.trim();
    const col = COL_MAP[t];
    return col ? t + ' ' + col.cn : t;
  }).join(', ');
}

function selectMethod(mid) {
  _selectedMethodId = mid;
  const method = ANALYSIS_METHODS.find(m => m.id === mid);
  if (!method) return;
  $$('.method-card').forEach(c => c.classList.toggle('selected', parseInt(c.dataset.mid) === mid));
  const exec = $('#analysisExec');
  exec.innerHTML =
    '<div class="exec-header">' +
    '<div class="exec-title">' +
    '<span class="exec-dim">' + DIM_NAMES[method.dim] + '</span>' +
    '<span class="exec-score-total"><span class="exec-score-label">综合评分</span>' + scoreBadge(method.score) + '</span>' +
    '</div>' +
    '<h3 class="exec-name">' + esc(method.name + ' - ' + method.desc) + '</h3>' +
    '<p class="exec-fields">涉及字段: ' + esc(fmtFields(method.fields)) + '</p>' +
    (method.guide ? '<p class="exec-guide-text">分析方法：' + esc(method.guide) + '</p>' : '') +
    '<p class="exec-conclusion" id="execConclusion">分析结论：分析中…</p>' +
    '</div>' +
    '<div class="exec-result" id="execResult"><p class="exec-hint">分析中…</p></div>';
  executeMethod(method);
}

function generateConclusion(methodId, tables) {
  if (!tables || !tables.length) return '无';
  const t = tables[0];
  const rows = t.rows;
  if (!rows || !rows.length) return '无';
  const num = s => parseFloat(String(s).replace(/[%,]/g, '')) || 0;
  const findRow = (kw) => rows.find(r => r.some(c => String(c).includes(kw)));
  const findVal = (kw) => { const r = findRow(kw); return r ? r[r.length - 1] : null; };

  switch (methodId) {
    case 1: { // 维修单总量与状态分布
      const total = num(rows[0][1]);
      const rej = rows.find(r => /驳回/.test(r[0]));
      const rejPct = rej ? num(rej[2]) : 0;
      if (rejPct > 15) return '驳回率 ' + rejPct.toFixed(1) + '% 偏高，建议统一维修申请标准并加强需求描述培训，减少无效审批流转。';
      if (rejPct > 5) return '驳回率 ' + rejPct.toFixed(1) + '%，处于正常偏高水平，建议关注驳回原因分布，优化审批流程。';
      return '维修单总量 ' + total + ' 条，驳回率 ' + rejPct.toFixed(1) + '%，整体流程健康度良好。';
    }
    case 2: { // 维修单类型分布
      const top = rows[0];
      return '主要维修类型为「' + top[0] + '」，占比 ' + top[2] + '，建议据此分配维修资源。';
    }
    case 3: { // 车间分类分布
      const top = rows[0];
      const second = rows[1] || null;
      if (second && num(top[1]) > num(second[1]) * 2) return '「' + top[0] + '」维修量 ' + top[1] + ' 远超其他分类，建议重点保障该分类的备件和人力。';
      return '各分类维修量分布较均衡，「' + top[0] + '」占比最高（' + top[1] + '）。';
    }
    case 4: { // 车间维修频次分布
      const top = rows[0];
      const total = rows.reduce((s, r) => s + num(r[1]), 0);
      const pct = (num(top[1]) / total * 100).toFixed(1);
      if (parseFloat(pct) > 30) return '「' + top[0] + '」维修量 ' + top[1] + ' 条（占 ' + pct + '%），集中度高，建议重点保障备件库存和维修人员配置。';
      return '「' + top[0] + '」维修量最高（' + top[1] + ' 条），各车间分布相对均衡。';
    }
    case 5: { // 产线维修频次分布
      const top = rows[0];
      const second = rows[1] || null;
      if (second && num(top[1]) > num(second[1]) * 2) return '产线「' + top[0] + '」维修量 ' + top[1] + ' 条，显著高于其他产线，建议排查该产线设备状态或操作规范。';
      return '产线「' + top[0] + '」维修量最高（' + top[1] + ' 条），整体分布较均匀。';
    }
    case 6: { // 工序维修频次分布
      const top = rows[0];
      return '工序「' + top[0] + '」故障频次最高（' + top[1] + ' 次），建议针对该工序安排预防性维护。';
    }
    case 7: { // 故障等级分布
      const aRow = rows.find(r => /A/.test(r[0]));
      const aPct = aRow ? num(aRow[2]) : 0;
      if (aPct > 20) return 'A级故障占比 ' + aPct.toFixed(1) + '%，偏高，设备可靠性风险较大，建议加强设备风险评估和预防性维护。';
      if (aPct > 10) return 'A级故障占比 ' + aPct.toFixed(1) + '%，需关注关键设备状态，建议加强巡检。';
      return 'A级故障占比 ' + aPct.toFixed(1) + '%，整体故障等级分布合理。';
    }
    case 8: { // 紧急程度分布
      const urgent = rows.filter(r => /突发|特急/.test(r[0]));
      const uPct = urgent.reduce((s, r) => s + num(r[2]), 0);
      if (uPct > 40) return '突发/特急维修占比 ' + uPct.toFixed(1) + '%，偏高，预防性维护不足，建议提高巡检频率和预防性维护计划覆盖率。';
      if (uPct > 25) return '突发/特急维修占比 ' + uPct.toFixed(1) + '%，建议关注预防性维护计划的执行效果。';
      return '突发/特急维修占比 ' + uPct.toFixed(1) + '%，紧急程度分布正常。';
    }
    case 9: { // 班次维修量对比
      if (rows.length < 2) return '无';
      const ratio = num(rows[0][1]) / Math.max(1, num(rows[rows.length - 1][1]));
      if (ratio > 1.5) return '班次维修量失衡，最高是最低的 ' + ratio.toFixed(1) + ' 倍，建议调整班次人员配置。';
      return '各班次维修量基本均衡，比值 ' + ratio.toFixed(2) + '。';
    }
    case 10: { // 维修时长统计
      const meanV = findVal('均值') || findVal('平均');
      const medV = findVal('中位数');
      const p90 = findVal('P90');
      if (meanV && medV) {
        const m = num(meanV), med = num(medV);
        if (m > med * 1.5) return '维修时长均值 ' + m.toFixed(2) + 'h 显著高于中位数 ' + med.toFixed(2) + 'h，存在异常长维修拉高均值，建议排查超长维修单的共性原因。';
        return '维修时长中位数 ' + med.toFixed(2) + 'h，均值 ' + m.toFixed(2) + 'h，分布较为集中。P90 为 ' + (p90 || '—') + 'h。';
      }
      return '无';
    }
    case 11: { // 等待时长统计
      const meanV = findVal('均值') || findVal('平均');
      if (meanV) {
        const m = num(meanV);
        if (m > 24) return '等待时长均值 ' + m.toFixed(2) + 'h，过长，维修调度效率低，建议优化派单流程。';
        if (m > 8) return '等待时长均值 ' + m.toFixed(2) + 'h，偏长，建议关注派单响应速度。';
        return '等待时长均值 ' + m.toFixed(2) + 'h，响应速度正常。';
      }
      return '无';
    }
    case 12: { // 维修费用统计
      const anMean = findVal('备件价值(AN) 均值');
      const apMean = findVal('委外费用(AP) 均值');
      if (anMean && apMean) {
        const an = num(anMean), ap = num(apMean);
        if (ap > an * 0.8) return '委外费用均值 ' + ap.toFixed(2) + ' 接近备件费用 ' + an.toFixed(2) + '，委外占比偏高，建议评估自修能力并增加维修技能培训。';
        return '备件费用均值 ' + an.toFixed(2) + '，委外费用均值 ' + ap.toFixed(2) + '，成本结构合理。';
      }
      return '无';
    }
    case 13: { // 设备故障频次帕累托
      const idx = rows.findIndex(r => parseFloat(String(r[3]).replace('%', '')) >= 80);
      if (idx >= 0) return '前 ' + (idx + 1) + ' 台设备累计故障占比达 80%，符合帕累托原则，建议将这 ' + (idx + 1) + ' 台设备列为重点维护对象并优先储备备件。';
      return '故障分布较为分散，未出现明显的帕累托集中趋势。';
    }
    case 14: { // 故障原因帕累托
      const idx = rows.findIndex(r => parseFloat(String(r[3]).replace('%', '')) >= 80);
      if (idx >= 0) return '前 ' + (idx + 1) + ' 个故障原因累计占比达 80%，集中解决这些原因可大幅降低维修量，建议制定根因消除计划。';
      return '故障原因分布分散，无明显的帕累托集中。';
    }
    case 15: { // 车间维修成本帕累托
      const top = rows[0];
      const idx = rows.findIndex(r => parseFloat(String(r[3]).replace('%', '')) >= 80);
      if (idx >= 0) return '「' + top[0] + '」维修成本最高（' + top[1] + '），前 ' + (idx + 1) + ' 个车间累计成本占比达 80%，建议重点控制这些车间的维修支出。';
      return '各车间成本分布较均匀，「' + top[0] + '」成本最高（' + top[1] + '）。';
    }
    case 16: { // 高频故障设备 Top 20
      const top = rows[0];
      return '故障频次最高的设备为「' + top[1] + '」（' + top[2] + ' 次），建议优先安排预防性维护和备件储备。';
    }
    case 17: { // 维修人工作量排名
      if (rows.length < 2) return '无';
      const top = rows[0], bottom = rows[rows.length - 1];
      const ratio = num(top[2]) / Math.max(1, num(bottom[2]));
      if (ratio > 3) return '维修人工作量分配不均，「' + top[1] + '」维修单数 ' + top[2] + '，是最低的 ' + ratio.toFixed(1) + ' 倍，建议优化维修人员任务分配。';
      return '维修人工作量分布较均衡，最高/最低比值为 ' + ratio.toFixed(1) + '。';
    }
    case 18: { // 维修人效率排名
      const top = rows[0], bottom = rows[rows.length - 1];
      const ratio = num(bottom[3]) / Math.max(0.01, num(top[3]));
      if (ratio > 2) return '维修人效率差异较大，最长平均时长「' + bottom[1] + '」（' + bottom[3] + 'h）是最短「' + top[1] + '」的 ' + ratio.toFixed(1) + ' 倍，建议对效率偏低的维修人进行技能培训。';
      return '维修人效率差异不大，平均时长分布较均匀。';
    }
    case 19: { // 故障部件频次排名
      const top = rows[0];
      return '故障频次最高的部件为「' + top[1] + '」（' + top[2] + ' 次），建议提高该部件备件库存。';
    }
    case 20: { // 高成本维修单 Top 20
      const top = rows[0];
      return '最高成本维修单为「' + top[1] + '」（' + top[4] + ' 元），设备「' + top[3] + '」，建议分析高成本维修单的共性原因并设定审批阈值。';
    }
    case 21: { // 月度维修单数量趋势
      if (rows.length < 2) return '无';
      const first = num(rows[0][1]), last = num(rows[rows.length - 1][1]);
      const change = ((last - first) / Math.max(1, first) * 100).toFixed(1);
      if (parseFloat(change) > 20) return '维修量从 ' + rows[0][0] + ' 的 ' + first + ' 条上升至 ' + rows[rows.length - 1][0] + ' 的 ' + last + ' 条（+' + change + '%），趋势上升明显，建议排查设备老化或维护不足问题。';
      if (parseFloat(change) < -20) return '维修量从 ' + first + ' 下降至 ' + last + '（' + change + '%），趋势下降，维护效果良好。';
      return '维修量整体趋势平稳，首末月变化 ' + change + '%。';
    }
    case 22: { // 周度维修单数量趋势
      if (rows.length < 2) return '无';
      const max = rows.reduce((m, r) => num(r[1]) > num(m[1]) ? r : m, rows[0]);
      return '维修量最高的周为「' + max[0] + '」（' + max[1] + ' 条），可能对应设备集中故障或季节性因素，建议关注该时段的维修资源调度。';
    }
    case 23: { // 月度维修时长趋势
      if (rows.length < 2) return '无';
      const first = num(rows[0][2]), last = num(rows[rows.length - 1][2]);
      const change = ((last - first) / Math.max(0.01, first) * 100).toFixed(1);
      if (parseFloat(change) > 20) return '平均维修时长从 ' + first + 'h 上升至 ' + last + 'h（+' + change + '%），维修难度增加或技能下降，建议评估维修团队能力。';
      return '平均维修时长趋势平稳，首末月变化 ' + change + '%。';
    }
    case 24: { // 月度维修费用趋势
      if (rows.length < 2) return '无';
      const maxRow = rows.reduce((m, r) => num(r[3]) > num(m[3]) ? r : m, rows[0]);
      return '维修费用最高的月份为「' + maxRow[0] + '」（总成本 ' + maxRow[3] + '），建议排查该月是否有大修或异常故障。';
    }
    case 25: { // 各车间月度故障趋势
      if (rows.length < 2) return '无';
      const wsCols = t.headers.slice(1);
      const lastRow = rows[rows.length - 1];
      const firstRow = rows[0];
      let maxWs = wsCols[0], maxChange = -Infinity;
      wsCols.forEach((ws, i) => {
        const change = num(lastRow[i + 1]) - num(firstRow[i + 1]);
        if (change > maxChange) { maxChange = change; maxWs = ws; }
      });
      if (maxChange > 0) return '「' + maxWs + '」维修量增长最多（+' + maxChange + ' 条），趋势恶化最快，建议优先干预。';
      return '各车间维修量趋势整体平稳或下降。';
    }
    case 26: { // 班次维修量月度趋势
      if (rows.length < 2) return '无';
      return '各班次月度维修量趋势已展示，建议关注两班趋势是否分化，夜班持续偏高需排查夜班设备巡检质量。';
    }
    case 27: { // 故障等级月度趋势
      if (rows.length < 2) return '无';
      const aCol = t.headers.findIndex(h => /A/.test(h));
      if (aCol < 0) return '无';
      const first = num(rows[0][aCol]), last = num(rows[rows.length - 1][aCol]);
      if (last > first * 1.5) return 'A级故障从 ' + first + ' 增至 ' + last + '，设备可靠性恶化，建议加强设备风险评估。';
      return 'A级故障月度趋势平稳。';
    }
    case 28: { // 维修评分月度趋势
      if (rows.length < 2) return '无';
      const first = num(rows[0][1]), last = num(rows[rows.length - 1][1]);
      const change = ((last - first) / Math.max(0.01, first) * 100).toFixed(1);
      if (parseFloat(change) < -10) return '维修评分从 ' + first + ' 下降至 ' + last + '（' + change + '%），维修质量下降，建议调查原因。';
      return '维修评分趋势平稳，首末月变化 ' + change + '%。';
    }
    case 29: { // MTBF 计算
      const low = rows[rows.length - 1];
      const lowMtbf = num(low[3]);
      if (lowMtbf < 7) return 'MTBF 最短的设备「' + low[1] + '」仅 ' + lowMtbf.toFixed(1) + ' 天，故障频繁，建议评估更换或大修。';
      if (lowMtbf < 30) return 'MTBF 最短的设备「' + low[1] + '」为 ' + lowMtbf.toFixed(1) + ' 天，建议关注并加强预防性维护。';
      return '各设备 MTBF 整体在合理范围内，最短为 ' + lowMtbf.toFixed(1) + ' 天。';
    }
    case 30: { // MTTR 计算
      const top = rows[0];
      const topMttr = num(top[3]);
      if (topMttr > 24) return 'MTTR 最长的设备「' + top[1] + '」达 ' + topMttr.toFixed(2) + 'h，可能备件不足或维修难度大，建议预储备件和优化维修方案。';
      return '各设备 MTTR 整体正常，最长为 ' + topMttr.toFixed(2) + 'h。';
    }
    case 31: { // 设备可用度估算
      const low = rows[rows.length - 1];
      const avail = parseFloat(String(low[5]).replace('%', ''));
      if (avail < 90) return '设备「' + low[1] + '」可用度仅 ' + avail.toFixed(1) + '%，严重影响产能，建议优先安排设备更新。';
      if (avail < 95) return '设备「' + low[1] + '」可用度 ' + avail.toFixed(1) + '%，低于 95% 阈值，建议关注。';
      return '各设备可用度均高于 95%，整体可靠性良好。';
    }
    case 32: { // 重复故障识别
      const top = rows[0];
      const cnt = num(top[2]);
      if (cnt > 5) return '设备「' + top[1] + '」重复故障 ' + cnt + ' 次，说明上次维修未解决根因，建议制定专项根因分析计划。';
      return '重复故障整体较少，最多 ' + cnt + ' 次。';
    }
    case 33: { // 设备故障频率趋势
      if (rows.length < 2) return '无';
      return 'Top 5 设备的月度故障频率趋势已展示，建议关注频率加速上升的设备，可能已进入衰退期。';
    }
    case 34: { // 设备故障模式分析
      return '设备×故障原因交叉表已展示，建议关注高频组合，某设备某原因集中说明系统性问题，需针对性维修方案。';
    }
    case 35: { // 故障部件频率分析
      return '设备×故障部件交叉表已展示，建议关注某设备频繁损坏的部件，可能是设计缺陷或操作不当。';
    }
    case 36: { // 设备维修分类对比
      return '各设备维修分类分布已展示，建议关注频繁小修的设备，可能预防性维护不足。';
    }
    case 37: { // 新设备故障率分析
      const early = rows.find(r => /0-30/.test(r[0]));
      if (early) {
        const total = rows.reduce((s, r) => s + num(r[1]), 0);
        const pct = (num(early[1]) / Math.max(1, total) * 100).toFixed(1);
        if (parseFloat(pct) > 30) return '0-30 天内故障占比 ' + pct + '%，偏高，说明安装调试或质量问题，建议优化新设备验收标准。';
        return '0-30 天内故障占比 ' + pct + '%，新设备故障率正常。';
      }
      return '无';
    }
    case 38: { // 设备健康度评分
      const top = rows[0];
      return '健康度评分最高的设备「' + top[1] + '」（评分 ' + top[5] + '），是优先关注和投入资源的目标，建议纳入设备分级管理。';
    }
    case 39: { // 等待 vs 维修时长占比
      const waitPct = findVal('等待占比');
      if (waitPct) {
        if (num(waitPct) > 30) return '等待占比 ' + waitPct + '%，超过 30%，派单流程需优化，建议缩短维修响应时间。';
        return '等待占比 ' + waitPct + '%，流程效率正常。';
      }
      return '无';
    }
    case 40: { // 班次维修效率对比
      if (rows.length < 2) return '无';
      const max = rows.reduce((m, r) => num(r[2]) > num(m[2]) ? r : m, rows[0]);
      const min = rows.reduce((m, r) => num(r[2]) < num(m[2]) ? r : m, rows[0]);
      const ratio = num(max[2]) / Math.max(0.01, num(min[2]));
      if (ratio > 1.5) return '班次效率差异较大，「' + max[0] + '」平均时长 ' + max[2] + 'h 是「' + min[0] + '」的 ' + ratio.toFixed(1) + ' 倍，建议优化某班技能或资源配置。';
      return '各班次维修效率基本一致。';
    }
    case 41: { // 维修人效率对比
      if (rows.length < 2) return '无';
      const max = rows.reduce((m, r) => num(r[3]) > num(m[3]) ? r : m, rows[0]);
      return '维修人「' + max[1] + '」平均维修时长最长（' + max[3] + 'h），建议针对性技能培训。';
    }
    case 42: { // 响应时间分析
      const totalMean = findVal('申请→开始(均值)');
      if (totalMean) return '从申请到维修开始的平均总时长为 ' + totalMean + 'h，建议关注耗时最长的阶段并针对性优化。';
      return '无';
    }
    case 43: { // 维修超时分析
      const p90 = findVal('P90阈值');
      const overP90 = findVal('超P90维修单数');
      if (p90 && overP90) return 'P90 阈值为 ' + p90 + 'h，超 P90 维修单 ' + overP90 + ' 条，建议分析超时维修单的共性设备或原因并建立超时预警机制。';
      return '无';
    }
    case 44: { // 驳回率分析
      const top = rows[0];
      const topRate = num(top[3]);
      if (topRate > 20) return '「' + top[0] + '」驳回率 ' + topRate.toFixed(1) + '%，偏高，可能需求描述不清或审批标准不一，建议统一维修申请标准并加强培训。';
      return '各车间驳回率整体正常，「' + top[0] + '」最高（' + topRate.toFixed(1) + '%）。';
    }
    case 45: { // 维修阶段时长分解
      const waitV = findVal('等待确认(均值)');
      const repairV = findVal('维修阶段(均值)');
      if (waitV && repairV) {
        const w = num(waitV), r = num(repairV);
        if (w > r) return '等待确认阶段均值 ' + w.toFixed(1) + 'h 大于维修阶段 ' + r.toFixed(1) + 'h，流程优化重点在缩短确认环节。';
        return '维修阶段均值 ' + r.toFixed(1) + 'h 是主要耗时环节，建议优化维修执行效率。';
      }
      return '无';
    }
    case 46: { // 及时性分析
      const meanV = findVal('均值');
      if (meanV) {
        const m = num(meanV);
        if (m < 3) return '及时性评分均值 ' + m.toFixed(2) + '，结单及时性良好。';
        return '及时性评分均值 ' + m.toFixed(2) + '，偏低，结单不及时，建议考核维修团队响应速度。';
      }
      return '无';
    }
    case 47: { // 维修成本按车间分布
      const top = rows[0];
      return '「' + top[0] + '」总成本最高（' + top[3] + '），建议重点控制该车间维修支出。';
    }
    case 48: { // 备件 vs 委外成本占比
      const partsPct = findVal('备件占比');
      const outPct = findVal('委外占比');
      if (outPct) {
        if (num(outPct) > 50) return '委外费用占比 ' + outPct + '%，超过备件费用，建议评估自修能力是否不足并增加维修技能培训。';
        return '备件占比 ' + partsPct + '%，委外占比 ' + outPct + '%，成本结构合理。';
      }
      return '无';
    }
    case 49: { // 产能影响量化
      const cnt = findVal('影响产能维修单数');
      if (cnt) return '影响产能的维修单 ' + cnt + ' 条，建议按车间分布优先保障高影响车间的设备可用性。';
      return '无';
    }
    case 50: { // 质量影响分析
      if (!rows.length) return '无影响质量的维修单。';
      const top = rows[0];
      return '共 ' + rows.length + ' 个车间有影响质量的维修单，「' + top[0] + '」最多（' + top[1] + ' 条），建议针对性改进。';
    }
    case 51: { // 成本趋势分析
      if (rows.length < 2) return '无';
      const first = num(rows[0][1]), last = num(rows[rows.length - 1][1]);
      const change = ((last - first) / Math.max(1, first) * 100).toFixed(1);
      if (parseFloat(change) > 30) return '维修总成本从 ' + first + ' 上升至 ' + last + '（+' + change + '%），需排查原因并加强月度成本监控。';
      return '维修成本趋势平稳，首末月变化 ' + change + '%。';
    }
    case 52: { // 高成本维修单识别
      const p95 = findVal('P95阈值');
      const cnt = findVal('超阈值维修单数');
      if (p95 && cnt) return 'P95 阈值为 ' + p95 + ' 元，超阈值维修单 ' + cnt + ' 条，建议审查这些异常高成本维修单并设定审批阈值。';
      return '无';
    }
    case 53: { // 成本 vs 故障等级
      if (rows.length < 2) return '无';
      const aRow = rows.find(r => /A/.test(r[0]));
      const cRow = rows.find(r => /C/.test(r[0]));
      if (aRow && cRow) {
        const ratio = num(aRow[2]) / Math.max(0.01, num(cRow[2]));
        if (ratio > 2) return 'A级故障平均成本 ' + aRow[2] + ' 是C级的 ' + ratio.toFixed(1) + ' 倍，高等级故障经济影响显著，建议加强A级故障预防。';
        return '各等级故障成本差异不大，A级/C级比值 ' + ratio.toFixed(1) + '。';
      }
      return '无';
    }
    case 54: { // 维修评分分布
      const minDim = rows.reduce((m, r) => num(r[2]) < num(m[2]) ? r : m, rows[0]);
      return '评分均值最低的维度为「' + minDim[0] + '」（' + minDim[2] + '），是该维度改进重点，建议制定针对性质量改进计划。';
    }
    case 55: { // 评分维度相关性
      return 'Pearson 相关系数矩阵已展示，建议关注与总分相关性最高的维度，该维度对总分影响最大，可优化评分权重分配。';
    }
    case 56: { // 维修人评分排名
      if (!rows.length) return '无';
      const bottom = rows[rows.length - 1];
      return '维修人「' + bottom[1] + '」平均总分最低（' + bottom[3] + '），建议针对性培训或调整岗位。';
    }
    case 57: { // 车间评分对比
      if (!rows.length) return '无';
      const bottom = rows[rows.length - 1];
      return '「' + bottom[0] + '」平均总分最低（' + bottom[2] + '），建议调查该车间维修质量原因。';
    }
    case 58: { // 低分维修单分析
      const cnt = findVal('低分维修单数');
      const pct = findVal('占总维修单比');
      if (cnt) return '低分维修单（<3分）' + cnt + ' 条，占总比 ' + pct + '，建议分析低分单的共性原因并制定质量改进专项计划。';
      return '无';
    }
    case 59: { // 评分趋势分析
      if (rows.length < 2) return '无';
      const first = num(rows[0][1]), last = num(rows[rows.length - 1][1]);
      const change = ((last - first) / Math.max(0.01, first) * 100).toFixed(1);
      if (parseFloat(change) < -10) return '维修评分从 ' + first + ' 下降至 ' + last + '（' + change + '%），质量下降趋势明显，建议调查原因。';
      return '维修评分趋势平稳，首末月变化 ' + change + '%。';
    }
    case 60: { // 维修时长 vs 故障等级
      if (rows.length < 2) return '无';
      const aRow = rows.find(r => /A/.test(r[0]));
      const cRow = rows.find(r => /C/.test(r[0]));
      if (aRow && cRow) {
        const ratio = num(aRow[2]) / Math.max(0.01, num(cRow[2]));
        if (ratio > 1.5) return 'A级故障平均维修时长 ' + aRow[2] + 'h 是C级的 ' + ratio.toFixed(1) + ' 倍，等级与时长关联较强，可用于维修工时预估。';
        return '各等级维修时长差异不大。';
      }
      return '无';
    }
    case 61: { // 等待时长 vs 维修评分
      const corr = findVal('Pearson相关系数');
      if (corr) {
        const c = num(corr);
        if (c < -0.3) return '等待时长与维修评分相关系数 ' + corr + '，呈负相关，等待越久评分越低，证明缩短等待时间的必要性。';
        return '等待时长与维修评分相关系数 ' + corr + '，相关性不显著。';
      }
      return '无';
    }
    case 62: { // 故障频率 vs 维修成本
      const corr = findVal('Pearson相关系数');
      if (corr) {
        const c = num(corr);
        if (c > 0.5) return '故障频率与维修成本相关系数 ' + corr + '，正相关显著，频率高的设备成本也高，建议关注高频设备的全生命周期成本。';
        return '故障频率与维修成本相关系数 ' + corr + '，线性关系不显著。';
      }
      return '无';
    }
    case 63: { // 班次 vs 故障类型
      return '班次×故障原因交叉表已展示，建议关注某班次是否集中某类故障，可能与操作习惯相关，可针对性开展班次操作培训。';
    }
    case 64: { // 设备类型 vs 故障模式
      return '设备分类×故障原因交叉表已展示，建议关注某类设备是否集中某类故障，用于设备选型和改进参考。';
    }
    case 65: { // 维修人数 vs 维修时长
      const corr = findVal('Pearson相关系数');
      if (corr) {
        const c = num(corr);
        if (c > -0.1 && c < 0.1) return '参与人数与维修时长相关系数 ' + corr + '，几乎无相关性，增加人数未缩短时长，说明协作效率低，建议优化维修人员配置。';
        if (c < -0.3) return '参与人数与维修时长相关系数 ' + corr + '，负相关，增加人数有助于缩短时长，但需注意边际效益。';
        return '参与人数与维修时长相关系数 ' + corr + '，存在一定正相关，人数增多反而时长增加，建议评估协作模式。';
      }
      return '无';
    }
    case 66: { // 维修时长预测
      const r2 = findVal('R²');
      if (r2) {
        const r = num(r2);
        if (r > 0.3) return '回归模型 R²=' + r2 + '，故障等级对维修时长有一定预测力，可用于维修工时预估和排程参考。';
        return '回归模型 R²=' + r2 + '，预测力较弱，故障等级对维修时长的解释力有限。';
      }
      return '无';
    }
    case 67: { // 维修成本预测
      const r2 = findVal('R²');
      if (r2) {
        const r = num(r2);
        if (r > 0.3) return '回归模型 R²=' + r2 + '，维修时长对成本有一定预测力，可用于维修成本预估和审批参考。';
        return '回归模型 R²=' + r2 + '，预测力较弱，维修时长对成本的解释力有限。';
      }
      return '无';
    }
    case 68: { // 评分影响因素回归
      return '各评分维度对总分的回归系数已展示，建议关注斜率最大的维度，该维度对总分影响最大，可优化评分权重或考核重点。';
    }
    case 69: { // 月度维修量预测
      const lastRow = rows[rows.length - 1];
      if (lastRow && /预测/.test(lastRow[0])) return '下月维修量预测值为 ' + lastRow[1] + ' 条，建议据此预分配维修资源。';
      return '预测数据不足。';
    }
    case 70: { // 设备故障概率预测
      const top = rows[0];
      if (top) return '设备「' + top[1] + '」下月预测故障数最高（' + top[5] + '），建议提前安排预防性维护。';
      return '无';
    }
    case 71: { // 问题描述词频分析
      const top = rows[0];
      return '最高频关键词为「' + top[1] + '」（' + top[2] + ' 次），代表最常见的故障现象，建议纳入故障知识库用于快速诊断。';
    }
    case 72: { // 故障原因描述词频
      const top = rows[0];
      return '最高频原因关键词为「' + top[1] + '」（' + top[2] + ' 次），集中解决该原因可降低维修量。';
    }
    case 73: { // 维修内容词频
      const top = rows[0];
      return '最高频维修操作词为「' + top[1] + '」（' + top[2] + ' 次），建议据此制定标准化维修SOP。';
    }
    case 74: { // 高频故障模式提取
      const top = rows[0];
      return '最高频故障模式为「' + top[1] + '」（' + top[2] + ' 次），同一模式反复出现说明系统性问题未解决，建议开展根因消除。';
    }
    case 75: { // 多维异常检测
      const cnt = findVal('异常维修单总数');
      const multi = findVal('多维度异常');
      if (cnt) return '检出异常维修单 ' + cnt + ' 条，其中多维度异常 ' + (multi || 0) + ' 条，建议重点审查这些在时长/等待/成本上同时偏离正常范围的维修单。';
      return '无';
    }
    case 76: { // 多维交叉透视
      return '车间×故障等级热力图已展示，建议关注颜色最深的格子，表示该车间该等级故障频次最高，用于识别组合风险。';
    }
    case 77: { // 关联规则挖掘
      if (!rows.length) return '未发现满足最小支持度的关联规则。';
      const top = rows[0];
      return '提升度最高的规则为「' + top[1] + ' → ' + top[2] + '」（Lift=' + top[6] + '），表示两个因素同时出现概率高于随机，建议关注该关联因素。';
    }
    case 78: { // 季节性分解
      if (rows.length < 6) return '数据不足，无法进行季节性分解。';
      return '趋势/季节性/残差三分解已展示，建议关注季节性指数大于 0 的月份，这些是维修高峰期，可据此制定季节性预防性维护计划。';
    }
    case 79: { // 设备K-Means聚类
      const highFreq = rows.find(r => /高频高耗/.test(r[0]));
      if (highFreq) return '「高频高耗」聚类包含 ' + highFreq[1] + ' 台设备，平均频次 ' + highFreq[2] + '，是重点关注设备群，建议优先分配维护资源。';
      return '设备聚类已展示，建议关注高频高耗聚类的设备，进行分级管理。';
    }
    case 80: { // ARIMA维修量预测
      const lastRow = rows[rows.length - 1];
      if (lastRow && /预测/.test(lastRow[0])) return 'ARIMA 模型预测下月维修量为 ' + lastRow[2] + ' 条，建议据此进行中长期维修资源规划。';
      return '数据不足以拟合 ARIMA 模型。';
    }
    case 81: { // TF-IDF关键词分析
      const top = rows[0];
      return 'TF-IDF 权重最高的关键词为「' + top[1] + '」（权重 ' + top[3] + '），该词在某类维修中特别突出，建议用于故障分类和知识库标签建设。';
    }
    default: return '无';
  }
}

function executeMethod(method) {
  const resultEl = $('#execResult');
  resultEl.innerHTML = '<p class="exec-running">分析中…</p>';
  setTimeout(() => {
    try {
      const filteredRows = getFilteredRows(S.rows, S.analysisFilter);
      const html = method.fn(filteredRows);
      resultEl.innerHTML = html;
      const scripts = resultEl.querySelectorAll('script');
      scripts.forEach(oldScript => {
        const newScript = document.createElement('script');
        newScript.textContent = oldScript.textContent;
        oldScript.parentNode.replaceChild(newScript, oldScript);
      });
      const tables = extractTables(html);
      const conclusion = generateConclusion(method.id, tables);
      const concEl = $('#execConclusion');
      if (concEl) concEl.textContent = '分析结论：' + conclusion;
    } catch (err) {
      console.error(err);
      resultEl.innerHTML = '<p class="exec-error">分析出错: ' + esc(err.message) + '</p>';
    }
  }, 50);
}

/* ---- 水平条形图 ---- */
function renderHBarChart(selector, counts, variant) {
  const el = $(selector);
  if (!el) return;
  const entries = Object.entries(counts).sort((a, b) => b[1] - a[1]);
  if (!entries.length) { el.innerHTML = '<p style="color:var(--ink-faded)">暂无数据</p>'; return; }
  const max = entries[0][1];
  const total = entries.reduce((s, [, v]) => s + v, 0);
  const fillClass = variant ? 'h-bar-fill ' + variant : 'h-bar-fill';
  let html = '<div class="h-bar-list">';
  entries.forEach(([label, val]) => {
    const pct = (val / max * 100).toFixed(1);
    const share = (val / total * 100).toFixed(1);
    html += '<div class="h-bar-row">' +
      '<span class="h-bar-label" title="' + esc(label) + '">' + esc(label) + '</span>' +
      '<div class="h-bar-track"><div class="' + fillClass + '" style="width:' + pct + '%"></div></div>' +
      '<span class="h-bar-val">' + val + ' (' + share + '%)</span>' +
      '</div>';
  });
  html += '</div>';
  el.innerHTML = html;
}

/* ---- 数据预览表 ---- */
function renderPreview() {
  const table = $('#previewTable');
  const previewCols = ['A', 'B', 'F', 'G', 'H', 'L', 'M', 'P', 'Q', 'S', 'T', 'W', 'AD', 'AK', 'AX'];
  let html = '<thead><tr>';
  previewCols.forEach(L => {
    const col = COL_MAP[L];
    html += '<th>' + L + ' ' + (col ? col.cn : L) + '</th>';
  });
  html += '</tr></thead><tbody>';
  S.rows.slice(0, 20).forEach(row => {
    html += '<tr>';
    previewCols.forEach(L => {
      let v = row[L];
      if (v instanceof Date) v = fmtDateTime(v);
      html += '<td title="' + esc(v) + '">' + esc(v) + '</td>';
    });
    html += '</tr>';
  });
  html += '</tbody>';
  table.innerHTML = html;
}

/* ---- 数据筛选器 ---- */
function getFilteredRows(rows, filter) {
  if (!filter.months.length && !filter.workshop) return rows;
  return rows.filter(row => {
    if (filter.months.length) {
      const d = parseDateVal(row['Y']);
      if (!d) return false;
      const mk = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
      if (!filter.months.includes(mk)) return false;
    }
    if (filter.workshop) {
      const v = row['F'];
      if (v == null || String(v).trim() !== filter.workshop) return false;
    }
    return true;
  });
}

function getAvailableMonths() {
  const set = new Set();
  S.rows.forEach(row => {
    const d = parseDateVal(row['Y']);
    if (d) set.add(d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0'));
  });
  return [...set].sort();
}

function getAvailableWorkshops() {
  const set = new Set();
  S.rows.forEach(row => {
    const v = row['F'];
    if (v != null && String(v).trim()) set.add(String(v).trim());
  });
  const priority = ['宁夏一车间', '宁夏二车间', '宁夏三车间', '宁夏六车间', '宁夏七车间', '宁夏八车间', '宁夏十车间'];
  return [...set].sort((a, b) => {
    const ia = priority.indexOf(a), ib = priority.indexOf(b);
    if (ia !== -1 && ib !== -1) return ia - ib;
    if (ia !== -1) return -1;
    if (ib !== -1) return 1;
    return a.localeCompare(b, 'zh');
  });
}

function initDataFilters() {
  const months = getAvailableMonths();
  const workshops = getAvailableWorkshops();
  initMonthMultiSelect('af', months);
  initWorkshopSelect('af', workshops);
  initMonthMultiSelect('ef', months);
  initWorkshopSelect('ef', workshops);
  S.analysisFilter = { months: [], workshop: '' };
  S.exportFilter = { months: [], workshop: '' };
  updateFilterCount('af');
}

function initMonthMultiSelect(prefix, months) {
  const panel = $('#' + prefix + 'MonthPanel');
  const btn = $('#' + prefix + 'MonthBtn');
  const wrap = $('#' + prefix + 'MonthWrap');
  panel.innerHTML = '<label class="df-all"><input type="checkbox" value="__all__" data-prefix="' + prefix + '"> 全部</label>' +
    months.map(m =>
      '<label><input type="checkbox" value="' + m + '" data-prefix="' + prefix + '"> ' + m + '</label>'
    ).join('');
  btn.addEventListener('click', e => {
    e.stopPropagation();
    panel.classList.toggle('open');
  });
  panel.addEventListener('change', e => {
    if (e.target.type === 'checkbox') {
      const p = e.target.dataset.prefix;
      if (e.target.value === '__all__') {
        const cbs = panel.querySelectorAll('input[type=checkbox]:not([value=__all__])');
        cbs.forEach(cb => { cb.checked = e.target.checked; });
      } else {
        const allCb = panel.querySelector('input[value=__all__]');
        const monthCbs = panel.querySelectorAll('input[type=checkbox]:not([value=__all__])');
        if (allCb) allCb.checked = [...monthCbs].every(cb => cb.checked);
      }
      const selected = [...panel.querySelectorAll('input:checked:not([value=__all__])')].map(cb => cb.value);
      const filter = p === 'af' ? S.analysisFilter : S.exportFilter;
      filter.months = selected;
      updateMonthBtnText(p);
      if (p === 'af') {
        updateFilterCount('af');
        if (_selectedMethodId) {
          const method = ANALYSIS_METHODS.find(m => m.id === _selectedMethodId);
          if (method) executeMethod(method);
        }
      }
    }
  });
  document.addEventListener('click', e => {
    if (!wrap.contains(e.target)) panel.classList.remove('open');
  });
}

function initWorkshopSelect(prefix, workshops) {
  const sel = $('#' + prefix + 'Workshop');
  sel.innerHTML = '<option value="">全部车间</option>' + workshops.map(w =>
    '<option value="' + esc(w) + '">' + esc(w) + '</option>'
  ).join('');
  sel.addEventListener('change', () => {
    const filter = prefix === 'af' ? S.analysisFilter : S.exportFilter;
    filter.workshop = sel.value;
    if (prefix === 'af') {
      updateFilterCount('af');
      if (_selectedMethodId) {
        const method = ANALYSIS_METHODS.find(m => m.id === _selectedMethodId);
        if (method) executeMethod(method);
      }
    }
  });
}

function updateMonthBtnText(prefix) {
  const panel = $('#' + prefix + 'MonthPanel');
  const btn = $('#' + prefix + 'MonthBtn');
  const checked = panel.querySelectorAll('input:checked:not([value=__all__])');
  const total = panel.querySelectorAll('input[type=checkbox]:not([value=__all__])');
  if (checked.length === 0 || checked.length === total.length) {
    btn.textContent = '全部年月';
  } else if (checked.length <= 3) {
    btn.textContent = [...checked].map(cb => cb.value).join(', ');
  } else {
    btn.textContent = '已选 ' + checked.length + ' 个月';
  }
}

function updateFilterCount(prefix) {
  if (prefix !== 'af') return;
  const el = $('#afCount');
  if (!el) return;
  const filtered = getFilteredRows(S.rows, S.analysisFilter);
  el.textContent = '筛选后：' + filtered.length + ' 条';
}

function syncFilterUI(prefix) {
  const filter = prefix === 'af' ? S.analysisFilter : S.exportFilter;
  const panel = $('#' + prefix + 'MonthPanel');
  const monthCbs = panel.querySelectorAll('input[type=checkbox]:not([value=__all__])');
  monthCbs.forEach(cb => { cb.checked = filter.months.includes(cb.value); });
  const allCb = panel.querySelector('input[value=__all__]');
  if (allCb) allCb.checked = filter.months.length === 0 || filter.months.length === monthCbs.length;
  updateMonthBtnText(prefix);
  $('#' + prefix + 'Workshop').value = filter.workshop;
}

/* ---- 导出 ---- */
function openExport() {
  S.exportFilter = { months: S.analysisFilter.months.slice(), workshop: S.analysisFilter.workshop };
  syncFilterUI('ef');
  $('#exportModal').classList.add('open');
  $('#exportModal').setAttribute('aria-hidden', 'false');
}
function closeExport() {
  $('#exportModal').classList.remove('open');
  $('#exportModal').setAttribute('aria-hidden', 'true');
}

function download(blob, name) {
  const a = document.createElement('a');
  const u = URL.createObjectURL(blob);
  a.href = u;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(u); a.remove(); }, 60000);
}

async function doExport() {
  closeExport();
  const fmt = document.querySelector('input[name="exportFmt"]:checked').value;
  const methods = ANALYSIS_METHODS;
  if (!methods.length) { toast('没有符合条件的方法'); return; }
  const ts = new Date();
  const pad = n => String(n).padStart(2, '0');
  const baseName = 'Maintenance_Data_Analysis_' + ts.getFullYear() + pad(ts.getMonth() + 1) + pad(ts.getDate()) + pad(ts.getHours()) + pad(ts.getMinutes());

  const doPdf = fmt === 'pdf' || fmt === 'both';
  const doExcel = fmt === 'excel' || fmt === 'both';

  doExportAnalysis(methods, fmt, baseName).catch(e => toast('导出失败: ' + e.message));
}

function sanitizeSheetName(name) {
  let n = name.replace(/[\\\/\?\*\[\]:]/g, '').slice(0, 31);
  return n || 'Sheet';
}

function extractTableData(html) {
  const tmp = document.createElement('div');
  tmp.innerHTML = html;
  const tables = tmp.querySelectorAll('table.result-table');
  if (!tables.length) return null;
  const table = tables[0];
  const headers = [...table.querySelectorAll('thead th')].map(th => th.textContent.trim());
  const rows = [...table.querySelectorAll('tbody tr')].map(tr =>
    [...tr.querySelectorAll('td')].map(td => td.textContent.trim())
  );
  return { headers, rows };
}

function renderAnalysisForExport(method, container) {
  container.innerHTML = '';
  const wrap = document.createElement('div');
  wrap.style.cssText = 'padding:20px;background:#f3ecdc;width:800px;';
  const title = document.createElement('h3');
  title.textContent = method.name + ' - ' + method.desc;
  title.style.cssText = 'font-family:serif;margin-bottom:.3rem;font-size:1.1rem;';
  wrap.appendChild(title);
  const result = document.createElement('div');
  result.className = 'exec-result';
  wrap.appendChild(result);
  container.appendChild(wrap);

  let html;
  try {
    html = method.fn(getFilteredRows(S.rows, S.exportFilter));
  } catch (err) {
    html = '<p class="exec-error">分析出错: ' + esc(err.message) + '</p>';
  }
  result.innerHTML = html;

  const scripts = result.querySelectorAll('script');
  scripts.forEach(oldScript => {
    const newScript = document.createElement('script');
    newScript.textContent = oldScript.textContent;
    oldScript.parentNode.replaceChild(newScript, oldScript);
  });

  const tables = extractTables(html);
  const conclusion = generateConclusion(method.id, tables);
  const concEl = document.createElement('p');
  concEl.style.cssText = 'font-size:.74rem;color:#646464;font-family:monospace;margin:.4rem 0 .6rem;padding:.4rem .6rem;background:#f6eed6;border-left:3px solid #b1281e;line-height:1.5;';
  concEl.textContent = '分析结论：' + conclusion;
  result.insertBefore(concEl, result.firstChild);

  return { wrap, result, conclusion };
}

function getChartImages(container) {
  const canvases = container.querySelectorAll('canvas');
  const images = [];
  const _doughnutLabelPlugin = {
    id: 'doughnutLabels',
    afterDatasetsDraw(chart) {
      if (chart.config.type !== 'doughnut' && chart.config.type !== 'pie') return;
      const meta = chart.getDatasetMeta(0);
      const total = chart.data.datasets[0].data.reduce((s, v) => s + v, 0);
      meta.data.forEach((arc, i) => {
        const val = chart.data.datasets[0].data[i];
        if (val == null || total === 0) return;
        const pct = (val / total * 100).toFixed(1);
        const angle = arc.endAngle - arc.startAngle;
        if (angle < 0.25) return;
        const midAngle = (arc.startAngle + arc.endAngle) / 2;
        const r = (arc.innerRadius + arc.outerRadius) / 2;
        const cx = arc.x + Math.cos(midAngle) * r;
        const cy = arc.y + Math.sin(midAngle) * r;
        const label = val + ' (' + pct + '%)';
        const ctx = chart.ctx;
        ctx.save();
        ctx.font = '11px "JetBrains Mono", monospace';
        const tw = ctx.measureText(label).width;
        if (tw > (arc.outerRadius - arc.innerRadius) * 2 * 0.8) {
          ctx.font = '9px "JetBrains Mono", monospace';
          const tw2 = ctx.measureText(label).width;
          if (tw2 > (arc.outerRadius - arc.innerRadius) * 2 * 0.8) { ctx.restore(); return; }
        }
        const bg = chart.data.datasets[0].backgroundColor[i];
        const hex = typeof bg === 'string' ? bg : '#1d4e89';
        const lum = parseInt(hex.slice(1, 3), 16) * 0.299 + parseInt(hex.slice(3, 5), 16) * 0.587 + parseInt(hex.slice(5, 7), 16) * 0.114;
        ctx.fillStyle = lum < 140 ? '#ffffff' : '#3a2f25';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(label, cx, cy);
        ctx.restore();
      });
    }
  };
  canvases.forEach(canvas => {
    const chart = Chart.getChart(canvas);
    if (!chart) return;
    const config = chart.config;
    const data = JSON.parse(JSON.stringify(chart.data));
    const opts = JSON.parse(JSON.stringify(chart.options));
    opts.animation = false;
    opts.responsive = false;
    opts.maintainAspectRatio = false;
    chart.destroy();
    const newCanvas = document.createElement('canvas');
    newCanvas.width = 760;
    newCanvas.height = 360;
    const plugins = (config.plugins || []).slice();
    if (config.type === 'doughnut' || config.type === 'pie') {
      plugins.push(_doughnutLabelPlugin);
    }
    if (config.type === 'bar') {
      plugins.push(Charts._barLabelPlugin());
    }
    const newChart = new Chart(newCanvas.getContext('2d'), {
      type: config.type,
      data: data,
      options: opts,
      plugins: plugins
    });
    const cw = newCanvas.width;
    const ch = newCanvas.height;
    const tmp = document.createElement('canvas');
    tmp.width = cw;
    tmp.height = ch;
    const ctx = tmp.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, cw, ch);
    ctx.drawImage(newCanvas, 0, 0);
    images.push({ data: tmp.toDataURL('image/jpeg', 0.9), w: 760, h: 360 });
    newChart.destroy();
  });
  return images;
}

function getTextContent(container) {
  const texts = [];
  container.querySelectorAll('p.analysis-info, p.no-data, p.exec-error, h4').forEach(el => {
    texts.push({ type: el.tagName.toLowerCase(), text: el.textContent.trim() });
  });
  return texts;
}

function getTableData(container) {
  const tables = container.querySelectorAll('table.result-table');
  const result = [];
  tables.forEach(table => {
    const headers = [...table.querySelectorAll('thead th')].map(th => th.textContent.trim());
    const rows = [...table.querySelectorAll('tbody tr')].map(tr =>
      [...tr.querySelectorAll('td')].map(td => td.textContent.trim())
    );
    result.push({ headers, rows });
  });
  return result;
}

async function doExportAnalysis(methods, fmt, baseName) {
  const exportArea = document.createElement('div');
  exportArea.style.cssText = 'position:fixed;left:-9999px;top:0;width:800px;';
  document.body.appendChild(exportArea);

  const doPdf = fmt === 'pdf' || fmt === 'both';
  const doExcel = fmt === 'excel' || fmt === 'both';

  let pdfPages = [];
  let excelSheets = [];

  const onBeforeUnload = e => { e.preventDefault(); e.returnValue = '正在导出，确定离开？'; };
  window.addEventListener('beforeunload', onBeforeUnload);

  for (let i = 0; i < methods.length; i++) {
    const method = methods[i];
    toast('正在导出 ' + (i + 1) + '/' + methods.length + '：' + method.name, true);

    try {
      const { result, conclusion } = renderAnalysisForExport(method, exportArea);
      await new Promise(r => requestAnimationFrame(r));

      const images = getChartImages(result);
      const texts = getTextContent(result);
      const tables = getTableData(result);

      if (doPdf) {
        pdfPages.push({ name: method.name, fields: '涉及字段: ' + fmtFields(method.fields), guide: method.guide ? '分析方法：' + method.guide : '', desc: '分析结论：' + conclusion, images, texts, tables });
      }

      if (doExcel) {
        const infoRows = [];
        infoRows.push(['涉及字段：' + fmtFields(method.fields)]);
        if (method.guide) infoRows.push(['分析方法：' + method.guide]);
        infoRows.push(['分析结论：' + conclusion]);
        if (tables.length) {
          const t = tables[0];
          const padded = infoRows.map(r => { while (r.length < t.headers.length) r.push(''); return r; });
          excelSheets.push({ name: sanitizeSheetName(method.name), headers: t.headers, rows: [...padded, ...t.rows], images });
        } else {
          excelSheets.push({ name: sanitizeSheetName(method.name), headers: ['说明'], rows: [...infoRows, ['该分析无表格数据，请查看 PDF 导出']], images });
        }
      }
    } catch (err) {
      console.error('导出失败: ' + method.name + ' - ' + method.desc, err);
      if (doExcel) {
        excelSheets.push({ name: sanitizeSheetName(method.name), headers: ['错误'], rows: [['分析出错: ' + err.message]], images: [] });
      }
      if (doPdf) {
        pdfPages.push({ name: method.name, desc: '', images: [], texts: [{ type: 'p', text: '分析出错: ' + err.message }], tables: [] });
      }
    }

    Charts.destroyAll();
  }

  window.removeEventListener('beforeunload', onBeforeUnload);
  document.body.removeChild(exportArea);

  // 智能分析建议作为第二页
  if (_lastSuggestions.length) {
    const sevLabel = { high: '🔴', medium: '🟡', low: '🟢' };
    const sorted = _lastSuggestions.slice().sort((a, b) => {
      const order = { high: 0, medium: 1, low: 2 };
      return (order[a.severity] ?? 3) - (order[b.severity] ?? 3);
    });
    if (doPdf) {
      const sugTexts = [];
      sorted.forEach(s => {
        sugTexts.push({ type: 'p', text: sevLabel[s.severity] + ' ' + s.text });
      });
      pdfPages.splice(0, 0, { name: '智能分析建议（' + sorted.length + ' 条）', desc: '', images: [], texts: sugTexts, tables: [] });
    }
    if (doExcel) {
      const sugRows = sorted.map(s => [sevLabel[s.severity] || '⚪', s.text, (s.methods || []).join('、')]);
      excelSheets.splice(0, 0, { name: '智能分析建议', headers: ['级别', '建议内容', '分析方法'], rows: sugRows, images: [] });
    }
  }

  if (doPdf) {
    await yieldUI();
    await exportPDF(pdfPages, baseName);
  }
  if (doExcel) {
    await yieldUI();
    await exportExcel(excelSheets, baseName);
  }

  toastResult('导出完成：共 ' + methods.length + ' 个分析', null);
}

const PDF_DPI = 200;
const PDF_DPI_MM = PDF_DPI / 25.4;

function generateCoverImage() {
  const w = Math.ceil(210 * PDF_DPI_MM);
  const h = Math.ceil(297 * PDF_DPI_MM);
  const canvas = document.createElement('canvas');
  canvas.width = w;
  canvas.height = h;
  const ctx = canvas.getContext('2d');

  const paper = '#f3ecdc', paper2 = '#ece3cf', ink = '#1a1410', inkSoft = '#3a2f25', inkFaded = '#6b6258';
  const blueInk = '#1d4e89', vermillion = '#b1281e', rule = '#b8a98a', ruleSoft = '#d6c9ad';

  ctx.fillStyle = paper;
  ctx.fillRect(0, 0, w, h);

  const grad = ctx.createLinearGradient(0, 0, 0, h);
  grad.addColorStop(0, '#faf5e8');
  grad.addColorStop(0.4, paper);
  grad.addColorStop(1, paper2);
  ctx.fillStyle = grad;
  ctx.fillRect(0, 0, w, h);

  const px = mm => mm * PDF_DPI_MM;

  ctx.strokeStyle = ink;
  ctx.lineWidth = px(1.5);
  ctx.strokeRect(px(12), px(12), px(186), px(273));

  ctx.strokeStyle = rule;
  ctx.lineWidth = px(0.5);
  ctx.strokeRect(px(15), px(15), px(180), px(267));

  ctx.strokeStyle = blueInk;
  ctx.lineWidth = px(3);
  ctx.beginPath();
  ctx.moveTo(px(15), px(35));
  ctx.lineTo(px(195), px(35));
  ctx.stroke();

  ctx.strokeStyle = ink;
  ctx.lineWidth = px(0.8);
  ctx.beginPath();
  ctx.moveTo(px(15), px(37));
  ctx.lineTo(px(195), px(37));
  ctx.stroke();

  const stampY = px(24);
  ctx.font = 'italic ' + px(3.5) + 'px "EB Garamond", serif';
  ctx.fillStyle = inkFaded;
  ctx.textAlign = 'right';
  ctx.textBaseline = 'middle';
  const now = new Date();
  const dateStr = now.getFullYear() + '.' + String(now.getMonth() + 1).padStart(2, '0') + '.' + String(now.getDate()).padStart(2, '0');
  ctx.fillText(dateStr, px(190), stampY);

  const titleY = px(110);
  ctx.font = '900 ' + px(11) + 'px "Noto Serif SC", serif';
  ctx.fillStyle = ink;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  const titleChars = '维修数据分析';
  const titleSpacing = px(14);
  let titleStartX = w / 2 - (titleChars.length * titleSpacing) / 2 + titleSpacing / 2;
  for (let i = 0; i < titleChars.length; i++) {
    ctx.fillText(titleChars[i], titleStartX + i * titleSpacing, titleY);
  }

  ctx.strokeStyle = ink;
  ctx.lineWidth = px(0.6);
  ctx.beginPath();
  ctx.moveTo(px(60), titleY + px(10));
  ctx.lineTo(px(150), titleY + px(10));
  ctx.stroke();

  ctx.strokeStyle = ink;
  ctx.lineWidth = px(0.3);
  ctx.beginPath();
  ctx.moveTo(px(60), titleY + px(12));
  ctx.lineTo(px(150), titleY + px(12));
  ctx.stroke();

  ctx.font = 'italic 500 ' + px(5) + 'px "EB Garamond", serif';
  ctx.fillStyle = inkFaded;
  ctx.textAlign = 'center';
  ctx.fillText('M A I N T E N A N C E   D A T A   A N A L Y S I S', w / 2, titleY + px(20));

  ctx.font = px(3.8) + 'px "Noto Serif SC", serif';
  ctx.fillStyle = inkSoft;
  const reportLine = '维修数据多维度分析报告';
  ctx.fillText(reportLine, w / 2, titleY + px(35));

  ctx.font = px(3.2) + 'px "Noto Serif SC", serif';
  ctx.fillStyle = inkFaded;
  ctx.fillText('共 81 种分析方法 · 涵盖描述性统计、帕累托、趋势、可靠性、效率、成本等维度', w / 2, titleY + px(45));

  const infoY = px(200);
  ctx.strokeStyle = ruleSoft;
  ctx.lineWidth = px(0.5);
  ctx.beginPath();
  ctx.moveTo(px(40), infoY);
  ctx.lineTo(px(170), infoY);
  ctx.stroke();

  ctx.font = px(3) + 'px "Noto Serif SC", serif';
  ctx.fillStyle = inkSoft;
  ctx.textAlign = 'left';
  const infoItems = [
    '数据来源：' + (S.fileName || '—'),
    '记录总数：' + S.rows.length + ' 条',
    '分析方法：' + ANALYSIS_METHODS.length + ' 种',
  ];
  const filterDesc = [];
  if (S.exportFilter.months.length) filterDesc.push('年月：' + (S.exportFilter.months.length <= 3 ? S.exportFilter.months.join(', ') : S.exportFilter.months.length + ' 个月'));
  if (S.exportFilter.workshop) filterDesc.push('车间：' + S.exportFilter.workshop);
  if (filterDesc.length) infoItems.push('数据筛选：' + filterDesc.join(' · '));
  infoItems.push('导出日期：' + dateStr);

  infoItems.forEach((line, i) => {
    const y = infoY + px(8) + i * px(7);
    ctx.fillText(line, px(40), y);
  });

  ctx.font = 'italic ' + px(2.5) + 'px "EB Garamond", serif';
  ctx.fillStyle = inkFaded;
  ctx.fillText('Maintenance Data Analysis · v1.0', w / 2, px(278));

  return canvas.toDataURL('image/jpeg', 0.92);
}

function addText(pdf, text, x, yMm, fontSizeMm, color, maxWidthMm) {
  const fontSizePx = fontSizeMm * PDF_DPI_MM;
  const maxWidthPx = maxWidthMm ? maxWidthMm * PDF_DPI_MM : null;
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d');
  const fontStr = fontSizePx + 'px "Noto Serif SC", serif';
  ctx.font = fontStr;
  const lines = [];
  if (maxWidthPx) {
    let cur = '';
    for (const ch of String(text)) {
      if (ctx.measureText(cur + ch).width > maxWidthPx && cur) { lines.push(cur); cur = ch; }
      else cur += ch;
    }
    if (cur) lines.push(cur);
  } else { lines.push(String(text)); }
  const lineHPx = fontSizePx * 1.35;
  const maxW = Math.max(0, ...lines.map(l => ctx.measureText(l).width));
  canvas.width = Math.max(2, Math.ceil(maxW + 2));
  canvas.height = Math.max(2, Math.ceil(lineHPx * lines.length));
  ctx.font = fontStr;
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.fillStyle = color || '#000000';
  ctx.textBaseline = 'top';
  lines.forEach((line, i) => ctx.fillText(line, 0, i * lineHPx));
  const wMm = canvas.width / PDF_DPI_MM;
  const hMm = canvas.height / PDF_DPI_MM;
  pdf.addImage(canvas.toDataURL('image/jpeg', 0.85), 'JPEG', x, yMm, wMm, hMm, undefined, 'FAST');
  return hMm;
}

function addTable(pdf, table, x, yMm, contentWMm, pageHMm, margin) {
  const fontSizeMm = 3.2;
  const fontSizePx = fontSizeMm * PDF_DPI_MM;
  const headerHPx = fontSizePx * 1.5;
  const rowHPx = fontSizePx * 1.4;
  if (!table.headers.length) return 0;
  const colW = contentWMm / table.headers.length;
  const colWPx = colW * PDF_DPI_MM;
  const maxRowsPerImg = Math.floor((pageHMm - margin * 2) * PDF_DPI_MM / rowHPx) - 1;
  let y = yMm;
  let rowIdx = 0;
  while (rowIdx < table.rows.length || rowIdx === 0) {
    const rowsInChunk = Math.min(maxRowsPerImg, table.rows.length - rowIdx);
    const totalRows = (rowIdx === 0 ? 1 : 0) + Math.max(rowsInChunk, 0);
    const canvasH = headerHPx + rowsInChunk * rowHPx + 4;
    const canvas = document.createElement('canvas');
    canvas.width = Math.ceil(contentWMm * PDF_DPI_MM);
    canvas.height = Math.ceil(canvasH);
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.font = fontSizePx + 'px "Noto Serif SC", serif';
    ctx.textBaseline = 'top';
    if (rowIdx === 0) {
      ctx.fillStyle = '#dcd2be';
      ctx.fillRect(0, 0, canvas.width, headerHPx);
      ctx.fillStyle = '#000000';
      table.headers.forEach((h, ci) => {
        ctx.fillText(String(h).slice(0, 20), ci * colWPx + 2, 2);
      });
    }
    for (let r = 0; r < rowsInChunk; r++) {
      const row = table.rows[rowIdx + r];
      const ry = (rowIdx === 0 ? headerHPx : 0) + r * rowHPx;
      row.forEach((cell, ci) => {
        ctx.fillText(String(cell).slice(0, 25), ci * colWPx + 2, ry + 2);
      });
    }
    const hMm = canvas.height / PDF_DPI_MM;
    pdf.addImage(canvas.toDataURL('image/jpeg', 0.85), 'JPEG', x, y, contentWMm, hMm, undefined, 'FAST');
    y += hMm + 2;
    rowIdx += rowsInChunk;
    if (rowIdx < table.rows.length) { pdf.addPage(); y = margin; }
    if (rowIdx === 0 && table.rows.length === 0) break;
  }
  return y - yMm;
}

async function exportPDF(pages, baseName) {
  if (!pages.length) return;
  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF('p', 'mm', 'a4');
  const pageW = 210, pageH = 297, margin = 10;
  const contentW = pageW - margin * 2;
  let y = margin;
  const pageStarts = [];

  toast('PDF：正在生成封面…', true);
  await yieldUI();
  const coverData = generateCoverImage();
  pdf.addImage(coverData, 'JPEG', 0, 0, pageW, pageH, undefined, 'FAST');

  for (let i = 0; i < pages.length; i++) {
    const page = pages[i];
    toast('PDF ' + (i + 1) + '/' + pages.length + '：' + page.name + '（' + page.images.length + ' 图, ' + page.tables.reduce((s,t) => s + t.rows.length, 0) + ' 行表）', true);
    await yieldUI();
    pdf.addPage(); y = margin;
    pageStarts.push({ name: page.name, pageNum: pdf.getNumberOfPages() });

    y += addText(pdf, page.name, margin, y, 5) + 2;
    if (page.fields) y += addText(pdf, page.fields, margin, y, 3.2, '#646464', contentW) + 1;
    if (page.guide) y += addText(pdf, page.guide, margin, y, 3.2, '#646464', contentW) + 1;
    if (page.desc) y += addText(pdf, page.desc, margin, y, 3.2, '#646464', contentW) + 2;

    page.texts.forEach(t => {
      if (t.type === 'h4') {
        if (y > pageH - margin - 10) { pdf.addPage(); y = margin; }
        y += addText(pdf, t.text, margin, y, 4) + 2;
      } else {
        if (y > pageH - margin - 8) { pdf.addPage(); y = margin; }
        y += addText(pdf, t.text, margin, y, 3.2, '#000000', contentW) + 1;
      }
    });

    for (let j = 0; j < page.images.length; j++) {
      const img = page.images[j];
      const imgW = contentW;
      const ratio = img.h / img.w;
      let imgH = imgW * ratio;
      if (imgH > pageH - margin * 2) imgH = pageH - margin * 2;
      if (y + imgH > pageH - margin) { pdf.addPage(); y = margin; }
      toast('PDF ' + (i + 1) + '/' + pages.length + '：' + page.name + ' — 写入图片 ' + (j + 1) + '/' + page.images.length, true);
      try { pdf.addImage(img.data, 'JPEG', margin, y, imgW, imgH, undefined, 'FAST'); } catch(e) {}
      y += imgH + 4;
      await yieldUI();
    }

    page.tables.forEach(table => {
      if (table.headers.length === 0) return;
      if (y > pageH - margin - 20) { pdf.addPage(); y = margin; }
      const tableH = addTable(pdf, table, margin, y, contentW, pageH, margin);
      y += tableH + 4;
    });

    await yieldUI();
  }

  toast('PDF：正在序列化 ' + pages.length + ' 页…', true);
  await yieldUI();

  pageStarts.forEach(ps => {
    pdf.outline.add(null, ps.name, { pageNumber: ps.pageNum });
  });

  const pdfBlob = pdf.output('blob');
  toast('PDF：正在写入文件（' + (pdfBlob.size / 1024 / 1024).toFixed(1) + ' MB）…', true);
  await yieldUI();
  download(pdfBlob, baseName + '.pdf');
}

async function exportExcel(sheets, baseName) {
  if (!sheets.length) return;
  const wb = new ExcelJS.Workbook();

  toast('Excel：正在生成封面…', true);
  await yieldUI();
  const coverData = generateCoverImage();
  const coverWs = wb.addWorksheet('封面');
  coverWs.columns = [{ width: 12 }];
  for (let r = 0; r < 40; r++) coverWs.mergeCells(r + 1, 1, r + 1, 12);
  const coverImgId = wb.addImage({
    buffer: base64ToUint8Array(coverData.split(',')[1]),
    extension: 'jpeg',
  });
  coverWs.addImage(coverImgId, {
    tl: { col: 0, row: 0 },
    ext: { width: 800, height: Math.round(800 * 297 / 210) },
  });

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    toast('Excel ' + (i + 1) + '/' + sheets.length + '：' + sheet.name, true);
    const ws = wb.addWorksheet(sheet.name);

    if (i === 0 && sheets.length > 1) {
      const tocWs = wb.addWorksheet('目录');
      tocWs.columns = [{ width: 6 }, { width: 50 }];
      tocWs.addRow(['序号', '分析方法']);
      tocWs.getRow(1).font = { bold: true };
      for (let j = 0; j < sheets.length; j++) {
        const row = tocWs.addRow([j + 1, '']);
        const cell = row.getCell(2);
        cell.value = { text: sheets[j].name, hyperlink: "#'" + sheets[j].name + "'!A1" };
        cell.font = { color: { argb: 'FF0563C1' }, underline: true };
      }
    }

    const hasBackLink = i >= 1;
    if (hasBackLink) {
      ws.addRow(['← 返回目录']);
      const backCell = ws.getCell('A1');
      backCell.value = { text: '← 返回目录', hyperlink: "#'目录'!A1" };
      backCell.font = { color: { argb: 'FF0563C1' }, underline: true };
    }

    const headerRow = hasBackLink ? 2 : 1;
    ws.addRow(sheet.headers);
    ws.getRow(headerRow).font = { bold: true };
    sheet.rows.forEach(row => ws.addRow(row));

    ws.columns.forEach(col => {
      let maxLen = 10;
      col.eachCell({ includeEmpty: true }, cell => {
        const v = cell.value;
        const text = v && typeof v === 'object' ? v.text : v;
        const len = String(text || '').length;
        if (len > maxLen) maxLen = len;
      });
      col.width = Math.min(maxLen + 2, 40);
    });

    if (sheet.images && sheet.images.length) {
      const startRow = sheet.rows.length + (i >= 1 ? 4 : 3);
      let curRow = startRow;
      for (let j = 0; j < sheet.images.length; j++) {
        const img = sheet.images[j];
        const base64Data = img.data.split(',')[1];
        const imgId = wb.addImage({
          buffer: base64ToUint8Array(base64Data),
          extension: 'jpeg',
        });
        const targetW = 800;
        const targetH = Math.round(targetW * img.h / img.w);
        const rowSpan = Math.ceil(targetH / 20);
        ws.addImage(imgId, {
          tl: { col: 0, row: curRow },
          ext: { width: targetW, height: targetH },
        });
        curRow += rowSpan + 1;
      }
    }

    await yieldUI();
  }

  toast('Excel：正在序列化…', true);
  await yieldUI();
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  download(blob, baseName + '.xlsx');
}

function base64ToUint8Array(base64) {
  const binary = atob(base64);
  const arr = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) arr[i] = binary.charCodeAt(i);
  return arr;
}

/* ---- 启动 ---- */
initLoad();
