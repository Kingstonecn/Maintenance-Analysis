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
    '<div class="mc-desc">' + esc(m.desc) + '</div>' +
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
    '<h3 class="exec-name">' + esc(method.name) + '</h3>' +
    '<p class="exec-desc">' + esc(method.desc) + '</p>' +
    '<p class="exec-fields">涉及字段: ' + esc(fmtFields(method.fields)) + '</p>' +
    (method.guide ? '<div class="exec-guide"><div class="exec-guide-toggle">📖 如何看懂 & 使用此分析</div><div class="exec-guide-content">' + esc(method.guide) + '</div></div>' : '') +
    '<div class="exec-scores">' +
    '<span class="exec-score-item"><span class="exec-score-label">业务价值</span>' + scoreBadge(method.bv, '业务价值') + '</span>' +
    '<span class="exec-score-item"><span class="exec-score-label">易用性</span>' + scoreBadge(method.ua, '易用性') + '</span>' +
    '</div>' +
    '</div>' +
    '<div class="exec-result" id="execResult"><p class="exec-hint">分析中…</p></div>';
  const guideToggle = $('.exec-guide-toggle', exec);
  if (guideToggle) guideToggle.addEventListener('click', () => guideToggle.parentElement.classList.toggle('open'));
  executeMethod(method);
}

function executeMethod(method) {
  const resultEl = $('#execResult');
  resultEl.innerHTML = '<p class="exec-running">分析中…</p>';
  setTimeout(() => {
    try {
      const html = method.fn(S.rows);
      resultEl.innerHTML = html;
      const scripts = resultEl.querySelectorAll('script');
      scripts.forEach(oldScript => {
        const newScript = document.createElement('script');
        newScript.textContent = oldScript.textContent;
        oldScript.parentNode.replaceChild(newScript, oldScript);
      });
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

/* ---- 导出 ---- */
function openExport() {
  const all = ANALYSIS_METHODS.length;
  const c5 = ANALYSIS_METHODS.filter(m => parseFloat(m.score) >= 5).length;
  const c4 = ANALYSIS_METHODS.filter(m => parseFloat(m.score) >= 4).length;
  $('#cntAll').textContent = '(' + all + ' 种)';
  $('#cnt5').textContent = '(' + c5 + ' 种)';
  $('#cnt4').textContent = '(' + c4 + ' 种)';
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
  const scope = document.querySelector('input[name="exportScope"]:checked').value;
  const methods = ANALYSIS_METHODS.filter(m => {
    if (scope === '5') return parseFloat(m.score) >= 5;
    if (scope === '4') return parseFloat(m.score) >= 4;
    return true;
  });
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
  title.textContent = method.name;
  title.style.cssText = 'font-family:serif;margin-bottom:.3rem;font-size:1.1rem;';
  wrap.appendChild(title);
  const desc = document.createElement('p');
  desc.textContent = method.desc;
  desc.style.cssText = 'font-size:.82rem;color:#666;margin-bottom:1rem;';
  wrap.appendChild(desc);
  const result = document.createElement('div');
  result.className = 'exec-result';
  wrap.appendChild(result);
  container.appendChild(wrap);

  let html;
  try {
    html = method.fn(S.rows);
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

  return { wrap, result };
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
      const { result } = renderAnalysisForExport(method, exportArea);
      await new Promise(r => requestAnimationFrame(r));

      const images = getChartImages(result);
      const texts = getTextContent(result);
      const tables = getTableData(result);

      if (doPdf) {
        pdfPages.push({ name: method.name, desc: method.desc, images, texts, tables });
      }

      if (doExcel) {
        if (tables.length) {
          const t = tables[0];
          excelSheets.push({ name: sanitizeSheetName(method.name), headers: t.headers, rows: t.rows, images });
        } else {
          excelSheets.push({ name: sanitizeSheetName(method.name), headers: ['说明'], rows: [['该分析无表格数据，请查看 PDF 导出']], images });
        }
      }
    } catch (err) {
      console.error('导出失败: ' + method.name, err);
      if (doExcel) {
        excelSheets.push({ name: sanitizeSheetName(method.name), headers: ['错误'], rows: [['分析出错: ' + err.message]], images: [] });
      }
      if (doPdf) {
        pdfPages.push({ name: method.name, desc: method.desc, images: [], texts: [{ type: 'p', text: '分析出错: ' + err.message }], tables: [] });
      }
    }

    Charts.destroyAll();
  }

  window.removeEventListener('beforeunload', onBeforeUnload);
  document.body.removeChild(exportArea);

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

  for (let i = 0; i < pages.length; i++) {
    const page = pages[i];
    toast('PDF ' + (i + 1) + '/' + pages.length + '：' + page.name + '（' + page.images.length + ' 图, ' + page.tables.reduce((s,t) => s + t.rows.length, 0) + ' 行表）', true);
    await yieldUI();
    if (i > 0) { pdf.addPage(); y = margin; }
    pageStarts.push({ name: page.name, pageNum: pdf.getNumberOfPages() });

    y += addText(pdf, page.name, margin, y, 5) + 2;
    y += addText(pdf, page.desc, margin, y, 3.2, '#646464', contentW) + 2;

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

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    toast('Excel ' + (i + 1) + '/' + sheets.length + '：' + sheet.name, true);
    const ws = wb.addWorksheet(sheet.name);

    ws.addRow(sheet.headers);
    ws.getRow(1).font = { bold: true };
    sheet.rows.forEach(row => ws.addRow(row));

    ws.columns.forEach(col => {
      let maxLen = 10;
      col.eachCell({ includeEmpty: true }, cell => {
        const len = String(cell.value || '').length;
        if (len > maxLen) maxLen = len;
      });
      col.width = Math.min(maxLen + 2, 40);
    });

    if (sheet.images && sheet.images.length) {
      const startRow = sheet.rows.length + 3;
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
