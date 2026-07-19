/* ===================================================================
   维修数据分析 — app.js v0.1
   纯前端、本地处理。加载 Excel → 数据概览 → 分析面板（待开发）
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

function toast(m) {
  const t = $('#toast');
  t.textContent = m;
  t.classList.add('show');
  clearTimeout(toast._t);
  toast._t = setTimeout(() => t.classList.remove('show'), 2400);
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
  $('#todayStamp').textContent = '版本时间：2026.07.19';
}

/* ---- 文件处理 ---- */
async function handleFile(file) {
  try {
    toast('正在解析…');
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellFormula: true, cellNF: true });
    const sheetName = wb.SheetNames[0];
    if (!sheetName) { toast('未找到工作表'); return; }
    S.fileName = file.name;
    parseData(wb.Sheets[sheetName]);
    computeStats();
    showOverview();
    toast('已载入：共 ' + S.rows.length + ' 条维修记录');
  } catch (err) {
    console.error(err);
    toast('载入失败：' + err.message);
  }
}

/* ---- 解析数据 ---- */
function parseData(ws) {
  const ref = ws['!ref'];
  if (!ref) return;
  const m = ref.match(/A\d+:(\w+?)(\d+)/);
  const maxRow = m ? parseInt(m[2]) : 0;
  const maxCol = m ? colToIdx(m[1]) + 1 : COLS.length;

  // Parse data rows
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
  setTimeout(() => { URL.revokeObjectURL(u); a.remove(); }, 2000);
}

function doExport() {
  closeExport();
  const exportSummary = $('#optSummary input').checked;
  const exportRaw = $('#optRaw input').checked;
  if (exportSummary) exportSummaryCSV();
  if (exportRaw) exportRawJSON();
  toast('已导出');
}

function exportSummaryCSV() {
  const rows = [['维度', '类别', '数量', '占比']];
  const addSection = (title, counts) => {
    const total = Object.values(counts).reduce((s, v) => s + v, 0);
    Object.entries(counts).sort((a, b) => b[1] - a[1]).forEach(([k, v]) => {
      rows.push([title, k, v, (v / total * 100).toFixed(2) + '%']);
    });
  };
  addSection('维修状态', S.stats.status);
  addSection('维修单类型', S.stats.ticketType);
  addSection('车间分类', S.stats.workshopCat);
  addSection('车间名称', S.stats.workshop);
  addSection('紧急程度', S.stats.urgency);
  addSection('班次', S.stats.shift);
  addSection('故障等级', S.stats.faultLevel);
  addSection('设备维修分类', S.stats.repairCat);
  addSection('是否影响产能', S.stats.affectsCap);
  addSection('是否影响质量', S.stats.affectsQual);
  addSection('维修质量评分', S.stats.qualityScore);
  const csv = '\ufeff' + rows.map(r => r.map(c => {
    const s = String(c == null ? '' : c);
    if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  }).join(',')).join('\r\n');
  download(new Blob([csv], { type: 'text/csv;charset=utf-8' }), S.fileName.replace(/\.xlsx?$/i, '') + '-summary.csv');
}

function exportRawJSON() {
  const data = S.rows.map(row => {
    const obj = {};
    for (const L in row) {
      const col = COL_MAP[L];
      const key = col ? col.cn : L;
      let v = row[L];
      if (v instanceof Date) v = fmtDateTime(v);
      obj[key] = v;
    }
    return obj;
  });
  const json = JSON.stringify(data, null, 2);
  download(new Blob([json], { type: 'application/json' }), S.fileName.replace(/\.xlsx?$/i, '') + '-data.json');
}

/* ---- 启动 ---- */
initLoad();
