'use strict';

/* ---- 工具函数 ---- */
function numVal(row, col) { const v = row[col]; if (v == null || v === '') return null; const n = Number(v); return isNaN(n) ? null : n; }
function strVal(row, col) { const v = row[col]; return v == null ? '' : String(v).trim(); }
function dateVal(row, col) { return parseDateVal(row[col]); }
function countByCol(rows, col) { const c = {}; rows.forEach(r => { const v = strVal(r, col); if (v) c[v] = (c[v] || 0) + 1; }); return c; }
function groupBy(rows, col) { const g = {}; rows.forEach(r => { const v = strVal(r, col); if (v) (g[v] || (g[v] = [])).push(r); }); return g; }
function monthKey(d) { if (!d) return ''; return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0'); }
function weekKey(d) { if (!d) return ''; const y = d.getFullYear(); const jan1 = new Date(y, 0, 1); const week = Math.ceil(((d - jan1) / 86400000 + jan1.getDay() + 1) / 7); return y + '-W' + String(week).padStart(2, '0'); }
function sortEntries(obj, desc = true) { return Object.entries(obj).sort((a, b) => desc ? b[1] - a[1] : a[1] - b[1]); }
function sum(arr) { return arr.reduce((s, v) => s + (v || 0), 0); }
function mean(arr) { return arr.length ? sum(arr) / arr.length : 0; }
function fmtPct(v) { return (v * 100).toFixed(1) + '%'; }
function histogram(vals, bins) { const min = ss.min(vals), max = ss.max(vals), bw = (max - min) / bins || 1; const hist = new Array(bins).fill(0); vals.forEach(v => { hist[Math.min(Math.floor((v - min) / bw), bins - 1)]++; }); return { labels: hist.map((_, i) => (min + i * bw).toFixed(2)), data: hist }; }
function crossTab(rows, colA, valsA, colB, valsB) { const groups = {}; rows.forEach(r => { const a = strVal(r, colA), b = strVal(r, colB); if (!a || !b) return; (groups[a] || (groups[a] = {}))[b] = (groups[a] && groups[a][b] || 0) + 1; }); return valsA.map(a => valsB.map(b => (groups[a] && groups[a][b]) || 0)); }

function renderTable(headers, rows, maxRows) {
  let html = '<div class="result-table-wrap"><table class="result-table"><thead><tr>';
  headers.forEach(h => html += '<th>' + esc(h) + '</th>');
  html += '</tr></thead><tbody>';
  const limit = maxRows ? Math.min(rows.length, maxRows) : rows.length;
  for (let i = 0; i < limit; i++) {
    html += '<tr>';
    rows[i].forEach(v => html += '<td>' + esc(v == null ? '' : v) + '</td>');
    html += '</tr>';
  }
  html += '</tbody></table></div>';
  if (maxRows && rows.length > maxRows) html += '<p class="table-note">显示前 ' + maxRows + ' 行，共 ' + rows.length + ' 行</p>';
  return html;
}

function chartCanvas(id) { return '<div class="chart-canvas-wrap"><canvas id="' + id + '"></canvas></div>'; }

function exportCsvBtn(name, headers, rows) {
  return '<button class="btn-chart-export" onclick="AnalysisExport.csv(\'' + name + '\',' + JSON.stringify(headers) + ',' + JSON.stringify(rows) + ')">导出CSV</button>';
}

const AnalysisExport = {
  csv(name, headers, rows) {
    const csv = '\ufeff' + [headers, ...rows].map(r => r.map(c => { const s = String(c == null ? '' : c); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }).join(',')).join('\r\n');
    download(new Blob([csv], { type: 'text/csv;charset=utf-8' }), name + '.csv');
  }
};

/* ---- 分析方法注册表 ---- */
const ANALYSIS_METHODS = [];
function reg(id, dim, complexity, name, fields, desc, fn, guide) {
  ANALYSIS_METHODS.push({ id, dim, complexity, name, fields, desc, fn, guide });
}

/* ===== ① 描述性统计 (12) ===== */
reg(1,1,1,'维修单总量与状态分布','B','统计维修中/已维修/驳回等各状态数量与占比',execStatusDist,'看各状态占比是否合理，驳回率过高说明审批严格或需求不清。用于评估整体维修流程健康度。');
reg(2,1,1,'维修单类型分布','C','设备维修与公共维修等类型数量对比',execTypeDist,'看设备维修与公共维修比例，判断资源分配是否合理。用于调整维修资源配比。');
reg(3,1,1,'车间分类分布','G','火法/湿法/环保等车间分类维修量占比',execWorkshopCat,'看各车间分类维修量占比，识别维修需求集中的区域。用于跨车间资源调配参考。');
reg(4,1,1,'车间维修频次分布','F','各车间维修单数量排名，识别高频车间',execWorkshopFreq,'看排名靠前车间的维修单数量，高频车间需重点保障备件和人力。用于制定车间级维修预算。');
reg(5,1,1,'产线维修频次分布','H','各产线维修单数量，定位高频产线',execLineFreq,'看高频产线是否集中，若某产线远超其他需排查设备或操作问题。用于产线级维护计划制定。');
reg(6,1,1,'工序维修频次分布','I','各工序维修单数量，识别故障高发工序',execProcessFreq,'看高发工序是否与特定工艺相关，结合设备分析定位根因。用于工序级预防性维护安排。');
reg(7,1,1,'故障等级分布','P','A/B/C级故障数量占比，评估严重程度',execFaultLevel,'看A级故障占比是否偏高，A级多说明设备可靠性差。用于评估整体设备风险等级。');
reg(8,1,1,'紧急程度分布','Q','突发性/一般性/特急等紧急程度占比',execUrgency,'看突发/特急占比，过高说明预防性维护不足。用于调整预防性维护计划频率。');
reg(9,1,1,'班次维修量对比','S','白班与夜班维修单数量对比',execShiftCompare,'看两班维修量是否均衡，夜班偏多可能说明夜班设备状态差。用于班次维修人员配置调整。');
reg(10,1,1,'维修时长统计','AK','维修时长均值/中位数/标准差/四分位数',execRepairDurationStat,'看中位数与均值差距，差距大说明有异常长维修拉高均值。用于设定维修时长基准和考核标准。');
reg(11,1,1,'等待时长统计','AI','等待时长均值/中位数/分布，评估响应速度',execWaitTimeStat,'看等待时长是否过长，过长说明维修调度效率低。用于优化维修派单流程。');
reg(12,1,1,'维修费用统计','AN,AP,AQ','备件/委外/总费用均值与分布统计',execCostStat,'看备件与委外费用比例，委外占比高需评估自修能力。用于制定维修成本预算基准。');

/* ===== ② 排名与帕累托 (8) ===== */
reg(13,2,1,'设备故障频次帕累托','L','识别20%设备贡献80%故障的关键少数',execEquipPareto,'看累计占比曲线在哪个设备处达到80%，这些设备是重点维护对象。用于制定重点设备清单和备件储备优先级。');
reg(14,2,1,'故障原因帕累托','W','故障原因Top N排名及累计占比',execCausePareto,'看前3-5个故障原因累计占比，集中解决这些原因可大幅降低维修量。用于制定根因消除计划。');
reg(15,2,2,'车间维修成本帕累托','F,AQ','各车间维修总成本排名与累计占比',execWsCostPareto,'看哪个车间成本占比最高，重点控制该车间维修支出。用于车间成本考核和预算分配。');
reg(16,2,1,'高频故障设备 Top 20','L','按故障次数排名前20设备',execEquipTop20,'看Top20设备的故障次数，这些设备应优先安排预防性维护。用于制定设备更新换代优先级清单。');
reg(17,2,2,'维修人工作量排名','X,AK','按维修单数和总维修时长排名',execRepairerWorkload,'看工作量是否均衡，某维修人过高说明人力分配不均。用于调整维修团队人员分工。');
reg(18,2,2,'维修人效率排名','X,AK','按平均维修时长评估维修人效率',execRepairerEfficiency,'看平均时长排名，偏长的维修人可能需要技能培训。用于识别培训需求和效率标杆。');
reg(19,2,1,'故障部件频次排名','AE','故障部件出现频次Top N排名',execPartFreq,'看高频故障部件，这些部件应提高备件库存。用于优化备件采购和库存管理。');
reg(20,2,2,'高成本维修单 Top 20','AQ','按维修总价值排名前20维修单',execHighCostTop20,'看高成本维修单的设备分布，分析是否有共性原因。用于成本控制和单笔维修审批阈值设定。');

/* ===== ③ 时间趋势 (8) ===== */
reg(21,3,2,'月度维修单数量趋势','Y','按月统计维修单数量，观察整体趋势',execMonthlyTrend,'看维修量是否逐月上升，上升说明设备老化或维护不足。用于制定年度维修计划。');
reg(22,3,2,'周度维修单数量趋势','Y','按周统计维修单数量，捕捉短期波动',execWeeklyTrend,'看某周是否突增，突增可能对应设备集中故障或季节性因素。用于短期维修资源调度。');
reg(23,3,2,'月度维修时长趋势','Y,AK','月度总时长与平均时长变化趋势',execMonthlyDurTrend,'看平均时长是否上升，上升说明维修难度增加或技能下降。用于评估维修团队能力趋势。');
reg(24,3,2,'月度维修费用趋势','Y,AN,AP,AQ','月度备件/委外/总费用变化趋势',execMonthlyCostTrend,'看费用是否突增，突增月份需排查大修或异常故障。用于月度成本监控和预算预警。');
reg(25,3,2,'各车间月度故障趋势','Y,F','多车间月度维修量同图对比',execWsMonthlyTrend,'看哪个车间趋势恶化最快，优先干预。用于车间级维修绩效考核。');
reg(26,3,2,'班次维修量月度趋势','Y,S','白班与夜班月度维修量对比',execShiftMonthlyTrend,'看两班趋势是否分化，夜班持续偏高需排查夜班设备巡检质量。用于班次管理优化。');
reg(27,3,2,'故障等级月度趋势','Y,P','A/B/C级故障月度数量变化',execFaultLevelTrend,'看A级故障是否逐月增加，增加说明设备可靠性恶化。用于设备风险评估和更新计划。');
reg(28,3,2,'维修评分月度趋势','Y,BD','月度平均维修评分变化趋势',execScoreMonthlyTrend,'看评分是否下降，下降趋势需调查维修质量原因。用于维修质量长期跟踪。');

/* ===== ④ 设备可靠性 (10) ===== */
reg(29,4,2,'MTBF 计算','L,Y','按设备计算平均故障间隔时间，评估可靠性',execMTBF,'看MTBF排名靠后的设备，MTBF短说明故障频繁。用于制定设备更换或大修计划。');
reg(30,4,2,'MTTR 计算','L,AK','按设备计算平均维修时间，评估修复能力',execMTTR,'看MTTR长的设备，可能备件不足或维修难度大。用于备件预储和维修方案优化。');
reg(31,4,3,'设备可用度估算','L,Y,AK','MTBF/(MTBF+MTTR)估算设备可用度',execAvailability,'看可用度低于95%的设备，这些设备影响产能最严重。用于制定设备更新优先级。');
reg(32,4,2,'重复故障识别','R,L','识别重复故障设备及关联记录',execRepeatFault,'看重复故障次数高的设备，说明上次维修未解决根因。用于制定根因分析计划。');
reg(33,4,2,'设备故障频率趋势','L,Y','单设备月度故障次数变化趋势',execEquipFaultTrend,'看Top5设备的故障频率是否加速上升，加速上升说明设备进入衰退期。用于设备换代决策。');
reg(34,4,3,'设备故障模式分析','L,W','同设备不同故障原因分布分析',execEquipFaultMode,'看交叉表中高频组合，某设备某原因集中说明系统性问题。用于针对性维修方案制定。');
reg(35,4,3,'故障部件频率分析','L,AE','按设备分组分析故障部件频率',execEquipPartFreq,'看某设备频繁损坏的部件，可能是设计缺陷或操作不当。用于部件级改进和备件储备。');
reg(36,4,2,'设备维修分类对比','L,AD','各设备维修分类分布对比',execEquipRepairCat,'看某设备是否以某类维修为主，如频繁小修说明预防性维护不足。用于调整维护策略。');
reg(37,4,3,'新设备故障率分析','M,Y','按首次出现时间分析新设备故障率',execNewEquipFault,'看0-30天故障占比，占比高说明安装调试或质量问题。用于新设备验收标准优化。');
reg(38,4,3,'设备风险度评分','L,AK,AQ','综合频率/时长/成本给设备风险打分',execEquipHealth,'看风险度评分最高的设备，这些是优先关注和投入资源的目标。用于设备分级管理。');

/* ===== ⑤ 维修效率 (8) ===== */
reg(39,5,2,'等待 vs 维修时长占比','AI,AK','等待与维修时长比例，评估流程效率',execWaitVsRepair,'看等待占比是否过高，超过30%说明派单流程需优化。用于优化维修调度流程。');
reg(40,5,2,'班次维修效率对比','S,AK','白班与夜班维修时长分布对比',execShiftEfficiency,'看两班平均时长差距，差距大说明某班技能或资源不足。用于班次人员配置优化。');
reg(41,5,2,'维修人效率对比','X,AK','各维修人平均维修时长对比',execRepairerCompare,'看各维修人平均时长分布，偏长的需技能培训。用于制定维修人培训计划。');
reg(42,5,3,'响应时间分析','Y,V,AF','申请到接收到维修开始各阶段时长',execResponseTime,'看哪个阶段耗时最长，针对性优化该环节。用于缩短维修响应时间的流程改进。');
reg(43,5,2,'维修超时分析','AK','识别超长维修单，分析超时分布',execOvertime,'看超P90的维修单特征，分析是否有共性设备或原因。用于制定超时预警机制。');
reg(44,5,2,'驳回率分析','B,F','各车间驳回维修单占比分析',execRejectRate,'看驳回率高的车间，可能需求描述不清或审批标准不一。用于统一维修申请标准。');
reg(45,5,3,'维修阶段时长分解','Y,AA,AF,AG','申请到确认到维修各阶段时长分解',execPhaseBreakdown,'看哪个阶段耗时最长，耗时最长的阶段是流程优化重点。用于缩短整体维修周期。');
reg(46,5,2,'及时性分析','BA','接单结单及时性评分分布',execTimeliness,'看低分占比，低分多说明结单不及时。用于考核维修团队响应速度。');

/* ===== ⑥ 成本与影响 (7) ===== */
reg(47,6,2,'维修成本按车间分布','F,AN,AP,AQ','各车间备件/委外/总成本对比',execWsCost,'看哪个车间总成本最高，备件和委外分别占比多少。用于车间成本控制目标制定。');
reg(48,6,2,'备件 vs 委外成本占比','AN,AP','备件与委外费用占比分析',execPartsVsOutsource,'看委外占比是否过高，过高需评估自修能力是否不足。用于委外策略调整。');
reg(49,6,2,'产能影响量化','AR,F,AT,AU','影响产能维修单的时长与数量分析',execCapacityImpact,'看影响产能的维修单数和时长，量化产能损失。用于产能保障优先级排序。');
reg(50,6,2,'质量影响分析','AS,F,AQ','影响质量的维修单数量与特征',execQualityImpact,'看影响质量的维修单集中在哪个车间，针对性改进。用于质量风险管理。');
reg(51,6,2,'成本趋势分析','Y,AQ','月度维修总成本变化趋势',execCostTrend,'看成本是否逐月上升，上升需排查原因。用于月度成本预警和预算控制。');
reg(52,6,3,'高成本维修单识别','AQ','识别超P95阈值的异常高成本维修单',execHighCostOutlier,'看超P95的维修单详情，分析是否有共性设备或原因。用于异常成本审查和审批。');
reg(53,6,3,'成本 vs 故障等级','P,AQ','不同故障等级的维修成本分布',execCostVsLevel,'看A级故障平均成本是否远超B/C级，评估高等级故障的经济影响。用于故障等级成本量化。');

/* ===== ⑦ 评分与质量 (6) ===== */
reg(54,7,2,'维修评分分布','AX,AY,AZ,BA','质量/现场/重复/及时性各维度评分分布',execScoreDist,'看哪个维度均值最低，最低维度是改进重点。用于制定维修质量改进方向。');
reg(55,7,3,'评分维度相关性','AX,AY,AZ,BA,BD','各评分维度与总分的相关系数',execScoreCorrelation,'看哪个维度与总分相关性最高，高相关的维度对总分影响最大。用于优化评分权重分配。');
reg(56,7,2,'维修人评分排名','X,BD','按平均总分对维修人排名',execRepairerScore,'看低分维修人，针对性培训或调整岗位。用于维修人员绩效考核。');
reg(57,7,2,'车间评分对比','F,BD','各车间平均维修评分对比',execWsScore,'看哪个车间评分最低，低分车间需调查维修质量原因。用于车间维修质量考核。');
reg(58,7,2,'低分维修单分析','BD,F','总分低于阈值的维修单特征分析',execLowScore,'看低分维修单集中在哪个车间，分析共性原因。用于质量改进专项计划。');
reg(59,7,2,'评分趋势分析','Y,BD','月度平均维修评分变化趋势',execScoreTrend,'看评分是否下降趋势，下降需调查原因。用于维修质量长期跟踪和预警。');

/* ===== ⑧ 相关性分析 (6) ===== */
reg(60,8,3,'维修时长 vs 故障等级','P,AK','不同故障等级维修时长分布对比',execDurVsLevel,'看A级维修时长是否显著高于C级，评估等级与时长的关联强度。用于维修工时预估。');
reg(61,8,3,'等待时长 vs 维修评分','AI,BD','等待时长与维修评分散点关系',execWaitVsScore,'看相关系数是否为负，负相关说明等待越久评分越低。用于证明缩短等待时间的必要性。');
reg(62,8,3,'故障频率 vs 维修成本','L,AQ','设备故障次数与维修成本关系',execFreqVsCost,'看散点是否有线性趋势，频率高的设备成本是否也高。用于设备全生命周期成本评估。');
reg(63,8,3,'班次 vs 故障类型','S,W','班次与故障原因交叉表',execShiftVsFault,'看某班次是否集中某类故障，可能与操作习惯相关。用于班次操作培训。');
reg(64,8,3,'设备类型 vs 故障模式','J,W','设备分类与故障原因交叉分析',execEquipTypeVsFault,'看某类设备是否集中某类故障，用于设备选型和改进。用于设备采购参考。');
reg(65,8,3,'维修人数 vs 维修时长','AH,AK','参与维修人数与维修时长关系',execParticipantsVsDur,'看人数增加是否缩短时长，如果不缩短说明协作效率低。用于维修人员配置优化。');

/* ===== ⑨ 回归与预测 (5) ===== */
reg(66,9,3,'维修时长预测','P,AK','基于故障等级预测维修时长',execDurRegression,'看R²值，R²>0.3说明等级有一定预测力。用于维修工时预估和排程参考。');
reg(67,9,3,'维修成本预测','AK,AQ','基于维修时长预测维修成本',execCostRegression,'看回归方程斜率，斜率大说明时长对成本影响大。用于维修成本预估和审批参考。');
reg(68,9,3,'评分影响因素回归','AX,AY,AZ,BA,BD','各评分因素对总分的回归系数',execScoreRegression,'看哪个维度斜率最大，该维度对总分影响最大。用于优化评分权重或考核重点。');
reg(69,9,4,'月度维修量预测','Y','移动平均与指数平滑预测未来维修量',execMonthlyForecast,'看预测值与实际值差距，差距大说明趋势不稳定。用于下月维修资源预分配。');
reg(70,9,4,'设备故障概率预测','L,Y','基于频率趋势预测设备故障概率',execFaultProbForecast,'看预测故障数高的设备，提前安排预防性维护。用于预防性维护计划制定。');

/* ===== ⑩ 文本分析 (4) ===== */
reg(71,10,2,'问题描述词频分析','T','问题描述关键词频次排名',execDescWordFreq,'看高频关键词，这些词代表最常见的故障现象。用于故障知识库建设和快速诊断。');
reg(72,10,2,'故障原因描述词频','AL','故障原因描述关键词频次排名',execCauseDescWordFreq,'看高频原因关键词，集中解决这些原因可降低维修量。用于根因分析参考。');
reg(73,10,2,'维修内容词频','AM','维修内容关键词频次排名',execContentWordFreq,'看高频维修操作词，了解最常执行的维修动作。用于维修标准化和SOP制定。');
reg(74,10,3,'高频故障模式提取','T,AL,AM','组合三列文本聚类提取故障模式',execFaultPattern,'看高频组合模式，同一模式反复出现说明系统性问题未解决。用于系统性故障根因消除。');

/* ---- 评分注入 ---- */
const _SCORES = {
  1:[5,5,5],2:[4,5,5],3:[4,5,5],4:[5,5,5],5:[4,5,5],6:[4,5,5],7:[5,5,5],8:[4,5,5],
  9:[4,5,5],10:[5,4,5],11:[4,4,5],12:[5,4,5],
  13:[5,5,5],14:[5,4,5],15:[5,4,5],16:[5,5,5],17:[4,4,5],18:[4,4,5],19:[4,5,5],20:[4,5,5],
  21:[5,4,5],22:[4,4,5],23:[4,4,5],24:[5,4,4],25:[4,3,4],26:[3,4,4],27:[4,4,4],28:[3,4,4],
  29:[5,4,5],30:[5,4,5],31:[5,3,5],32:[5,4,5],33:[4,3,4],34:[4,3,4],35:[3,3,4],36:[3,4,4],
  37:[3,3,4],38:[5,3,4],
  39:[5,4,5],40:[4,4,5],41:[4,4,5],42:[5,3,4],43:[5,5,5],44:[4,4,5],45:[4,2,4],46:[3,4,5],
  47:[5,4,5],48:[4,5,5],49:[5,4,4],50:[4,5,5],51:[5,4,4],52:[4,3,5],53:[4,3,4],
  54:[4,4,5],55:[3,2,5],56:[4,4,5],57:[4,4,5],58:[4,4,5],59:[3,4,4],
  60:[4,3,4],61:[3,2,4],62:[4,3,4],63:[3,2,3],64:[3,2,3],65:[2,3,4],
  66:[4,2,3],67:[4,2,3],68:[3,2,3],69:[5,2,3],70:[4,1,2],
  71:[4,4,4],72:[3,4,4],73:[3,4,4],74:[4,2,2],
  75:[5,5,5],76:[4,5,5],77:[4,4,5],78:[4,3,4],79:[4,3,3],80:[4,2,3],81:[3,3,4],
};

/* ===== 实现函数 ===== */

function execStatusDist(rows) {
  const counts = countByCol(rows, 'B');
  const entries = sortEntries(counts);
  const labels = entries.map(e => e[0]), data = entries.map(e => e[1]);
  const total = sum(data);
  const tblRows = entries.map(e => [e[0], e[1], fmtPct(e[1] / total)]);
  return chartCanvas('ac1') + '<div class="chart-row">' + renderTable(['状态','数量','占比'], tblRows) + '</div>' +
    '<script>Charts.doughnut("ac1",' + JSON.stringify(labels) + ',' + JSON.stringify(data) + ')</script>' +
    exportCsvBtn('status-dist', ['状态','数量','占比'], tblRows);
}

function execTypeDist(rows) {
  const counts = countByCol(rows, 'C');
  const entries = sortEntries(counts);
  const labels = entries.map(e => e[0]), data = entries.map(e => e[1]);
  const total = sum(data);
  const tblRows = entries.map(e => [e[0], e[1], fmtPct(e[1] / total)]);
  return chartCanvas('ac2') + renderTable(['类型','数量','占比'], tblRows) +
    '<script>Charts.pie("ac2",' + JSON.stringify(labels) + ',' + JSON.stringify(data) + ')</script>' +
    exportCsvBtn('type-dist', ['类型','数量','占比'], tblRows);
}

function execWorkshopCat(rows) {
  const counts = countByCol(rows, 'G');
  const entries = sortEntries(counts);
  const labels = entries.map(e => e[0]), data = entries.map(e => e[1]);
  return chartCanvas('ac3') + renderTable(['分类','数量'], entries.map(e => [e[0], e[1]])) +
    '<script>Charts.bar("ac3",' + JSON.stringify(labels) + ',[{label:"数量",data:' + JSON.stringify(data) + '}])</script>' +
    exportCsvBtn('ws-cat', ['分类','数量'], entries.map(e => [e[0], e[1]]));
}

function execWorkshopFreq(rows) {
  const counts = countByCol(rows, 'F');
  const entries = sortEntries(counts);
  const labels = entries.map(e => e[0]), data = entries.map(e => e[1]);
  return chartCanvas('ac4') + renderTable(['车间','数量'], entries.map(e => [e[0], e[1]])) +
    '<script>Charts.horizontalBar("ac4",' + JSON.stringify(labels) + ',[{label:"维修单数",data:' + JSON.stringify(data) + '}])</script>' +
    exportCsvBtn('ws-freq', ['车间','数量'], entries.map(e => [e[0], e[1]]));
}

function execLineFreq(rows) {
  const counts = countByCol(rows, 'H');
  const entries = sortEntries(counts).slice(0, 20);
  const labels = entries.map(e => e[0]), data = entries.map(e => e[1]);
  return chartCanvas('ac5') + renderTable(['产线','数量'], entries.map(e => [e[0], e[1]]), 20) +
    '<script>Charts.horizontalBar("ac5",' + JSON.stringify(labels) + ',[{label:"维修单数",data:' + JSON.stringify(data) + '}])</script>' +
    exportCsvBtn('line-freq', ['产线','数量'], entries.map(e => [e[0], e[1]]));
}

function execProcessFreq(rows) {
  const counts = countByCol(rows, 'I');
  const entries = sortEntries(counts).slice(0, 20);
  const labels = entries.map(e => e[0]), data = entries.map(e => e[1]);
  return chartCanvas('ac6') + renderTable(['工序','数量'], entries.map(e => [e[0], e[1]]), 20) +
    '<script>Charts.horizontalBar("ac6",' + JSON.stringify(labels) + ',[{label:"维修单数",data:' + JSON.stringify(data) + '}])</script>' +
    exportCsvBtn('process-freq', ['工序','数量'], entries.map(e => [e[0], e[1]]));
}

function execFaultLevel(rows) {
  const counts = countByCol(rows, 'P');
  const entries = sortEntries(counts);
  const labels = entries.map(e => e[0]), data = entries.map(e => e[1]);
  return chartCanvas('ac7') + renderTable(['等级','数量','占比'], entries.map(e => [e[0], e[1], fmtPct(e[1] / sum(data))])) +
    '<script>Charts.doughnut("ac7",' + JSON.stringify(labels) + ',' + JSON.stringify(data) + ')</script>' +
    exportCsvBtn('fault-level', ['等级','数量','占比'], entries.map(e => [e[0], e[1], fmtPct(e[1] / sum(data))]));
}

function execUrgency(rows) {
  const counts = countByCol(rows, 'Q');
  const entries = sortEntries(counts);
  const labels = entries.map(e => e[0]), data = entries.map(e => e[1]);
  return chartCanvas('ac8') + renderTable(['紧急程度','数量','占比'], entries.map(e => [e[0], e[1], fmtPct(e[1] / sum(data))])) +
    '<script>Charts.bar("ac8",' + JSON.stringify(labels) + ',[{label:"数量",data:' + JSON.stringify(data) + '}])</script>' +
    exportCsvBtn('urgency', ['紧急程度','数量','占比'], entries.map(e => [e[0], e[1], fmtPct(e[1] / sum(data))]));
}

function execShiftCompare(rows) {
  const counts = countByCol(rows, 'S');
  const entries = sortEntries(counts);
  const labels = entries.map(e => e[0]), data = entries.map(e => e[1]);
  return chartCanvas('ac9') + renderTable(['班次','数量','占比'], entries.map(e => [e[0], e[1], fmtPct(e[1] / sum(data))])) +
    '<script>Charts.bar("ac9",' + JSON.stringify(labels) + ',[{label:"维修单数",data:' + JSON.stringify(data) + '}])</script>' +
    exportCsvBtn('shift-compare', ['班次','数量','占比'], entries.map(e => [e[0], e[1], fmtPct(e[1] / sum(data))]));
}

function execRepairDurationStat(rows) {
  const vals = rows.map(r => numVal(r, 'AK')).filter(v => v != null && v > 0);
  if (!vals.length) return '<p class="no-data">无有效维修时长数据</p>';
  const sorted = vals.slice().sort((a, b) => a - b);
  const q = p => ss.quantile(sorted, p);
  const stats = [['样本数', vals.length], ['均值', mean(vals).toFixed(3)], ['中位数', ss.median(vals).toFixed(3)],
    ['标准差', ss.standardDeviation(vals).toFixed(3)], ['最小值', ss.min(vals).toFixed(3)], ['最大值', ss.max(vals).toFixed(3)],
    ['Q1(25%)', q(0.25).toFixed(3)], ['Q3(75%)', q(0.75).toFixed(3)], ['P90', q(0.9).toFixed(3)], ['P95', q(0.95).toFixed(3)]];
  const h = histogram(vals, 20);
  return chartCanvas('ac10') + renderTable(['统计量','值'], stats) +
    '<script>Charts.bar("ac10",' + JSON.stringify(h.labels) + ',[{label:"频次",data:' + JSON.stringify(h.data) + '}],{plugins:{title:{display:true,text:"维修时长分布直方图"}}})</script>' +
    exportCsvBtn('dur-stat', ['统计量','值'], stats);
}

function execWaitTimeStat(rows) {
  const vals = rows.map(r => numVal(r, 'AI')).filter(v => v != null && v > 0);
  if (!vals.length) return '<p class="no-data">无有效等待时长数据</p>';
  const sorted = vals.slice().sort((a, b) => a - b);
  const stats = [['样本数', vals.length], ['均值', mean(vals).toFixed(3)], ['中位数', ss.median(vals).toFixed(3)],
    ['标准差', ss.standardDeviation(vals).toFixed(3)], ['最小值', ss.min(vals).toFixed(3)], ['最大值', ss.max(vals).toFixed(3)],
    ['P90', ss.quantile(sorted, 0.9).toFixed(3)], ['P95', ss.quantile(sorted, 0.95).toFixed(3)]];
  const h = histogram(vals, 20);
  return chartCanvas('ac11') + renderTable(['统计量','值'], stats) +
    '<script>Charts.bar("ac11",' + JSON.stringify(h.labels) + ',[{label:"频次",data:' + JSON.stringify(h.data) + '}],{plugins:{title:{display:true,text:"等待时长分布直方图"}}})</script>' +
    exportCsvBtn('wait-stat', ['统计量','值'], stats);
}

function execCostStat(rows) {
  const an = rows.map(r => numVal(r, 'AN')).filter(v => v != null && v > 0);
  const ap = rows.map(r => numVal(r, 'AP')).filter(v => v != null && v > 0);
  const aq = rows.map(r => numVal(r, 'AQ')).filter(v => v != null && v > 0);
  if (!an.length && !ap.length && !aq.length) return '<p class="no-data">无有效费用数据</p>';
  const safe = arr => arr.length ? arr : [0];
  const stats = [['备件价值(AN) 均值', mean(safe(an)).toFixed(2)], ['备件价值(AN) 中位数', ss.median(safe(an)).toFixed(2)],
    ['委外费用(AP) 均值', mean(safe(ap)).toFixed(2)], ['委外费用(AP) 中位数', ss.median(safe(ap)).toFixed(2)],
    ['维修总价值(AQ) 均值', mean(safe(aq)).toFixed(2)], ['维修总价值(AQ) 中位数', ss.median(safe(aq)).toFixed(2)],
    ['维修总价值(AQ) 最大值', ss.max(safe(aq)).toFixed(2)], ['维修总价值(AQ) 总和', sum(aq).toFixed(2)]];
  return chartCanvas('ac12') + renderTable(['统计量','值'], stats) +
    '<script>Charts.bar("ac12",["备件价值","委外费用","维修总价值"],[{label:"均值",data:[' + mean(safe(an)).toFixed(2) + ',' + mean(safe(ap)).toFixed(2) + ',' + mean(safe(aq)).toFixed(2) + ']},{label:"中位数",data:[' + ss.median(safe(an)).toFixed(2) + ',' + ss.median(safe(ap)).toFixed(2) + ',' + ss.median(safe(aq)).toFixed(2) + ']}])</script>' +
    exportCsvBtn('cost-stat', ['统计量','值'], stats);
}

/* ===== ② 帕累托 ===== */
function execEquipPareto(rows) {
  const counts = countByCol(rows, 'L');
  const entries = sortEntries(counts);
  const total = sum(entries.map(e => e[1]));
  let cum = 0; const cumData = entries.map(e => { cum += e[1]; return cum / total * 100; });
  const labels = entries.map(e => e[0]).slice(0, 30);
  const barData = entries.map(e => e[1]).slice(0, 30);
  const lineData = cumData.slice(0, 30);
  const tblRows = entries.slice(0, 30).map((e, i) => [e[0], e[1], fmtPct(e[1] / total), lineData[i].toFixed(1) + '%']);
  return chartCanvas('ac13') + renderTable(['设备','故障数','占比','累计占比'], tblRows, 30) +
    '<script>(function(){Charts._destroy("ac13");var ctx=document.getElementById("ac13").getContext("2d");new Chart(ctx,{type:"bar",data:{labels:' + JSON.stringify(labels) + ',datasets:[{type:"bar",label:"故障数",data:' + JSON.stringify(barData) + ',backgroundColor:"#1d4e89"},{type:"line",label:"累计占比%",data:' + JSON.stringify(lineData) + ',borderColor:"#b1281e",borderWidth:2,fill:false,yAxisID:"y1"}]},options:{responsive:true,maintainAspectRatio:false,scales:{y:{beginAtZero:true,title:{display:true,text:"故障数"}},y1:{position:"right",max:100,title:{display:true,text:"累计占比%"},grid:{drawOnChartArea:false}}}}})})()</script>' +
    exportCsvBtn('equip-pareto', ['设备','故障数','占比','累计占比'], tblRows);
}

function execCausePareto(rows) {
  const counts = countByCol(rows, 'W');
  const entries = sortEntries(counts).filter(e => e[0]);
  const total = sum(entries.map(e => e[1]));
  let cum = 0; const cumData = entries.map(e => { cum += e[1]; return cum / total * 100; });
  const labels = entries.map(e => e[0]).slice(0, 20);
  const barData = entries.map(e => e[1]).slice(0, 20);
  const lineData = cumData.slice(0, 20);
  const tblRows = entries.slice(0, 20).map((e, i) => [e[0], e[1], fmtPct(e[1] / total), lineData[i].toFixed(1) + '%']);
  return chartCanvas('ac14') + renderTable(['故障原因','数量','占比','累计占比'], tblRows, 20) +
    '<script>(function(){Charts._destroy("ac14");var ctx=document.getElementById("ac14").getContext("2d");new Chart(ctx,{type:"bar",data:{labels:' + JSON.stringify(labels) + ',datasets:[{type:"bar",label:"数量",data:' + JSON.stringify(barData) + ',backgroundColor:"#1d4e89"},{type:"line",label:"累计占比%",data:' + JSON.stringify(lineData) + ',borderColor:"#b1281e",borderWidth:2,fill:false,yAxisID:"y1"}]},options:{responsive:true,maintainAspectRatio:false,scales:{y:{beginAtZero:true},y1:{position:"right",max:100,grid:{drawOnChartArea:false}}}}})})()</script>' +
    exportCsvBtn('cause-pareto', ['故障原因','数量','占比','累计占比'], tblRows);
}

function execWsCostPareto(rows) {
  const groups = groupBy(rows, 'F');
  const entries = Object.entries(groups).map(([ws, rs]) => [ws, sum(rs.map(r => numVal(r, 'AQ') || 0))]).sort((a, b) => b[1] - a[1]);
  const total = sum(entries.map(e => e[1]));
  let cum = 0; const cumData = entries.map(e => { cum += e[1]; return cum / total * 100; });
  const labels = entries.map(e => e[0]), barData = entries.map(e => e[1]), lineData = cumData;
  const tblRows = entries.map((e, i) => [e[0], e[1].toFixed(2), fmtPct(e[1] / total), lineData[i].toFixed(1) + '%']);
  return chartCanvas('ac15') + renderTable(['车间','维修总成本','占比','累计占比'], tblRows) +
    '<script>(function(){Charts._destroy("ac15");var ctx=document.getElementById("ac15").getContext("2d");new Chart(ctx,{type:"bar",data:{labels:' + JSON.stringify(labels) + ',datasets:[{type:"bar",label:"成本",data:' + JSON.stringify(barData) + ',backgroundColor:"#1d4e89"},{type:"line",label:"累计占比%",data:' + JSON.stringify(lineData) + ',borderColor:"#b1281e",borderWidth:2,fill:false,yAxisID:"y1"}]},options:{responsive:true,maintainAspectRatio:false,scales:{y:{beginAtZero:true},y1:{position:"right",max:100,grid:{drawOnChartArea:false}}}}})})()</script>' +
    exportCsvBtn('ws-cost-pareto', ['车间','维修总成本','占比','累计占比'], tblRows);
}

function execEquipTop20(rows) {
  const counts = countByCol(rows, 'L');
  const entries = sortEntries(counts).slice(0, 20);
  const labels = entries.map(e => e[0]), data = entries.map(e => e[1]);
  const tblRows = entries.map((e, i) => [i + 1, e[0], e[1]]);
  return chartCanvas('ac16') + renderTable(['排名','设备','故障次数'], tblRows) +
    '<script>Charts.horizontalBar("ac16",' + JSON.stringify(labels) + ',[{label:"故障次数",data:' + JSON.stringify(data) + '}])</script>' +
    exportCsvBtn('equip-top20', ['排名','设备','故障次数'], tblRows);
}

function execRepairerWorkload(rows) {
  const groups = groupBy(rows, 'X');
  const entries = Object.entries(groups).map(([p, rs]) => [p, rs.length, sum(rs.map(r => numVal(r, 'AK') || 0))]).filter(e => e[0]).sort((a, b) => b[1] - a[1]).slice(0, 30);
  const labels = entries.map(e => e[0]);
  const tblRows = entries.map((e, i) => [i + 1, e[0], e[1], e[2].toFixed(2)]);
  return chartCanvas('ac17') + renderTable(['排名','维修人','维修单数','总维修时长'], tblRows, 30) +
    '<script>Charts.horizontalBar("ac17",' + JSON.stringify(labels) + ',[{label:"维修单数",data:' + JSON.stringify(entries.map(e => e[1])) + '}])</script>' +
    exportCsvBtn('repairer-workload', ['排名','维修人','维修单数','总维修时长'], tblRows);
}

function execRepairerEfficiency(rows) {
  const groups = groupBy(rows, 'X');
  const entries = Object.entries(groups).map(([p, rs]) => { const durs = rs.map(r => numVal(r, 'AK')).filter(v => v != null && v > 0); if (!durs.length) return null; return [p, rs.length, mean(durs), ss.median(durs)]; }).filter(e => e && e[0] && e[1] >= 5).sort((a, b) => a[2] - b[2]).slice(0, 30);
  const labels = entries.map(e => e[0]);
  const tblRows = entries.map((e, i) => [i + 1, e[0], e[1], e[2].toFixed(3), e[3].toFixed(3)]);
  return chartCanvas('ac18') + renderTable(['排名','维修人','维修单数','平均维修时长','中位数时长'], tblRows, 30) +
    '<script>Charts.horizontalBar("ac18",' + JSON.stringify(labels) + ',[{label:"平均维修时长",data:' + JSON.stringify(entries.map(e => +e[2].toFixed(3))) + '}])</script>' +
    exportCsvBtn('repairer-efficiency', ['排名','维修人','维修单数','平均维修时长','中位数时长'], tblRows);
}

function execPartFreq(rows) {
  const counts = countByCol(rows, 'AE');
  const entries = sortEntries(counts).filter(e => e[0]).slice(0, 20);
  const labels = entries.map(e => e[0]), data = entries.map(e => e[1]);
  const tblRows = entries.map((e, i) => [i + 1, e[0], e[1]]);
  return chartCanvas('ac19') + renderTable(['排名','故障部件','频次'], tblRows) +
    '<script>Charts.horizontalBar("ac19",' + JSON.stringify(labels) + ',[{label:"频次",data:' + JSON.stringify(data) + '}])</script>' +
    exportCsvBtn('part-freq', ['排名','故障部件','频次'], tblRows);
}

function execHighCostTop20(rows) {
  const entries = rows.map(r => ({ id: strVal(r, 'A'), ws: strVal(r, 'F'), equip: strVal(r, 'L'), cost: numVal(r, 'AQ') || 0, desc: strVal(r, 'T').slice(0, 40) })).filter(r => r.cost > 0).sort((a, b) => b.cost - a.cost).slice(0, 20);
  const tblRows = entries.map((e, i) => [i + 1, e.id, e.ws, e.equip, e.cost.toFixed(2), e.desc]);
  return renderTable(['排名','维修单号','车间','设备','维修总价值','问题描述'], tblRows) + exportCsvBtn('high-cost-top20', ['排名','维修单号','车间','设备','维修总价值','问题描述'], tblRows);
}

/* ===== ③ 时间趋势 ===== */
function execMonthlyTrend(rows) {
  const months = {};
  rows.forEach(r => { const d = dateVal(r, 'Y'); if (!d) return; const k = monthKey(d); months[k] = (months[k] || 0) + 1; });
  const sorted = Object.keys(months).sort();
  const data = sorted.map(k => months[k]);
  const tblRows = sorted.map((k, i) => [k, data[i]]);
  return chartCanvas('ac21') + renderTable(['月份','维修单数'], tblRows) +
    '<script>Charts.line("ac21",' + JSON.stringify(sorted) + ',[{label:"维修单数",data:' + JSON.stringify(data) + '}])</script>' +
    exportCsvBtn('monthly-trend', ['月份','维修单数'], tblRows);
}

function execWeeklyTrend(rows) {
  const weeks = {};
  rows.forEach(r => { const d = dateVal(r, 'Y'); if (!d) return; const k = weekKey(d); weeks[k] = (weeks[k] || 0) + 1; });
  const sorted = Object.keys(weeks).sort();
  const data = sorted.map(k => weeks[k]);
  const tblRows = sorted.map((k, i) => [k, data[i]]);
  return chartCanvas('ac22') + renderTable(['周','维修单数'], tblRows) +
    '<script>Charts.line("ac22",' + JSON.stringify(sorted) + ',[{label:"维修单数",data:' + JSON.stringify(data) + '}])</script>' +
    exportCsvBtn('weekly-trend', ['周','维修单数'], tblRows);
}

function execMonthlyDurTrend(rows) {
  const months = {};
  rows.forEach(r => { const d = dateVal(r, 'Y'); const dur = numVal(r, 'AK'); if (!d || dur == null || dur <= 0) return; const k = monthKey(d); (months[k] || (months[k] = [])).push(dur); });
  const sorted = Object.keys(months).sort();
  const totalData = sorted.map(k => sum(months[k]).toFixed(2));
  const avgData = sorted.map(k => mean(months[k]).toFixed(3));
  const tblRows = sorted.map((k, i) => [k, totalData[i], avgData[i]]);
  return chartCanvas('ac23') + renderTable(['月份','总维修时长','平均维修时长'], tblRows) +
    '<script>Charts.line("ac23",' + JSON.stringify(sorted) + ',[{label:"总维修时长",data:' + JSON.stringify(totalData) + '},{label:"平均维修时长",data:' + JSON.stringify(avgData) + '}])</script>' +
    exportCsvBtn('monthly-dur-trend', ['月份','总维修时长','平均维修时长'], tblRows);
}

function execMonthlyCostTrend(rows) {
  const months = {};
  rows.forEach(r => { const d = dateVal(r, 'Y'); if (!d) return; const k = monthKey(d); if (!months[k]) months[k] = { an: 0, ap: 0, aq: 0 }; months[k].an += numVal(r, 'AN') || 0; months[k].ap += numVal(r, 'AP') || 0; months[k].aq += numVal(r, 'AQ') || 0; });
  const sorted = Object.keys(months).sort();
  const anData = sorted.map(k => +months[k].an.toFixed(2));
  const apData = sorted.map(k => +months[k].ap.toFixed(2));
  const aqData = sorted.map(k => +months[k].aq.toFixed(2));
  const tblRows = sorted.map((k, i) => [k, anData[i], apData[i], aqData[i]]);
  return chartCanvas('ac24') + renderTable(['月份','备件成本','委外费用','总成本'], tblRows) +
    '<script>Charts.line("ac24",' + JSON.stringify(sorted) + ',[{label:"备件成本",data:' + JSON.stringify(anData) + ',fill:true},{label:"委外费用",data:' + JSON.stringify(apData) + ',fill:true},{label:"总成本",data:' + JSON.stringify(aqData) + '}])</script>' +
    exportCsvBtn('monthly-cost-trend', ['月份','备件成本','委外费用','总成本'], tblRows);
}

function execWsMonthlyTrend(rows) {
  const topWs = sortEntries(countByCol(rows, 'F')).slice(0, 5).map(e => e[0]);
  const months = {};
  rows.forEach(r => { const d = dateVal(r, 'Y'); const ws = strVal(r, 'F'); if (!d || !topWs.includes(ws)) return; const k = monthKey(d); if (!months[k]) months[k] = {}; months[k][ws] = (months[k][ws] || 0) + 1; });
  const sorted = Object.keys(months).sort();
  const datasets = topWs.map((ws, i) => ({ label: ws, data: sorted.map(k => months[k][ws] || 0) }));
  const tblRows = sorted.map((k, i) => [k, ...topWs.map(ws => months[k][ws] || 0)]);
  return chartCanvas('ac25') + renderTable(['月份', ...topWs], tblRows) +
    '<script>Charts.line("ac25",' + JSON.stringify(sorted) + ',' + JSON.stringify(datasets) + ')</script>' +
    exportCsvBtn('ws-monthly-trend', ['月份', ...topWs], tblRows);
}

function execShiftMonthlyTrend(rows) {
  const months = {};
  rows.forEach(r => { const d = dateVal(r, 'Y'); const s = strVal(r, 'S'); if (!d || !s) return; const k = monthKey(d); if (!months[k]) months[k] = {}; months[k][s] = (months[k][s] || 0) + 1; });
  const sorted = Object.keys(months).sort();
  const shifts = [...new Set(rows.map(r => strVal(r, 'S')).filter(Boolean))].sort();
  const datasets = shifts.map(s => ({ label: s, data: sorted.map(k => (months[k][s] || 0)) }));
  const tblRows = sorted.map((k, i) => [k, shifts.map(s => months[k][s] || 0)].flat());
  return chartCanvas('ac26') + renderTable(['月份', ...shifts], tblRows) +
    '<script>Charts.line("ac26",' + JSON.stringify(sorted) + ',' + JSON.stringify(datasets) + ')</script>' +
    exportCsvBtn('shift-monthly-trend', ['月份', ...shifts], tblRows);
}

function execFaultLevelTrend(rows) {
  const months = {};
  rows.forEach(r => { const d = dateVal(r, 'Y'); const p = strVal(r, 'P'); if (!d || !p) return; const k = monthKey(d); if (!months[k]) months[k] = {}; months[k][p] = (months[k][p] || 0) + 1; });
  const sorted = Object.keys(months).sort();
  const levels = [...new Set(rows.map(r => strVal(r, 'P')).filter(Boolean))].sort();
  const datasets = levels.map(l => ({ label: l, data: sorted.map(k => months[k][l] || 0) }));
  const tblRows = sorted.map(k => [k, ...levels.map(l => months[k][l] || 0)]);
  return chartCanvas('ac27') + renderTable(['月份', ...levels], tblRows) +
    '<script>Charts.bar("ac27",' + JSON.stringify(sorted) + ',' + JSON.stringify(datasets) + ')</script>' +
    exportCsvBtn('fault-level-trend', ['月份', ...levels], tblRows);
}

function execScoreMonthlyTrend(rows) {
  const months = {};
  rows.forEach(r => { const d = dateVal(r, 'Y'); const s = numVal(r, 'BD'); if (!d || s == null) return; const k = monthKey(d); (months[k] || (months[k] = [])).push(s); });
  const sorted = Object.keys(months).sort();
  const data = sorted.map(k => mean(months[k]).toFixed(2));
  const tblRows = sorted.map((k, i) => [k, data[i]]);
  return chartCanvas('ac28') + renderTable(['月份','平均总分'], tblRows) +
    '<script>Charts.line("ac28",' + JSON.stringify(sorted) + ',[{label:"平均总分",data:' + JSON.stringify(data) + '}])</script>' +
    exportCsvBtn('score-monthly-trend', ['月份','平均总分'], tblRows);
}

/* ===== ④ 设备可靠性 ===== */
function execMTBF(rows) {
  const groups = groupBy(rows, 'L');
  const entries = Object.entries(groups).map(([equip, rs]) => {
    const dates = rs.map(r => dateVal(r, 'Y')).filter(d => d).sort((a, b) => a - b);
    if (dates.length < 2) return [equip, dates.length, null];
    const totalDays = (dates[dates.length - 1] - dates[0]) / 86400000;
    const mtbf = totalDays / (dates.length - 1);
    return [equip, dates.length, mtbf.toFixed(2)];
  }).filter(e => e[2] != null).sort((a, b) => b[2] - a[2]);
  const tblRows = entries.slice(0, 50).map((e, i) => [i + 1, e[0], e[1], e[2]]);
  return '<p class="analysis-info">MTBF = 总运行天数 / (故障次数-1)，单位：天</p>' + renderTable(['排名','设备','故障次数','MTBF(天)'], tblRows, 50) + exportCsvBtn('mtbf', ['排名','设备','故障次数','MTBF(天)'], tblRows);
}

function execMTTR(rows) {
  const groups = groupBy(rows, 'L');
  const entries = Object.entries(groups).map(([equip, rs]) => {
    const durs = rs.map(r => numVal(r, 'AK')).filter(v => v != null && v > 0);
    if (!durs.length) return [equip, 0, null, null];
    return [equip, durs.length, mean(durs).toFixed(3), ss.median(durs).toFixed(3)];
  }).filter(e => e[2] != null).sort((a, b) => b[2] - a[2]);
  const tblRows = entries.slice(0, 50).map((e, i) => [i + 1, e[0], e[1], e[2], e[3]]);
  return '<p class="analysis-info">MTTR = 平均维修时间，单位：小时</p>' + renderTable(['排名','设备','维修次数','MTTR(均值)','MTTR(中位数)'], tblRows, 50) + exportCsvBtn('mttr', ['排名','设备','维修次数','MTTR(均值)','MTTR(中位数)'], tblRows);
}

function execAvailability(rows) {
  const groups = groupBy(rows, 'L');
  const entries = Object.entries(groups).map(([equip, rs]) => {
    const dates = rs.map(r => dateVal(r, 'Y')).filter(d => d).sort((a, b) => a - b);
    const durs = rs.map(r => numVal(r, 'AK')).filter(v => v != null && v > 0);
    if (dates.length < 2 || !durs.length) return [equip, dates.length, null, null, null];
    const totalDays = (dates[dates.length - 1] - dates[0]) / 86400000;
    const mtbf = totalDays / (dates.length - 1);
    const mttr = mean(durs) / 24;
    const avail = mtbf / (mtbf + mttr);
    return [equip, dates.length, mtbf.toFixed(2), mttr.toFixed(4), (avail * 100).toFixed(2) + '%'];
  }).filter(e => e[2] != null).sort((a, b) => parseFloat(b[4]) - parseFloat(a[4]));
  const tblRows = entries.slice(0, 50).map((e, i) => [i + 1, e[0], e[1], e[2], e[3], e[4]]);
  return '<p class="analysis-info">可用度 = MTBF / (MTBF + MTTR)，MTBF单位天，MTTR单位天</p>' + renderTable(['排名','设备','故障次数','MTBF(天)','MTTR(天)','可用度'], tblRows, 50) + exportCsvBtn('availability', ['排名','设备','故障次数','MTBF(天)','MTTR(天)','可用度'], tblRows);
}

function execRepeatFault(rows) {
  const repeats = rows.filter(r => { const v = strVal(r, 'R'); return v && v !== ''; });
  const counts = countByCol(repeats, 'L');
  const entries = sortEntries(counts).slice(0, 20);
  const tblRows = entries.map((e, i) => [i + 1, e[0], e[1]]);
  return '<p class="analysis-info">共 ' + repeats.length + ' 条重复故障记录（R列非空）</p>' + renderTable(['排名','设备','重复故障次数'], tblRows) + exportCsvBtn('repeat-fault', ['排名','设备','重复故障次数'], tblRows);
}

function execEquipFaultTrend(rows) {
  const topEquip = sortEntries(countByCol(rows, 'L')).slice(0, 5).map(e => e[0]);
  const months = {};
  rows.forEach(r => { const d = dateVal(r, 'Y'); const eq = strVal(r, 'L'); if (!d || !topEquip.includes(eq)) return; const k = monthKey(d); if (!months[k]) months[k] = {}; months[k][eq] = (months[k][eq] || 0) + 1; });
  const sorted = Object.keys(months).sort();
  const datasets = topEquip.map(eq => ({ label: eq, data: sorted.map(k => months[k][eq] || 0) }));
  const tblRows = sorted.map(k => [k, ...topEquip.map(eq => months[k][eq] || 0)]);
  return chartCanvas('ac33') + renderTable(['月份', ...topEquip], tblRows) +
    '<script>Charts.line("ac33",' + JSON.stringify(sorted) + ',' + JSON.stringify(datasets) + ')</script>' +
    exportCsvBtn('equip-fault-trend', ['月份', ...topEquip], tblRows);
}

function execEquipFaultMode(rows) {
  const topEquip = sortEntries(countByCol(rows, 'L')).slice(0, 10).map(e => e[0]);
  const causes = sortEntries(countByCol(rows, 'W')).filter(e => e[0]).slice(0, 10).map(e => e[0]);
  const matrix = crossTab(rows, 'L', topEquip, 'W', causes);
  const tblRows = topEquip.map((eq, i) => [eq, ...matrix[i]]);
  return '<p class="analysis-info">设备 × 故障原因 交叉表（Top 10 × Top 10）</p>' + renderTable(['设备', ...causes], tblRows) + exportCsvBtn('equip-fault-mode', ['设备', ...causes], tblRows);
}

function execEquipPartFreq(rows) {
  const topEquip = sortEntries(countByCol(rows, 'L')).slice(0, 10).map(e => e[0]);
  const parts = sortEntries(countByCol(rows, 'AE')).filter(e => e[0]).slice(0, 10).map(e => e[0]);
  const matrix = crossTab(rows, 'L', topEquip, 'AE', parts);
  const tblRows = topEquip.map((eq, i) => [eq, ...matrix[i]]);
  return '<p class="analysis-info">设备 × 故障部件 交叉表（Top 10 × Top 10）</p>' + renderTable(['设备', ...parts], tblRows) + exportCsvBtn('equip-part-freq', ['设备', ...parts], tblRows);
}

function execEquipRepairCat(rows) {
  const topEquip = sortEntries(countByCol(rows, 'L')).slice(0, 15).map(e => e[0]);
  const cats = [...new Set(rows.map(r => strVal(r, 'AD')).filter(Boolean))];
  const data = crossTab(rows, 'L', topEquip, 'AD', cats);
  const datasets = cats.map((c, ci) => ({ label: c, data: data.map(d => d[ci]) }));
  const tblRows = topEquip.map((eq, i) => [eq, ...data[i]]);
  return chartCanvas('ac36') + renderTable(['设备', ...cats], tblRows) +
    '<script>Charts.horizontalBar("ac36",' + JSON.stringify(topEquip) + ',' + JSON.stringify(datasets) + ')</script>' +
    exportCsvBtn('equip-repair-cat', ['设备', ...cats], tblRows);
}

function execNewEquipFault(rows) {
  const firstSeen = {};
  rows.forEach(r => { const code = strVal(r, 'M'); const d = dateVal(r, 'Y'); if (!code || !d) return; if (!firstSeen[code] || d < firstSeen[code]) firstSeen[code] = d; });
  const months = {};
  rows.forEach(r => { const code = strVal(r, 'M'); const d = dateVal(r, 'Y'); if (!code || !d || !firstSeen[code]) return; const diff = (d - firstSeen[code]) / 86400000; const bucket = diff < 30 ? '0-30天' : diff < 90 ? '30-90天' : diff < 180 ? '90-180天' : '180天+'; months[bucket] = (months[bucket] || 0) + 1; });
  const order = ['0-30天', '30-90天', '90-180天', '180天+'];
  const labels = order.filter(k => months[k]);
  const data = labels.map(k => months[k]);
  const tblRows = labels.map((k, i) => [k, data[i]]);
  return chartCanvas('ac37') + renderTable(['首次故障时间段','故障次数'], tblRows) +
    '<script>Charts.bar("ac37",' + JSON.stringify(labels) + ',[{label:"故障次数",data:' + JSON.stringify(data) + '}])</script>' +
    exportCsvBtn('new-equip-fault', ['首次故障时间段','故障次数'], tblRows);
}

function execEquipHealth(rows) {
  const groups = groupBy(rows, 'L');
  const raw = Object.entries(groups).map(([equip, rs]) => {
    const freq = rs.length;
    const durs = rs.map(r => numVal(r, 'AK')).filter(v => v != null && v > 0);
    const avgDur = durs.length ? mean(durs) : 0;
    const costs = rs.map(r => numVal(r, 'AQ')).filter(v => v != null && v > 0);
    const avgCost = costs.length ? mean(costs) : 0;
    return { equip, freq, avgDur, avgCost };
  });
  const maxFreq = Math.max(...raw.map(d => d.freq)) || 1;
  const maxDur = Math.max(...raw.map(d => d.avgDur)) || 1;
  const maxCost = Math.max(...raw.map(d => d.avgCost)) || 1;
  const entries = raw.map(d => {
    const score = (d.freq / maxFreq) * 0.4 + (d.avgDur / maxDur) * 0.3 + (d.avgCost / maxCost) * 0.3;
    return [d.equip, d.freq, d.avgDur.toFixed(2), d.avgCost.toFixed(2), (score * 100).toFixed(2)];
  }).sort((a, b) => parseFloat(b[4]) - parseFloat(a[4]));
  const tblRows = entries.slice(0, 30).map((e, i) => [i + 1, e[0], e[1], e[2], e[3], e[4]]);
  return '<p class="analysis-info">风险度评分 = 频率归一×0.4 + 时长归一×0.3 + 成本归一×0.3（各维度按最大值归一化到0-1，再乘100），分数越高越需关注</p>' + renderTable(['排名','设备','故障次数','平均维修时长','平均成本','风险度评分'], tblRows, 30) + exportCsvBtn('equip-health', ['排名','设备','故障次数','平均维修时长','平均成本','风险度评分'], tblRows);
}

/* ===== ⑤ 维修效率 ===== */
function execWaitVsRepair(rows) {
  const data = rows.map(r => ({ wait: numVal(r, 'AI') || 0, repair: numVal(r, 'AK') || 0 })).filter(d => d.wait > 0 || d.repair > 0);
  const waitSum = sum(data.map(d => d.wait)), repairSum = sum(data.map(d => d.repair));
  const total = waitSum + repairSum || 1;
  const tblRows = [['总等待时长', waitSum.toFixed(2)], ['总维修时长', repairSum.toFixed(2)], ['等待占比', fmtPct(waitSum / total)], ['维修占比', fmtPct(repairSum / total)], ['平均等待', mean(data.map(d => d.wait)).toFixed(3)], ['平均维修', mean(data.map(d => d.repair)).toFixed(3)]];
  return chartCanvas('ac39') + renderTable(['指标','值'], tblRows) +
    '<script>Charts.doughnut("ac39",["总等待时长","总维修时长"],[' + waitSum.toFixed(2) + ',' + repairSum.toFixed(2) + '])</script>' +
    exportCsvBtn('wait-vs-repair', ['指标','值'], tblRows);
}

function execShiftEfficiency(rows) {
  const groups = groupBy(rows, 'S');
  const entries = Object.entries(groups).map(([s, rs]) => { const durs = rs.map(r => numVal(r, 'AK')).filter(v => v != null && v > 0); if (!durs.length) return null; return [s, durs.length, mean(durs).toFixed(3), ss.median(durs).toFixed(3), ss.min(durs).toFixed(3), ss.max(durs).toFixed(3)]; }).filter(e => e && e[0]);
  const tblRows = entries;
  return chartCanvas('ac40') + renderTable(['班次','样本数','平均维修时长','中位数','最小值','最大值'], tblRows) +
    '<script>Charts.bar("ac40",' + JSON.stringify(entries.map(e => e[0])) + ',[{label:"平均维修时长",data:' + JSON.stringify(entries.map(e => +e[2])) + '},{label:"中位数",data:' + JSON.stringify(entries.map(e => +e[3])) + '}])</script>' +
    exportCsvBtn('shift-efficiency', ['班次','样本数','平均维修时长','中位数','最小值','最大值'], tblRows);
}

function execRepairerCompare(rows) {
  const groups = groupBy(rows, 'X');
  const entries = Object.entries(groups).map(([p, rs]) => { const durs = rs.map(r => numVal(r, 'AK')).filter(v => v != null && v > 0); if (!durs.length) return null; return [p, durs.length, mean(durs).toFixed(3), ss.median(durs).toFixed(3)]; }).filter(e => e && e[0] && e[1] >= 5).sort((a, b) => b[2] - a[2]).slice(0, 20);
  const tblRows = entries.map((e, i) => [i + 1, ...e]);
  return chartCanvas('ac41') + renderTable(['排名','维修人','样本数','平均维修时长','中位数'], tblRows) +
    '<script>Charts.horizontalBar("ac41",' + JSON.stringify(entries.map(e => e[0])) + ',[{label:"平均维修时长",data:' + JSON.stringify(entries.map(e => +e[2])) + '}])</script>' +
    exportCsvBtn('repairer-compare', ['排名','维修人','样本数','平均维修时长','中位数'], tblRows);
}

function execResponseTime(rows) {
  const stages = rows.map(r => { const y = dateVal(r, 'Y'), v = dateVal(r, 'V'), af = dateVal(r, 'AF'); if (!y) return null; return { applyToRecv: v ? (v - y) / 3600000 : null, recvToStart: (af && v) ? (af - v) / 3600000 : null, total: af ? (af - y) / 3600000 : null }; }).filter(Boolean);
  const safeMed = arr => arr.length ? ss.median(arr).toFixed(2) : '—';
  const safeMean = arr => arr.length ? mean(arr).toFixed(2) : '—';
  const a2r = stages.map(s => s.applyToRecv).filter(Boolean), r2s = stages.map(s => s.recvToStart).filter(Boolean), tot = stages.map(s => s.total).filter(Boolean);
  const stats = [['申请→接收(均值)', safeMean(a2r)], ['申请→接收(中位数)', safeMed(a2r)], ['接收→开始(均值)', safeMean(r2s)], ['接收→开始(中位数)', safeMed(r2s)], ['申请→开始(均值)', safeMean(tot)], ['申请→开始(中位数)', safeMed(tot)]];
  return chartCanvas('ac42') + renderTable(['阶段','时长(小时)'], stats) +
    '<script>Charts.bar("ac42",["申请→接收(均值)","接收→开始(均值)","申请→开始(均值)"],[{label:"小时",data:[' + (+a2r.length ? mean(a2r).toFixed(2) : 0) + ',' + (+r2s.length ? mean(r2s).toFixed(2) : 0) + ',' + (+tot.length ? mean(tot).toFixed(2) : 0) + ']}])</script>' +
    exportCsvBtn('response-time', ['阶段','时长(小时)'], stats);
}

function execOvertime(rows) {
  const durs = rows.map(r => ({ id: strVal(r, 'A'), equip: strVal(r, 'L'), dur: numVal(r, 'AK') || 0, ws: strVal(r, 'F') })).filter(d => d.dur > 0);
  if (!durs.length) return '<p class="no-data">无有效维修时长数据</p>';
  const sorted = durs.sort((a, b) => b.dur - a.dur);
  const ascDurs = durs.map(d => d.dur).sort((a, b) => a - b);
  const p90 = ss.quantile(ascDurs, 0.9);
  const p95 = ss.quantile(ascDurs, 0.95);
  const overP90 = sorted.filter(d => d.dur >= p90);
  const tblRows = [['P90阈值', p90.toFixed(3)], ['P95阈值', p95.toFixed(3)], ['超P90维修单数', overP90.length], ['超P95维修单数', sorted.filter(d => d.dur >= p95).length]];
  const topRows = sorted.slice(0, 20).map((d, i) => [i + 1, d.id, d.ws, d.equip, d.dur.toFixed(3)]);
  return renderTable(['指标','值'], tblRows) + '<h4>Top 20 超长维修单</h4>' + renderTable(['排名','维修单号','车间','设备','维修时长'], topRows) + exportCsvBtn('overtime', ['排名','维修单号','车间','设备','维修时长'], topRows);
}

function execRejectRate(rows) {
  const rejects = rows.filter(r => { const s = strVal(r, 'B'); return s === '已驳回' || s === '驳回关闭'; });
  const groups = groupBy(rows, 'F');
  const entries = Object.entries(groups).map(([ws, rs]) => { const rej = rs.filter(r => { const s = strVal(r, 'B'); return s === '已驳回' || s === '驳回关闭'; }).length; return [ws, rs.length, rej, (rej / rs.length * 100).toFixed(2) + '%']; }).sort((a, b) => parseFloat(b[3]) - parseFloat(a[3]));
  const totalRate = (rejects.length / rows.length * 100).toFixed(2) + '%';
  return '<p class="analysis-info">总驳回率：' + totalRate + '（' + rejects.length + '/' + rows.length + '）</p>' + renderTable(['车间','总维修单','驳回数','驳回率'], entries) + exportCsvBtn('reject-rate', ['车间','总维修单','驳回数','驳回率'], entries);
}

function execPhaseBreakdown(rows) {
  const phases = rows.map(r => { const y = dateVal(r, 'Y'), aa = dateVal(r, 'AA'), af = dateVal(r, 'AF'), ag = dateVal(r, 'AG'); return { wait: aa ? (aa - y) / 3600000 : null, confirm: af && aa ? (af - aa) / 3600000 : null, repair: ag && af ? (ag - af) / 3600000 : null }; }).filter(p => p.wait != null || p.confirm != null || p.repair != null);
  const stats = [['等待确认(均值)', mean(phases.map(p => p.wait).filter(Boolean)).toFixed(2)], ['确认→开始(均值)', mean(phases.map(p => p.confirm).filter(Boolean)).toFixed(2)], ['维修阶段(均值)', mean(phases.map(p => p.repair).filter(Boolean)).toFixed(2)]];
  return chartCanvas('ac45') + renderTable(['阶段','时长(小时)'], stats) +
    '<script>Charts.bar("ac45",["等待确认","确认→开始","维修阶段"],[{label:"小时",data:[' + stats[0][1] + ',' + stats[1][1] + ',' + stats[2][1] + ']}])</script>' +
    exportCsvBtn('phase-breakdown', ['阶段','时长(小时)'], stats);
}

function execTimeliness(rows) {
  const vals = rows.map(r => numVal(r, 'BA')).filter(v => v != null && v > 0);
  if (!vals.length) return '<p class="no-data">无及时性数据</p>';
  const stats = [['样本数', vals.length], ['均值', mean(vals).toFixed(3)], ['中位数', ss.median(vals).toFixed(3)], ['标准差', ss.standardDeviation(vals).toFixed(3)], ['P90', ss.quantile(vals.sort((a, b) => a - b), 0.9).toFixed(3)], ['P95', ss.quantile(vals.sort((a, b) => a - b), 0.95).toFixed(3)]];
  const h = histogram(vals, 15);
  return chartCanvas('ac46') + renderTable(['统计量','值'], stats) +
    '<script>Charts.bar("ac46",' + JSON.stringify(h.labels) + ',[{label:"频次",data:' + JSON.stringify(h.data) + '}])</script>' +
    exportCsvBtn('timeliness', ['统计量','值'], stats);
}

/* ===== ⑥ 成本与影响 ===== */
function execWsCost(rows) {
  const groups = groupBy(rows, 'F');
  const entries = Object.entries(groups).map(([ws, rs]) => [ws, sum(rs.map(r => numVal(r, 'AN') || 0)).toFixed(2), sum(rs.map(r => numVal(r, 'AP') || 0)).toFixed(2), sum(rs.map(r => numVal(r, 'AQ') || 0)).toFixed(2)]).sort((a, b) => parseFloat(b[3]) - parseFloat(a[3]));
  return chartCanvas('ac47') + renderTable(['车间','备件成本','委外费用','总成本'], entries) +
    '<script>Charts.horizontalBar("ac47",' + JSON.stringify(entries.map(e => e[0])) + ',[{label:"备件成本",data:' + JSON.stringify(entries.map(e => +e[1])) + '},{label:"委外费用",data:' + JSON.stringify(entries.map(e => +e[2])) + '},{label:"总成本",data:' + JSON.stringify(entries.map(e => +e[3])) + '}])</script>' +
    exportCsvBtn('ws-cost', ['车间','备件成本','委外费用','总成本'], entries);
}

function execPartsVsOutsource(rows) {
  const anSum = sum(rows.map(r => numVal(r, 'AN') || 0));
  const apSum = sum(rows.map(r => numVal(r, 'AP') || 0));
  const total = anSum + apSum || 1;
  const tblRows = [['备件成本总额', anSum.toFixed(2)], ['委外费用总额', apSum.toFixed(2)], ['总计', (anSum + apSum).toFixed(2)], ['备件占比', fmtPct(anSum / total)], ['委外占比', fmtPct(apSum / total)]];
  return chartCanvas('ac48') + renderTable(['指标','值'], tblRows) +
    '<script>Charts.doughnut("ac48",["备件成本","委外费用"],[' + anSum.toFixed(2) + ',' + apSum.toFixed(2) + '])</script>' +
    exportCsvBtn('parts-vs-outsource', ['指标','值'], tblRows);
}

function execCapacityImpact(rows) {
  const affected = rows.filter(r => strVal(r, 'AR') === '是');
  const totalDur = sum(affected.map(r => numVal(r, 'AT') || 0));
  const totalQty = sum(affected.map(r => numVal(r, 'AU') || 0));
  const groups = groupBy(affected, 'F');
  const entries = Object.entries(groups).map(([ws, rs]) => [ws, rs.length, sum(rs.map(r => numVal(r, 'AT') || 0)).toFixed(2), sum(rs.map(r => numVal(r, 'AU') || 0)).toFixed(2)]).sort((a, b) => b[1] - a[1]);
  const summary = [['影响产能维修单数', affected.length], ['总产量影响时长(h)', totalDur.toFixed(2)], ['总产能影响数量(kg)', totalQty.toFixed(2)]];
  return renderTable(['指标','值'], summary) + '<h4>按车间分布</h4>' + renderTable(['车间','维修单数','影响时长(h)','影响数量(kg)'], entries) + exportCsvBtn('capacity-impact', ['车间','维修单数','影响时长(h)','影响数量(kg)'], entries);
}

function execQualityImpact(rows) {
  const affected = rows.filter(r => strVal(r, 'AS') === '是');
  const groups = groupBy(affected, 'F');
  const entries = Object.entries(groups).map(([ws, rs]) => [ws, rs.length, sum(rs.map(r => numVal(r, 'AQ') || 0)).toFixed(2)]).sort((a, b) => b[1] - a[1]);
  return '<p class="analysis-info">共 ' + affected.length + ' 条影响质量的维修单</p>' + renderTable(['车间','维修单数','维修总成本'], entries) + exportCsvBtn('quality-impact', ['车间','维修单数','维修总成本'], entries);
}

function execCostTrend(rows) {
  const months = {};
  rows.forEach(r => { const d = dateVal(r, 'Y'); const c = numVal(r, 'AQ'); if (!d || c == null) return; const k = monthKey(d); months[k] = (months[k] || 0) + c; });
  const sorted = Object.keys(months).sort();
  const data = sorted.map(k => +months[k].toFixed(2));
  const tblRows = sorted.map((k, i) => [k, data[i]]);
  return chartCanvas('ac51') + renderTable(['月份','维修总成本'], tblRows) +
    '<script>Charts.line("ac51",' + JSON.stringify(sorted) + ',[{label:"维修总成本",data:' + JSON.stringify(data) + ',fill:true}])</script>' +
    exportCsvBtn('cost-trend', ['月份','维修总成本'], tblRows);
}

function execHighCostOutlier(rows) {
  const costs = rows.map(r => numVal(r, 'AQ')).filter(v => v != null && v > 0).sort((a, b) => a - b);
  if (!costs.length) return '<p class="no-data">无有效维修成本数据</p>';
  const p95 = ss.quantile(costs, 0.95);
  const outliers = rows.filter(r => { const c = numVal(r, 'AQ'); return c != null && c >= p95; }).map(r => ({ id: strVal(r, 'A'), ws: strVal(r, 'F'), equip: strVal(r, 'L'), cost: numVal(r, 'AQ'), desc: strVal(r, 'T').slice(0, 40) }));
  const tblRows = [['P95阈值', p95.toFixed(2)], ['超阈值维修单数', outliers.length]];
  const outlierRows = outliers.sort((a, b) => b.cost - a.cost).slice(0, 20).map((o, i) => [i + 1, o.id, o.ws, o.equip, o.cost.toFixed(2), o.desc]);
  return renderTable(['指标','值'], tblRows) + '<h4>Top 20 高成本维修单</h4>' + renderTable(['排名','维修单号','车间','设备','维修总价值','问题描述'], outlierRows) + exportCsvBtn('high-cost-outlier', ['排名','维修单号','车间','设备','维修总价值','问题描述'], outlierRows);
}

function execCostVsLevel(rows) {
  const groups = groupBy(rows, 'P');
  const entries = Object.entries(groups).map(([level, rs]) => { const costs = rs.map(r => numVal(r, 'AQ')).filter(v => v != null && v > 0); if (!costs.length) return null; return [level, costs.length, mean(costs).toFixed(2), ss.median(costs).toFixed(2), ss.min(costs).toFixed(2), ss.max(costs).toFixed(2)]; }).filter(e => e && e[0] && e[1] > 0);
  const tblRows = entries;
  return chartCanvas('ac53') + renderTable(['故障等级','样本数','平均成本','中位数','最小值','最大值'], tblRows) +
    '<script>Charts.bar("ac53",' + JSON.stringify(entries.map(e => e[0])) + ',[{label:"平均成本",data:' + JSON.stringify(entries.map(e => +e[2])) + '},{label:"中位数",data:' + JSON.stringify(entries.map(e => +e[3])) + '}])</script>' +
    exportCsvBtn('cost-vs-level', ['故障等级','样本数','平均成本','中位数','最小值','最大值'], tblRows);
}

/* ===== ⑦ 评分与质量 ===== */
function execScoreDist(rows) {
  const cols = [['AX', '维修质量评分'], ['AY', '现场维修评分'], ['AZ', '重复维修评分'], ['BA', '接单结单及时性']];
  const stats = cols.map(([col, name]) => { const vals = rows.map(r => numVal(r, col)).filter(v => v != null); if (!vals.length) return [name, 0, '—', '—', '—']; return [name, vals.length, mean(vals).toFixed(3), ss.median(vals).toFixed(3), ss.standardDeviation(vals).toFixed(3)]; });
  return chartCanvas('ac54') + renderTable(['评分维度','样本数','均值','中位数','标准差'], stats) +
    '<script>Charts.bar("ac54",' + JSON.stringify(cols.map(c => c[1])) + ',[{label:"均值",data:' + JSON.stringify(stats.map(s => isNaN(+s[2]) ? 0 : +s[2])) + '},{label:"中位数",data:' + JSON.stringify(stats.map(s => isNaN(+s[3]) ? 0 : +s[3])) + '}])</script>' +
    exportCsvBtn('score-dist', ['评分维度','样本数','均值','中位数','标准差'], stats);
}

function execScoreCorrelation(rows) {
  const cols = ['AX', 'AY', 'AZ', 'BA', 'BD'];
  const names = ['维修质量', '现场维修', '重复维修', '及时性', '总分'];
  const validRows = rows.filter(r => cols.every(c => numVal(r, c) != null));
  if (validRows.length < 2) return '<p class="no-data">有效数据不足，无法计算相关系数</p>';
  const matrix = cols.map(c1 => cols.map(c2 => { const x = validRows.map(r => numVal(r, c1)); const y = validRows.map(r => numVal(r, c2)); return ss.sampleCorrelation(x, y).toFixed(3); }));
  const tblRows = matrix.map((row, i) => [names[i], ...row]);
  return '<p class="analysis-info">Pearson 相关系数矩阵（' + validRows.length + ' 条有效数据）</p>' + renderTable(['', ...names], tblRows) + exportCsvBtn('score-correlation', ['', ...names], tblRows);
}

function execRepairerScore(rows) {
  const groups = groupBy(rows, 'X');
  const entries = Object.entries(groups).map(([p, rs]) => { const scores = rs.map(r => numVal(r, 'BD')).filter(v => v != null); if (!scores.length) return null; return [p, scores.length, mean(scores).toFixed(3), ss.median(scores).toFixed(3)]; }).filter(e => e && e[0] && e[1] >= 5).sort((a, b) => parseFloat(b[2]) - parseFloat(a[2])).slice(0, 30);
  const tblRows = entries.map((e, i) => [i + 1, ...e]);
  return chartCanvas('ac56') + renderTable(['排名','维修人','评分样本数','平均总分','中位数'], tblRows, 30) +
    '<script>Charts.horizontalBar("ac56",' + JSON.stringify(entries.map(e => e[0])) + ',[{label:"平均总分",data:' + JSON.stringify(entries.map(e => +e[2])) + '}])</script>' +
    exportCsvBtn('repairer-score', ['排名','维修人','评分样本数','平均总分','中位数'], tblRows);
}

function execWsScore(rows) {
  const groups = groupBy(rows, 'F');
  const entries = Object.entries(groups).map(([ws, rs]) => { const scores = rs.map(r => numVal(r, 'BD')).filter(v => v != null); if (!scores.length) return null; return [ws, scores.length, mean(scores).toFixed(3)]; }).filter(e => e && e[1] >= 5).sort((a, b) => parseFloat(b[2]) - parseFloat(a[2]));
  const tblRows = entries;
  return chartCanvas('ac57') + renderTable(['车间','评分样本数','平均总分'], tblRows) +
    '<script>Charts.bar("ac57",' + JSON.stringify(entries.map(e => e[0])) + ',[{label:"平均总分",data:' + JSON.stringify(entries.map(e => +e[2])) + '}])</script>' +
    exportCsvBtn('ws-score', ['车间','评分样本数','平均总分'], tblRows);
}

function execLowScore(rows) {
  const low = rows.filter(r => { const s = numVal(r, 'BD'); return s != null && s < 3; });
  const groups = groupBy(low, 'F');
  const entries = Object.entries(groups).map(([ws, rs]) => [ws, rs.length]).sort((a, b) => b[1] - a[1]);
  const summary = [['低分维修单数(<3)', low.length], ['占总维修单比', fmtPct(low.length / rows.length)]];
  return renderTable(['指标','值'], summary) + '<h4>按车间分布</h4>' + renderTable(['车间','低分数'], entries) + exportCsvBtn('low-score', ['车间','低分数'], entries);
}

function execScoreTrend(rows) {
  const months = {};
  rows.forEach(r => { const d = dateVal(r, 'Y'); const s = numVal(r, 'BD'); if (!d || s == null) return; const k = monthKey(d); (months[k] || (months[k] = [])).push(s); });
  const sorted = Object.keys(months).sort();
  const data = sorted.map(k => mean(months[k]).toFixed(3));
  const tblRows = sorted.map((k, i) => [k, data[i]]);
  return chartCanvas('ac59') + renderTable(['月份','平均总分'], tblRows) +
    '<script>Charts.line("ac59",' + JSON.stringify(sorted) + ',[{label:"平均总分",data:' + JSON.stringify(data) + '}])</script>' +
    exportCsvBtn('score-trend', ['月份','平均总分'], tblRows);
}

/* ===== ⑧ 相关性分析 ===== */
function execDurVsLevel(rows) {
  const groups = groupBy(rows, 'P');
  const entries = Object.entries(groups).map(([level, rs]) => { const durs = rs.map(r => numVal(r, 'AK')).filter(v => v != null && v > 0); return [level, durs.length, mean(durs).toFixed(3), ss.median(durs).toFixed(3)]; }).filter(e => e[0] && e[1] > 0);
  const tblRows = entries;
  return chartCanvas('ac60') + renderTable(['故障等级','样本数','平均维修时长','中位数'], tblRows) +
    '<script>Charts.bar("ac60",' + JSON.stringify(entries.map(e => e[0])) + ',[{label:"平均维修时长",data:' + JSON.stringify(entries.map(e => +e[2])) + '},{label:"中位数",data:' + JSON.stringify(entries.map(e => +e[3])) + '}])</script>' +
    exportCsvBtn('dur-vs-level', ['故障等级','样本数','平均维修时长','中位数'], tblRows);
}

function execWaitVsScore(rows) {
  const data = rows.map(r => ({ wait: numVal(r, 'AI'), score: numVal(r, 'BD') })).filter(d => d.wait != null && d.score != null);
  if (!data.length) return '<p class="no-data">无有效数据</p>';
  const corr = ss.sampleCorrelation(data.map(d => d.wait), data.map(d => d.score));
  const sample = data.slice(0, 200).map(d => ({ x: d.wait, y: d.score }));
  const tblRows = [['样本数', data.length], ['Pearson相关系数', corr.toFixed(4)], ['等待时长均值', mean(data.map(d => d.wait)).toFixed(3)], ['评分均值', mean(data.map(d => d.score)).toFixed(3)]];
  return chartCanvas('ac61') + renderTable(['指标','值'], tblRows) +
    '<script>Charts.scatter("ac61",[{label:"等待vs评分",data:' + JSON.stringify(sample) + '}],{scales:{x:{title:{display:true,text:"等待时长"}},y:{title:{display:true,text:"维修评分"}}}})</script>' +
    exportCsvBtn('wait-vs-score', ['指标','值'], tblRows);
}

function execFreqVsCost(rows) {
  const groups = groupBy(rows, 'L');
  const data = Object.entries(groups).map(([equip, rs]) => ({ x: rs.length, y: sum(rs.map(r => numVal(r, 'AQ') || 0)) })).filter(d => d.y > 0);
  if (data.length < 2) return '<p class="no-data">有效设备数据不足</p>';
  const corr = ss.sampleCorrelation(data.map(d => d.x), data.map(d => d.y));
  const tblRows = [['设备数', data.length], ['Pearson相关系数', corr.toFixed(4)]];
  return chartCanvas('ac62') + renderTable(['指标','值'], tblRows) +
    '<script>Charts.scatter("ac62",[{label:"故障频率vs成本",data:' + JSON.stringify(data) + '}],{scales:{x:{title:{display:true,text:"故障次数"}},y:{title:{display:true,text:"维修总成本"}}}})</script>' +
    exportCsvBtn('freq-vs-cost', ['指标','值'], tblRows);
}

function execShiftVsFault(rows) {
  const shifts = [...new Set(rows.map(r => strVal(r, 'S')).filter(Boolean))];
  const causes = sortEntries(countByCol(rows, 'W')).filter(e => e[0]).slice(0, 10).map(e => e[0]);
  const matrix = crossTab(rows, 'S', shifts, 'W', causes);
  const tblRows = shifts.map((s, i) => [s, ...matrix[i]]);
  return '<p class="analysis-info">班次 × 故障原因 交叉表</p>' + renderTable(['班次', ...causes], tblRows) + exportCsvBtn('shift-vs-fault', ['班次', ...causes], tblRows);
}

function execEquipTypeVsFault(rows) {
  const types = sortEntries(countByCol(rows, 'J')).filter(e => e[0]).slice(0, 10).map(e => e[0]);
  const causes = sortEntries(countByCol(rows, 'W')).filter(e => e[0]).slice(0, 10).map(e => e[0]);
  const matrix = crossTab(rows, 'J', types, 'W', causes);
  const tblRows = types.map((t, i) => [t, ...matrix[i]]);
  return '<p class="analysis-info">设备分类 × 故障原因 交叉表（Top 10 × Top 10）</p>' + renderTable(['设备分类', ...causes], tblRows) + exportCsvBtn('equip-type-vs-fault', ['设备分类', ...causes], tblRows);
}

function execParticipantsVsDur(rows) {
  const data = rows.map(r => { const ah = strVal(r, 'AH'); const cnt = ah ? ah.split(',').length : 1; const dur = numVal(r, 'AK'); return { x: cnt, y: dur }; }).filter(d => d.y != null && d.y > 0);
  if (data.length < 2) return '<p class="no-data">有效数据不足</p>';
  const corr = ss.sampleCorrelation(data.map(d => d.x), data.map(d => d.y));
  const sample = data.slice(0, 200);
  const tblRows = [['样本数', data.length], ['Pearson相关系数', corr.toFixed(4)], ['平均参与人数', mean(data.map(d => d.x)).toFixed(2)], ['平均维修时长', mean(data.map(d => d.y)).toFixed(3)]];
  return chartCanvas('ac65') + renderTable(['指标','值'], tblRows) +
    '<script>Charts.scatter("ac65",[{label:"参与人数vs维修时长",data:' + JSON.stringify(sample) + '}],{scales:{x:{title:{display:true,text:"参与维修人数"}},y:{title:{display:true,text:"维修时长"}}}})</script>' +
    exportCsvBtn('participants-vs-dur', ['指标','值'], tblRows);
}

/* ===== ⑨ 回归与预测 ===== */
function execDurRegression(rows) {
  const data = rows.map(r => { const dur = numVal(r, 'AK'); if (dur == null || dur <= 0) return null; const level = strVal(r, 'P'); return { dur, level, levelNum: level === 'A级' ? 3 : level === 'B级' ? 2 : level === 'C级' ? 1 : 0 }; }).filter(Boolean);
  if (data.length < 2) return '<p class="no-data">有效数据不足，无法进行回归分析</p>';
  const y = data.map(d => d.dur);
  const result = ss.linearRegression(data.map((d, i) => ({ x: d.levelNum, y: y[i] })));
  const corr = ss.sampleCorrelation(data.map(d => d.levelNum), y);
  const tblRows = [['样本数', data.length], ['回归方程', '维修时长 ≈ ' + result.m.toFixed(3) + ' × 故障等级 + ' + result.b.toFixed(3)], ['R²', (corr * corr).toFixed(4)], ['A级平均时长', mean(data.filter(d => d.level === 'A级').map(d => d.dur)).toFixed(3)], ['B级平均时长', mean(data.filter(d => d.level === 'B级').map(d => d.dur)).toFixed(3)], ['C级平均时长', mean(data.filter(d => d.level === 'C级').map(d => d.dur)).toFixed(3)]];
  return '<p class="analysis-info">基于故障等级的线性回归</p>' + renderTable(['指标','值'], tblRows) + exportCsvBtn('dur-regression', ['指标','值'], tblRows);
}

function execCostRegression(rows) {
  const data = rows.map(r => { const cost = numVal(r, 'AQ'); if (cost == null) return null; const dur = numVal(r, 'AK') || 0; return { cost, dur }; }).filter(Boolean);
  if (data.length < 2) return '<p class="no-data">有效数据不足，无法进行回归分析</p>';
  const result = ss.linearRegression(data.map(d => ({ x: d.dur, y: d.cost })));
  const corr = ss.sampleCorrelation(data.map(d => d.dur), data.map(d => d.cost));
  const tblRows = [['样本数', data.length], ['回归方程', '维修成本 ≈ ' + result.m.toFixed(3) + ' × 维修时长 + ' + result.b.toFixed(2)], ['R²', (corr * corr).toFixed(4)], ['相关系数', corr.toFixed(4)]];
  return '<p class="analysis-info">维修时长 → 维修成本的线性回归</p>' + renderTable(['指标','值'], tblRows) + exportCsvBtn('cost-regression', ['指标','值'], tblRows);
}

function execScoreRegression(rows) {
  const cols = ['AX', 'AY', 'AZ', 'BA'];
  const names = ['维修质量评分', '现场维修评分', '重复维修评分', '接单结单及时性'];
  const validRows = rows.filter(r => cols.every(c => numVal(r, c) != null) && numVal(r, 'BD') != null);
  if (validRows.length < 2) return '<p class="no-data">有效数据不足，无法进行回归分析</p>';
  const bd = validRows.map(r => numVal(r, 'BD'));
  const results = cols.map((c, i) => { const x = validRows.map(r => numVal(r, c)); const corr = ss.sampleCorrelation(x, bd); const reg = ss.linearRegression(x.map((xi, j) => ({ x: xi, y: bd[j] }))); return [names[i], corr.toFixed(4), (corr * corr).toFixed(4), reg.m.toFixed(3), reg.b.toFixed(3)]; });
  return '<p class="analysis-info">各评分维度对总分的回归系数（' + validRows.length + ' 条有效数据）</p>' + renderTable(['评分维度','相关系数','R²','斜率','截距'], results) + exportCsvBtn('score-regression', ['评分维度','相关系数','R²','斜率','截距'], results);
}

function execMonthlyForecast(rows) {
  const months = {};
  rows.forEach(r => { const d = dateVal(r, 'Y'); if (!d) return; const k = monthKey(d); months[k] = (months[k] || 0) + 1; });
  const sorted = Object.keys(months).sort();
  const actual = sorted.map(k => months[k]);
  if (!actual.length) return '<p class="no-data">无有效月度数据</p>';
  const ma = actual.map((_, i) => i < 2 ? null : mean(actual.slice(i - 2, i + 1)));
  const alpha = 0.3;
  const es = []; let prev = actual[0];
  actual.forEach((v, i) => { if (i === 0) { es.push(v); return; } prev = alpha * v + (1 - alpha) * prev; es.push(+prev.toFixed(2)); });
  const lastMa = ma[ma.length - 1];
  const forecast = lastMa ? +lastMa.toFixed(1) : null;
  const tblRows = sorted.map((k, i) => [k, actual[i], ma[i] ? ma[i].toFixed(2) : '—', es[i]]);
  tblRows.push(['下月预测', forecast, forecast, forecast]);
  return chartCanvas('ac69') + renderTable(['月份','实际值','3月移动平均','指数平滑(α=0.3)'], tblRows) +
    '<script>Charts.line("ac69",' + JSON.stringify(sorted.concat(['预测'])) + ',[{label:"实际值",data:' + JSON.stringify(actual.concat([null])) + '},{label:"移动平均",data:' + JSON.stringify(ma.concat([forecast])) + '},{label:"指数平滑",data:' + JSON.stringify(es.concat([forecast])) + '}])</script>' +
    exportCsvBtn('monthly-forecast', ['月份','实际值','3月移动平均','指数平滑(α=0.3)'], tblRows);
}

function execFaultProbForecast(rows) {
  const groups = groupBy(rows, 'L');
  const entries = Object.entries(groups).map(([equip, rs]) => {
    const dates = rs.map(r => dateVal(r, 'Y')).filter(d => d).sort((a, b) => a - b);
    if (dates.length < 3) return null;
    const months = {};
    dates.forEach(d => { const k = monthKey(d); months[k] = (months[k] || 0) + 1; });
    const sortedMonths = Object.keys(months).sort();
    const counts = sortedMonths.map(k => months[k]);
    const avgFreq = mean(counts);
    const trend = ss.linearRegression(counts.map((c, i) => ({ x: i, y: c })));
    const nextMonthPred = Math.max(0, trend.m * counts.length + trend.b);
    return [equip, dates.length, avgFreq.toFixed(2), trend.m.toFixed(3), nextMonthPred.toFixed(2)];
  }).filter(Boolean).sort((a, b) => parseFloat(b[4]) - parseFloat(a[4]));
  const tblRows = entries.slice(0, 30).map((e, i) => [i + 1, ...e]);
  return '<p class="analysis-info">基于月度故障频率趋势的线性预测（仅展示前30个设备）</p>' + renderTable(['排名','设备','总故障数','月均故障','趋势斜率','下月预测故障数'], tblRows, 30) + exportCsvBtn('fault-prob-forecast', ['排名','设备','总故障数','月均故障','趋势斜率','下月预测故障数'], tblRows);
}

/* ===== ⑩ 文本分析 ===== */
function extractKeywords(text, minLen) {
  if (!text) return [];
  const ml = minLen || 2;
  const result = [];
  const parts = text.split(/([\u4e00-\u9fa5]+|[a-zA-Z0-9]+)/);
  parts.forEach(part => {
    if (!part) return;
    if (/^[\u4e00-\u9fa5]+$/.test(part)) {
      if (part.length === 1) { result.push(part); return; }
      for (let i = 0; i < part.length - 1; i++) result.push(part.slice(i, i + 2));
    } else if (/^[a-zA-Z0-9]+$/.test(part) && part.length >= ml) {
      result.push(part);
    }
  });
  return result;
}

function wordFreq(rows, col, topN) {
  const freq = {};
  rows.forEach(r => { const text = strVal(r, col); if (!text) return; const words = extractKeywords(text); words.forEach(w => { freq[w] = (freq[w] || 0) + 1; }); });
  return sortEntries(freq).slice(0, topN || 30);
}

function execDescWordFreq(rows) {
  const entries = wordFreq(rows, 'T', 30);
  const labels = entries.map(e => e[0]), data = entries.map(e => e[1]);
  const tblRows = entries.map((e, i) => [i + 1, e[0], e[1]]);
  return chartCanvas('ac71') + renderTable(['排名','关键词','频次'], tblRows) +
    '<script>Charts.horizontalBar("ac71",' + JSON.stringify(labels) + ',[{label:"频次",data:' + JSON.stringify(data) + '}])</script>' +
    exportCsvBtn('desc-word-freq', ['排名','关键词','频次'], tblRows);
}

function execCauseDescWordFreq(rows) {
  const entries = wordFreq(rows, 'AL', 30);
  const labels = entries.map(e => e[0]), data = entries.map(e => e[1]);
  const tblRows = entries.map((e, i) => [i + 1, e[0], e[1]]);
  return chartCanvas('ac72') + renderTable(['排名','关键词','频次'], tblRows) +
    '<script>Charts.horizontalBar("ac72",' + JSON.stringify(labels) + ',[{label:"频次",data:' + JSON.stringify(data) + '}])</script>' +
    exportCsvBtn('cause-desc-word-freq', ['排名','关键词','频次'], tblRows);
}

function execContentWordFreq(rows) {
  const entries = wordFreq(rows, 'AM', 30);
  const labels = entries.map(e => e[0]), data = entries.map(e => e[1]);
  const tblRows = entries.map((e, i) => [i + 1, e[0], e[1]]);
  return chartCanvas('ac73') + renderTable(['排名','关键词','频次'], tblRows) +
    '<script>Charts.horizontalBar("ac73",' + JSON.stringify(labels) + ',[{label:"频次",data:' + JSON.stringify(data) + '}])</script>' +
    exportCsvBtn('content-word-freq', ['排名','关键词','频次'], tblRows);
}

function execFaultPattern(rows) {
  const patterns = {};
  rows.forEach(r => { const t = strVal(r, 'T'), al = strVal(r, 'AL'), am = strVal(r, 'AM'); const combined = [t, al, am].filter(Boolean).join(' | '); if (!combined) return; const words = extractKeywords(combined).slice(0, 3).sort().join('+'); if (!words) return; patterns[words] = (patterns[words] || 0) + 1; });
  const entries = sortEntries(patterns).slice(0, 30);
  const tblRows = entries.map((e, i) => [i + 1, e[0], e[1]]);
  return '<p class="analysis-info">T+AL+AM组合文本的关键词聚类（取每条记录前3个关键词组合）</p>' + renderTable(['排名','关键词组合','出现次数'], tblRows) + exportCsvBtn('fault-pattern', ['排名','关键词组合','出现次数'], tblRows);
}

/* ===== ⑪ 大数据分析 ===== */

/* --- 多维异常检测 (IQR + Z-score) --- */
function execMultiOutlier(rows) {
  const dims = [
    { col: 'AK', name: '维修时长' },
    { col: 'AI', name: '等待时长' },
    { col: 'AQ', name: '维修总价值' },
    { col: 'AT', name: '产量影响时长' },
  ];
  const dimStats = dims.map(d => {
    const all = rows.map(rr => numVal(rr, d.col)).filter(x => x != null && x > 0).sort((a, b) => a - b);
    if (all.length < 4) return null;
    const q1 = ss.quantile(all, 0.25);
    const q3 = ss.quantile(all, 0.75);
    return { col: d.col, name: d.name, upper: q3 + 1.5 * (q3 - q1), avg: mean(all), sd: ss.standardDeviation(all) || 1 };
  }).filter(Boolean);
  const anomalies = [];
  rows.forEach(r => {
    const id = strVal(r, 'A'), equip = strVal(r, 'L'), ws = strVal(r, 'F');
    let flags = [], score = 0;
    dimStats.forEach(ds => {
      const v = numVal(r, ds.col);
      if (v == null || v <= 0) return;
      if (v > ds.upper) { flags.push(ds.name + '超IQR'); score++; }
      const z = Math.abs((v - ds.avg) / ds.sd);
      if (z > 2.5) { flags.push(ds.name + 'Z=' + z.toFixed(1)); score++; }
    });
    if (score > 0) anomalies.push({ id, equip, ws, score, flags: flags.join('; ') });
  });
  anomalies.sort((a, b) => b.score - a.score);
  const tblRows = anomalies.slice(0, 50).map((a, i) => [i + 1, a.id, a.ws, a.equip, a.score, a.flags]);
  const summary = [
    ['异常维修单总数', anomalies.length],
    ['占总比', fmtPct(anomalies.length / rows.length)],
    ['多维度异常(≥2)', anomalies.filter(a => a.score >= 2).length],
  ];
  return '<p class="analysis-info">基于IQR(1.5倍四分位距)和Z-score(|z|>2.5)对维修时长/等待时长/维修成本/产量影响四维度检测异常</p>' +
    renderTable(['指标', '值'], summary) + '<h4>异常维修单 Top 50</h4>' +
    renderTable(['排名', '维修单号', '车间', '设备', '异常维度数', '异常标记'], tblRows, 50) +
    exportCsvBtn('multi-outlier', ['排名', '维修单号', '车间', '设备', '异常维度数', '异常标记'], tblRows);
}

/* --- 多维交叉透视 (热力图) --- */
function execCrossPivot(rows) {
  const rowDim = 'F', colDim = 'P', valDim = 'AK';
  const rowVals = [...new Set(rows.map(r => strVal(r, rowDim)).filter(Boolean))].sort();
  const colVals = [...new Set(rows.map(r => strVal(r, colDim)).filter(Boolean))].sort();
  const matrix = {};
  rowVals.forEach(rv => { matrix[rv] = {}; colVals.forEach(cv => { matrix[rv][cv] = { count: 0, sumDur: 0 }; }); });
  rows.forEach(r => {
    const rv = strVal(r, rowDim), cv = strVal(r, colDim);
    const dur = numVal(r, valDim) || 0;
    if (rv && cv) { matrix[rv][cv].count++; matrix[rv][cv].sumDur += dur; }
  });
  const maxCount = Math.max(...rowVals.map(rv => Math.max(...colVals.map(cv => matrix[rv][cv].count))), 1);
  let html = '<div class="heatmap-wrap"><table class="result-table heatmap"><thead><tr><th>车间 \\ 故障等级</th>';
  colVals.forEach(cv => { html += '<th>' + esc(cv) + '</th>'; });
  html += '</tr></thead><tbody>';
  rowVals.forEach(rv => {
    html += '<tr><td class="heatmap-label">' + esc(rv) + '</td>';
    colVals.forEach(cv => {
      const c = matrix[rv][cv].count;
      const avgDur = c ? (matrix[rv][cv].sumDur / c).toFixed(1) : '—';
      const intensity = c / maxCount;
      const bg = c > 0 ? 'background:rgba(177,40,30,' + (0.15 + intensity * 0.7).toFixed(2) + ')' : '';
      html += '<td style="' + bg + '" title="频次:' + c + ' 均时长:' + avgDur + '">' + c + (c > 0 ? '<span class="heatmap-sub">' + avgDur + 'h</span>' : '') + '</td>';
    });
    html += '</tr>';
  });
  html += '</tbody></table></div>';
  const flatRows = [];
  rowVals.forEach(rv => colVals.forEach(cv => { if (matrix[rv][cv].count) flatRows.push([rv, cv, matrix[rv][cv].count, (matrix[rv][cv].sumDur / matrix[rv][cv].count).toFixed(2)]); }));
  flatRows.sort((a, b) => b[2] - a[2]);
  return '<p class="analysis-info">车间×故障等级交叉透视，单元格显示频次和平均维修时长，颜色深浅表示频次强度</p>' + html +
    '<h4>高频组合 Top 20</h4>' + renderTable(['车间', '故障等级', '频次', '平均维修时长'], flatRows.slice(0, 20), 20) +
    exportCsvBtn('cross-pivot', ['车间', '故障等级', '频次', '平均维修时长'], flatRows);
}

/* --- 关联规则挖掘 (共现矩阵) --- */
function execAssocRules(rows) {
  const minSupport = Math.max(3, Math.ceil(rows.length * 0.005));
  const itemsets = rows.map(r => {
    const items = [];
    const ws = strVal(r, 'F'), cause = strVal(r, 'W'), part = strVal(r, 'AE'), cat = strVal(r, 'AD');
    if (ws) items.push('车间:' + ws);
    if (cause) items.push('原因:' + cause);
    if (part) items.push('部件:' + part);
    if (cat) items.push('分类:' + cat);
    return items;
  }).filter(a => a.length >= 2);
  const pairCount = {};
  itemsets.forEach(items => {
    for (let i = 0; i < items.length; i++) {
      for (let j = i + 1; j < items.length; j++) {
        const key = items[i] + ' → ' + items[j];
        pairCount[key] = (pairCount[key] || 0) + 1;
      }
    }
  });
  const itemTotal = {};
  itemsets.forEach(items => { items.forEach(it => { itemTotal[it] = (itemTotal[it] || 0) + 1; }); });
  const rules = Object.entries(pairCount).filter(([, c]) => c >= minSupport).map(([key, c]) => {
    const [a, b] = key.split(' → ');
    const support = c / rows.length;
    const confidence = c / (itemTotal[a] || 1);
    const lift = confidence / ((itemTotal[b] || 1) / itemsets.length);
    return { key, a, b, count: c, support, confidence, lift };
  }).sort((x, y) => y.lift - x.lift);
  const tblRows = rules.slice(0, 30).map((r, i) => [i + 1, r.a, r.b, r.count, fmtPct(r.support), fmtPct(r.confidence), r.lift.toFixed(2)]);
  return '<p class="analysis-info">挖掘故障原因/部件/分类/车间之间的关联规则（最小支持度=' + minSupport + '），按提升度(Lift)排序，Lift>1表示正相关</p>' +
    renderTable(['排名', '前项', '后项', '共现次数', '支持度', '置信度', '提升度'], tblRows, 30) +
    exportCsvBtn('assoc-rules', ['排名', '前项', '后项', '共现次数', '支持度', '置信度', '提升度'], tblRows);
}

/* --- 季节性分解 (趋势+周期+残差) --- */
function execSeasonalDecomp(rows) {
  const months = {};
  rows.forEach(r => { const d = dateVal(r, 'Y'); if (!d) return; const k = monthKey(d); months[k] = (months[k] || 0) + 1; });
  const sorted = Object.keys(months).sort();
  const actual = sorted.map(k => months[k]);
  if (actual.length < 6) return '<p class="no-data">数据不足，至少需要6个月数据才能进行季节性分解</p>';
  const period = Math.min(12, actual.length);
  const trend = actual.map((_, i) => {
    const half = Math.floor(period / 2);
    const lo = Math.max(0, i - half), hi = Math.min(actual.length, i + half + 1);
    return mean(actual.slice(lo, hi));
  });
  const detrended = actual.map((v, i) => v - trend[i]);
  const seasonal = [];
  for (let p = 0; p < period; p++) {
    const vals = [];
    for (let i = p; i < detrended.length; i += period) vals.push(detrended[i]);
    seasonal[p] = mean(vals);
  }
  const seasonalAvg = mean(seasonal);
  const seasonalNorm = seasonal.map(s => s - seasonalAvg);
  const residual = actual.map((v, i) => v - trend[i] - seasonalNorm[i % period]);
  const tblRows = sorted.map((k, i) => [k, actual[i].toFixed(0), trend[i].toFixed(1), seasonalNorm[i % period].toFixed(2), residual[i].toFixed(2)]);
  const labels = sorted;
  return chartCanvas('ac75') + renderTable(['月份', '实际值', '趋势', '季节性', '残差'], tblRows) +
    '<script>Charts.line("ac75",' + JSON.stringify(labels) + ',[{label:"实际值",data:' + JSON.stringify(actual) + '},{label:"趋势",data:' + JSON.stringify(trend.map(t => +t.toFixed(2))) + '},{label:"季节性",data:' + JSON.stringify(actual.map((_, i) => +seasonalNorm[i % period].toFixed(2))) + '},{label:"残差",data:' + JSON.stringify(residual.map(r => +r.toFixed(2))) + '}])</script>' +
    exportCsvBtn('seasonal-decomp', ['月份', '实际值', '趋势', '季节性', '残差'], tblRows);
}

/* --- K-Means 设备聚类 --- */
function _kmeans(data, k, maxIter) {
  const n = data.length;
  if (n < k) k = n;
  const centroids = [];
  for (let i = 0; i < k; i++) centroids.push(data[Math.floor(i * n / k)].slice());
  const labels = new Array(n).fill(0);
  for (let iter = 0; iter < (maxIter || 50); iter++) {
    let changed = false;
    for (let i = 0; i < n; i++) {
      let best = 0, bestDist = Infinity;
      for (let c = 0; c < k; c++) {
        let d = 0;
        for (let d2 = 0; d2 < data[i].length; d2++) d += (data[i][d2] - centroids[c][d2]) ** 2;
        if (d < bestDist) { bestDist = d; best = c; }
      }
      if (labels[i] !== best) { labels[i] = best; changed = true; }
    }
    const sums = Array.from({ length: k }, () => new Array(data[0].length).fill(0));
    const counts = new Array(k).fill(0);
    for (let i = 0; i < n; i++) { counts[labels[i]]++; for (let d = 0; d < data[i].length; d++) sums[labels[i]][d] += data[i][d]; }
    for (let c = 0; c < k; c++) { if (counts[c]) for (let d = 0; d < data[0].length; d++) centroids[c][d] = sums[c][d] / counts[c]; }
    if (!changed) break;
  }
  return { labels, centroids };
}

function execKmeansCluster(rows) {
  const groups = groupBy(rows, 'L');
  const equipData = Object.entries(groups).map(([equip, rs]) => {
    const durs = rs.map(r => numVal(r, 'AK')).filter(v => v != null && v > 0);
    const costs = rs.map(r => numVal(r, 'AQ')).filter(v => v != null && v >= 0);
    if (!durs.length) return null;
    return {
      equip,
      freq: rs.length,
      avgDur: mean(durs),
      avgCost: mean(costs),
    };
  }).filter(Boolean);
  if (equipData.length < 3) return '<p class="no-data">设备数量不足，至少需要3台设备才能聚类</p>';
  const maxFreq = Math.max(...equipData.map(d => d.freq)) || 1;
  const maxDur = Math.max(...equipData.map(d => d.avgDur)) || 1;
  const maxCost = Math.max(...equipData.map(d => d.avgCost)) || 1;
  const features = equipData.map(d => [d.freq / maxFreq, d.avgDur / maxDur, d.avgCost / maxCost]);
  const k = Math.min(4, Math.max(2, Math.floor(Math.sqrt(equipData.length / 2))));
  const { labels } = _kmeans(features, k, 50);
  const clusterStats = Array.from({ length: k }, () => ({ freqSum: 0, durSum: 0, costSum: 0, count: 0 }));
  equipData.forEach((d, i) => {
    const c = labels[i];
    clusterStats[c].freqSum += d.freq;
    clusterStats[c].durSum += d.avgDur;
    clusterStats[c].costSum += d.avgCost;
    clusterStats[c].count++;
  });
  const clusterRank = clusterStats.map((cs, i) => ({
    idx: i,
    score: cs.count > 0 ? (cs.freqSum / cs.count / maxFreq + cs.durSum / cs.count / maxDur + cs.costSum / cs.count / maxCost) : 0
  })).sort((a, b) => b.score - a.score);
  const namePool = ['高频高耗', '中频中耗', '低频低耗', '特殊模式'];
  const clusterNames = new Array(k);
  clusterRank.forEach((cr, i) => { clusterNames[cr.idx] = namePool[i] || ('簇' + (i + 1)); });
  const clusters = Array.from({ length: k }, () => []);
  equipData.forEach((d, i) => clusters[labels[i]].push(d));
  const tblRows = [];
  clusters.forEach((c, ci) => {
    c.sort((a, b) => b.freq - a.freq);
    c.slice(0, 15).forEach(d => tblRows.push([clusterNames[ci] || ('簇' + (ci + 1)), d.equip, d.freq, d.avgDur.toFixed(2), d.avgCost.toFixed(2)]));
  });
  const summary = clusters.map((c, ci) => [clusterNames[ci] || ('簇' + (ci + 1)), c.length, mean(c.map(d => d.freq)).toFixed(1), mean(c.map(d => d.avgDur)).toFixed(2), mean(c.map(d => d.avgCost)).toFixed(2)]);
  return '<p class="analysis-info">对设备按故障频率/平均维修时长/平均成本三维特征做K-Means聚类(k=' + k + ')，自动识别设备风险群组</p>' +
    renderTable(['聚类', '设备数', '平均频次', '平均时长', '平均成本'], summary) +
    '<h4>各聚类设备明细 (每簇Top 15)</h4>' +
    renderTable(['聚类', '设备名称', '故障次数', '平均维修时长', '平均成本'], tblRows, 60) +
    exportCsvBtn('kmeans-cluster', ['聚类', '设备名称', '故障次数', '平均维修时长', '平均成本'], tblRows);
}

/* --- ARIMA 时序预测 --- */
function _diff(arr, d) {
  let r = arr.slice();
  for (let i = 0; i < d; i++) r = r.slice(1).map((v, j) => v - r[j]);
  return r;
}
function _arma(data, p, q) {
  const n = data.length;
  const phi = new Array(p).fill(0);
  const theta = new Array(q).fill(0);
  const resid = new Array(n).fill(0);
  for (let iter = 0; iter < 30; iter++) {
    for (let t = Math.max(p, q); t < n; t++) {
      let pred = 0;
      for (let i = 0; i < p; i++) pred += phi[i] * data[t - 1 - i];
      for (let i = 0; i < q; i++) pred += theta[i] * resid[t - 1 - i];
      resid[t] = data[t] - pred;
    }
    for (let i = 0; i < p; i++) {
      let num = 0, den = 0;
      for (let t = p; t < n; t++) { num += data[t - 1 - i] * resid[t]; den += data[t - 1 - i] ** 2; }
      phi[i] = den > 0 ? num / den : 0;
    }
    for (let i = 0; i < q; i++) {
      let num = 0, den = 0;
      for (let t = q; t < n; t++) { num += resid[t - 1 - i] * resid[t]; den += resid[t - 1 - i] ** 2; }
      theta[i] = den > 0 ? -num / den : 0;
    }
  }
  return { phi, theta, resid };
}

function execArimaForecast(rows) {
  const months = {};
  rows.forEach(r => { const d = dateVal(r, 'Y'); if (!d) return; const k = monthKey(d); months[k] = (months[k] || 0) + 1; });
  const sorted = Object.keys(months).sort();
  const actual = sorted.map(k => months[k]);
  if (actual.length < 8) return '<p class="no-data">数据不足，至少需要8个月数据才能拟合ARIMA模型</p>';
  const d = 1, p = 2, q = 1;
  const diffed = _diff(actual, d);
  const { phi, theta, resid } = _arma(diffed, p, q);
  const forecasts = [];
  let lastDiff = diffed.slice(-p);
  let lastResid = resid.slice(-q);
  let prevValue = actual[actual.length - 1];
  for (let step = 0; step < 3; step++) {
    let pred = 0;
    for (let i = 0; i < p; i++) pred += phi[i] * lastDiff[lastDiff.length - 1 - i];
    for (let i = 0; i < q; i++) pred += theta[i] * lastResid[lastResid.length - 1 - i];
    const forecastVal = prevValue + pred;
    forecasts.push(Math.max(0, Math.round(forecastVal)));
    prevValue = forecastVal;
    lastDiff = [...lastDiff.slice(1), pred];
    lastResid = [...lastResid.slice(1), 0];
  }
  const fitted = actual.map((v, i) => {
    if (i < Math.max(p, q)) return null;
    let pred = 0;
    for (let j = 0; j < p; j++) pred += phi[j] * diffed[i - 1 - j];
    for (let j = 0; j < q; j++) pred += theta[j] * resid[i - 1 - j];
    return Math.round(actual[i - 1] + pred);
  });
  const rmse = Math.sqrt(mean(actual.slice(Math.max(p, q)).map((v, i) => (v - fitted[i + Math.max(p, q)]) ** 2)));
  const futureLabels = [];
  const lastDate = new Date(sorted[sorted.length - 1] + '-01');
  for (let i = 1; i <= 3; i++) { const d2 = new Date(lastDate.getFullYear(), lastDate.getMonth() + i, 1); futureLabels.push(monthKey(d2)); }
  const tblRows = sorted.map((k, i) => [k, actual[i], fitted[i] != null ? fitted[i] : '—']).concat(futureLabels.map((k, i) => [k + '(预测)', '—', forecasts[i]]));
  return chartCanvas('ac76') + renderTable(['月份', '实际值', 'ARIMA拟合'], tblRows) +
    '<script>Charts.line("ac76",' + JSON.stringify(sorted.concat(futureLabels)) + ',[{label:"实际值",data:' + JSON.stringify(actual.concat([null, null, null])) + '},{label:"ARIMA拟合",data:' + JSON.stringify(fitted.concat(forecasts)) + '}])</script>' +
    '<p class="analysis-info">ARIMA(' + p + ',' + d + ',' + q + ') 模型，RMSE=' + rmse.toFixed(2) + '，预测未来3个月维修量</p>' +
    exportCsvBtn('arima-forecast', ['月份', '实际值', 'ARIMA拟合'], tblRows);
}

/* --- TF-IDF 文本分析 --- */
function execTfidfAnalysis(rows) {
  const docs = rows.map(r => strVal(r, 'T')).filter(Boolean);
  if (!docs.length) return '<p class="no-data">无问题描述文本数据</p>';
  const docWords = docs.map(text => extractKeywords(text));
  const N = docs.length;
  const df = {};
  docWords.forEach(words => { [...new Set(words)].forEach(w => { df[w] = (df[w] || 0) + 1; }); });
  const tfidf = {};
  docWords.forEach(words => {
    const tf = {};
    words.forEach(w => { tf[w] = (tf[w] || 0) + 1; });
    words.forEach(w => {
      const score = (tf[w] / words.length) * Math.log((N + 1) / (df[w] + 1) + 1);
      tfidf[w] = (tfidf[w] || 0) + score;
    });
  });
  const entries = Object.entries(tfidf).sort((a, b) => b[1] - a[1]).slice(0, 40);
  const tblRows = entries.map((e, i) => [i + 1, e[0], df[e[0]], e[1].toFixed(3)]);
  const labels = entries.map(e => e[0]), data = entries.map(e => +e[1].toFixed(3));
  return chartCanvas('ac77') + renderTable(['排名', '关键词', '文档频率', 'TF-IDF总分'], tblRows, 40) +
    '<script>Charts.horizontalBar("ac77",' + JSON.stringify(labels) + ',[{label:"TF-IDF",data:' + JSON.stringify(data) + '}])</script>' +
    '<p class="analysis-info">TF-IDF综合考虑词频(TF)和逆文档频率(IDF)，相比纯词频能识别更有区分度的关键词</p>' +
    exportCsvBtn('tfidf-analysis', ['排名', '关键词', '文档频率', 'TF-IDF总分'], tblRows);
}

/* --- 注册大数据方法 --- */
reg(75,11,3,'多维异常检测','AK,AI,AQ,AT','IQR与Z-score多维度识别异常维修单',execMultiOutlier,'看被标记为异常的维修单，这些单据在时长/等待/成本上同时偏离正常范围。用于异常维修单审查和根因调查。');
reg(76,11,3,'多维交叉透视','F,P,AK','车间×故障等级交叉频次与平均时长热力图',execCrossPivot,'看热力图中颜色最深的格子，表示该车间该等级故障频次最高或时长最长。用于识别车间与故障等级的组合风险。');
reg(77,11,3,'关联规则挖掘','F,W,AE,AD','挖掘故障原因/部件/车间间关联规则与提升度',execAssocRules,'看提升度>1的规则，表示两个因素同时出现概率高于随机。用于发现隐藏的故障关联因素。');
reg(78,11,3,'季节性分解','Y','月度维修量趋势/季节性/残差三分解',execSeasonalDecomp,'看季节性分量大于0的月份，这些是维修高峰期。用于季节性预防性维护计划制定。');
reg(79,11,3,'设备K-Means聚类','L,AK,AQ','按频率/时长/成本三维聚类识别设备风险群组',execKmeansCluster,'看各聚类中心特征，高频率高成本聚类是重点关注设备群。用于设备分级管理和资源优先分配。');
reg(80,11,4,'ARIMA维修量预测','Y','ARIMA(2,1,1)模型预测未来3个月维修量',execArimaForecast,'看预测趋势是否上升，上升需提前准备资源。用于中长期维修资源规划。');
reg(81,11,3,'TF-IDF关键词分析','T','TF-IDF权重提取问题描述区分性关键词',execTfidfAnalysis,'看TF-IDF权重高的词，这些词在某类维修中特别突出而非普遍出现。用于故障分类和知识库标签建设。');

/* ---- 评分注入（必须在所有 reg() 之后） ---- */
ANALYSIS_METHODS.forEach(m => {
  const s = _SCORES[m.id] || [3,3,3];
  m.bv = s[0]; m.ua = s[1]; m.cc = s[2];
  m.score = (0.6 * s[0] + 0.4 * s[1]).toFixed(2);
});
ANALYSIS_METHODS.sort((a, b) => parseFloat(b.score) - parseFloat(a.score));

