'use strict';
/* 回录与错报评估
 *
 * 系数先对 AICPA 公布的货币单元抽样表（3.00 / 4.75 / 6.30 / 7.76 / 9.16），
 * 再用「生成抽凭表 → 填结论 → 读回 → 评估」跑一遍闭环。
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('../vendor/xlsx.full.min.js');
const Core = require('../core.js');
const gt = require('./fixture/ground_truth.json');

let count = 0;
const fails = [];
function ok(cond, msg) { count++; if (!cond) fails.push(msg); }
function eq(a, b, msg) {
  count++;
  if (a !== b) fails.push(msg + ' — 期望 ' + JSON.stringify(b) + '，实际 ' + JSON.stringify(a));
}
function near(a, b, tol, msg) {
  count++;
  if (Math.abs(a - b) > tol) fails.push(msg + ' — 期望 ' + b + ' ± ' + tol + '，实际 ' + a);
}
function section(t) { console.log('\n[' + t + ']'); }

/* ---------------- 1 上界系数 ---------------- */
section('1 泊松上界系数：与 AICPA 表一致');
const PUB95 = [3.00, 4.75, 6.30, 7.76, 9.16];
PUB95.forEach((want, k) => {
  near(Core.poissonUpperFactor(k, 0.05), want, 0.02, '95% 置信 R' + k);
});
near(Core.poissonUpperFactor(0, 0.10), 2.31, 0.02, '90% 置信 R0');
near(Core.poissonUpperFactor(0, 0.01), 4.61, 0.03, '99% 置信 R0');
eq(Core.poissonUpperFactor(0, 0.05) === -Math.log(0.05), true, 'R0 应走闭式解');
// 系数必须单调递增：错报越多，上界越大
for (let k = 1; k < 6; k++) {
  ok(Core.poissonUpperFactor(k, 0.05) > Core.poissonUpperFactor(k - 1, 0.05), 'R' + k + ' 应大于 R' + (k - 1));
}

/* ---------------- 2 Stringer 上界 ---------------- */
section('2 Stringer 上界：三个可手算的情形');
const BV = 100000000; // 1,000,000.00 元
const N = 50;

const zero = Core.stringerUpperBound([], N, BV, 0.05);
eq(zero.upperCents, Math.round(BV / N * Core.poissonUpperFactor(0, 0.05)), '零错报：BV/n × R0');
eq(zero.factorTotal, Core.poissonUpperFactor(0, 0.05), '零错报的系数就是 R0');

const all = Core.stringerUpperBound([1, 1, 1], N, BV, 0.05);
eq(all.upperCents, Math.round(BV / N * Core.poissonUpperFactor(3, 0.05)), '三笔全额错报：退化成 R3');

const half = Core.stringerUpperBound([0.5], N, BV, 0.05);
const r0 = Core.poissonUpperFactor(0, 0.05), r1 = Core.poissonUpperFactor(1, 0.05);
eq(half.upperCents, Math.round(BV / N * (r0 + 0.5 * (r1 - r0))), '一笔 50% 错报：R0 + 0.5×(R1−R0)');

ok(all.upperCents > half.upperCents && half.upperCents > zero.upperCents, '上界应随错报增加而增加');
eq(Core.stringerUpperBound([], 0, BV, 0.05), null, '样本量为 0 时不计算');
eq(Core.stringerUpperBound([], N, 0, 0.05), null, '总体金额为 0 时不计算');

/* ---------------- 3 结论归一 ---------------- */
section('3 结论写法归一：认不出就不猜');
eq(Core.classifyConclusion('相符'), 'ok', '相符');
eq(Core.classifyConclusion('√'), 'ok', '打勾');
eq(Core.classifyConclusion(' 不符 '), 'exception', '带空格的不符');
eq(Core.classifyConclusion('×'), 'exception', '叉号');
eq(Core.classifyConclusion('未取得'), 'missing', '未取得');
eq(Core.classifyConclusion('没有拿到'), 'other', '同义但未登记的说法应报 other，不硬猜');
eq(Core.classifyConclusion('不适用'), 'na', '不适用');
eq(Core.classifyConclusion(''), '', '空着就是空着');
eq(Core.classifyConclusion('已取得发票'), 'other', '长句不猜');

/* ---------------- 4 闭环：生成 → 填 → 读回 → 评估 ---------------- */
section('4 闭环：抽凭表填完再读回来评估');
const wbSrc = XLSX.read(
  fs.readFileSync(path.join(__dirname, 'fixture', 'sample_ledger.xlsx')),
  { type: 'buffer', cellDates: false }
);
const wsSrc = wbSrc.Sheets['应收账款明细账'];
const range = XLSX.utils.decode_range(wsSrc['!ref']);
const rows = XLSX.utils.sheet_to_json(wsSrc, { header: 1, raw: true, defval: null, blankrows: true });
const header = Core.detectHeader(rows);
const parsed = Core.parseRows(rows, header, { firstExcelRow: range.s.r + 1 });
const pop = Core.buildPopulation(parsed.records, {
  dateFrom: { y: 2026, m: 1, d: 1 }, dateTo: { y: 2026, m: 12, d: 31 }, direction: 'all'
});
const smp = Core.buildSample(pop.included, {
  thresholdCents: gt.threshold_cents, n: 8, seed: 20260101, method: 'systematic'
});
const box = {
  client: gt.client, indexNo: 'B-3', preparer: '实习生', reviewer: '项目经理',
  dateText: '2026-11-21', periodText: '2026-01-01 至 2026-12-31',
  accountText: '1122 应收账款', presetName: '销售/应收',
  sample: smp.sample, meta: smp.meta, populationStats: pop.stats,
  excluded: pop.excluded, problems: parsed.problems, skipped: parsed.skipped,
  warnings: Object.assign({ parsedRows: parsed.records.length }, parsed.warnings),
  mapNotes: []
};

const sheets = Core.buildSheets(box);
const mainRows = sheets[0].rows.map(r => r.slice());
const headIdx = mainRows.findIndex(r => r[0] === '序号');
const head = mainRows[headIdx];
const cConc = head.indexOf('核对结论');
const cMis = head.indexOf('错报金额（元）');
ok(cConc > 0 && cMis > 0, '抽凭表应带核对结论与错报金额两列');

// 现场填：第 1 行相符、第 2 行不符且低估 100 元、第 3 行未取得，其余不填
const row1 = headIdx + 1, row2 = headIdx + 2, row3 = headIdx + 3;
mainRows[row1][cConc] = '相符';
mainRows[row2][cConc] = '不符';
mainRows[row2][cMis] = -100;
mainRows[row3][cConc] = '未取得';

const back = Core.parseVouchingRows(mainRows);
eq(back.headerRowIndex, headIdx, '能定位到抽凭表表头');
eq(back.records.length, smp.meta.sampleCount, '读回的笔数与样本一致');
eq(back.warnings.length, 0, '自带列全，不应有警告');
eq(back.records[0].seq, 1, '序号从 1 开始');
eq(back.records[0].conclusion, 'ok', '第 1 行结论');
eq(back.records[1].conclusion, 'exception', '第 2 行结论');
eq(back.records[1].misCents, -10000, '错报金额按分读回');
eq(back.records[2].conclusion, 'missing', '第 3 行结论');
eq(back.records[3].conclusion, '', '第 4 行未填');
ok(!!back.records[0].voucher, '凭证号读回');
ok(back.records[0].absCents > 0, '账面金额读回');

const info = Core.parseInfoRows(sheets[1].rows);
eq(info.popCount, pop.stats.count, '从抽样信息页读回总体笔数');
eq(info.popCents, pop.stats.totalAbsCents, '从抽样信息页读回总体金额');

const ev = Core.evaluateVouching(back.records, {
  alpha: 0.05,
  tolerableCents: 3000000,
  populationCents: info.popCents,
  sampleMethod: info.sampleMethod,
  sampleMethodLabel: info.sampleMethodLabel
});

eq(ev.counts.ok, 1, '相符 1 笔');
eq(ev.counts.exception, 1, '不符 1 笔');
eq(ev.counts.missing, 1, '未取得 1 笔');
eq(ev.counts.blank, smp.meta.sampleCount - 3, '未填的笔数');
eq(ev.examined, 3, '参与评估的样本量');
eq(ev.grossUnderCents, 10000, '低估金额');
eq(ev.grossOverCents, 0, '高估金额');
eq(ev.netCents, -10000, '净额');
eq(ev.taintCount, 2, 'taint 数 = 1 笔不符 + 1 笔未取得');

// 这份样本是等距抽的、还留了没填的行：按新规矩不许算 MUS 上限
eq(ev.upper, null, '等距样本 + 有未填行时不应给出错报上限');
ok(/金额加权/.test(ev.umBlocked), '应说明样本不是金额加权抽的，实际：' + ev.umBlocked);
ok(/没填核对结论/.test(ev.umBlocked), '应说明还有行没填，实际：' + ev.umBlocked);

// 换成「金额加权 + 全部评价」的样本，才允许算上限
const filled = back.records.map(r => Object.assign({}, r, { conclusion: 'ok', misCents: 0 }));
filled[1] = Object.assign({}, filled[1], { conclusion: 'exception', misCents: 10000 });
filled[2] = Object.assign({}, filled[2], { conclusion: 'missing', misCents: 0 });
const ppsOpts = { alpha: 0.05, tolerableCents: 3000000, populationCents: info.popCents, sampleMethod: 'pps' };
const evPps = Core.evaluateVouching(filled, ppsOpts);

ok(evPps.upper && evPps.upper.upperCents > 0, '金额加权的完整样本应算出错报上限');
eq(evPps.upper.upperCents, Math.min(Math.round(info.popCents / filled.length * evPps.upper.factorTotal), info.popCents),
   '上限 = 总体/样本量 × 系数，且不超过总体账面');
ok(evPps.upper.upperCents <= info.popCents, '错报上限不应超过总体账面金额');
ok(evPps.upper.factorTotal > Core.poissonUpperFactor(0, 0.05), '有错报时系数应大于 R0');

// 有全查层时同样不给 MUS 上限：全查项不是概率样本
const withFull = Core.evaluateVouching(filled, Object.assign({}, ppsOpts, { fullCount: 2 }));
eq(withFull.upper, null, '有全查层时不应算 MUS 上限');
ok(/全查/.test(withFull.umBlocked), '应说明全查层不参与，实际：' + withFull.umBlocked);

// 错报金额读不出来：不能当 0，也不能算上限
const badMis = filled.map(r => Object.assign({}, r));
badMis[0] = Object.assign({}, badMis[0], { misUnreadable: true });
const evBad = Core.evaluateVouching(badMis, ppsOpts);
eq(evBad.upper, null, '错报金额读不出时不应给上限');
ok(/读不出来/.test(evBad.umBlocked), '应说明哪类行读不出，实际：' + evBad.umBlocked);

// 样本极小又全错时，理论上限会超过总体，必须封顶
const tiny = Core.stringerUpperBound([1, 1, 1, 1, 1], 5, 1000000, 0.05);
eq(tiny.capped, true, '样本极小且全错时应标记封顶');
eq(tiny.upperCents, 1000000, '封顶为总体账面金额');
ok(tiny.rawCents > 1000000, '未封顶的理论值应大于总体账面');

const evSheet = Core.buildEvaluationSheet(evPps, box);
eq(evSheet.name, '评估表', '评估表名');
ok(evSheet.rows.some(r => r[0] === '错报上限'), '评估表应写明错报上限');
ok(evSheet.rows.some(r => r[0] === '置信水平'), '评估表应写明置信水平');
ok(evSheet.rows.some(r => r[0] === '可容忍错报'), '评估表应写明可容忍错报');
const evText = evSheet.rows.map(r => r.join('')).join('\n');
ok(evText.indexOf('不代替判断') >= 0, '评估表应写明不代替判断');
ok(evText.indexOf('未取得') >= 0 && evText.indexOf('100% 错报') >= 0, '应说明未取得按 100% 计入');

// 全部填「相符」时不该有上限虚高
const clean = back.records.map(r => Object.assign({}, r, { conclusion: 'ok', misCents: 0 }));
const ev2 = Core.evaluateVouching(clean, ppsOpts);
eq(ev2.counts.exception, 0, '全相符时无例外');
eq(ev2.upper.upperCents, Math.round(info.popCents / clean.length * Core.poissonUpperFactor(0, 0.05)),
   '全相符时上限退化成零错报上界');

// 留一份评估表给人眼看（走界面同一个装配入口，避免「示例不是发出来的那个」）
const evOut = Core.buildEvaluationWorkbook(XLSX, evPps, box);
eq(evOut.sheetNames.join(','), '评估表', '评估表工作簿只有一页');
ok(/^抽查评估_/.test(evOut.fileName) && /\.xlsx$/.test(evOut.fileName), '文件名应是「抽查评估_…xlsx」，实际 ' + evOut.fileName);
eq(evOut.wb.SheetNames.join(','), '评估表', '工作簿里的页名');

const evBack = XLSX.utils.sheet_to_json(evOut.wb.Sheets['评估表'], { header: 1, defval: null });
ok(evBack.some(r => r[0] === '推断错报（比率法）'), '评估表应含比率法推断错报');
ok(evBack.some(r => r[0] === '推断错报（差异法）'), '评估表应含差异法推断错报');

// 金额列必须是数值且带格式，不然贴进底稿是一片文本
const moneyRow = evBack.findIndex(r => r[0] === '错报上限');
ok(moneyRow > 0, '应有错报上限行');
const moneyCellAddr = XLSX.utils.encode_cell({ r: moneyRow, c: 1 });
const moneyCell = XLSX.utils.sheet_to_json(evOut.wb.Sheets['评估表'], { header: 1 })[moneyRow];
const backCell = XLSX.read(
  Buffer.from(XLSX.write(evOut.wb, { type: 'buffer', bookType: 'xlsx' })),
  { type: 'buffer', cellNF: true }
).Sheets['评估表'][moneyCellAddr];
ok(backCell && backCell.t === 'n', '错报上限应是数值而不是文本');
ok(backCell && backCell.z && backCell.z.indexOf('#,##0.00') >= 0, '错报上限应带 #,##0.00 格式，实际 ' + (backCell && backCell.z));

fs.writeFileSync(path.join(__dirname, 'fixture', 'example_evaluation.xlsx'),
  Buffer.from(XLSX.write(evOut.wb, { type: 'buffer', bookType: 'xlsx' })));

/* ---------------- 汇总 ---------------- */
console.log('\n' + '='.repeat(56));
if (fails.length) {
  console.log('失败 ' + fails.length + ' / ' + count + ' 项：');
  fails.forEach((f, i) => console.log('  ' + (i + 1) + '. ' + f));
  process.exit(1);
}
console.log('全部通过：' + count + ' 项断言');
console.log('='.repeat(56));
