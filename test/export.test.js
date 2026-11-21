'use strict';
/* 抽凭助手 · 导出往返
 * 走产品真正用的那条路：core.buildWorkbook → 写成 xlsx 字节 → 再读回来核对。
 * 只断言「文件里的数」与「抽样元数据」一致，不重复 core.test.js 的口径断言。
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
function section(t) { console.log('\n[' + t + ']'); }

/* ---------- 从夹具跑一遍完整流程 ---------- */
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
  client: gt.client,
  indexNo: 'B-3',
  preparer: '实习生',
  reviewer: '项目经理',
  dateText: '2026-11-21',
  periodText: '2026-01-01 至 2026-12-31',
  accountText: '1122 应收账款',
  presetName: '销售/应收',
  sample: smp.sample,
  meta: smp.meta,
  populationStats: pop.stats,
  excluded: pop.excluded,
  problems: parsed.problems,
  skipped: parsed.skipped,
  warnings: Object.assign({ parsedRows: parsed.records.length }, parsed.warnings),
  mapNotes: ['工作表：应收账款明细账；表头位于第 3 行']
};

/* ---------- 写出字节并读回 ---------- */
section('1 写出的文件本身可用');
const out = Core.buildWorkbook(XLSX, box);
const buf = XLSX.write(out.wb, { type: 'buffer', bookType: 'xlsx' });
// 留一份给人眼看，用的是合成夹具，不含任何真实数据
fs.writeFileSync(path.join(__dirname, 'fixture', 'example_output.xlsx'), Buffer.from(buf));
ok(buf && buf.length > 3000, 'xlsx 字节应非空，实际 ' + (buf ? buf.length : 0) + ' 字节');
eq(Buffer.from(buf).slice(0, 2).toString(), 'PK', 'xlsx 应是 zip 容器');
ok(/\.xlsx$/.test(out.fileName), '文件名应以 .xlsx 结尾，实际 ' + out.fileName);

eq(out.sheetNames.join(','), '抽凭表,抽样信息,排除清单,数据检查,说明', '五个工作表，顺序固定');
eq(out.wb.SheetNames.join(','), out.sheetNames.join(','), '工作簿里的表名与声明一致');

// cellNF:true 才会把单元格的数字格式读进 cell.z，否则永远是 undefined
const wb2 = XLSX.read(buf, { type: 'buffer', cellDates: false, cellNF: true });
eq(wb2.SheetNames.filter(n => out.sheetNames.indexOf(n) >= 0).length, 5, '读回后仍是五个表');

function back(name) {
  return XLSX.utils.sheet_to_json(wb2.Sheets[name], { header: 1, raw: true, defval: null, blankrows: true });
}

/* ---------- 抽凭表 ---------- */
section('2 抽凭表：金额逐笔与样本一致');
const mainBack = back('抽凭表');
const headIdx = mainBack.findIndex(r => r[0] === '序号');
ok(headIdx > 0, '读回后仍能找到表头行');

const dataStart = headIdx + 1;
eq(mainBack[dataStart][0], 1, '首条数据序号为 1');
eq(mainBack[dataStart + smp.meta.sampleCount - 1][0], smp.meta.sampleCount, '末条数据序号连续');

// 逐笔核对借方/贷方/绝对额 —— 这是「金额可复现」的核心断言
let mismatch = 0;
for (let i = 0; i < smp.meta.sampleCount; i++) {
  const row = mainBack[dataStart + i];
  const rec = smp.sample[i].rec;
  const dr = rec.debitCents === null ? '' : rec.debitCents / 100;
  const cr = rec.creditCents === null ? '' : rec.creditCents / 100;
  const abs = Math.abs(rec.signedCents) / 100;
  if (row[7] !== dr || row[8] !== cr || row[9] !== abs) mismatch++;
}
eq(mismatch, 0, '读回后逐笔借贷金额应完全一致（不一致 ' + mismatch + ' 笔）');

const totalRow = mainBack[dataStart + smp.meta.sampleCount];
eq(totalRow[0], '合计', '合计行紧跟样本');
eq(Math.round(totalRow[9] * 100), smp.meta.sampleAbsCents, '合计绝对额 = 样本金额（分）');
eq(Math.round(totalRow[7] * 100) - Math.round(totalRow[8] * 100),
   smp.sample.reduce((s, x) => s + x.rec.signedCents, 0), '合计借 − 贷 = 样本净额');

// 金额单元格带千分位格式，打开就能看（挑一个确实是数字的单元格）
const numRowOffset = mainBack.slice(dataStart, dataStart + smp.meta.sampleCount)
  .findIndex(r => typeof r[7] === 'number' || typeof r[8] === 'number');
ok(numRowOffset >= 0, '样本里应至少有一笔借方或贷方为数字');
const numCol = typeof mainBack[dataStart + numRowOffset][7] === 'number' ? 'H' : 'I';
const moneyCell = wb2.Sheets['抽凭表'][numCol + (dataStart + numRowOffset + 1)];
ok(moneyCell && moneyCell.z === '#,##0.00', '金额列应带 #,##0.00 格式，实际 ' + (moneyCell && moneyCell.z));

// 抽样参数写进了表头块，复核的人能照着重跑
const headText = mainBack.slice(0, headIdx).map(r => (r || []).join('')).join('\n');
ok(headText.indexOf(String(smp.meta.seed)) >= 0, '表头块应写明随机种子');
ok(headText.indexOf('起点 ' + smp.meta.start) >= 0, '表头块应写明起点');

/* ---------- 排除清单 ---------- */
section('3 排除清单：一条不少');
const exclBack = back('排除清单');
eq(exclBack.length, 2 + pop.excluded.length, '排除清单行数 = 标题 + 表头 + 每条排除');
const exclSrcRows = exclBack.slice(2).map(r => r[0]);
eq(exclSrcRows.join(','), pop.excluded.map(e => e.srcRow).join(','), '排除记录按原始行号逐条落到表上');
eq(exclBack.filter(r => r[1] === '日期缺失').length, gt.excluded.date_missing, '排除原因随行写出');

/* ---------- 数据检查 ---------- */
section('4 数据检查：跳过的行也说清楚');
const chkBack = back('数据检查');
const findRow = (label) => chkBack.find(r => r[0] === label);
eq((findRow('空行') || [])[1], parsed.warnings.blankRows, '空行数写出');
eq((findRow('汇总/合计行') || [])[1], parsed.warnings.subtotalRows, '汇总行数写出');
eq((findRow('明细行数') || [])[1], parsed.records.length, '明细行数写出');

/* ---------- 抽样信息 ---------- */
section('5 抽样信息：参数齐备，能独立复核');
const infoBack = back('抽样信息');
const getInfo = (k) => { const r = infoBack.find(x => x[0] === k); return r ? r[1] : undefined; };
eq(getInfo('总体笔数'), pop.stats.count, '总体笔数');
eq(getInfo('随机种子'), smp.meta.seed, '随机种子');
eq(getInfo('起点'), smp.meta.start, '起点');
eq(getInfo('等距'), smp.meta.k, '间距');
eq(getInfo('样本笔数'), smp.meta.sampleCount, '样本笔数');
ok(String(getInfo('工具版本')).indexOf('v0.6') >= 0, '版本号写出');

/* ---------- 重跑一致 ---------- */
section('6 同一份数据重跑，样本不变');
const smp2 = Core.buildSample(pop.included, {
  thresholdCents: gt.threshold_cents, n: 8, seed: 20260101, method: 'systematic'
});
eq(smp2.sample.map(x => x.rec.srcRow).join(','), smp.sample.map(x => x.rec.srcRow).join(','),
   '同数据同参数重跑，逐笔相同');

const other = Core.buildSample(pop.included, {
  thresholdCents: gt.threshold_cents, n: 8, seed: 777, method: 'systematic'
});
ok(other.meta.start !== smp.meta.start || other.sample.map(x => x.rec.srcRow).join(',') !== smp.sample.map(x => x.rec.srcRow).join(','),
   '换种子应换出一批不同的样本（或不同起点）');

/* ---------- 汇总 ---------- */
console.log('\n' + '='.repeat(56));
if (fails.length) {
  console.log('失败 ' + fails.length + ' / ' + count + ' 项：');
  fails.forEach((f, i) => console.log('  ' + (i + 1) + '. ' + f));
  process.exit(1);
}
console.log('全部通过：' + count + ' 项断言');
console.log('抽查文件：' + out.fileName);
console.log('='.repeat(56));
