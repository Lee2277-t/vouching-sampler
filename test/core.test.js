'use strict';
/* 抽凭助手 · 核心逻辑回归
 * 数据来源：test/fixture/sample_ledger.xlsx（Python 生成，全合成）
 * 期望值来源：test/fixture/ground_truth.json（Python 用 Decimal 独立算）
 * 两边是独立路径：Python 写 → SheetJS 读 → core.js 算 → 与 Python 的答案比。
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('../vendor/xlsx.full.min.js');
const Core = require('../core.js');
const gt = require('./fixture/ground_truth.json');

let count = 0;
const fails = [];

function ok(cond, msg) { count++; if (!cond) fails.push(msg); }
function eq(actual, expect, msg) {
  count++;
  if (actual !== expect) fails.push(msg + ' — 期望 ' + JSON.stringify(expect) + '，实际 ' + JSON.stringify(actual));
}
function section(t) { console.log('\n[' + t + ']'); }

/* ---------------- 读夹具 ---------------- */
const wbPath = path.join(__dirname, 'fixture', 'sample_ledger.xlsx');
// 这个构建里没有注入 fs，readFile 会直接抛错；自己读 buffer 再 read 即可。
const wb = XLSX.read(fs.readFileSync(wbPath), { type: 'buffer', cellDates: false });

function sheetRows(name) {
  const ws = wb.Sheets[name];
  if (!ws) throw new Error('夹具缺少工作表：' + name);
  const range = XLSX.utils.decode_range(ws['!ref']);
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  return { rows, firstExcelRow: range.s.r + 1 };
}

/* ---------------- 1 金额 ---------------- */
section('1 金额：一律转成「分」整数');
eq(Core.toCents('1,234.56').cents, gt.specific.text_amount_1234_56, '文本千分位');
eq(Core.toCents('1,000.005').cents, gt.specific.three_decimal, '三位小数四舍五入');
eq(Core.toCents('1,000.005').extra, true, '三位小数应标记 extra');
eq(Core.toCents('(800.00)').cents, gt.specific.paren_negative, '括号负数');
eq(Core.toCents('500.00-').cents, gt.specific.trailing_minus, '尾随负号');
eq(Core.toCents('1,200元').cents, gt.specific.yuan_suffix, '带「元」后缀');
eq(Core.toCents(500000).cents, gt.specific.large, '数字直接传');
eq(Core.toCents(0).cents, 0, '零');
eq(Core.toCents('').reason, 'empty', '空值');
eq(Core.toCents('   ').reason, 'empty', '全空格');
eq(Core.toCents('abc').reason, 'bad', '非数字');
eq(Core.toCents('1.2.3').reason, 'bad', '多个小数点');
eq(Core.moneyText(123456), '1,234.56', '回显千分位');
eq(Core.moneyText(-80000), '-800.00', '回显负数');
eq(Core.moneyText(0), '0.00', '回显零');
eq(Core.moneyText(100000000), '1,000,000.00', '回显百万');
eq(Core.toCents(Core.moneyText(61777957)).cents, 61777957, '文本 ↔ 分 往返一致');

/* ---------------- 2 日期 ---------------- */
section('2 日期：多种写法归一');
eq(Core.dateText(Core.toDate('2026年1月5日')), '2026-01-05', '中文年月日');
eq(Core.dateText(Core.toDate('2026/7/5')), '2026-07-05', '斜杠');
eq(Core.dateText(Core.toDate('2026-1-5')), '2026-01-05', '短横');
eq(Core.dateText(Core.toDate('20260105')), '2026-01-05', '紧凑数字串');
eq(Core.dateText(Core.toDate(20260105)), '2026-01-05', '紧凑数字');
const serial = Core.toDate(46000);
ok(serial && serial.y === 2025 && serial.m >= 11, 'Excel 序列号 → 日期，实际 ' + JSON.stringify(serial));
eq(Core.toDate(''), null, '空值返回 null');
eq(Core.toDate('不是日期'), null, '乱码返回 null');
eq(Core.toDate('2026-02-30'), null, '不存在的日期返回 null');

/* ---------------- 3 表头识别 ---------------- */
section('3 表头识别：跨过标题行找到真表头');
const main = sheetRows('应收账款明细账');
const header = Core.detectHeader(main.rows);
eq(header.rowIndex, 2, '表头应落在第 3 行（下标 2）');
eq(header.map.date, 0, '日期列');
eq(header.map.voucher, 1, '凭证号列');
eq(header.map.summary, 2, '摘要列');
eq(header.map.account, 3, '科目列');
eq(header.map.debit, 4, '借方列');
eq(header.map.credit, 5, '贷方列');
eq(header.map.amount, undefined, '不该把借方/贷方误认成单列金额');
eq(header.map.balance, 6, '余额列');
eq(header.map.maker, 7, '制单人列');

/* ---------------- 4 行解析 ---------------- */
section('4 行解析：明细与非明细分开，不静默丢行');
const parsed = Core.parseRows(main.rows, header, { firstExcelRow: main.firstExcelRow });
eq(parsed.records.length, gt.row_count_total, '明细行数（与 Python 侧一致）');
eq(parsed.problems.length, 0, '不应有金额无法识别的行');
eq(parsed.warnings.blankRows, 1, '空行计数');
eq(parsed.warnings.subtotalRows, 2, '小计/合计行计数');
eq(parsed.skipped.length, 3, '跳过的非明细行全部留痕');
ok(parsed.records.every(r => r.srcRow >= 4), '原始行号应指向 Excel 真实行');

/* ---------------- 5 总体过滤 ---------------- */
section('5 总体过滤：逐笔可核，排除必留痕');
const pop = Core.buildPopulation(parsed.records, {
  dateFrom: { y: 2026, m: 1, d: 1 },
  dateTo: { y: 2026, m: 12, d: 31 },
  direction: 'all'
});
eq(pop.stats.count, gt.included, '总体笔数');
eq(pop.stats.totalAbsCents, gt.total_abs_cents, '总体金额（绝对值合计，分）');
eq(pop.stats.debitCents, gt.debit_cents, '借方合计');
eq(pop.stats.creditCents, gt.credit_cents, '贷方合计');
eq(pop.stats.netCents, gt.net_cents, '净额');
eq(pop.stats.excludedCount, 3, '排除笔数');
eq(pop.stats.debitCents - pop.stats.creditCents, pop.stats.netCents, '借贷差应等于净额');

const reasons = pop.excluded.map(e => e.reason);
eq(reasons.filter(r => r === '日期缺失').length, gt.excluded.date_missing, '排除：日期缺失');
eq(reasons.filter(r => r.indexOf('晚于') === 0).length, gt.excluded.out_of_period, '排除：期间外');
eq(reasons.filter(r => r === '金额为零').length, gt.excluded.zero, '排除：零金额');
ok(pop.excluded.every(e => e.srcRow > 0 && e.reason), '每条排除都有原始行号与原因');

/* ---------------- 6 抽样 ---------------- */
section('6 抽样：可复现、等距不变量、分层不重不漏');
const s1 = Core.buildSample(pop.included, {
  thresholdCents: gt.threshold_cents, n: 8, seed: 20260101, method: 'systematic'
});
eq(s1.meta.fullCount, gt.full_count, '全查层笔数');
eq(s1.meta.restCount, gt.rest_count, '抽样层笔数');
eq(s1.meta.fullAbsCents, gt.full_abs_cents, '全查层金额');
eq(s1.meta.restAbsCents, gt.rest_abs_cents, '抽样层金额');
eq(s1.meta.fullAbsCents + s1.meta.restAbsCents, gt.total_abs_cents, '两层金额之和 = 总体金额');
eq(s1.meta.sampleCount, gt.full_count + 8, '样本笔数 = 全查 + 抽样');
// 间距用小数 N/n：老实现用 floor，n 超过 N/2 时会退化成 1，样本变成「前 n 笔」
eq(s1.meta.k, Math.round(gt.rest_count / 8 * 100) / 100, '等距间距 = N/n（不取整）');
ok(s1.meta.start >= 0 && s1.meta.start < s1.meta.k, '起点应落在 [0,k)');
eq(s1.sample.filter(x => x.layer === '全查').length, gt.full_count, '全查层全部入样');
eq(s1.sample.filter(x => x.layer === '抽样').length, 8, '抽样层笔数正确');

const rowsA = s1.sample.map(x => x.rec.srcRow).join(',');
const s2 = Core.buildSample(pop.included, {
  thresholdCents: gt.threshold_cents, n: 8, seed: 20260101, method: 'systematic'
});
eq(s2.sample.map(x => x.rec.srcRow).join(','), rowsA, '同种子 + 同参数 → 逐笔相同');

const starts = new Set();
for (let sd = 1; sd <= 20; sd++) {
  starts.add(Core.buildSample(pop.included, { thresholdCents: gt.threshold_cents, n: 8, seed: sd }).meta.start);
}
ok(starts.size >= 2, '不同种子应能给出不同起点，实际只有 ' + starts.size + ' 种');

const fake = [];
for (let i = 0; i < 100; i++) fake.push({ signedCents: (i + 1) * 100, srcRow: i + 1, date: null, voucher: '', summary: '' });
const sys = Core.sampleSystematic(fake, 7, 42);
eq(sys.indices.length, 7, '等距抽样笔数');
eq(new Set(sys.indices).size, 7, '等距抽样不得重复');
eq(sys.k, 14.29, '等距间距 = 100/7');
ok(sys.indices.every((v, i) => i === 0 || v > sys.indices[i - 1]), '等距抽样递增');
ok(sys.indices[sys.indices.length - 1] < 100, '等距抽样不越界');

// n > N/2：老实现 k=floor(N/n)=1 会变成「前 60 笔」，覆盖不到期后
const sysDense = Core.sampleSystematic(fake, 60, 7);
eq(sysDense.indices.length, 60, '密集抽样笔数');
eq(new Set(sysDense.indices).size, 60, '密集抽样不得重复');
ok(sysDense.indices[0] < 5, '密集抽样起点应在前段，实际 ' + sysDense.indices[0]);
ok(sysDense.indices[59] > 90, '密集抽样必须覆盖到后段，实际最后一个是 ' + sysDense.indices[59]);

const rnd = Core.sampleRandom(fake, 7, 42);
eq(rnd.indices.length, 7, '随机抽样笔数');
eq(new Set(rnd.indices).size, 7, '随机抽样不得重复');
eq(rnd.indices.join(','), Core.sampleRandom(fake, 7, 42).indices.join(','), '随机抽样同样可复现');

const all = Core.buildSample(fake, { thresholdCents: 0, n: 200, seed: 1 });
eq(all.meta.sampleCount, 100, '样本量大于总体时全部抽取');

// 同一张凭证被抽中多行：现场只需翻一次，所以单独统计凭证张数
eq(s1.meta.voucherCount, s1.meta.sampleCount, '夹具里每行一张凭证，凭证数应等于笔数');
const dupPop = [];
for (let i = 0; i < 6; i++) {
  dupPop.push({ signedCents: 100, srcRow: i + 1, date: null, voucher: i < 3 ? '记-A' : '记-B', summary: '', account: '' });
}
const dup = Core.buildSample(dupPop, { thresholdCents: 0, n: 6, seed: 5 });
eq(dup.meta.sampleCount, 6, '六行全抽');
eq(dup.meta.voucherCount, 2, '六行只涉及两张凭证，应报 2 而不是 6');

/* ---------------- 7 抽凭表输出 ---------------- */
section('7 抽凭表输出：合计与样本一致，说明齐全');
const sheets = Core.buildSheets({
  client: gt.client,
  indexNo: 'B-3',
  preparer: '实习生',
  reviewer: '项目经理',
  dateText: '2026-11-21',
  periodText: '2026-01-01 至 2026-12-31',
  accountText: '1122 应收账款',
  presetName: '销售/应收',
  sample: s1.sample,
  meta: s1.meta,
  populationStats: pop.stats,
  excluded: pop.excluded,
  problems: parsed.problems,
  skipped: parsed.skipped,
  warnings: Object.assign({ parsedRows: parsed.records.length }, parsed.warnings),
  mapNotes: ['日期列按「记账日期」识别']
});
eq(sheets.length, 5, '应输出 5 个工作表');
eq(sheets.map(s => s.name).join(','), '抽凭表,抽样信息,排除清单,数据检查,说明', '工作表顺序');

const mainRows = sheets[0].rows;
const headIdx = mainRows.findIndex(r => r[0] === '序号');
ok(headIdx > 0, '抽凭表应有表头行');
const head = mainRows[headIdx];
eq(head[0], '序号', '首列序号');
eq(head[1], '原始行号', '第 2 列原始行号（凭证号丢失时靠它回溯）');
eq(head[11], '抽样类别', '第 12 列抽样类别');
ok(head.indexOf('核对结论') > 10, '应带核对结论列供回录');
ok(head.indexOf('错报金额（元）') > 10, '应带错报金额列供回录');
ok(head.indexOf('发票') > 10, '销售预设应含发票列');
ok(head.indexOf('签收单') > 10, '销售预设应含签收单列');

const dataRows = mainRows.slice(headIdx + 1, headIdx + 1 + s1.meta.sampleCount);
eq(dataRows.length, s1.meta.sampleCount, '数据行数 = 样本笔数');
eq(dataRows[0][0], 1, '序号从 1 开始');
eq(dataRows[dataRows.length - 1][0], s1.meta.sampleCount, '序号连续到末行');

const totalRow = mainRows[headIdx + 1 + s1.meta.sampleCount];
eq(totalRow[0], '合计', '末行合计');
const sumDr = s1.sample.reduce((s, x) => s + (x.rec.debitCents || 0), 0);
const sumAbs = s1.sample.reduce((s, x) => s + Math.abs(x.rec.signedCents), 0);
eq(Math.round(totalRow[7] * 100), sumDr, '合计借方 = 样本借方之和');
eq(Math.round(totalRow[9] * 100), sumAbs, '合计绝对额 = 样本绝对额之和');
eq(Math.round(totalRow[9] * 100), s1.meta.sampleAbsCents, '合计绝对额 = 抽样元数据');

const infoRows = sheets[1].rows;
ok(infoRows.some(r => r[0] === '随机种子' && r[1] === 20260101), '抽样信息应记录随机种子');
ok(infoRows.some(r => r[0] === '起点'), '抽样信息应记录起点');
ok(infoRows.some(r => r[0] === '等距'), '抽样信息应记录间距');

const exclRows = sheets[2].rows;
eq(exclRows.length, 2 + pop.excluded.length, '排除清单行数 = 表头 + 每条排除');

const helpText = sheets[4].rows.map(r => r.join('')).join('\n');
ok(helpText.indexOf('不联网') >= 0, '说明应写明数据不出本机');
ok(helpText.indexOf('不代替判断') >= 0 || helpText.indexOf('不代替') >= 0, '说明应写明不代替判断');

/* ---------------- 8 单列金额 + 方向 ---------------- */
section('8 另一种账表：单列金额 + 借贷方向');
const am = sheetRows('单列金额');
const h2 = Core.detectHeader(am.rows);
eq(h2.map.amount, 3, '单列金额列');
eq(h2.map.direction, 4, '方向列');
const p2 = Core.parseRows(am.rows, h2, { firstExcelRow: am.firstExcelRow });
eq(p2.records.length, gt.amount_mode.count, '单列金额明细行数');
const pop2 = Core.buildPopulation(p2.records, {});
eq(pop2.stats.totalAbsCents, gt.amount_mode.total_abs_cents, '单列金额总体绝对额');
eq(pop2.stats.netCents, gt.amount_mode.net_cents, '单列金额净额（方向列生效）');

/* ---------------- 9 文件名 ---------------- */
section('9 文件名：非法字符替换');
const fn = Core.fileName({ client: '甲/乙 公司', accountText: '应收账款' });
ok(/^抽凭表_甲_乙_公司_应收账款_\d{8}\.xlsx$/.test(fn), '非法字符应替换，实际 ' + fn);

/* ---------------- 汇总 ---------------- */
console.log('\n' + '='.repeat(56));
if (fails.length) {
  console.log('失败 ' + fails.length + ' / ' + count + ' 项：');
  fails.forEach((f, i) => console.log('  ' + (i + 1) + '. ' + f));
  process.exit(1);
}
console.log('全部通过：' + count + ' 项断言');
console.log('='.repeat(56));
