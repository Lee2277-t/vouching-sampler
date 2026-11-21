'use strict';
/* 真实世界格式回归
 *
 * 四张表的结构照抄自真实底稿，数据全是合成的（见 make_fixture.py）：
 *   序时账     标题占两行，第 3 行才是表头，凭证号拆成「字」+「号」
 *   明细表     SAP 式扁平导出，科目和借贷散在很后面
 *   填制凭证   日期列叫「制单日期」
 *   科目余额表 两行标题 + 两级表头，没有日期 —— 绝对不能当抽凭总体
 *
 * 每一条都对应一次真实的翻车。
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('../vendor/xlsx.full.min.js');
const Core = require('../core.js');
const gt = require('./fixture/ground_truth.json').realworld;

let count = 0;
const fails = [];
function ok(cond, msg) { count++; if (!cond) fails.push(msg); }
function eq(a, b, msg) {
  count++;
  if (a !== b) fails.push(msg + ' — 期望 ' + JSON.stringify(b) + '，实际 ' + JSON.stringify(a));
}
function section(t) { console.log('\n[' + t + ']'); }

const wb = XLSX.read(
  fs.readFileSync(path.join(__dirname, 'fixture', 'sample_ledger.xlsx')),
  { type: 'buffer', cellDates: false }
);

function load(name) {
  const ws = wb.Sheets[name];
  if (!ws) throw new Error('夹具缺工作表：' + name);
  const range = XLSX.utils.decode_range(ws['!ref']);
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  const header = Core.detectHeader(rows);
  return {
    header,
    verdict: Core.sheetVerdict(header),
    parsed: Core.parseRows(rows, header, { firstExcelRow: range.s.r + 1 }),
    pop: Core.buildPopulation(Core.parseRows(rows, header, { firstExcelRow: range.s.r + 1 }).records, {})
  };
}

/* ---------------- 1 金蝶式序时账 ---------------- */
section('1 标题占两行 + 凭证号拆列：往里翻两行才找到表头');
const xs = load('序时账');

eq(xs.header.rowIndex, 2, '表头应在第 3 行');
eq(xs.header.map.date, 0, '日期列');
eq(xs.header.map.voucherWord, 1, '「字」列');
eq(xs.header.map.voucher, 2, '「号」列');
eq(xs.header.map.accountCode, 4, '「科目代码」进编码字段');
eq(xs.header.map.account, 5, '「科目名称」进科名字段');
eq(xs.header.map.debit, 6, '借方列');
eq(xs.header.map.credit, 7, '贷方列');
eq(xs.verdict.usable, true, '应判定为可用');
eq(xs.parsed.records.length, gt.xushizhang.records, '明细行数');
eq(xs.pop.stats.totalAbsCents, gt.xushizhang.total_abs_cents, '金额合计');
eq(xs.parsed.records[0].voucher, gt.xushizhang.first_voucher, '「记」+1 应拼成「记-1」');
eq(xs.parsed.records[0].account, gt.xushizhang.account, '科目取名称');
eq(xs.parsed.records[0].accountCode, gt.xushizhang.account_code, '科目代码单独存');
ok(xs.parsed.records[0].date !== null, '日期应解析出来');

/* ---------------- 2 SAP 式扁平导出 ---------------- */
section('2 SAP 式扁平导出：科目和借贷散在十几列之后');
const sap = load('明细表');

eq(sap.header.map.date, 8, '「凭证日期」列');
eq(sap.header.map.voucher, 2, '「凭证编号」列');
eq(sap.header.map.account, 12, '「科目」列');
eq(sap.header.map.debit, 14, '「借方」列');
eq(sap.header.map.credit, 15, '「贷方」列');
eq(sap.verdict.usable, true, '应判定为可用');
eq(sap.parsed.records.length, gt.sap.records, '明细行数');
eq(sap.pop.stats.totalAbsCents, gt.sap.total_abs_cents, '金额合计');

/* ---------------- 3 制单日期 ---------------- */
section('3 日期列叫「制单日期」：以前这里认不出来');
const tz = load('填制凭证');

eq(tz.header.map.date, 0, '「制单日期」应被认成日期列');
eq(tz.verdict.usable, true, '应判定为可用');
eq(tz.parsed.records.length, gt.tianzhi.records, '明细行数');
eq(tz.pop.stats.totalAbsCents, gt.tianzhi.total_abs_cents, '金额合计');
eq(tz.parsed.records[0].voucher, gt.tianzhi.first_voucher, '凭证字号');
eq(tz.header.map.debit, 9, '「借方(本币)」应认成借方');
eq(tz.header.map.credit, 10, '「贷方(本币)」应认成贷方');

/* ---------------- 4 科目余额表必须被拒绝 ---------------- */
section('4 科目余额表：有借贷金额也不能当抽凭总体');
const bs = load('科目余额表');

eq(bs.verdict.usable, gt.balance_sheet_usable, '应判定为不可用');
eq(bs.verdict.hasDate, false, '不应认出日期列');
ok(bs.verdict.problems.indexOf('没有日期列') >= 0, '应明确指出没有日期列');
ok(bs.verdict.problems.some(p => p.indexOf('科目余额表') >= 0 || p.indexOf('报表') >= 0),
   '应提示看起来像余额表/报表');

/* ---------------- 5 自动挑表：挑对，且不选余额表 ---------------- */
section('5 同一个工作簿里挑表：能用的排在前面');
const names = ['序时账', '明细表', '填制凭证', '科目余额表'];
const ranked = names.map(n => load(n)).map((x, i) => ({ name: names[i], verdict: x.verdict }));
const sorted = ranked.slice().sort((a, b) => b.verdict.score - a.verdict.score);

eq(sorted[sorted.length - 1].name, '科目余额表', '余额表应排在最后');
ok(sorted.slice(0, 3).every(r => r.verdict.usable), '前三名都应是可用表');
ok(sorted[0].verdict.score > sorted[sorted.length - 1].verdict.score, '可用表得分应显著高于余额表');

/* ---------------- 汇总 ---------------- */
console.log('\n' + '='.repeat(56));
if (fails.length) {
  console.log('失败 ' + fails.length + ' / ' + count + ' 项：');
  fails.forEach((f, i) => console.log('  ' + (i + 1) + '. ' + f));
  process.exit(1);
}
console.log('全部通过：' + count + ' 项断言');
console.log('='.repeat(56));
