'use strict';
/* 真实序时账回归
 *
 * 数据是公开仓库里的样例（不是客户数据）：
 *   aging      github.com/Cathywang1988/aging-generator
 *   valuation  github.com/zer0-lyz/valuation-detail-table
 * 期望值由 test/make_real_truth.py 用 openpyxl 独立算，见 fixture/real_truth.json。
 *
 * 这两份表的价值在于「脏得真实」：一份没有凭证号列，一份把凭证号拆成「字」+「号」，
 * 还把科目编码和科目全路径放在同一张表里。它们打回过三个真实缺陷，留着防复发。
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('../vendor/xlsx.full.min.js');
const Core = require('../core.js');
const GT = require('./fixture/real_truth.json');

let count = 0;
const fails = [];
function ok(cond, msg) { count++; if (!cond) fails.push(msg); }
function eq(a, b, msg) {
  count++;
  if (a !== b) fails.push(msg + ' — 期望 ' + JSON.stringify(b) + '，实际 ' + JSON.stringify(a));
}
function section(t) { console.log('\n[' + t + ']'); }

function load(name) {
  const p = path.join(__dirname, 'fixture', name);
  const wb = XLSX.read(fs.readFileSync(p), { type: 'buffer', cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const range = XLSX.utils.decode_range(ws['!ref']);
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  const header = Core.detectHeader(rows);
  const parsed = Core.parseRows(rows, header, { firstExcelRow: range.s.r + 1 });
  const pop = Core.buildPopulation(parsed.records, {});
  return { rows, header, parsed, pop };
}

/* ---------------- 1 只有科目、没有凭证号的表 ---------------- */
section('1 无凭证号列：金额要照样算对，缺列要说出来');
const a = load('aging_xushizhang.xlsx');
const gtA = GT['aging_xushizhang.xlsx'];

eq(a.header.map.date, 2, '日期列');
eq(a.header.map.debit, 4, '借方列');
eq(a.header.map.credit, 5, '贷方列');
eq(a.parsed.records.length, gtA.transactions, '明细行数');
eq(a.parsed.warnings.noVoucher, true, '应标记「未识别到凭证号」');
eq(a.pop.stats.totalAbsCents, gtA.total_abs_cents, '金额合计');
eq(a.pop.stats.debitCents, gtA.debit_cents, '借方合计');
eq(a.pop.stats.creditCents, gtA.credit_cents, '贷方合计');
eq(a.parsed.skipped.length, gtA.skipped_no_amount, '无金额行数');

// 科目必须落到名称而不是编码，否则按「应收账款」筛不出来
eq(a.parsed.records[0].account, '应收账款', '科目取的是名称，不是编码');
eq(a.parsed.records[0].accountCode, '1122', '科目编码单独存放');
const filtered = Core.buildPopulation(a.parsed.records, { accountText: '应收' });
eq(filtered.stats.count, gtA.account_filter_receivable, '按「应收」筛的命中数（若认成编码会是 0 笔）');

/* ---------------- 2 凭证号拆成「字」+「号」的表 ---------------- */
section('2 凭证号拆两列：要拼成「记账凭证-1」');
const b = load('valuation_xushizhang.xlsx');
const gtB = GT['valuation_xushizhang.xlsx'];

eq(b.parsed.records.length, gtB.transactions, '明细行数（期初余额行应被剔除）');
eq(b.parsed.skipped.length, gtB.skipped_no_amount, '无金额行数');
eq(b.parsed.skipped[0].reason, '无金额（借贷两列都为空）', '期初余额应记为「无金额」而不是「零金额」');
eq(b.parsed.records[0].voucher, '记账凭证-1', '字 + 号 应拼接');
eq(b.parsed.records[2].voucher, '付款凭证-2', '第二张凭证的拼接');
eq(b.pop.stats.totalAbsCents, gtB.total_abs_cents, '金额合计');
eq(b.pop.stats.debitCents, gtB.debit_cents, '借方合计');
eq(b.pop.stats.creditCents, gtB.credit_cents, '贷方合计');
eq(b.parsed.records[0].account, '银行存款/工商银行', '科目取末级科目全路径');
eq(b.parsed.records[0].accountCode, '100201', '末级科目编号单独存放');

// 这份表是筛选出来的单科目摘录，借贷本就不平；记下来，免得日后误当缺陷
ok(gtB.debit_cents !== gtB.credit_cents, '摘录表借贷不平是正常的，不要按「不平就是错」处理');

/* ---------------- 汇总 ---------------- */
console.log('\n' + '='.repeat(56));
if (fails.length) {
  console.log('失败 ' + fails.length + ' / ' + count + ' 项：');
  fails.forEach((f, i) => console.log('  ' + (i + 1) + '. ' + f));
  process.exit(1);
}
console.log('全部通过：' + count + ' 项断言');
console.log('='.repeat(56));
