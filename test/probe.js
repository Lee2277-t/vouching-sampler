'use strict';
/* 诊断脚本：把任意一份账表丢给 core，看它认成了什么。
 * 用法： node test/probe.js <文件> [工作表名] [--show]
 *   --show  额外打印前 5 条明细（含摘要与金额）。处理真实底稿时不要加这个开关。
 * 不写任何文件，只打印。拿不准某张表能不能用时先跑这个。
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('../vendor/xlsx.full.min.js');
const Core = require('../core.js');

const args = process.argv.slice(2);
const SHOW = args.indexOf('--show') >= 0;
const rest = args.filter(a => a !== '--show');
const file = rest[0];
const only = rest[1];
if (!file) {
  console.error('用法: node test/probe.js <xlsx文件> [工作表名]');
  process.exit(2);
}

const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer', cellDates: false });
console.log('文件:', path.basename(file));
console.log('工作表:', wb.SheetNames.join(' | '));

const names = only ? [only] : wb.SheetNames;

for (const name of names) {
  const ws = wb.Sheets[name];
  if (!ws) { console.log(`\n[${name}] 不存在`); continue; }
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });

  console.log('\n' + '='.repeat(72));
  console.log(`[${name}] 尺寸 ${ws['!ref']}  起始行 ${range.s.r + 1}  共 ${rows.length} 行`);
  console.log('前 3 行原始内容:');
  rows.slice(0, 3).forEach((r, i) => console.log('  ', i + 1, JSON.stringify((r || []).slice(0, 12))));

  const header = Core.detectHeader(rows);
  const verdict = Core.sheetVerdict(header);
  console.log(`\n表头识别: 第 ${header.rowIndex + range.s.r + 1} 行 (下标 ${header.rowIndex}), 得分 ${header.score}`);
  console.log(`可用性: ${verdict.usable ? '可用 ✓' : '不可用 ✗'}${verdict.problems.length ? '  ' + verdict.problems.join('；') : ''}`);
  console.log('表头原文:', JSON.stringify((rows[header.rowIndex] || []).slice(0, 12)));
  const mapped = {};
  for (const f in header.map) mapped[f] = { 列: header.map[f] + 1, 表头: (rows[header.rowIndex] || [])[header.map[f]] };
  console.log('列映射:', JSON.stringify(mapped, null, 0));

  const parsed = Core.parseRows(rows, header, { firstExcelRow: range.s.r + 1 });
  console.log(`\n解析: 明细 ${parsed.records.length} 条, 跳过 ${parsed.skipped.length} 行, 问题 ${parsed.problems.length} 行`);
  if (parsed.skipped.length) {
    const byReason = {};
    parsed.skipped.forEach(s => { byReason[s.reason] = (byReason[s.reason] || 0) + 1; });
    console.log('  跳过原因:', JSON.stringify(byReason));
  }
  if (parsed.problems.length) {
    console.log('  问题行:', parsed.problems.slice(0, 5).map(p => `第${p.srcRow}行 ${p.reason}`).join('; '));
  }

  const pop = Core.buildPopulation(parsed.records, {});
  console.log(`总体(不过滤): ${pop.stats.count} 笔, 金额合计 ${Core.moneyText(pop.stats.totalAbsCents)}`);
  if (SHOW) {
    console.log('前 5 条:');
    parsed.records.slice(0, 5).forEach(r => {
      console.log(`   行${r.srcRow} ${r.dateText || '(无日期)'} | 凭证「${r.voucher}」 | ${r.account} | ${r.summary} | 借 ${Core.moneyText(r.debitCents)} 贷 ${Core.moneyText(r.creditCents)}`);
    });
  } else {
    console.log('  （明细含客户信息，默认不打印；需要看就加 --show）');
    const withVoucher = parsed.records.filter(r => r.voucher).length;
    const withAccount = parsed.records.filter(r => r.account).length;
    const withDate = parsed.records.filter(r => r.date).length;
    console.log(`  字段完整度：凭证号 ${withVoucher}/${parsed.records.length}，科目 ${withAccount}，日期 ${withDate}`);
  }

  const missing = [];
  if (header.map.date === undefined) missing.push('日期');
  if (header.map.voucher === undefined) missing.push('凭证号');
  if (header.map.debit === undefined && header.map.credit === undefined && header.map.amount === undefined) missing.push('金额');
  if (missing.length) console.log('\n!! 缺列:', missing.join('、'));
}
