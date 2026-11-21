'use strict';
/* v0.4 / v0.5：逐条消灭已知限制
 *
 * 每条限制一个 section，断言就是「这条限制已经解决」的证据。
 *   1 一个科目区间   → 多关键词（空格/逗号分隔，命中任一）
 *   2 多主体不分     → 主体列识别 + 按主体筛选
 *   3 抽样单元只有行 → 按凭证抽样（合并同一凭证的多行）
 *   4 两级表头       → 表头两行拼接
 *   5 判断抽样无推断 → 比率法 / 差异法推断错报
 *   6 结论写法认不出 → 扩充词表
 *   7 英文表头认不出 → 中英别名混排
 *   8 跨文件合并     → 在 ui-smoke.test.js §6（要假 DOM 才跑得动）
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

/* 造一批可控的明细行 */
function rec(srcRow, day, voucher, account, debit, credit, extra) {
  const dr = debit === null ? null : debit;
  const cr = credit === null ? null : credit;
  const o = {
    srcRow, date: { y: 2026, m: 1, d: day }, dateText: '2026-01-' + String(day).padStart(2, '0'),
    voucher, summary: 's' + srcRow, account, accountCode: '',
    debitCents: dr, creditCents: cr, signedCents: (dr || 0) - (cr || 0)
  };
  return Object.assign(o, extra || {});
}

const RECS = [
  rec(2, 5, '记-1', '1122 应收账款', 100000, null),
  rec(3, 5, '记-1', '6001 主营业务收入', null, 100000),
  rec(4, 6, '记-2', '1122 应收账款', 50000, null),
  rec(5, 7, '记-3', '1221 其他应收款', 30000, null),
  rec(6, 8, '记-4', '2202 应付账款', null, 20000)
];

/* ---------------- 1 多科目关键词 ---------------- */
section('1 一个科目区间 → 多关键词');
eq(Core.buildPopulation(RECS, { accountText: '应收账款' }).stats.count, 2, '单个关键词');
eq(Core.buildPopulation(RECS, { accountText: '应收账款 其他应收款' }).stats.count, 3, '空格分隔两个关键词');
eq(Core.buildPopulation(RECS, { accountText: '应收，其他' }).stats.count, 3, '逗号也认');
eq(Core.buildPopulation(RECS, { accountText: '应收账款 1221' }).stats.count, 3, '关键词可以混编码');
eq(Core.buildPopulation(RECS, { accountText: '应付、收入' }).stats.count, 2, '顿号分隔');
const multiPop = Core.buildPopulation(RECS, { accountText: '应收账款 其他应收款' });
eq(multiPop.stats.excludedCount, 2, '没命中的走排除清单');
ok(multiPop.excluded.every(e => e.reason === '科目不含任一关键词'), '排除原因写明是关键词');

/* ---------------- 2 多主体 ---------------- */
section('2 多主体账表 → 识别主体列并按主体筛');
const withEntity = RECS.map((r, i) => Object.assign({}, r, { entity: i < 3 ? '甲公司' : '乙公司' }));

eq(Core.buildPopulation(withEntity, { entity: '甲公司' }).stats.count, 3, '按主体筛甲公司');
eq(Core.buildPopulation(withEntity, { entity: '乙公司' }).stats.count, 2, '按主体筛乙公司');
eq(Core.buildPopulation(withEntity, {}).stats.count, 5, '不选主体就是全部');
eq(Core.buildPopulation(withEntity, { entity: '丙公司' }).stats.count, 0, '不存在的主体得到 0 笔');
const entExcl = Core.buildPopulation(withEntity, { entity: '甲公司' }).excluded;
ok(entExcl.every(e => e.reason === '不属于所选主体'), '排除原因写明主体不符');

// 列识别：公司代码描述优先于公司代码
const entHeader = Core.detectHeader([
  ['公司代码', '公司代码描述', '凭证日期', '凭证编号', '科目', '借方', '贷方'],
  ['V100', '合成甲公司', '2026-01-05', '2000001', '100201 银行存款', 100, null]
]);
eq(entHeader.map.entity, 1, '主体优先取「公司代码描述」这种可读名字');

/* ---------------- 3 按凭证抽样 ---------------- */
section('3 抽样单元只有明细行 → 可按凭证抽样');
const linePop = Core.buildPopulation(RECS, { unit: 'line' });
const vouPop = Core.buildPopulation(RECS, { unit: 'voucher' });

eq(linePop.unit, 'line', '默认按明细行');
eq(linePop.stats.count, 5, '明细行 5 笔');
eq(linePop.stats.totalAbsCents, 300000, '明细行金额合计 3,000.00');

eq(vouPop.unit, 'voucher', '按凭证');
eq(vouPop.stats.count, 4, '4 张凭证');
// 记-1 借 1000 贷 1000：凭证金额取借方 1000，不是净额 0
eq(vouPop.stats.totalAbsCents, 100000 + 50000 + 30000 + 20000, '凭证金额合计（借贷平衡的凭证不塌成 0）');
const g1 = vouPop.included.filter(g => g.voucher === '记-1')[0];
eq(g1.lineCount, 2, '记-1 有 2 行');
eq(g1.signedCents, 100000, '借贷平衡的凭证金额取借方合计');
eq(g1.rows.join(','), '2,3', '凭证保留全部原始行号');
eq(g1.debitCents, 100000, '借方合计');
eq(g1.creditCents, 100000, '贷方合计');

// 按凭证抽出来的样本，凭证号不重复
const vouSample = Core.buildSample(vouPop.included, { thresholdCents: 0, n: 4, seed: 7 });
eq(vouSample.meta.voucherCount, vouSample.meta.sampleCount, '按凭证抽样时不会重复抽到同一张');
eq(vouSample.sample.filter(x => x.rec.lineCount > 1).length, 1, '多行凭证只出现一次');

// 科目筛完再归凭证：只归命中科目的行
const filtered = Core.buildPopulation(RECS, { accountText: '应收账款', unit: 'voucher' });
eq(filtered.stats.count, 2, '应收账款只有记-1、记-2 两张凭证');

/* ---------------- 4 两级表头 ---------------- */
section('4 两级合并表头 → 两行拼成一层');
const wb = XLSX.read(
  fs.readFileSync(path.join(__dirname, 'fixture', 'sample_ledger.xlsx')),
  { type: 'buffer', cellDates: false }
);
const ws = wb.Sheets['两级表头'];
const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
const h = Core.detectHeader(rows);

eq(h.rowIndex, 0, '表头从第 1 行开始');
eq(h.span, 2, '占用两行');
eq(h.map.date, 0, '日期列');
eq(h.map.voucher, 1, '凭证号列');
eq(h.map.debit, 4, '「本期发生额 借方金额」→ 借方');
eq(h.map.credit, 5, '「本期发生额 贷方金额」→ 贷方');
ok(h.labels[4].indexOf('借方') >= 0, '拼接后的标签带上下级信息：' + h.labels[4]);

const parsed = Core.parseRows(rows, h, { firstExcelRow: 1 });
eq(parsed.records.length, gt.twolevel.records, '数据从第 3 行开始读，条数正确');
eq(Core.buildPopulation(parsed.records, {}).stats.totalAbsCents, gt.twolevel.total_abs_cents, '金额合计');
eq(Core.sheetVerdict(h).usable, true, '两级表头也判定为可用');

/* ---------------- 5 判断抽样推断 ---------------- */
section('5 判断抽样 → 比率法与差异法推断错报');
const evRecords = [
  { seq: 1, srcRow: 2, voucher: '记-1', account: '应收账款', absCents: 100000, conclusion: 'ok', misCents: 0 },
  { seq: 2, srcRow: 3, voucher: '记-2', account: '应收账款', absCents: 50000, conclusion: 'exception', misCents: -5000 },
  { seq: 3, srcRow: 4, voucher: '记-3', account: '应收账款', absCents: 50000, conclusion: 'ok', misCents: 0 },
  { seq: 4, srcRow: 5, voucher: '记-4', account: '应收账款', absCents: 50000, conclusion: 'ok', misCents: 0 }
];
const ev = Core.evaluateVouching(evRecords, {
  alpha: 0.05, populationCents: 25000000, populationCount: 40, sampleMethod: 'pps'
});
eq(ev.projection.sampleBookCents, 250000, '样本账面合计');
eq(ev.projection.sampleMisAbsCents, 5000, '样本错报合计');
// 5000 / 250000 × 25,000,000 = 500,000
eq(ev.projection.ratioCents, 500000, '比率法推断');
// 5000 / 4 × 40 = 50,000
eq(ev.projection.differenceCents, 50000, '差异法推断');
ok(ev.upper && ev.upper.upperCents > 0, '统计上限照样算，两种口径并存');

const sheet = Core.buildEvaluationSheet(ev, {});
const text = sheet.rows.map(r => r.join('')).join('\n');
ok(text.indexOf('推断错报（比率法）') >= 0, '评估表写出比率法');
ok(text.indexOf('推断错报（差异法）') >= 0, '评估表写出差异法');
ok(text.indexOf('由审计人员判断') >= 0, '两种方法并存时说明由人判断');

/* ---------------- 6 结论写法 ---------------- */
section('6 结论写法：常见说法都要认');
const cases = {
  '相符': 'ok', '符合': 'ok', '无误': 'ok', '正常': 'ok', '通过': 'ok', '√': 'ok',
  '不符': 'exception', '不符合': 'exception', '不一致': 'exception', '异常': 'exception', '有误': 'exception',
  '未取得': 'missing', '未回函': 'missing', '未收到回函': 'missing', '无法取得': 'missing',
  '不适用': 'na', '本期无发生': 'na', '—': 'na'
};
Object.keys(cases).forEach(k => eq(Core.classifyConclusion(k), cases[k], '结论「' + k + '」'));
eq(Core.classifyConclusion('不符合规定'), 'exception', '带后缀的否定说法应为不符，不能因为含「符合」判成相符');
eq(Core.classifyConclusion('已取得发票'), 'other', '长句不猜，报 other 让人看');

/* ---------------- 7 英文表头 ---------------- */
section('7 英文表头：SAP / 英文界面的导出也要认');
const enRows = [
  ['Company Code', 'Posting Date', 'Document Number', 'Line Text', 'G/L Account',
   'Account Number', 'Debit', 'Credit', 'Amount in local currency', 'Vendor Name'],
  ['1000', '2026-01-05', '2000001', 'invoice 001', '100201 银行存款', '100201', '100.00', '', '100.00', 'ACME'],
  ['1000', '2026-01-06', '2000002', 'invoice 002', '112201 应收账款', '112201', '', '250.00', '250.00', 'ACME']
];
const enH = Core.detectHeader(enRows);

eq(enH.rowIndex, 0, '英文表头在第 1 行');
eq(enH.map.date, 1, 'Posting Date → 日期');
eq(enH.map.voucher, 2, 'Document Number → 凭证号');
eq(enH.map.summary, 3, 'Line Text → 摘要');
eq(enH.map.account, 4, 'G/L Account → 科目（斜杠要去掉）');
eq(enH.map.accountCode, 5, 'Account Number → 科目编码，不能和科目抢');
eq(enH.map.debit, 6, 'Debit → 借方');
eq(enH.map.credit, 7, 'Credit → 贷方');
eq(enH.map.entity, 0, 'Company Code → 主体');
eq(enH.map.counterpart, 9, 'Vendor Name → 对方单位');
eq(Core.sheetVerdict(enH).usable, true, '英文表也判定为可用');

const enParsed = Core.parseRows(enRows, enH, { firstExcelRow: 1 });
eq(enParsed.records.length, 2, '英文表解析出 2 行');
eq(Core.buildPopulation(enParsed.records, {}).stats.totalAbsCents, 35000, '100 + 250 = 350.00');
eq(enParsed.records[0].dateText, '2026-01-05', 'Posting Date 的日期要能解析');

// 中英混排（常见于外资所的底稿模板）
const mixH = Core.detectHeader([['期间', 'Posting Date', '凭证号', 'Description', '科目', 'Debit', 'Credit']]);
eq(mixH.map.date, 1, '中英混排：Posting Date 仍认成日期');
eq(mixH.map.voucher, 2, '中英混排：中文凭证号优先');
eq(mixH.map.summary, 3, '中英混排：Description 认成摘要');

// 不能被英文别名带偏
eq(Core.detectHeader([['记账日期', '凭证号', '科目', '借方金额', '贷方金额']]).map.date, 0,
   '中文表头不应因为加了英文别名而认错');

/* ---------------- 8 竖排表 ---------------- */
section('8 竖排表：一笔业务一列，转置后才能抽');
const wbV = XLSX.read(
  fs.readFileSync(path.join(__dirname, 'fixture', 'sample_ledger.xlsx')),
  { type: 'buffer', cellDates: false }
);
const vRows = XLSX.utils.sheet_to_json(wbV.Sheets['竖排凭证'], {
  header: 1, raw: true, defval: null, blankrows: true
});

// 直接读：表头只有「日期」一个字段能认出来，没有金额列
const direct = Core.detectHeader(vRows);
eq(Core.sheetVerdict(direct).usable, false, '原方向读不出金额列，判定不可用');

// 转置后：表头是原来最左边那一列（日期/凭证号/摘要/科目/借方）
const tr = Core.transposeRows(vRows);
eq(tr.length, 4, '5 行 4 列转置成 4 行 5 列');
eq(tr[0][0], '日期', '转置后第 1 行第 1 格还是「日期」');
eq(tr[1][1], '记-01', '转置后第 2 行第 2 格是原表第一笔的凭证号');

const picked = Core.maybeTranspose(vRows, false);
eq(picked.transposed, true, '不可用时会自动试转置');
eq(picked.verdict.usable, true, '转置后判定为可用');
eq(picked.header.map.date, 0, '转置后能认出日期列');
eq(picked.header.map.voucher, 1, '转置后能认出凭证号列');
eq(picked.header.map.debit, 4, '转置后能认出借方列');

const vParsed = Core.parseRows(picked.rows, picked.header, { firstExcelRow: 1 });
eq(vParsed.records.length, gt.vertical.records, '转置后解析出 3 笔');
eq(Core.buildPopulation(vParsed.records, {}).stats.totalAbsCents, gt.vertical.total_abs_cents,
   '100 + 200 + 300 = 600.00');

// 本来就是横表时不许乱转
const normalRows = [
  ['日期', '凭证号', '科目', '借方'],
  ['2026-01-05', '记-1', '应收账款', '100.00']
];
const keep = Core.maybeTranspose(normalRows, false);
eq(keep.transposed, false, '正常横表不应被转置');
eq(keep.header.map.date, 0, '正常横表原样返回');

// 强制转置也要守规矩：转完不可用就不转
const junk = [['甲', '乙'], ['丙', '丁']];
eq(Core.maybeTranspose(junk, false).transposed, false, '怎么转都没用的表不折腾');

/* ---------------- 9 记住冷门列名 ---------------- */
section('9 冷门列名：手工指定一次，下次自动认');
Core.forgetAliases();

const oddHeader = ['SZRQ', 'PZH', 'ZHAIYAO', 'KM', 'JFJE', 'DFJE'];
const before = Core.detectHeader([oddHeader, ['2026-01-05', '记-1', 'x', '应收', '100', '']]);
eq(before.map.date, undefined, '没教过之前认不出来');

ok(Core.learnAlias('SZRQ', 'date'), '可以教它 SZRQ 是日期');
ok(Core.learnAlias('PZH', 'voucher'), '可以教它 PZH 是凭证号');
ok(Core.learnAlias('JFJE', 'debit'), '教 JFJE 是借方');
ok(Core.learnAlias('DFJE', 'credit'), '教 DFJE 是贷方');
eq(Core.learnAlias('甲', 'date'), false, '单字不收，太容易误伤');
eq(Core.learnAlias('XYZ', '不存在的字段'), false, '不认识的字段名不收');

const after = Core.detectHeader([oddHeader, ['2026-01-05', '记-1', 'x', '应收', '100', '']]);
eq(after.map.date, 0, '教过之后 SZRQ 认成日期');
eq(after.map.voucher, 1, 'PZH 认成凭证号');
eq(after.map.debit, 4, 'JFJE 认成借方');
eq(Core.sheetVerdict(after).usable, true, '教完就可用');

// 存回去、读回来（app 用 localStorage 干这件事）
const saved = Core.learnedAliases();
eq(Object.keys(saved).length, 4, '记住 4 条');
Core.forgetAliases();
eq(Object.keys(Core.learnedAliases()).length, 0, '可以全部忘掉');
eq(Core.detectHeader([oddHeader, ['2026-01-05', '记-1', 'x', '应收', '100', '']]).map.date, undefined,
   '忘掉之后又认不出来');
Core.setLearnedAliases(saved);
eq(Core.detectHeader([oddHeader, ['2026-01-05', '记-1', 'x', '应收', '100', '']]).map.date, 0,
   '从存档恢复后应重新认得');

// 教过的写法优先于内置别名，但不能影响别的表头
Core.learnAlias('摘要', 'voucher');
eq(Core.detectHeader([['摘要', '科目', '金额']]).map.voucher, 0, '教过的优先于内置别名');
eq(Core.detectHeader([['日期', '凭证号', '科目', '借方金额']]).map.date, 0, '不影响没教过的表头');
Core.forgetAliases();

/* ---------------- 10 金额越大越容易被抽中 ---------------- */
section('10 按金额加权：大额必进样本，不是随机撒点');
// 1 笔 100 万 + 99 笔各 1 万：总额 199 万，抽 10 笔 → 间隔 19.9 万
const ppsPop = [];
ppsPop.push({ signedCents: 100000000, srcRow: 1, voucher: '记-大', date: { y: 2026, m: 1, d: 1 },
              dateText: '2026-01-01', summary: '', account: '', debitCents: 100000000, creditCents: null });
for (let i = 2; i <= 100; i++) {
  ppsPop.push({ signedCents: 1000000, srcRow: i, voucher: '记-' + i, date: { y: 2026, m: 1, d: (i % 28) + 1 },
                dateText: '', summary: '', account: '', debitCents: 1000000, creditCents: null });
}

const p1 = Core.samplePPS(ppsPop, 10, 20260101);
eq(Math.round(p1.interval), 19900000, '间隔 = 总体金额 ÷ 样本量（199 万 ÷ 10）');
ok(p1.indices.indexOf(0) >= 0, '100 万那笔必须被抽中（它比间隔大 5 倍）');
eq(p1.hits[0], 5, '它横跨 5 个间隔，被命中 5 次');
ok(p1.indices.length < 10, '大额占多个间隔，实抽笔数会少于目标笔数');

const p2 = Core.samplePPS(ppsPop, 10, 20260101);
eq(p1.indices.join(','), p2.indices.join(','), '同种子结果完全一致（可复核）');
const p3 = Core.samplePPS(ppsPop, 10, 999);
ok(p1.start !== p3.start, '换种子起点会变');

const ppsSample = Core.buildSample(ppsPop, { n: 10, method: 'pps', seed: 20260101 });
eq(ppsSample.meta.method, 'pps', '方法标记为 pps');
ok(/金额加权/.test(ppsSample.meta.methodLabel), '方法名要写得让人看懂');
eq(ppsSample.meta.fullCount, 0, '金额加权不另外分层');
eq(ppsSample.sample[0].layer, '金额加权', '样本层名');
ok(/少于目标/.test(ppsSample.meta.note), '实抽少于目标时必须解释，不能让人以为漏抽');
eq(ppsSample.sample.length, ppsSample.meta.sampleCount, '样本与元数据一致');

// 对照：等距抽样只看行序、完全不看金额 —— 把金额全改成一样，下标必须一模一样
const flatPop = ppsPop.map(r => Object.assign({}, r, { signedCents: 1000000 }));
eq(Core.sampleSystematic(flatPop, 10, 20260101).indices.join(','),
   Core.sampleSystematic(ppsPop, 10, 20260101).indices.join(','),
   '等距抽样与金额无关：金额改了样本不变（这就是它不能当 MUS 样本的原因）');


// 分层：阈值以上全查，仍然好用
const layered = Core.buildSample(ppsPop, { n: 10, method: 'systematic', thresholdCents: 5000000, seed: 1 });
eq(layered.meta.fullCount, 1, '超过阈值的 1 笔全查');
eq(layered.sample.filter(x => x.layer === '全查').length, 1, '全查层有且只有 1 笔');

/* ---------------- 11 独立审核发现的问题：逐条守住 ---------------- */
section('11 审核发现：静默吞掉例外、凭证撞号、方向丢失');

// 11.1 「不相符」不能被「相符」抢走（原来判成 ok，例外从评估里消失）
eq(Core.classifyConclusion('不相符'), 'exception', '「不相符」必须判成例外');
eq(Core.classifyConclusion('未相符'), 'exception', '「未相符」必须判成例外');
eq(Core.classifyConclusion('不相符合'), 'exception', '「不相符合」必须判成例外');
eq(Core.classifyConclusion('未见异常'), 'ok', '「未见异常」是相符，不能被「异常」抢走');
eq(Core.classifyConclusion('无异常'), 'ok', '「无异常」是相符');
eq(Core.classifyConclusion('否'), 'exception', '「否」是例外');

// 11.2 按凭证抽样：不同月份的「记-1」不能合成一张
const crossMonth = [];
[['2026-01-05', '记-1', 1000], ['2026-01-05', '记-1', 1000],
 ['2026-02-06', '记-1', 2000], ['2026-02-06', '记-1', 2000],
 ['2026-02-07', '记-2', 500]].forEach((r, i) => {
  const parts = r[0].split('-').map(Number);
  crossMonth.push({
    srcRow: i + 2, date: { y: parts[0], m: parts[1], d: parts[2] }, dateText: r[0],
    voucher: r[1], summary: '', account: '应收账款', accountCode: '',
    debitCents: r[2], creditCents: null, signedCents: r[2]
  });
});
const cmPop = Core.buildPopulation(crossMonth, { unit: 'voucher' });
eq(cmPop.stats.count, 3, '跨月同号凭证应算 3 张，不是 2 张');
eq(cmPop.included.filter(x => x.lineCount === 2).length, 2, '两张「记-1」各含 2 行');

// 11.3 方向列写 DR/CR 时不能一律当借方
const dirRows = [['日期', '凭证号', '科目', '金额', '方向'],
  ['2026-01-05', '记-1', '应收', '100', 'CR'],
  ['2026-01-06', '记-2', '应收', '200', 'DR'],
  ['2026-01-07', '记-3', '应收', '300', '借方'],
  ['2026-01-08', '记-4', '应收', '400', '带']];
const dirParsed = Core.parseRows(dirRows, Core.detectHeader(dirRows), { firstExcelRow: 1 });
eq(dirParsed.records[0].creditCents, 10000, 'CR 应进贷方');
eq(dirParsed.records[0].debitCents, null, 'CR 不应进借方');
eq(dirParsed.records[1].debitCents, 20000, 'DR 应进借方');
eq(dirParsed.records[2].debitCents, 30000, '「借方」应进借方');
eq(dirParsed.problems.length, 1, '认不出的方向要报出来，不能默认借方');
ok(/方向/.test(dirParsed.problems[0].reason), '原因应指明是方向读不出，实际：' + dirParsed.problems[0].reason);

// 11.4 错报金额写错不能悄悄当 0
const vtRows = [['序号', '原始行号', '凭证号', '科目', '金额（绝对值）', '核对结论', '错报金额（元）'],
  ['1', '5', '记-1', '应收', '1000', '不符', '约500元'],
  ['2', '6', '记-2', '应收', '1000', '相符', '']];
const vt = Core.parseVouchingRows(vtRows);
eq(vt.records[0].misCents, null, '读不出的错报金额应为 null，不是 0');
eq(vt.records[0].misUnreadable, true, '应标记这一行读不出来');
eq(vt.records[1].misCents, 0, '空着的错报金额才是 0');
ok(vt.warnings.some(w => /读不出来/.test(w)), '应有警告说明哪一行读不出来');

// 11.5 多表合并：行号必须带来源，否则两张表的「第 5 行」分不清
const sheetA = [{ srcRow: 5, sheet: '一月', date: { y: 2026, m: 1, d: 5 }, dateText: '2026-01-05',
  voucher: '记-1', summary: '', account: '应收', accountCode: '', debitCents: 10000, creditCents: null, signedCents: 10000 }];
const sheetB = [{ srcRow: 5, sheet: '二月', date: { y: 2026, m: 2, d: 5 }, dateText: '2026-02-05',
  voucher: '记-1', summary: '', account: '应收', accountCode: '', debitCents: 20000, creditCents: null, signedCents: 20000 }];
const multiSrcPop = Core.buildPopulation(sheetA.concat(sheetB), {});
const multiSrcSmp = Core.buildSample(multiSrcPop.included, { n: 10, thresholdCents: 0, seed: 1 });
const multiSrcBox = {
  client: '测试', dateText: '2026-11-21', periodText: '全部', accountText: '', presetName: '销售/应收',
  sample: multiSrcSmp.sample, meta: Object.assign({}, multiSrcSmp.meta, { unit: 'line', multiSheet: true }),
  sheetNames: ['一月', '二月'], populationStats: multiSrcPop.stats, excluded: multiSrcPop.excluded,
  problems: [], skipped: [], warnings: { parsedRows: 2 }, mapNotes: [],
  filter: { accountText: '应收', entity: '', direction: 'debit', minCents: 100, maxCents: null },
  headerRow: 1
};
const multiSrcSheets = Core.buildSheets(multiSrcBox);
const multiSrcHead = multiSrcSheets[0].rows.filter(r => r[0] === '序号')[0];
const multiSrcRow1 = multiSrcSheets[0].rows.filter(r => r[0] === 1)[0];
ok(/一月/.test(String(multiSrcRow1[1])) && /第5行/.test(String(multiSrcRow1[1])),
   '合并读取时「原始行号」应写成「表名 第N行」，实际：' + multiSrcRow1[1]);
eq(multiSrcHead[1], '原始行号', '列名不变，仍是原始行号');

// 抽样信息必须把筛选条件写全，否则别人拿不到同样的总体边界
const mInfo = multiSrcSheets[1].rows.map(r => r.join('|')).join('\n');
ok(/筛选条件·科目\|应收/.test(mInfo), '抽样信息应记录科目条件');
ok(/筛选条件·方向\|只看借方/.test(mInfo), '抽样信息应记录方向条件');
ok(/筛选条件·金额下限\|1\.00/.test(mInfo), '抽样信息应记录金额下限');
ok(/筛选条件·期间/.test(mInfo), '抽样信息应记录期间条件');

// 11.6 汇总行的各种写法都要认出来（否则会虚增总体）
const subRows = [['日期', '凭证号', '科目', '借方'],
  ['2026-01-05', '记-1', '应收', '100'],
  ['本期发生额合计', '', '', '100'],
  ['合计（本期）', '', '', '100'],
  ['本期小计', '', '', '100'],
  ['应收合计', '', '', '100'],
  ['2026-01-06', '记-2', '应收', '50']];
const subParsed = Core.parseRows(subRows, Core.detectHeader(subRows), { firstExcelRow: 1 });
eq(subParsed.records.length, 2, '4 种汇总行写法都要跳过，只剩 2 笔明细');

// 11.7 三位小数：数值与文本走同一套四舍五入
eq(Core.toCents(1.005).cents, Core.toCents('1.005').cents, '数值与文本 1.005 应一致');
eq(Core.toCents(1.005).cents, 101, '1.005 应进位成 1.01');
eq(Core.toCents(1234.567).extra, true, '数值超两位小数也要进数据检查清单');

/* ---------------- 汇总 ---------------- */
console.log('\n' + '='.repeat(56));
if (fails.length) {
  console.log('失败 ' + fails.length + ' / ' + count + ' 项：');
  fails.forEach((f, i) => console.log('  ' + (i + 1) + '. ' + f));
  process.exit(1);
}
console.log('全部通过：' + count + ' 项断言');
console.log('='.repeat(56));
