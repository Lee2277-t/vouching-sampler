'use strict';
/* 接线检查：界面层与核心层之间的缝，静态验一遍。
 *
 * 浏览器里点不了的地方，至少守住这三条：
 *   1. app.js 取的每个 id，index.html 里都得有（拼错就是整块没反应）
 *   2. app.js 调的每个 Core.xxx，core.js 都得导出
 *   3. 页面上不许出现任何外链 —— 这是「数据不出本机」的硬保证
 */

const fs = require('fs');
const path = require('path');
const Core = require('../core.js');

const ROOT = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
const app = fs.readFileSync(path.join(ROOT, 'app.js'), 'utf8');
const css = fs.readFileSync(path.join(ROOT, 'app.css'), 'utf8');

let count = 0;
const fails = [];
function ok(cond, msg) { count++; if (!cond) fails.push(msg); }
function section(t) { console.log('\n[' + t + ']'); }

/* ---------- 1 id 对齐 ---------- */
section('1 界面 id：app.js 取的每个 id，页面里都得有');

const htmlIds = new Set();
let m;
const idRe = /\bid="([^"]+)"/g;
while ((m = idRe.exec(html))) htmlIds.add(m[1]);

const usedIds = new Set();
const useRe = /\$\('([^']+)'\)/g;
while ((m = useRe.exec(app))) usedIds.add(m[1]);
const gidRe = /getElementById\('([^']+)'\)/g;
while ((m = gidRe.exec(app))) usedIds.add(m[1]);

ok(usedIds.size > 20, '应当从 app.js 里抓到足够多的 id 引用，实际 ' + usedIds.size);

const missing = [...usedIds].filter(id => !htmlIds.has(id));
ok(missing.length === 0, '页面缺这些 id：' + missing.join(', '));

// 反向：页面上的 id 如果没人用，多半是改名后的残骸（class 除外，不算错）
const orphan = [...htmlIds].filter(id => !usedIds.has(id));
if (orphan.length) console.log('  提示：这些 id 没被 app.js 引用 —— ' + orphan.join(', '));

/* ---------- 2 Core 方法对齐 ---------- */
section('2 核心方法：app.js 调的每个 Core.xxx 都得导出');

const apiKeys = new Set(Object.keys(Core));
const coreCalls = new Set();
const callRe = /\bCore\.([A-Za-z_$][\w$]*)/g;
while ((m = callRe.exec(app))) coreCalls.add(m[1]);

ok(coreCalls.size > 5, '应当从 app.js 里抓到足够多的 Core 调用，实际 ' + coreCalls.size);
const badCalls = [...coreCalls].filter(k => !apiKeys.has(k));
ok(badCalls.length === 0, 'core.js 没有导出：' + badCalls.join(', '));

/* ---------- 3 无外链 ---------- */
section('3 数据不出本机：页面不得引用任何外部资源');

const externalPatterns = [
  [/<script[^>]+src\s*=\s*["']https?:/i, '外链脚本'],
  [/<link[^>]+href\s*=\s*["']https?:/i, '外链样式'],
  [/@import\s+url\(\s*["']?https?:/i, 'CSS @import 外链'],
  [/url\(\s*["']?https?:\/\//i, 'CSS 里的外链资源'],
  [/fonts\.googleapis|fonts\.gstatic|cdn\.jsdelivr|unpkg\.com|cdnjs/i, '常见 CDN 域名']
];
externalPatterns.forEach(([re, name]) => {
  ok(!re.test(html) && !re.test(css) && !re.test(app), '不应出现' + name);
});

// 唯一允许的外链是本地 vendored 的库
ok(/<script src="vendor\/xlsx\.full\.min\.js"><\/script>/.test(html), '应本地引用 vendor 里的 SheetJS');
ok(fs.existsSync(path.join(ROOT, 'vendor', 'xlsx.full.min.js')), 'vendor 库文件应存在');

// 运行时代码里不应有 fetch / XMLHttpRequest
ok(!/\bfetch\s*\(/.test(app), 'app.js 不应调用 fetch');
ok(!/XMLHttpRequest/.test(app), 'app.js 不应使用 XMLHttpRequest');

/* ---------- 4 关键控件的接线 ---------- */
section('4 关键控件：三步流程都接上了事件');

['fileInput', 'sheetSelect', 'headerRowInput', 'redetectBtn', 'mapTable',
 'dateFrom', 'dateTo', 'entitySelect', 'unitSelect', 'thresholdYuan', 'sampleN', 'seedInput',
 'genBtn', 'evalFile', 'alphaSelect', 'tolerableYuan', 'popYuan'].forEach(id => {
  ok(htmlIds.has(id), '页面应有 #' + id);
  ok(app.indexOf("'" + id + "'") >= 0, 'app.js 应引用 #' + id);
});

// v0.4 新能力必须在界面上真的接上，不能只躺在 core 里
ok(/id="sheetSelect"[^>]*multiple/.test(html), '工作表应可多选');
ok(app.indexOf('loadSelected') >= 0, 'app.js 应走多表读取入口 loadSelected');
ok(app.indexOf('state.sheets') >= 0, '应使用多表状态 state.sheets');
ok(/entity:\s*\$\('entitySelect'\)\.value/.test(app), '筛选条件应带上主体');
ok(/unit:\s*\$\('unitSelect'\)\.value/.test(app), '筛选条件应带上抽样单元');
ok(app.indexOf('renderEntitySelect') >= 0, '应按数据填充主体下拉');
ok(/populationCount:/.test(app), '评估应把总体笔数传给 core（差异法要用）');
ok(html.indexOf('id="unitSelect"') >= 0 && html.indexOf('value="voucher"') >= 0, '抽样单元应有「凭证」选项');
ok(/id="fileInput"[^>]*multiple/.test(html), '文件输入应可多选（跨文件合并）');
ok(app.indexOf('function readWorkbook') >= 0, 'app.js 应有逐文件读取的 readWorkbook');
ok(app.indexOf('state.books') >= 0, '应使用多文件状态 state.books');
ok(app.indexOf('function analyzeBooks') >= 0, '应对所有文件的表一起打分');
ok(app.indexOf('maybeTranspose') >= 0, '应能把竖排表转置后读');
ok(/orientSelect/.test(html) && /orientSelect/.test(app), '应有「表的走向」选择并接线');
ok(app.indexOf('learnAlias') >= 0 && app.indexOf('saveLearned') >= 0, '手工改的列名应被记住并保存');
ok(/localStorage/.test(app) && /chouping\.learned/.test(app), '记住的列名应存在本机 localStorage');
ok(/try\s*\{[\s\S]{0,200}localStorage/.test(app), 'localStorage 存取必须包 try（file:// 下可能被禁）');
ok(htmlIds.has('evalExportBtn') && app.indexOf("'evalExportBtn'") >= 0, '评估区应有导出按钮并接线');
ok(app.indexOf('buildEvaluationWorkbook') >= 0, '导出评估表应走 core.buildEvaluationWorkbook');
ok(html.indexOf('id="verText"') >= 0 && app.indexOf('verText') >= 0, '页脚版本号应来自 Core.VERSION，不能写死');
// 简易/高级： newcomer 只要 4 项能跑通
['modeSimple', 'modeExpert', 'demoBtn', 'simpleSummary'].forEach(id => {
  ok(htmlIds.has(id), '页面应有 #' + id);
  ok(app.indexOf("'" + id + "'") >= 0, 'app.js 应引用 #' + id);
});
ok(/body\.mode-simple\s+\.expert-only/.test(css), '简易模式必须靠 CSS 把高级项藏掉');
ok((html.match(/expert-only/g) || []).length >= 6, '应有足够多的高级项标了 expert-only');
ok(/setMode\(true\)/.test(app) && /mode-simple/.test(app), '应能在简易与高级之间切换');
// 「抽大的」这件事必须在界面上真的可选
ok(/<option value="pps">/.test(html), '「怎么抽」应有金额加权选项');
ok(/<option value="systematic">/.test(html), '默认应是「大额全查 + 其余均匀抽」');
ok(htmlIds.has('thresholdField'), '阈值那一格要能单独藏起来（金额加权时不需要它）');
ok(app.indexOf('suggestThresholdCents') >= 0, '阈值应按总体自动给建议值');
ok(app.indexOf('thresholdTouched') >= 0, '用户手工改过阈值后不许再被自动覆盖');
// 手工修了列映射之后，可用性必须用新映射重算，否则程序自己给的出路走不通
ok(/sheetVerdict\(\{\s*map:\s*s0\.map/.test(app), '改完列识别后应重算可用性');
ok(app.indexOf('misUnreadable') >= 0 || app.indexOf('badNumberFields') >= 0,
   '读不出的金额/结论必须报出来，不能静默当 0');
ok(/umBlocked/.test(app), '算不了上限时要说明原因，不能只显示空白');
ok(app.indexOf('sampleMethod') >= 0, '评估要带上抽样方法，等距样本不能套 MUS 上限');

const genIdx = app.indexOf("'genBtn'");
const buildIdx = app.indexOf('buildWorkbook');
ok(genIdx >= 0 && buildIdx > genIdx, '生成按钮的处理函数应在 genBtn 之后');
ok(buildIdx - genIdx < 4000, '生成按钮里应直接调用 buildWorkbook（距离 ' + (buildIdx - genIdx) + ' 字符）');
ok(app.indexOf("evalFile')") >= 0 && /evalFile'\)\.addEventListener\('change'/.test(app), '评估区应监听文件选择');

/* ---------- 汇总 ---------- */
console.log('\n' + '='.repeat(56));
if (fails.length) {
  console.log('失败 ' + fails.length + ' / ' + count + ' 项：');
  fails.forEach((f, i) => console.log('  ' + (i + 1) + '. ' + f));
  process.exit(1);
}
console.log('全部通过：' + count + ' 项断言');
console.log('='.repeat(56));
