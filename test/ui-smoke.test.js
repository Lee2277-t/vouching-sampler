'use strict';
/* app.js 冒烟测试：用假 DOM 真跑一遍界面层
 *
 * 静态接线检查只能证明 id 和 Core 方法对得上，证明不了「刷新时有个字段名写错了」。
 * 这里把 app.js 原样 require 进来，喂一份真实格式的 xlsx，走完
 *   读文件 → 认表 → 筛总体 → 抽样 → 生成工作簿
 * 任何一步抛异常都算失败。
 *
 * 假 DOM 只实现 app.js 真正用到的那点东西：getElementById、innerHTML 里的 option、
 * addEventListener / dispatch、FileReader。
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');

let count = 0;
const fails = [];
function ok(cond, msg) { count++; if (!cond) fails.push(msg); }
function eq(a, b, msg) {
  count++;
  if (a !== b) fails.push(msg + ' — 期望 ' + JSON.stringify(b) + '，实际 ' + JSON.stringify(a));
}
function section(t) { console.log('\n[' + t + ']'); }

/* ---------------- 假 DOM ---------------- */

function parseOptions(markup) {
  const out = [];
  const re = /<option\b([^>]*)>([\s\S]*?)<\/option>/gi;
  let m2;
  while ((m2 = re.exec(markup))) {
    const attrs = m2[1];
    const vm = /value="([^"]*)"/i.exec(attrs);
    out.push({
      value: vm ? vm[1] : m2[2].trim(),
      text: m2[2].replace(/<[^>]+>/g, '').trim(),
      selected: /\bselected\b/i.test(attrs)
    });
  }
  return out;
}

const els = new Map();

/** 有状态的 classList，断言才看得见 */
function makeClassList(owner) {
  owner._cls = {};
  return {
    add: function (c) { owner._cls[c] = 1; },
    remove: function (c) { delete owner._cls[c]; },
    contains: function (c) { return !!owner._cls[c]; },
    toggle: function (c, force) {
      var on = (force === undefined) ? !owner._cls[c] : !!force;
      if (on) owner._cls[c] = 1; else delete owner._cls[c];
      return on;
    }
  };
}

function makeEl(id) {
  const el = {
    id: id,
    _html: '',
    _listeners: {},
    value: '',
    textContent: '',
    disabled: false,
    checked: false,
    files: null,
    options: [],
    className: '',
    addEventListener: function (type, fn) {
      (this._listeners[type] = this._listeners[type] || []).push(fn);
    },
    removeEventListener: function () {},
    dispatch: function (type, ev) {
      (this._listeners[type] || []).forEach(function (fn) { fn.call(this, ev || { target: this }); }, this);
    },
    querySelectorAll: function () { return []; },
    appendChild: function () {},
    focus: function () {},
    getAttribute: function () { return null; }
  };
  el.classList = makeClassList(el);

  Object.defineProperty(el, 'innerHTML', {
    get: function () { return el._html; },
    set: function (v) {
      el._html = String(v);
      if (el._html.indexOf('<option') >= 0) {
        el.options = parseOptions(el._html);
        // 真浏览器里 innerHTML 写进去后，第一个 option 会自动被选中
        const first = el.options.filter(function (o) { return o.selected; })[0] || el.options[0];
        if (first) {
          el.value = first.value;
          el.options.forEach(function (o) { o.selected = (o === first); });
        }
      }
    }
  });

  return el;
}

global.document = {
  getElementById: function (id) {
    if (!els.has(id)) els.set(id, makeEl(id));
    return els.get(id);
  },
  querySelectorAll: function () { return []; },
  createElement: function (tag) { return makeEl('<' + tag + '>'); },
  body: (function () { var b = { appendChild: function () {} }; b.classList = makeClassList(b); return b; })()
};

global.window = global;

/* 从 index.html 里带出静态初值：<input value="..."> 和 <select> 的第一个选项 */
(function seed() {
  const selRe = /<select\b[^>]*\bid="([^"]+)"[^>]*>([\s\S]*?)<\/select>/gi;
  let m3;
  while ((m3 = selRe.exec(html))) {
    const opts = parseOptions(m3[2]);
    if (opts.length) {
      const el = global.document.getElementById(m3[1]);
      const chosen = opts.filter(function (o) { return o.selected; })[0] || opts[0];
      el.value = chosen.value;
      el.options = opts;
    }
  }
  const inpRe = /<input\b([^>]*)>/gi;
  while ((m3 = inpRe.exec(html))) {
    const attrs = m3[1];
    const idm = /\bid="([^"]+)"/i.exec(attrs);
    const vm = /\bvalue="([^"]*)"/i.exec(attrs);
    if (idm && vm) global.document.getElementById(idm[1]).value = vm[1];
  }
})();

/* ---------------- 假 FileReader ---------------- */

global.FileReader = function () {
  const self = this;
  this.result = null;
  this.onload = null;
  this.readAsArrayBuffer = function (file) {
    setTimeout(function () {
      // app.js 读的是 reader.result（不是事件里的 target.result），别搞错
      self.result = file.__buffer;
      self.onload({ target: self });
    }, 0);
  };
};

/* ---------------- 加载被测对象 ---------------- */

const XLSX = require('../vendor/xlsx.full.min.js');
const Core = require('../core.js');

let captured = null;
XLSX.writeFile = function (wb, name) { captured = { wb: wb, name: name }; };

global.XLSX = XLSX;
global.AuditCore = Core;

require('../app.js');

/* ---------------- 用例 ---------------- */

function tick() { return new Promise(function (r) { setTimeout(r, 5); }); }
/** 输入类控件有 180ms 防抖，等它跑完 */
function settle() { return new Promise(function (r) { setTimeout(r, 260); }); }

const fixture = path.join(__dirname, 'fixture', 'sample_ledger.xlsx');

/** 下拉的 value 是「文件序号 + \u0001 + 表名」，按表名挑 */
function optionsOfSheet(name) {
  const sel = global.document.getElementById('sheetSelect');
  return sel.options.filter(function (o) {
    return o.value.split('\u0001')[1] === name;
  });
}
function selectSheets(names) {
  const sel = global.document.getElementById('sheetSelect');
  sel.options.forEach(function (o) {
    o.selected = names.indexOf(o.value.split('\u0001')[1]) >= 0;
  });
  sel.dispatch('change');
}
function countInStats() {
  const m = global.document.getElementById('statsBox').innerHTML
    .match(/总体(?:笔数|凭证数)[\s\S]*?<div class="v">([\d,]+)/);
  return m ? parseInt(m[1].replace(/,/g, ''), 10) : null;
}
function statsHtml() {
  return global.document.getElementById('statsBox').innerHTML;
}

(async function () {
  section('1 读文件：整本书进来，默认选中可用的那张表');
  global.document.getElementById('fileInput').dispatch('change', {
    target: { files: [{ name: 'sample_ledger.xlsx', __buffer: fs.readFileSync(fixture) }] }
  });
  await tick();

  const sheetSel = global.document.getElementById('sheetSelect');
  ok(sheetSel.options.length >= 6, '应列出工作簿里的全部工作表，实际 ' + sheetSel.options.length);

  const stats = global.document.getElementById('statsBox');
  ok(/总体/.test(stats.innerHTML), '刷新后应显示总体统计');
  ok(/总体笔数|总体凭证数/.test(stats.innerHTML), '应显示总体笔数或凭证数');

  section('2 切到金蝶式序时账：多工作表各认各的表头');
  ok(optionsOfSheet('序时账').length === 1, '下拉里应有「序时账」');
  selectSheets(['序时账']);
  await tick();

  ok(/总体笔数/.test(stats.innerHTML), '单表模式显示「总体笔数」');
  eq(countInStats(), 12, '序时账 12 行');

  section('3 多选：序时账 + 填制凭证合并读取');
  selectSheets(['序时账', '填制凭证']);
  await tick();
  ok(/总体笔数/.test(stats.innerHTML), '多表合并后仍是明细行模式');
  eq(countInStats(), 16, '12 + 4 = 16 行');

  section('4 抽样单元切到凭证：统计口径跟着变');
  const unitSel = global.document.getElementById('unitSelect');
  unitSel.value = 'voucher';
  unitSel.dispatch('change');
  await settle();
  ok(/总体凭证数/.test(stats.innerHTML), '凭证模式下标题变成「总体凭证数」');
  ok(/张/.test(stats.innerHTML), '凭证模式单位显示「张」');

  section('5 生成：走 core.buildWorkbook 出真工作簿');
  global.document.getElementById('genBtn').dispatch('click');
  await tick();

  ok(!!captured, '生成按钮应产出工作簿');
  if (captured) {
    const names = captured.wb.SheetNames.join(',');
    eq(names, '抽凭表,抽样信息,排除清单,数据检查,说明', '五页齐全');
    const rows = XLSX.utils.sheet_to_json(captured.wb.Sheets['抽样信息'], { header: 1, defval: null });
    const flat = rows.map(function (r) { return r.join('|'); }).join('\n');
    ok(/抽样单元/.test(flat), '抽样信息应写明抽样单元');
    ok(/多个工作表合并/.test(flat), '多表读取时抽样信息应写明数据来源');
    ok(/凭证/.test(flat), '凭证模式应写进说明');
  }

  const result = global.document.getElementById('resultBox');
  ok(/已生成/.test(result.innerHTML), '页面应提示已生成');

  section('6 跨文件合并：两份文件各一张序时账');
  const buf = fs.readFileSync(fixture);
  global.document.getElementById('fileInput').dispatch('change', {
    target: { files: [
      { name: '1月序时账.xlsx', __buffer: buf },
      { name: '2月序时账.xlsx', __buffer: buf }
    ] }
  });
  await tick();

  const fileInfo = global.document.getElementById('fileInfo').innerHTML;
  ok(/2 个文件/.test(fileInfo), '文件信息应说明这是两个文件，实际：' + fileInfo);
  eq(optionsOfSheet('序时账').length, 2, '两张序时账（每个文件一张）');

  unitSel.value = 'line';
  unitSel.dispatch('change');
  await settle();

  selectSheets(['序时账']);
  await tick();
  eq(countInStats(), 24, '两边的 12 行应合并成 24 行');

  captured = null;
  global.document.getElementById('genBtn').dispatch('click');
  await tick();
  ok(!!captured, '跨文件也应能生成');
  if (captured) {
    const rows = XLSX.utils.sheet_to_json(captured.wb.Sheets['抽样信息'], { header: 1, defval: null });
    const flat = rows.map(function (r) { return r.join('|'); }).join('\n');
    ok(/1月序时账\.xlsx/.test(flat), '抽样信息应写明来自哪个文件');
    ok(/2月序时账\.xlsx/.test(flat), '两个文件都应列出');
  }

  section('7 空文件/无表：不许抛异常');
  global.document.getElementById('fileInput').dispatch('change', {
    target: { files: [{ name: 'empty.csv', __buffer: Buffer.from('a,b\n1,2\n', 'utf8') }] }
  });
  await tick();
  // 不能只看「没抛异常」：要确认它给出的结论是有意义的（总体为 0，且拦住了生成）
  eq(countInStats(), 0, '只有表头没有数据时，总体应为 0 笔');
  captured = null;
  global.document.getElementById('genBtn').dispatch('click');
  await tick();
  ok(!captured, '总体为 0 时不许生成抽凭表');
  ok(/总体是空的|不能当抽凭总体/.test(global.document.getElementById('resultBox').innerHTML),
     '应给出人看得懂的原因，实际：' + global.document.getElementById('resultBox').innerHTML.slice(0, 80));

  section('8 竖排表：自动转置，并把「行号」改成「列号」');
  global.document.getElementById('fileInput').dispatch('change', {
    target: { files: [{ name: '竖排.xlsx', __buffer: fs.readFileSync(fixture) }] }
  });
  await tick();

  const vOpt = global.document.getElementById('sheetSelect').options
    .filter(function (o) { return o.value.split('\u0001')[1] === '竖排凭证'; })[0];
  ok(!!vOpt, '下拉里应有「竖排凭证」');
  ok(/竖排已转置/.test(vOpt.text), '下拉里应标出这张表是竖排转置来的，实际：' + (vOpt ? vOpt.text : ''));

  selectSheets(['竖排凭证']);
  await tick();
  eq(countInStats(), 3, '转置后应识别出 3 笔');

  captured = null;
  global.document.getElementById('genBtn').dispatch('click');
  await tick();
  ok(!!captured, '竖排表也应能生成');
  if (captured) {
    const rows = XLSX.utils.sheet_to_json(captured.wb.Sheets['抽凭表'], { header: 1, defval: null });
    const headRow = rows.filter(function (r) { return r[0] === '序号'; })[0];
    ok(!!headRow, '应有表头行');
    eq(headRow[1], '原始列号', '竖排表不应把列位置说成「原始行号」');
    const infoFlat = XLSX.utils.sheet_to_json(captured.wb.Sheets['说明'], { header: 1, defval: null })
      .map(function (r) { return r.join(''); }).join('\n');
    ok(/竖排/.test(infoFlat), '说明页应写明这张表是转置读的');
  }
  section('9 第 4 步闭环：填结论 → 回录 → 导出评估表');
  ok(!!captured && captured.wb.Sheets['抽凭表'], '前面应已生成一份抽凭表');

  const outWb = captured.wb;
  const mainRows = XLSX.utils.sheet_to_json(outWb.Sheets['抽凭表'], {
    header: 1, defval: null, blankrows: true
  });
  const hIdx = mainRows.findIndex(function (r) { return r[0] === '序号'; });
  ok(hIdx >= 0, '抽凭表里应有表头行');
  const hdr = mainRows[hIdx];
  const cConc = hdr.indexOf('核对结论');
  const cMis = hdr.indexOf('错报金额（元）');
  ok(cConc >= 0, '抽凭表应有「核对结论」列');
  ok(cMis >= 0, '抽凭表应有「错报金额（元）」列');

  // 模仿现场：第一笔相符、第二笔不符且账面多计 12.34、第三笔未取得
  mainRows[hIdx + 1][cConc] = '相符';
  mainRows[hIdx + 2][cConc] = '不符';
  mainRows[hIdx + 2][cMis] = 12.34;
  mainRows[hIdx + 3][cConc] = '未取得';
  outWb.Sheets['抽凭表'] = XLSX.utils.aoa_to_sheet(mainRows);

  const evalBuf = XLSX.write(outWb, { type: 'buffer', bookType: 'xlsx' });
  global.document.getElementById('evalFile').dispatch('change', {
    target: { files: [{ name: '填好的抽凭表.xlsx', __buffer: evalBuf }] }
  });
  await tick();

  const evalInfo = global.document.getElementById('evalInfo').innerHTML;
  ok(/读到 3 笔/.test(evalInfo), '应读到 3 笔回录，实际：' + evalInfo);
  ok(global.document.getElementById('popYuan').value !== '', '总体金额应从「抽样信息」页自动读回');

  const evalHtml = global.document.getElementById('evalResult').innerHTML;
  ok(/抽查结果/.test(evalHtml), '应显示抽查结果');
  ok(/不相符|不符/.test(evalHtml), '应统计到「不符」');
  ok(/错报上限|算不出上限/.test(evalHtml), '应给出错报上限或说明为何算不出');

  captured = null;
  global.document.getElementById('evalExportBtn').dispatch('click');
  await tick();

  ok(!!captured, '点导出应产出评估表');
  if (captured) {
    eq(captured.wb.SheetNames.join(','), '评估表', '评估表只有一页');
    const evRows = XLSX.utils.sheet_to_json(captured.wb.Sheets['评估表'], { header: 1, defval: null });
    const flat = evRows.map(function (r) { return r.join('|'); }).join('\n');
    ok(/不符/.test(flat) && /未取得/.test(flat), '评估表应含各项计数');
    ok(/错报上限|上界系数/.test(flat), '评估表应含错报上限');
    ok(/推断错报/.test(flat), '评估表应含推断错报');
    ok(/不代替判断/.test(flat), '评估表应写明不代替判断');
  }

  section('10 页脚版本号来自 Core.VERSION');
  eq(global.document.getElementById('verText').textContent, 'v' + Core.VERSION,
     '页脚版本不应写死（以前写死成 v0.1 一直没人发现）');

  section('11 简易模式：默认只留四项，一句话说清楚');
  const body = global.document.body;
  ok(body.classList.contains('mode-simple'), '默认应是简易模式');

  global.document.getElementById('demoBtn').dispatch('click');
  await tick();

  const sum = global.document.getElementById('simpleSummary');
  ok(!sum.classList.contains('hidden'), '简易模式应显示一句话摘要');
  ok(/已认出/.test(sum.innerHTML), '摘要应说明认出多少笔，实际：' + sum.innerHTML);
  ok(/示例数据/.test(sum.innerHTML), '摘要应说明数据来自哪个文件');
  eq(countInStats(), 48, '示例 60 笔里 48 笔是应收账款');

  // 简易模式下照样能生成
  captured = null;
  global.document.getElementById('genBtn').dispatch('click');
  await tick();
  ok(!!captured, '简易模式下应能生成抽凭表');

  section('12 切到高级：参数全露，简易摘要收起');
  global.document.getElementById('modeExpert').dispatch('click');
  await tick();
  ok(!body.classList.contains('mode-simple'), '点高级后应退出简易模式');
  ok(sum.classList.contains('hidden'), '高级模式下摘要应收起');

  global.document.getElementById('modeSimple').dispatch('click');
  await tick();
  ok(body.classList.contains('mode-simple'), '还能切回简易');

  section('13 大额全查：阈值按总体自动建议，可改');
  const thr = global.document.getElementById('thresholdYuan');
  ok(thr.value !== '' && thr.value !== '0', '导入数据后应自动建议一个阈值，实际：' + thr.value);
  const suggested = parseFloat(thr.value);

  // 手工改过就不许再被覆盖
  thr.value = '12345';
  thr.dispatch('input');
  await settle();
  eq(thr.value, '12345', '手工改过的阈值不能被自动建议覆盖');
  ok(Math.abs(suggested - 12345) > 0.01, '建议值与手工值不同，说明确实建议过');

  section('14 按金额加权：自动藏掉阈值，统计口径跟着变');
  const method = global.document.getElementById('methodSelect');
  method.value = 'pps';
  method.dispatch('change');
  await settle();
  ok(global.document.getElementById('thresholdField').classList.contains('hidden'),
     '金额加权时不该再显示阈值');
  ok(/金额间隔/.test(statsHtml()), '统计区应显示金额间隔，实际：' + statsHtml().slice(0, 120));

  captured = null;
  global.document.getElementById('genBtn').dispatch('click');
  await tick();
  ok(!!captured, '金额加权也应能生成');
  if (captured) {
    const infoFlat = XLSX.utils.sheet_to_json(captured.wb.Sheets['抽样信息'], { header: 1, defval: null })
      .map(function (r) { return r.join('|'); }).join('\n');
    ok(/金额加权/.test(infoFlat), '抽样信息应写明金额加权');
    ok(/金额间隔/.test(infoFlat), '抽样信息应写出金额间隔（复核要用）');
  }

  section('15 阈值写了「10万」这种读不出的值：拦住，不许当没填');
  global.document.getElementById('modeExpert').dispatch('click');
  await tick();
  const thrEl = global.document.getElementById('thresholdYuan');
  thrEl.value = '10万';
  thrEl.dispatch('input');
  await settle();
  captured = null;
  global.document.getElementById('genBtn').dispatch('click');
  await tick();
  ok(!captured, '金额读不出来时不许生成');
  ok(/不是能识别的金额/.test(global.document.getElementById('resultBox').innerHTML),
     '应说清楚哪一项读不出来，实际：' + global.document.getElementById('resultBox').innerHTML.slice(0, 90));
  ok(/10万/.test(global.document.getElementById('resultBox').innerHTML), '应把原值回显出来');

  thrEl.value = '5000';
  thrEl.dispatch('input');
  await settle();
  captured = null;
  global.document.getElementById('genBtn').dispatch('click');
  await tick();
  ok(!!captured, '改成正常数字后应恢复可生成');

  /* ---------------- 汇总 ---------------- */
  console.log('\n' + '='.repeat(56));
  if (fails.length) {
    console.log('失败 ' + fails.length + ' / ' + count + ' 项：');
    fails.forEach(function (f, i) { console.log('  ' + (i + 1) + '. ' + f); });
    process.exit(1);
  }
  console.log('全部通过：' + count + ' 项断言');
  console.log('='.repeat(56));
})();
