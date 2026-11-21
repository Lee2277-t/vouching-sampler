/*!
 * 抽凭助手 · core.js
 * 纯逻辑层：表头识别 → 行解析 → 总体过滤 → 抽样 → 抽凭表结构。
 * 不依赖 DOM：浏览器挂 window.AuditCore，node 下直接 require 做断言。
 *
 * 三条硬口径（改代码前先读）：
 *   1. 金额一律以「分」整数存储与运算，展示时才除 100；不得浮点累加。
 *   2. 任何被排除的记录都必须落进排除清单，附原始行号；不许静默丢行。
 *   3. 抽样可复现：同数据 + 同参数 + 同种子 → 同一批样本，起点与间距写进输出。
 */
(function (root, factory) {
  'use strict';
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.AuditCore = factory();
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  var VERSION = '0.6.0';

  /* ============================================================
   * 一、基础：文本、金额、日期
   * ============================================================ */

  function str(v) { return (v === null || v === undefined) ? '' : String(v); }

  /** 表头/标签归一化：去空白与括号冒号逗号，转小写 */
  function norm(v) {
    return str(v)
      .replace(/[\s\u3000]/g, '')
      .replace(/[（）()【】\[\]〔〕:：、,，*\/]/g, '')
      .toLowerCase();
  }

  /**
   * 金额 → 分（整数）。
   * 支持：数字、"1,234.56"、"¥1,234.56"、"(800.00)"、"−500.00"、"500.00-"、"1000元"
   * 返回 {cents, reason, extra}  reason: null(成功) | 'empty' | 'bad'
   * extra=true 表示原值超过两位小数，已按四舍五入处理（会进数据检查清单）。
   */
  function toCents(v) {
    if (v === null || v === undefined) return { cents: null, reason: 'empty', extra: false };

    if (typeof v === 'number') {
      if (!isFinite(v)) return { cents: null, reason: 'bad', extra: false };
      // 转成最短十进制字符串再走同一套解析。
      // 不能用 v.toFixed(2)：那是二进制就近舍入，1.005 会变成 1.00（应为 1.01），
      // 而且会丢掉「超过两位小数」这个该进数据检查清单的信号。
      var sNum = String(v);
      if (/e/i.test(sNum)) return toCents(v.toFixed(2));      // 科学计数法交回字符串路径
      var dot = sNum.indexOf('.');
      if (dot >= 0 && sNum.length - dot - 1 > 6) {             // 浮点噪声（0.30000000000000004）
        return toCents(v.toFixed(2));
      }
      return toCents(sNum);
    }

    var raw = str(v).trim();
    if (!raw) return { cents: null, reason: 'empty', extra: false };

    var s = raw.replace(/[\s\u3000,，¥￥$]/g, '').replace(/(元|人民币|rmb)/gi, '');
    if (!s) return { cents: null, reason: 'empty', extra: false };
    // 全角括号也要认成负数写法（从 PDF/网页抄过来常见）
    s = s.replace(/（/g, '(').replace(/）/g, ')')
         .replace(/[－—–]/g, '-');

    var neg = false;
    if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
    if (/^[-−－—]/.test(s)) { neg = !neg; s = s.slice(1); }
    if (/[-−－—]$/.test(s)) { neg = !neg; s = s.slice(0, -1); }
    if (/^\+/.test(s)) { s = s.slice(1); }

    if (!/^\d*\.?\d*$/.test(s) || s === '' || s === '.') return { cents: null, reason: 'bad', extra: false };

    var m = /^(\d*)(?:\.(\d*))?$/.exec(s);
    if (!m) return { cents: null, reason: 'bad', extra: false };

    var ip = m[1] || '0';
    var frac = m[2] || '';
    var extra = false;
    var roundUp = false;
    if (frac.length > 2) {
      extra = true;
      roundUp = (frac.charCodeAt(2) - 48) >= 5;
      frac = frac.slice(0, 2);
    }
    frac = (frac + '00').slice(0, 2);

    var cents = Number(ip) * 100 + Number(frac);
    if (roundUp) cents += 1;
    if (!isFinite(cents)) return { cents: null, reason: 'bad', extra: false };

    return { cents: neg ? -cents : cents, reason: null, extra: extra };
  }

  /** 分 → "1,234.56"（负数带前置减号） */
  function moneyText(cents) {
    if (cents === null || cents === undefined || !isFinite(cents)) return '';
    var neg = cents < 0;
    var a = Math.abs(Math.round(cents));
    var ip = Math.floor(a / 100);
    var fp = a % 100;
    var s = String(ip).replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + String(fp).padStart(2, '0');
    return (neg ? '-' : '') + s;
  }

  function pad2(n) { return String(n).padStart(2, '0'); }

  function dateText(d) {
    if (!d) return '';
    return d.y + '-' + pad2(d.m) + '-' + pad2(d.d);
  }

  function mkDate(y, m, d) {
    if (!(y >= 1900 && y <= 2999)) return null;
    if (!(m >= 1 && m <= 12)) return null;
    if (!(d >= 1 && d <= 31)) return null;
    var dt = new Date(Date.UTC(y, m - 1, d));
    if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
    return { y: y, m: m, d: d };
  }

  function serialToDate(n) {
    var ms = Math.round((Math.floor(n) - 25569) * 86400000);
    var dt = new Date(ms);
    if (isNaN(dt.getTime())) return null;
    return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
  }

  /** 日期 → {y,m,d}；支持 Excel 日期、Excel 序列号、20260105、"2026年1月5日"、"2026/1/5" */
  function toDate(v) {
    if (v === null || v === undefined || v === '') return null;

    if (v instanceof Date) {
      if (isNaN(v.getTime())) return null;
      return { y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate() };
    }

    if (typeof v === 'number') {
      if (!isFinite(v)) return null;
      if (v >= 20000 && v <= 80000) return serialToDate(v);
      if (v >= 19000101 && v <= 29991231) {
        var t = String(Math.floor(v));
        return mkDate(+t.slice(0, 4), +t.slice(4, 6), +t.slice(6, 8));
      }
      return null;
    }

    var s = str(v).trim();
    if (!s) return null;
    // 先切掉时间部分再删空格：「2026-1-5 10:30:00」删空格后会变成 2026-1-510:30:00，
    // 日会被读成 51
    s = s.replace(/[Tt]\d[\d:.]*$/, '').replace(/\s+\d{1,2}:\d{2}(:\d{2})?(\.\d+)?$/, '');
    s = s.replace(/[\s\u3000]/g, '');

    var m = /^(\d{4})[年\-\/\.](\d{1,2})[月\-\/\.](\d{1,2})/.exec(s);
    if (m) return mkDate(+m[1], +m[2], +m[3]);

    m = /^(\d{4})(\d{2})(\d{2})$/.exec(s);
    if (m) return mkDate(+m[1], +m[2], +m[3]);

    return null;
  }

  function dnum(d) { return d ? d.y * 10000 + d.m * 100 + d.d : null; }

  /* ============================================================
   * 二、表头识别
   * ============================================================ */

  var FIELDS = [
    { key: 'date',        label: '日期',     weight: 3 },
    { key: 'voucher',     label: '凭证号',   weight: 3 },
    { key: 'voucherWord', label: '凭证字',   weight: 2 },
    { key: 'summary',     label: '摘要',     weight: 1 },
    { key: 'account',     label: '科目',     weight: 2 },
    { key: 'accountCode', label: '科目编码', weight: 1 },
    { key: 'entity',      label: '主体',     weight: 2 },
    { key: 'debit',       label: '借方金额', weight: 4 },
    { key: 'credit',      label: '贷方金额', weight: 4 },
    { key: 'amount',      label: '金额（单列）', weight: 4 },
    { key: 'direction',   label: '借贷方向', weight: 2 },
    { key: 'balance',     label: '余额',     weight: 1 },
    { key: 'counterpart', label: '对方科目', weight: 1 },
    { key: 'maker',       label: '制单人',   weight: 1 }
  ];

  /* 别名表按「越具体越靠前」排。真实 ERP 导出里同一个概念至少十几种写法，
   * 而且往往同时给「科目编码」和「科目全路径」——编码必须落到 accountCode，
   * 否则用户按「应收账款」筛的时候会因为列里只有 1122 而一笔都筛不出来。
   *
   * 中英混排在一个表里：norm 会去掉空格和斜杠并转小写，所以 "G/L Account" 归一成
   * "glaccount"、"Posting Date" 归一成 "postingdate"。英文放后面，同分时中文优先。 */
  var ALIASES = {
    date: ['记账日期', '凭证日期', '制单日期', '填制日期', '制证日期', '业务日期', '交易日期', '发生日期', '过账日期', '单据日期', '入账日期', '日期',
           'postingdate', 'documentdate', 'docdate', 'entrydate', 'gldate', 'transdate', 'date'],
    voucher: ['记账凭证号', '凭证字号', '凭证编号', '凭证号码', '凭证号', '单据号', '单号', '号码', '号',
              'documentnumber', 'docnumber', 'vouchernumber', 'voucherno', 'docno', 'referenceno', 'journalnumber', 'voucher'],
    voucherWord: ['凭证字', '凭证类别', '凭证种类', '字', 'doctype', 'documenttype', 'vouchertype'],
    summary: ['业务摘要', '摘要说明', '摘要', '事由', '说明', '内容',
              'linetext', 'description', 'narration', 'particulars', 'memo', 'text'],
    account: ['末级科目全路径', '科目全路径', '科目全称', '会计科目', '一级科目', '明细科目', '末级科目', '科目名称', '科目',
              'glaccount', 'accountname', 'accountdescription', 'account', 'nominal'],
    accountCode: ['末级科目编号', '原科目编码', '科目编码', '科目代码', '科目编号',
                  'accountnumber', 'accountno', 'accountcode', 'glcode', 'glnumber'],
    entity: ['公司代码描述', '企业主体名称', '核算主体名称', '主体名称', '公司名称', '单位名称', '公司代码', '核算主体', '主体',
             'companycodename', 'companyname', 'companydescription', 'companycode', 'entity'],
    debit: ['借方金额', '借方发生额', '借方本币', '本位币借方', '原币借方', '借方合计', '借方',
            'debitamount', 'debitlocal', 'debit'],
    credit: ['贷方金额', '贷方发生额', '贷方本币', '本位币贷方', '原币贷方', '贷方合计', '贷方',
             'creditamount', 'creditlocal', 'credit'],
    amount: ['本位币金额', '原币金额', '发生额', '交易金额', '金额',
             'amountinlocalcurrency', 'localamount', 'amount', 'value'],
    direction: ['借贷方向', '借贷标志', '方向', '借贷', 'debitcreditindicator', 'drcr', 'dcflag'],
    balance: ['期末余额', '期初余额', '本位币余额', '余额',
              'closingbalance', 'endingbalance', 'balance'],
    counterpart: ['对方科目', '对应科目', '客户档案名称', '供应商档案名称', '客户名称', '供应商名称', '对方往来单位名称', '往来单位名称', '往来单位', '对方单位',
                  'offsettingaccount', 'customername', 'vendorname', 'businesspartner'],
    maker: ['制单员', '制单人', '制单', '录入人', '操作员', 'createdby', 'enteredby', 'username', 'maker']
  };

  var LABEL_OF = {};
  var WEIGHT_OF = {};
  FIELDS.forEach(function (f) { LABEL_OF[f.key] = f.label; WEIGHT_OF[f.key] = f.weight; });

  /* 用户手工指定的「表头写法 → 字段」，只存表头文字，不存任何数据。
   * 冷门 ERP 的私有字段名教一次就够，下次自动认。 */
  var LEARNED = {};

  /** 记住一个表头写法。单字太容易误伤，不收。 */
  function learnAlias(headerText, fieldKey) {
    var t = norm(headerText);
    if (!t || t.length < 2) return false;
    if (!fieldKey || !LABEL_OF.hasOwnProperty(fieldKey)) return false;
    LEARNED[t] = fieldKey;
    return true;
  }

  function forgetAliases() { LEARNED = {}; }

  function learnedAliases() {
    var out = {};
    for (var k in LEARNED) if (LEARNED.hasOwnProperty(k)) out[k] = LEARNED[k];
    return out;
  }

  /** 从本地存回来的字典（app 用 localStorage 持久化） */
  function setLearnedAliases(dict) {
    LEARNED = {};
    if (!dict) return;
    for (var k in dict) {
      if (!dict.hasOwnProperty(k)) continue;
      if (!LABEL_OF.hasOwnProperty(dict[k])) continue;
      var t = norm(k);
      if (t && t.length >= 2) LEARNED[t] = dict[k];
    }
  }

  /** 单个表头单元格 → {field, score}
   * 排序三条：精确优于包含；别名越靠前越优先（表里写在前面的才是常用叫法）；
   * 最后才比长度。没有这三条，「对方往来单位名称」会把「对方科目」挤掉。 */
  function matchField(t) {
    if (!t) return null;
    // 教过的写法直接命中，优先于所有内置别名
    if (LEARNED[t]) return { field: LEARNED[t], score: 1000 };
    var best = null;
    for (var key in ALIASES) {
      if (!ALIASES.hasOwnProperty(key)) continue;
      var list = ALIASES[key];
      for (var i = 0; i < list.length; i++) {
        var a = norm(list[i]);
        if (!a) continue;
        var order = (list.length - i) * 10;
        var score = 0;
        if (t === a) score = 100 + a.length + order;
        else if (t.indexOf(a) >= 0 && a.length >= 2) score = 10 + a.length + order;
        if (score > 0 && (!best || score > best.score)) best = { field: key, score: score };
      }
    }
    return best;
  }

  /** 一行候选表头 → 字段映射 */
  function mapRow(row) {
    var cands = [];
    var labels = [];
    for (var c = 0; c < row.length; c++) {
      var t = norm(row[c]);
      labels.push(str(row[c]).trim());
      if (!t) continue;
      var hit = matchField(t);
      if (hit) cands.push({ col: c, field: hit.field, score: hit.score });
    }
    cands.sort(function (a, b) { return (b.score - a.score) || (a.col - b.col); });

    var map = {};
    var taken = {};
    cands.forEach(function (k) {
      if (map[k.field] !== undefined) return;
      map[k.field] = k.col;
      taken[k.col] = true;
    });

    var score = 0;
    for (var f in map) if (map.hasOwnProperty(f)) score += WEIGHT_OF[f] || 1;

    var hasAmountSide = (map.debit !== undefined) || (map.credit !== undefined) || (map.amount !== undefined);
    if (hasAmountSide) score += 6;
    if (map.date !== undefined && map.voucher !== undefined) score += 4;

    return { map: map, score: score, labels: labels, hasAmountSide: hasAmountSide };
  }

  /** 合并单元格只在左上角留值，其余是 null。把非空的上一格向右补齐。 */
  function carryMerged(row) {
    var nonEmpty = 0;
    for (var i = 0; i < row.length; i++) if (str(row[i]).trim()) nonEmpty++;
    if (nonEmpty < 2) return null; // 只有一个格有字，那是标题行不是表头行

    var out = [];
    var last = '';
    for (var c = 0; c < row.length; c++) {
      var t = str(row[c]).trim();
      if (t) { last = t; out[c] = t; } else out[c] = last;
    }
    return out;
  }

  /** 两行拼成一层表头：上一行给分类，下一行给具体列名 */
  function combineHeaderRows(parent, child) {
    var p = carryMerged(parent);
    if (!p) return null;
    var n = Math.max(p.length, (child || []).length);
    var labels = [];
    for (var i = 0; i < n; i++) {
      var a = p[i] || '';
      var b = (child || [])[i] ? str(child[i]).trim() : '';
      labels.push(b ? (a ? a + ' ' + b : b) : a);
    }
    return labels;
  }

  /** 行列互换。有些系统把一笔业务导成一列（竖排），转置后才是「一行一条明细」。 */
  function transposeRows(rows) {
    var maxLen = 0, i;
    for (i = 0; i < rows.length; i++) {
      var len = rows[i] ? rows[i].length : 0;
      if (len > maxLen) maxLen = len;
    }
    var out = [];
    for (var c = 0; c < maxLen; c++) {
      var line = [];
      for (var r = 0; r < rows.length; r++) {
        var v = rows[r] ? rows[r][c] : undefined;
        line.push(v === undefined ? null : v);
      }
      out.push(line);
    }
    return out;
  }

  /** 竖表判定与转置：只有转置后明显更像明细账才采用，否则原样返回 */
  function maybeTranspose(rows, force) {
    var normal = detectHeader(rows);
    var normalVerdict = sheetVerdict(normal);

    if (normalVerdict.usable && !force) {
      return { rows: rows, header: normal, verdict: normalVerdict, transposed: false };
    }

    // 转置代价是行×列，太大就不试了
    var maxLen = 0;
    for (var i = 0; i < rows.length; i++) {
      var len = rows[i] ? rows[i].length : 0;
      if (len > maxLen) maxLen = len;
    }
    if (rows.length * Math.max(1, maxLen) > 2000000) {
      return { rows: rows, header: normal, verdict: normalVerdict, transposed: false };
    }

    var tr = transposeRows(rows);
    var h2 = detectHeader(tr);
    var v2 = sheetVerdict(h2);

    if (v2.usable && (force || v2.score > normalVerdict.score)) {
      return { rows: tr, header: h2, verdict: v2, transposed: true };
    }
    return { rows: rows, header: normal, verdict: normalVerdict, transposed: false };
  }

  /** 在前 maxScan 行里找最像表头的一行；单行赢不了时再试两行拼接 */
  function detectHeader(rows, maxScan) {
    var limit = Math.min(rows.length || 0, maxScan || 25);
    var best = { rowIndex: -1, score: 0, map: {}, labels: [], hasAmountSide: false, span: 1 };

    for (var r = 0; r < limit; r++) {
      var cand = mapRow(rows[r] || []);
      if (cand.score > best.score) {
        best = { rowIndex: r, score: cand.score, map: cand.map, labels: cand.labels, hasAmountSide: cand.hasAmountSide, span: 1 };
      }
    }

    // 两级表头：上一行是「本期发生额」这种分类，下一行才是「借方金额」。两行拼起来明显更强才用。
    for (var r2 = 0; r2 < limit - 1; r2++) {
      var labels = combineHeaderRows(rows[r2] || [], rows[r2 + 1] || []);
      if (!labels) continue;
      var cand2 = mapRow(labels);
      if (cand2.score > best.score + 2) {
        best = { rowIndex: r2, score: cand2.score, map: cand2.map, labels: labels, hasAmountSide: cand2.hasAmountSide, span: 2 };
      }
    }

    return best;
  }

  /* ============================================================
   * 三、行解析
   * ============================================================ */

  // 汇总行：整词命中（本期发生额合计、合计（本期）、本期合计（借方）等都要认出来）
  var SUBTOTAL_RE = /^(小计|合计|总计|本期合计|本年累计|本月合计|累计|以上合计|期初余额|期末余额|承前页|结转下页|结转下年|过次页)$/;
  var SUBTOTAL_PARTS = ['合计', '小计', '总计', '累计'];

  /** 某一格是不是「汇总行」的标签：容忍前后缀与括号（本期发生额合计 / 合计（本期）） */
  function isSubtotalLabel(text) {
    var t = norm(text);
    if (!t) return false;
    if (SUBTOTAL_RE.test(str(text).replace(/[\s\u3000]/g, ''))) return true;
    // 归一化后（去括号/冒号/逗号、转小写）只留汉字与数字，避免「应收账款合计」以外的噪声
    if (t.length > 12) return false;              // 长句子不是标签
    if (/说明|备注|提示|注释|注$/.test(t)) return false;   // 「合计差异说明」这类是注释，不是汇总行
    for (var i = 0; i < SUBTOTAL_PARTS.length; i++) {
      if (t.indexOf(SUBTOTAL_PARTS[i]) >= 0) {
        // 「合计」必须出现在开头或结尾附近，避免把「合计差异说明」当汇总行
        if (t.indexOf(SUBTOTAL_PARTS[i]) <= 2 || t.slice(-4).indexOf(SUBTOTAL_PARTS[i]) >= 0) return true;
      }
    }
    return false;
  }

  /**
   * 判断一张表能不能当抽凭总体用。
   * 抽凭至少要有「日期 + 金额」：科目余额表、报表也有借贷金额列，但没有日期，
   * 硬抽出来的是没有期间的假样本 —— 宁可不给，也不能给错。
   */
  function sheetVerdict(header) {
    var map = (header && header.map) || {};
    var hasAmount = map.debit !== undefined || map.credit !== undefined || map.amount !== undefined;
    var hasDate = map.date !== undefined;
    var hasVoucher = map.voucher !== undefined;

    var problems = [];
    if (!hasAmount) problems.push('没有金额列');
    if (!hasDate) problems.push('没有日期列');

    var usable = hasAmount && hasDate;
    if (usable && !hasVoucher) problems.push('没有凭证号（只能靠原始行号定位）');
    if (!usable && hasAmount && !hasDate) problems.push('看着像科目余额表或报表，不是明细账');

    return {
      usable: usable,
      hasAmount: hasAmount,
      hasDate: hasDate,
      hasVoucher: hasVoucher,
      problems: problems,
      score: (usable ? 1000 : 0) + (hasDate ? 100 : 0) + (hasVoucher ? 100 : 0) + (hasAmount ? 100 : 0) + ((header && header.score) || 0)
    };
  }

  function cellTexts(row) {
    var out = [];
    for (var c = 0; c < row.length; c++) {
      var t = str(row[c]).trim();
      if (t) out.push(t);
    }
    return out;
  }

  function isBlankRow(row) {
    return cellTexts(row).length === 0;
  }

  function isSubtotalRow(row) {
    var texts = cellTexts(row).slice(0, 4);
    for (var i = 0; i < texts.length; i++) {
      var raw = texts[i].replace(/[\s\u3000]/g, '');
      if (SUBTOTAL_RE.test(raw.replace(/[（）()【】\[\]:：]/g, ''))) return raw;
      if (isSubtotalLabel(texts[i])) return raw;      // 本期发生额合计 / 合计（本期）这类
    }
    return null;
  }

  function readAmount(row, map) {
    var out = { debit: null, credit: null, signed: 0, bad: null, extra: false, empty: false };

    if (map.debit !== undefined || map.credit !== undefined) {
      if (map.debit !== undefined) {
        var d = toCents(row[map.debit]);
        if (d.reason === 'bad') out.bad = '借方';
        if (d.cents !== null) out.debit = d.cents;
        if (d.extra) out.extra = true;
      }
      if (map.credit !== undefined) {
        var c = toCents(row[map.credit]);
        if (c.reason === 'bad') out.bad = out.bad ? out.bad + '/贷方' : '贷方';
        if (c.cents !== null) out.credit = c.cents;
        if (c.extra) out.extra = true;
      }
      // 借贷两列都空，是「期初余额」「结转下年」这类非明细行，不是零金额交易
      if (out.debit === null && out.credit === null) { out.empty = true; return out; }
      out.signed = (out.debit || 0) - (out.credit || 0);
      return out;
    }

    if (map.amount !== undefined) {
      var a = toCents(row[map.amount]);
      if (a.reason === 'bad') out.bad = '金额';
      if (a.extra) out.extra = true;
      if (a.cents !== null) {
        var v = a.cents;
        if (map.direction !== undefined) {
          var rawDir = str(row[map.direction]).trim();
          var dir = norm(rawDir);
          // 国产账写「借/贷」，英文账写 DR/CR、C/D、Debit/Credit，还有写 +/- 的
          if (!dir) {
            /* 空方向：保持金额原符号 */
          } else if (dir.indexOf('贷') >= 0 || dir === 'cr' || dir === 'c' ||
                     dir === 'credit' || dir.indexOf('credit') >= 0 ||
                     dir === '-' || dir === '负' || dir.indexOf('减') >= 0) {
            v = -Math.abs(v);
          } else if (dir.indexOf('借') >= 0 || dir === 'dr' || dir === 'd' ||
                     dir === 'debit' || dir.indexOf('debit') >= 0 ||
                     dir === '+' || dir === '正' || dir.indexOf('加') >= 0) {
            v = Math.abs(v);
          } else {
            // 认不出的方向不能默认当借方，否则借贷会一路错到净额和抽凭表
            out.bad = '方向（读不出「' + rawDir + '」是借还是贷）';
            return out;
          }
        }
        out.signed = v;
        if (v >= 0) out.debit = v; else out.credit = -v;
      } else if (a.reason === 'empty') {
        out.empty = true;
      }
      return out;
    }

    out.bad = '未识别金额列';
    return out;
  }

  /**
   * 解析明细行。
   * @param rows    二维数组（含表头）
   * @param header  detectHeader 的结果，rowIndex 为表头所在下标
   * @param opts    { map?: 覆盖字段映射, firstExcelRow?: 该 rows[0] 对应的 Excel 行号, 默认 1 }
   */
  function parseRows(rows, header, opts) {
    opts = opts || {};
    var map = opts.map || header.map || {};
    var firstExcelRow = opts.firstExcelRow || 1;
    var headerIndex = header.rowIndex;

    var records = [];
    var problems = [];   // 金额无法识别：既不能进总体，也不是正常排除
    var skipped = [];    // 非明细行（空行/合计行/无金额行）
    var warnings = { extraDecimals: 0, blankRows: 0, subtotalRows: 0, emptyAmountRows: 0 };

    var dataStart = headerIndex + Math.max(1, header.span || 1);
    for (var r = dataStart; r < rows.length; r++) {
      var row = rows[r] || [];
      var excelRow = firstExcelRow + r;

      if (isBlankRow(row)) { warnings.blankRows++; skipped.push({ srcRow: excelRow, reason: '空行', text: '' }); continue; }

      var sub = isSubtotalRow(row);
      if (sub) { warnings.subtotalRows++; skipped.push({ srcRow: excelRow, reason: '汇总行（非明细，未纳入总体）', text: sub }); continue; }

      var amt = readAmount(row, map);
      if (amt.bad) {
        var why = amt.bad.indexOf('方向') === 0 ? amt.bad : ('金额无法识别：' + amt.bad);
        problems.push({ srcRow: excelRow, reason: why, text: cellTexts(row).join(' | ').slice(0, 80) });
        continue;
      }
      if (amt.empty) {
        warnings.emptyAmountRows++;
        skipped.push({ srcRow: excelRow, reason: '无金额（借贷两列都为空）', text: cellTexts(row).join(' | ').slice(0, 60) });
        continue;
      }
      if (amt.extra) warnings.extraDecimals++;

      var dateRaw = map.date !== undefined ? row[map.date] : null;
      var d = toDate(dateRaw);

      // 「字」+「号」分两列是国产 ERP 的常见做法（记账凭证 / 1），拼成一个凭证号才认得出
      var vNo = map.voucher !== undefined ? str(row[map.voucher]).trim() : '';
      var vWord = map.voucherWord !== undefined ? str(row[map.voucherWord]).trim() : '';
      var voucher = vNo ? (vWord ? vWord + '-' + vNo : vNo) : vWord;

      records.push({
        srcRow: excelRow,
        date: d,
        dateText: d ? dateText(d) : str(dateRaw).trim(),
        voucher: voucher,
        summary: map.summary !== undefined ? str(row[map.summary]).trim() : '',
        account: map.account !== undefined ? str(row[map.account]).trim() : '',
        accountCode: map.accountCode !== undefined ? str(row[map.accountCode]).trim() : '',
        entity: map.entity !== undefined ? str(row[map.entity]).trim() : '',
        debitCents: amt.debit,
        creditCents: amt.credit,
        signedCents: amt.signed
      });
    }

    if (records.length && records.every(function (x) { return !x.voucher; })) warnings.noVoucher = true;

    return { records: records, problems: problems, skipped: skipped, warnings: warnings, map: map };
  }

  /* ============================================================
   * 四、总体过滤
   * ============================================================ */

  function sumAbs(list) {
    var s = 0;
    for (var i = 0; i < list.length; i++) s += Math.abs(list[i].signedCents);
    return s;
  }

  function sumSigned(list) {
    var s = 0;
    for (var i = 0; i < list.length; i++) s += list[i].signedCents;
    return s;
  }

  /** 科目关键词支持多个：空格、逗号、顿号、分号分隔，命中任一即可 */
  function splitKeywords(s) {
    return str(s).split(/[\s,，、;；|]+/).map(norm).filter(function (x) { return !!x; });
  }

  /**
   * 把若干明细行按凭证归成一组：金额相加，行号留全部。
   * 凭证号必须带上期间与来源表：国产 ERP 凭证号按月重排，1 月和 2 月都有「记-1」，
   * 只按凭证号分组会把两张不同月份的凭证合成一张（金额虚增、行号串月）。
   */
  function groupByVoucher(lines) {
    var map = {};
    var order = [];
    for (var i = 0; i < lines.length; i++) {
      var r = lines[i];
      var ym = r.date ? (r.date.y + '-' + pad2(r.date.m)) : '';
      var key = (r.entity ? r.entity + '|' : '') +
                (ym ? ym + '#' : '') +
                (r.sheet ? r.sheet + '@' : '') +
                (r.voucher || ('#行' + r.srcRow));
      var g = map[key];
      if (!g) {
        g = map[key] = {
          key: key, isGroup: true,
          srcRow: r.srcRow,
          date: r.date, dateText: r.dateText,
          voucher: r.voucher, entity: r.entity,
          summary: r.summary, account: r.account, accountCode: r.accountCode,
          debitCents: null, creditCents: null, signedCents: 0,
          lineCount: 0, rows: []
        };
        order.push(g);
      }
      g.lineCount++;
      g.rows.push(r.srcRow);
      if (r.debitCents !== null) g.debitCents = (g.debitCents || 0) + r.debitCents;
      if (r.creditCents !== null) g.creditCents = (g.creditCents || 0) + r.creditCents;
    }

    // 凭证的「金额」是发生额，不是净额：一张借贷平衡的凭证净额为 0，
    // 拿来分层会被全部归到「金额为零」。取借/贷较大的一方作为凭证金额。
    order.forEach(function (g) {
      var dr = g.debitCents || 0;
      var cr = g.creditCents || 0;
      g.signedCents = dr >= cr ? dr : -cr;
      g.rows.sort(function (a, b) { return a - b; });   // 原始行号按 Excel 顺序
    });
    return order;
  }

  function statsOf(list, excludedCount) {
    var totalAbs = 0, net = 0, debit = 0, credit = 0;
    for (var i = 0; i < list.length; i++) {
      var r = list[i];
      totalAbs += Math.abs(r.signedCents);
      net += r.signedCents;
      debit += r.debitCents || 0;
      credit += r.creditCents || 0;
    }
    return {
      count: list.length, totalAbsCents: totalAbs, netCents: net,
      debitCents: debit, creditCents: credit, excludedCount: excludedCount
    };
  }

  /**
   * @param records parseRows 的 records
   * @param f { dateFrom, dateTo: {y,m,d}|null,
   *            accountText: 关键词，空格/逗号分隔，命中任一即可
   *            entity: 主体名称（精确匹配）
   *            direction: 'all'|'debit'|'credit',
   *            minCents, maxCents: number|null,
   *            unit: 'line'|'voucher' }
   */
  function buildPopulation(records, f) {
    f = f || {};
    var from = f.dateFrom ? dnum(f.dateFrom) : null;
    var to = f.dateTo ? dnum(f.dateTo) : null;
    var keys = splitKeywords(f.accountText);
    var entity = f.entity ? norm(f.entity) : '';
    var dir = f.direction || 'all';
    var unit = f.unit === 'voucher' ? 'voucher' : 'line';
    var min = (f.minCents === null || f.minCents === undefined) ? null : f.minCents;
    var max = (f.maxCents === null || f.maxCents === undefined) ? null : f.maxCents;

    var included = [];
    var excluded = [];

    for (var i = 0; i < records.length; i++) {
      var r = records[i];
      var abs = Math.abs(r.signedCents);
      var dn = r.date ? dnum(r.date) : null;
      var reason = null;

      var acctHit = keys.length === 0;
      if (keys.length) {
        var a1 = norm(r.account), a2 = norm(r.accountCode);
        for (var k = 0; k < keys.length; k++) {
          if ((a1 && a1.indexOf(keys[k]) >= 0) || (a2 && a2.indexOf(keys[k]) >= 0)) { acctHit = true; break; }
        }
      }

      if ((from !== null || to !== null) && dn === null) reason = '日期缺失';
      else if (from !== null && dn < from) reason = '早于期间起点';
      else if (to !== null && dn > to) reason = '晚于期间终点';
      else if (entity && norm(r.entity) !== entity) reason = '不属于所选主体';
      else if (keys.length && !acctHit) reason = '科目不含任一关键词';
      else if (dir === 'debit' && r.signedCents <= 0) reason = '方向不符（要求借方）';
      else if (dir === 'credit' && r.signedCents >= 0) reason = '方向不符（要求贷方）';
      else if (r.signedCents === 0) reason = '金额为零';
      else if (min !== null && abs < min) reason = '金额低于下限';
      else if (max !== null && abs > max) reason = '金额高于上限';

      if (reason) {
        excluded.push({
          srcRow: r.srcRow, reason: reason, dateText: r.dateText, voucher: r.voucher,
          summary: r.summary, account: r.account, entity: r.entity,
          debitCents: r.debitCents, creditCents: r.creditCents
        });
      } else {
        included.push(r);
      }
    }

    included.sort(cmpForSampling);

    if (unit === 'voucher') {
      var groups = groupByVoucher(included);
      groups.sort(cmpForSampling);
      return { included: groups, excluded: excluded, unit: 'voucher', stats: statsOf(groups, excluded.length) };
    }

    return { included: included, excluded: excluded, unit: 'line', stats: statsOf(included, excluded.length) };
  }

  /** 抽样用的稳定排序：日期 → 凭证号 → 金额 → 摘要 → 行号。保证同一份数据重导出后顺序一致。 */
  function cmpForSampling(a, b) {
    var ka = [dnum(a.date) || 0, a.voucher || '', a.signedCents, a.summary || '', a.srcRow];
    var kb = [dnum(b.date) || 0, b.voucher || '', b.signedCents, b.summary || '', b.srcRow];
    for (var i = 0; i < ka.length; i++) {
      if (ka[i] < kb[i]) return -1;
      if (ka[i] > kb[i]) return 1;
    }
    return 0;
  }

  /* ============================================================
   * 五、抽样（全部种子化，可复现）
   * ============================================================ */

  function mulberry32(a) {
    a = a >>> 0;
    return function () {
      a = (a + 0x6D2B79F5) >>> 0;
      var t = a;
      t = Math.imul(t ^ (t >>> 15), 1 | t);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  /**
   * 等距抽样：间距 k=N/n 用小数，起点 ∈ [0,k)。
   * k 不能用 floor(N/n)：n 超过 N/2 时 floor 会退化成 1，样本变成「前 n 笔」，
   * 总体按日期排序，等于只抽了最早一段（小科目 40 笔抽 25 笔就会踩到）。
   * 复核时用「起点 + 间距」仍能手算。
   */
  function sampleSystematic(pop, n, seed) {
    var N = pop.length;
    if (n <= 0) return { indices: [], k: 0, start: 0, note: '' };
    if (n >= N) return { indices: pop.map(function (_, i) { return i; }), k: 1, start: 0, note: '样本量不小于总体，全部抽取' };

    var interval = N / n;
    var rnd = mulberry32(seed);
    var startF = rnd() * interval;

    var idx = [];
    for (var i = 0; i < n; i++) {
      var p = Math.floor(startF + i * interval);
      if (p >= N) p = N - 1;
      if (idx.length && idx[idx.length - 1] === p) continue;   // 间距接近 1 时防重
      idx.push(p);
    }
    return {
      indices: idx,
      k: Math.round(interval * 100) / 100,
      start: Math.round(startF * 100) / 100,
      note: ''
    };
  }

  /** 随机抽样：Fisher-Yates 部分洗牌，种子固定则结果固定。 */
  function sampleRandom(pop, n, seed) {
    var N = pop.length;
    if (n <= 0) return { indices: [], k: 0, start: 0, note: '' };
    if (n >= N) return { indices: pop.map(function (_, i) { return i; }), k: 0, start: 0, note: '样本量不小于总体，全部抽取' };
    var idx = [];
    for (var i = 0; i < N; i++) idx.push(i);
    var rnd = mulberry32(seed);
    for (var j = 0; j < n; j++) {
      var pick = j + Math.floor(rnd() * (N - j));
      var tmp = idx[j]; idx[j] = idx[pick]; idx[pick] = tmp;
    }
    var chosen = idx.slice(0, n).sort(function (a, b) { return a - b; });
    return { indices: chosen, k: 0, start: 0, note: '' };
  }

  /**
   * 金额加权抽样（PPS / 货币单元抽样选样）：
   * 把总体想成一条按金额排的「金额轴」，等距取点，落到哪一笔就抽哪一笔。
   * 结果：金额越大，占的轴越长，被抽中的概率越高；单笔金额 ≥ 间隔的必然被抽中。
   * @returns {indices, interval, start, hits}  hits[i] 表示第 i 笔被命中几次
   */
  function samplePPS(pop, n, seed) {
    var N = pop.length;
    if (n <= 0 || !N) return { indices: [], interval: 0, start: 0, hits: {}, note: '' };

    var total = sumAbs(pop);
    if (total <= 0) return { indices: [], interval: 0, start: 0, hits: {}, note: '总体金额为 0' };

    if (n >= N) {
      var all = [];
      for (var q = 0; q < N; q++) all.push(q);
      return { indices: all, interval: 0, start: 0, hits: {}, note: '样本量不小于总体，全部抽取' };
    }

    var interval = total / n;
    var rnd = mulberry32(seed);
    var start = rnd() * interval;

    var cum = 0, idx = 0;
    var picked = [], hits = {};
    for (var t = start; t < total && picked.length < n; t += interval) {
      while (idx < N - 1 && cum + Math.abs(pop[idx].signedCents) <= t) {
        cum += Math.abs(pop[idx].signedCents);
        idx++;
      }
      picked.push(idx);
      hits[idx] = (hits[idx] || 0) + 1;
    }

    var uniq = Object.keys(hits).map(Number).sort(function (a, b) { return a - b; });
    return { indices: uniq, interval: interval, start: start, hits: hits, note: '' };
  }

  /**
   * 分层抽样：|金额| ≥ 阈值 的全部抽（全查层），其余按等距/随机抽 n 笔。
   * @param pop  已排序的总体
   * @param opts { thresholdCents, n, seed, method: 'systematic'|'random'|'pps' }
   */
  function buildSample(pop, opts) {
    opts = opts || {};
    var method = opts.method === 'random' ? 'random' : (opts.method === 'pps' ? 'pps' : 'systematic');
    var seed = (opts.seed === undefined || opts.seed === null) ? 20260101 : (opts.seed >>> 0);
    var threshold = Math.max(0, Math.round(opts.thresholdCents || 0));
    var n = Math.max(0, Math.round(opts.n || 0));

    // 金额加权不分层：金额本身已经决定了被抽中的概率，再套一层阈值只会重复照顾大额
    if (method === 'pps') {
      var pps = samplePPS(pop, n, seed);
      var pSample = pps.indices.map(function (i) {
        return { rec: pop[i], layer: '金额加权', hits: pps.hits[i] || 1 };
      });
      pSample.sort(function (a, b) { return cmpForSampling(a.rec, b.rec); });

      var pAbs = pSample.reduce(function (s, x) { return s + Math.abs(x.rec.signedCents); }, 0);
      var pPopAbs = sumAbs(pop);
      var pVouchers = {};
      var pCount = 0;
      pSample.forEach(function (x) {
        var v = x.rec.voucher;
        if (v && !pVouchers[v]) { pVouchers[v] = 1; pCount++; }
      });

      var pNote = pps.note || '';
      if (!pNote && n > 0 && pSample.length < n) {
        pNote = '一笔大额可能横跨多个金额间隔，因此实抽 ' + pSample.length +
          ' 笔少于目标 ' + n + ' 笔——这是金额加权的正常结果，不是漏抽。';
      }

      return {
        sample: pSample,
        meta: {
          method: 'pps',
          methodLabel: '金额加权抽样（金额越大越容易被抽中）',
          seed: seed,
          n: n,
          thresholdCents: 0,
          popCount: pop.length,
          popAbsCents: pPopAbs,
          fullCount: 0,
          fullAbsCents: 0,
          restCount: pop.length,
          restAbsCents: pPopAbs,
          sampleCount: pSample.length,
          sampleAbsCents: pAbs,
          voucherCount: pCount,
          coverage: pPopAbs > 0 ? pAbs / pPopAbs : 0,
          k: 0,
          start: pps.start ? Math.round(pps.start) : 0,
          intervalCents: pps.interval,
          note: pNote
        }
      };
    }

    var full = [];
    var rest = [];
    for (var i = 0; i < pop.length; i++) {
      if (threshold > 0 && Math.abs(pop[i].signedCents) >= threshold) full.push(pop[i]);
      else rest.push(pop[i]);
    }

    var pick = n > 0
      ? (method === 'random' ? sampleRandom(rest, n, seed) : sampleSystematic(rest, n, seed))
      : { indices: [], k: 0, start: 0, note: '样本量为 0' };

    var sample = [];
    full.forEach(function (r) { sample.push({ rec: r, layer: '全查' }); });
    pick.indices.forEach(function (i) { if (rest[i]) sample.push({ rec: rest[i], layer: '抽样' }); });
    sample.sort(function (a, b) { return cmpForSampling(a.rec, b.rec); });

    var sampleAbs = sample.reduce(function (s, x) { return s + Math.abs(x.rec.signedCents); }, 0);
    var popAbs = sumAbs(pop);
    var coverage = popAbs > 0 ? sampleAbs / popAbs : 0;

    // 同一张凭证可能被抽中多行。现场是按凭证翻的，所以要单独报「涉及凭证数」。
    var voucherSet = {};
    var voucherCount = 0;
    sample.forEach(function (x) {
      var v = x.rec.voucher;
      if (v && !voucherSet[v]) { voucherSet[v] = 1; voucherCount++; }
    });

    return {
      sample: sample,
      meta: {
        method: method,
        methodLabel: threshold > 0
          ? (method === 'random' ? '大额全查 + 其余随机抽样' : '大额全查 + 其余等距抽样')
          : (method === 'random' ? '随机抽样（不分层）' : '等距抽样（不分层）'),
        seed: seed,
        n: n,
        thresholdCents: threshold,
        popCount: pop.length,
        popAbsCents: popAbs,
        fullCount: full.length,
        fullAbsCents: sumAbs(full),
        restCount: rest.length,
        restAbsCents: sumAbs(rest),
        sampleCount: sample.length,
        sampleAbsCents: sampleAbs,
        voucherCount: voucherCount,
        coverage: coverage,
        k: pick.k,
        start: pick.start,
        note: pick.note || ''
      }
    };
  }

  /* ============================================================
   * 六、抽凭表（输出结构）
   * ============================================================ */

  var EVIDENCE_PRESETS = {
    '采购/应付': ['合同', '发票', '验收单', '入库单', '付款审批', '银行回单'],
    '销售/应收': ['合同', '发票', '出库单', '签收单', '收款回单', '信用审批'],
    '费用/报销': ['发票', '报销单', '审批单', '付款回单', '预算依据'],
    '资金/银行': ['银行回单', '对账单', '审批单', '协议/合同'],
    '存货/成本': ['入库单', '出库单', '发票', '领料单', '盘点表'],
    '固定资产': ['验收单', '发票', '合同', '折旧计算表', '盘点表'],
    '通用': ['原始单据', '审批', '合同', '其他']
  };

  function presetColumns(name) {
    return EVIDENCE_PRESETS[name] || EVIDENCE_PRESETS['通用'];
  }

  /**
   * 生成四个工作表的数据（二维数组，纯数据，不碰 xlsx 库）。
   * @param box { client, indexNo, preparer, reviewer, dateText, periodText, accountText,
   *              presetName, sample, meta, populationStats, excluded, problems, skipped, warnings,
   *              mapNotes: string[] }
   */
  function buildSheets(box) {
    var evid = presetColumns(box.presetName);
    var m = box.meta;
    var ps = box.populationStats;

    /* ---- 抽凭表 ---- */
    var main = [];
    main.push(['抽凭表', '', '', '', '索引号：' + (box.indexNo || '')]);
    main.push(['被审计单位：' + (box.client || ''), '', '', '', '编制人：' + (box.preparer || ''), '', '复核人：' + (box.reviewer || '')]);
    main.push(['科目/循环：' + (box.accountText || ''), '', '', '', '编制日期：' + (box.dateText || '')]);
    main.push(['抽样期间：' + (box.periodText || ''), '', '', '', '抽样总体：' + ps.count + ' 笔，金额合计 ' + moneyText(ps.totalAbsCents)]);
    main.push(['抽样方法：' + m.methodLabel +
      (m.method === 'pps'
        ? '（金额间隔 ' + moneyText(Math.round(m.intervalCents || 0)) + '，单笔达到间隔即必然抽中）'
        : '（阈值 ' + moneyText(m.thresholdCents) + ' 以上全查，其余 ' + m.n + ' 笔）')]);
    main.push(['样本：' + m.sampleCount + ' 笔，覆盖金额 ' + moneyText(m.sampleAbsCents) + '（占总体 ' + (m.coverage * 100).toFixed(2) + '%）']);
    main.push(['抽样参数（复核用）：随机种子 ' + m.seed + '；等距 ' + (m.k || '—') + '；起点 ' + (m.start || 0) + '；全查层 ' + m.fullCount + ' 笔 / 抽样层 ' + m.restCount + ' 笔']);
    main.push(['复核提示：同人员用同参数重跑本项目，样本应与本表逐笔一致；不一致则以本表参数为准重新核对。']);
    main.push([]);

    var isVoucherUnit = m.unit === 'voucher';
    var rowCol = box.transposed ? '原始列号' : '原始行号';
    // 多表/多文件合并后行号会重复（每张表的第 5 行都叫 5），必须带上来源才追得回原文
    var multiSrc = !!(box.meta && box.meta.multiSheet);
    function where(r) {
      return multiSrc ? ((r.sheet || '') + ' 第' + r.srcRow + (box.transposed ? '列' : '行')) : r.srcRow;
    }
    var head = ['序号', rowCol, '凭证日期', '凭证号', '主体', '科目', '摘要',
                '借方金额', '贷方金额', '金额（绝对值）', '分录行数', '抽样类别']
      .concat(evid)
      .concat(['核对结论', '错报金额（元）', '核对说明']);
    main.push(head);

    var seq = 0;
    box.sample.forEach(function (x) {
      seq++;
      var r = x.rec;
      main.push([
        seq, where(r), r.dateText, r.voucher, r.entity || '', r.account, r.summary,
        r.debitCents === null ? '' : r.debitCents / 100,
        r.creditCents === null ? '' : r.creditCents / 100,
        Math.abs(r.signedCents) / 100,
        r.isGroup ? r.lineCount : 1,
        x.layer
      ].concat(evid.map(function () { return ''; }))
        .concat(['', '', '']));
    });

    var sumDr = 0, sumCr = 0, sumAbs = 0;
    box.sample.forEach(function (x) {
      sumDr += x.rec.debitCents || 0;
      sumCr += x.rec.creditCents || 0;
      sumAbs += Math.abs(x.rec.signedCents);
    });
    // 合计行必须和表头一样宽，否则以后加列会错位
    var sumRow = new Array(head.length);
    for (var si = 0; si < sumRow.length; si++) sumRow[si] = '';
    sumRow[0] = '合计';
    sumRow[7] = sumDr / 100;
    sumRow[8] = sumCr / 100;
    sumRow[9] = sumAbs / 100;
    main.push(sumRow);
    main.push([]);
    main.push(['说明：抽查结论由审计人员判断并签字，本表只提供样本清单与金额，不代替判断。勾稽列由现场核对后手工勾选。']);
    main.push(['回录：现场核对完，在「核对结论」填 相符／不符／未取得／不适用，「错报金额」填差异金额（元，正数=账面多计）。' +
               '回到工具第二页选这份文件，可算出例外与错报上限。']);

    /* ---- 抽样信息 ---- */
    var info = [];
    info.push(['抽样信息', '']);
    info.push(['工具版本', '抽凭助手 v' + VERSION]);
    if (box.sheetNames && box.sheetNames.length > 1) {
      info.push(['数据来源', box.sheetNames.join('、') + '（多个工作表合并）']);
    } else if (box.sheetNames && box.sheetNames.length === 1) {
      info.push(['数据来源', box.sheetNames[0]]);
    }
    info.push(['总体笔数', ps.count]);
    info.push(['总体金额（绝对值合计）', moneyText(ps.totalAbsCents)]);
    info.push(['总体借方合计', moneyText(ps.debitCents)]);
    info.push(['总体贷方合计', moneyText(ps.creditCents)]);
    info.push(['总体净额', moneyText(ps.netCents)]);
    info.push(['抽样单元', m.unit === 'voucher' ? '凭证（同一张凭证合并为一笔）' : '明细行']);

    // 筛选条件必须全部入档：总体边界比种子更影响结果，缺了就没法重跑出同一份样本
    var f = box.filter || {};
    function dirName(v) { return v === 'debit' ? '只看借方' : (v === 'credit' ? '只看贷方' : '借贷都看'); }
    function yuanOr(c) { return (c === null || c === undefined) ? '不限' : moneyText(c); }
    info.push(['筛选条件·期间', box.periodText || '全部日期']);
    info.push(['筛选条件·科目', f.accountText || '全部科目']);
    info.push(['筛选条件·主体', f.entity || '全部主体']);
    info.push(['筛选条件·方向', dirName(f.direction)]);
    info.push(['筛选条件·金额下限', yuanOr(f.minCents)]);
    info.push(['筛选条件·金额上限', yuanOr(f.maxCents)]);
    if (box.headerRow) info.push(['表头所在行', '第 ' + box.headerRow + ' 行']);

    info.push(['抽样方法', m.methodLabel]);
    info.push(['全查阈值', moneyText(m.thresholdCents)]);
    info.push(['全查层笔数 / 金额', m.fullCount + ' 笔 / ' + moneyText(m.fullAbsCents)]);
    info.push(['抽样层笔数 / 金额', m.restCount + ' 笔 / ' + moneyText(m.restAbsCents)]);
    info.push(['抽样层样本量', m.n]);
    info.push(['随机种子', m.seed]);
    info.push(['等距', m.k || '—']);
    info.push([m.method === 'pps' ? '起点（金额）' : '起点',
      m.method === 'pps' ? moneyText(m.start) : m.start,
      m.method === 'pps' ? '金额轴上的起点（元）；等距的起点是行序号，两者单位不同' : '']);
    if (m.intervalCents) {
      info.push(['金额间隔', moneyText(Math.round(m.intervalCents)),
        '金额加权抽样的间距；单笔金额超过它的必然被抽中']);
    }
    info.push(['样本笔数', m.sampleCount]);
    info.push(['样本涉及凭证数', m.voucherCount]);
    if (m.voucherCount && m.voucherCount < m.sampleCount) {
      info.push(['提示', '有 ' + (m.sampleCount - m.voucherCount) + ' 行与别的样本同属一张凭证，' +
        '现场只需翻 ' + m.voucherCount + ' 张凭证。']);
    }
    info.push(['样本金额', moneyText(m.sampleAbsCents)]);
    info.push(['覆盖率', (m.coverage * 100).toFixed(2) + '%']);
    info.push(['排除记录数', ps.excludedCount]);
    info.push(['生成时间', box.dateText]);
    if (m.note) info.push(['备注', m.note]);

    /* ---- 排除清单 ---- */
    var excl = [];
    excl.push(['排除清单（未纳入抽样总体的记录，逐条列出，可回溯到原始行号）']);
    excl.push([box.transposed ? '原始列号' : '原始行号', '排除原因', '凭证日期', '凭证号', '主体', '科目', '摘要', '借方', '贷方']);
    box.excluded.forEach(function (e) {
      excl.push([
        (multiSrc ? ((e.sheet || '') + ' 第' + e.srcRow + '行') : e.srcRow), e.reason, e.dateText, e.voucher, e.entity || '', e.account, e.summary,
        e.debitCents === null ? '' : e.debitCents / 100,
        e.creditCents === null ? '' : e.creditCents / 100
      ]);
    });
    if (!box.excluded.length) excl.push(['（无）']);

    /* ---- 数据检查 ---- */
    var chk = [];
    chk.push(['数据检查（导入阶段发现的问题，先看这里再看结果）']);
    chk.push(['项目', '数量', '说明']);
    chk.push(['明细行数', box.warnings.parsedRows, '成功解析为明细凭证的行数']);
    chk.push(['空行', box.warnings.blankRows, '已跳过，非明细']);
    chk.push(['汇总/合计行', box.warnings.subtotalRows, '已跳过，非明细']);
    chk.push(['无金额行', box.warnings.emptyAmountRows || 0, '借贷两列都为空，已跳过（不是零金额交易，如「期初余额」行）']);
    chk.push(['金额/方向无法识别', box.problems.length, '这些行读不出金额或借贷方向，既不进总体、也不进「排除清单」，需回 Excel 核对：下方逐条列出']);
    chk.push(['金额超两位小数', box.warnings.extraDecimals, '已按四舍五入到分，建议回原始数据确认']);
    chk.push(['凭证号', box.warnings.noVoucher ? '未识别到' : '已识别',
      box.warnings.noVoucher ? '抽凭表靠「原始行号」定位，请回原始 Excel 按行号核对' : '']);
    chk.push([]);
    if (box.mapNotes && box.mapNotes.length) {
      chk.push(['列识别提示']);
      box.mapNotes.forEach(function (t) { chk.push([t]); });
      chk.push([]);
    }
    if (box.problems.length) {
      chk.push(['无法识别的行', '原始行号', '内容']);
      box.problems.forEach(function (p) { chk.push([p.reason, p.srcRow, p.text]); });
      chk.push([]);
    }
    if (box.skipped && box.skipped.length) {
      chk.push(['跳过的非明细行', '原始行号', '内容']);
      box.skipped.forEach(function (s) { chk.push([s.reason, s.srcRow, s.text]); });
    }

    /* ---- 使用说明 ---- */
    var help = [];
    help.push(['口径与复核说明']);
    help.push(['1', '金额：全部按「分」整数计算，展示才四舍五入到两位小数；任何情况下不修改原始金额。']);
    help.push(['2', '总体：只有同时满足期间、科目、方向、金额条件的明细行进入总体。被排除的记录逐条列在「排除清单」；解析阶段读不出的行（空行、汇总行、无金额行、金额或方向无法识别）另在「数据检查」页逐条列出，两处合起来才是全部。']);
    if (box.transposed) {
      help.push(['2b', '这张表原本是竖排（一笔业务占一列），已转置后读取；「原始列号」指原表的列位置，不是行号。']);
    }
    if (m.unit === 'voucher') {
      help.push(['2a', '抽样单元：凭证。同一张凭证的多行合并为一笔，金额取借/贷发生额较大的一方；「分录行数」列写明该凭证合了几行。']);
    } else {
      help.push(['2a', '抽样单元：明细行。同一张凭证的多行各自成样本，「分录行数」列会提示是否与别的样本同凭证。']);
    }
    help.push(['3', '抽样：可复现。同种子 + 同参数 + 同一份导出数据 → 同一批样本；种子、起点、间距已写入本表。']);
    help.push(['4', '排除不等于没问题：排除只说明这些记录不在本次抽样范围，其性质仍需按审计程序另行确定。']);
    help.push(['5', '结论：本表不代替判断，只提供样本清单与金额；抽查结论、例外评价由审计人员判断并签字。']);
    help.push(['6', '回录：现场核对完，在抽凭表「核对结论」填 相符／不符／未取得／不适用，「错报金额」填差异（元，正数=账面多计），回到工具第二页选这份文件即可评估。']);
    help.push(['7', '数据安全：全部计算在本机浏览器内完成，不联网、不上传、不留副本。']);
    help.push([]);
    help.push(['常见问题']);
    help.push(['抽出来的笔数和预期不符', '先看「数据检查」里的空行/汇总行/金额无法识别三项，再看「排除清单」的排除原因计数。']);
    help.push(['科目关键词排除过多', '科目列可能是多级编码；把关键词改短（如「应收账款」改「应收」）或清空科目条件重试。']);
    help.push(['日期全部缺失', '多半是日期列没识别对；在「列识别」里手工指定日期列后重新生成。']);

    return [
      { name: '抽凭表', rows: main, moneyCols: [7, 8, 9, 13 + evid.length] },
      { name: '抽样信息', rows: info },
      { name: '排除清单', rows: excl, moneyCols: [7, 8] },
      { name: '数据检查', rows: chk },
      { name: '说明', rows: help }
    ];
  }

  /** 输出文件名：抽凭表_客户_科目_YYYYMMDD.xlsx（非法字符替换为下划线） */
  function fileName(box) {
    var parts = ['抽凭表'];
    if (box.client) parts.push(box.client);
    if (box.accountText) parts.push(box.accountText);
    var d = new Date();
    parts.push(d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate()));
    return parts.join('_').replace(/[\\\/:*?"<>|\s]/g, '_') + '.xlsx';
  }

  /* ============================================================
   * 七、回录与错报评估
   *
   * 抽样抽样，抽完要能评。这里只算数字，不下结论：
   *   - 例外汇总：笔数、金额
   *   - 货币单元抽样的错报上限（Stringer 上界），口径同 AICPA 审计抽样指南
   * 上界是一个统计推断值，不是审计结论。是否可接受由审计人员判断。
   * ============================================================ */

  /** 泊松分布 P(X ≤ k; λ) */
  function poissonCdf(k, lambda) {
    if (lambda <= 0) return 1;
    var term = Math.exp(-lambda);
    var sum = term;
    for (var i = 1; i <= k; i++) {
      term = term * lambda / i;
      sum += term;
    }
    return sum;
  }

  /**
   * 上界系数 R_k：使 P(X ≤ k; λ) = alpha 的 λ。
   * 含义是「k 笔错报时，总体错报率不超过 λ/n」的置信上界系数。
   * 注意分布函数随 λ 递减，方程是 = alpha 而不是 = 1 − alpha：
   *   k=0 时 λ = −ln(alpha)（95% → 3.00，90% → 2.31）
   *   k=1 时 λ ≈ 4.75，k=2 时 ≈ 6.30
   */
  function poissonUpperFactor(k, alpha) {
    if (k < 0) return 0;
    if (k === 0) return -Math.log(alpha);
    var lo = 0, hi = 1;
    while (poissonCdf(k, hi) > alpha && hi < 500) hi *= 2;
    for (var it = 0; it < 200; it++) {
      var mid = (lo + hi) / 2;
      if (poissonCdf(k, mid) > alpha) lo = mid; else hi = mid;
    }
    return (lo + hi) / 2;
  }

  /**
   * Stringer 上界（MUS 错报上限）。
   *   UM = (BV / n) × [ R_0 + Σ t_j × (R_j − R_{j−1}) ]，t_j 按降序。
   * 当所有 taint 都是 1 时，方括号退化成 R_k，与「k 笔偏差」的上界一致。
   *
   * @param taints  每笔例外的错报比例（|错报| / 账面），降序不需要，函数内排序
   * @param n       实际检查的样本量
   * @param bookCents 总体账面金额（分）
   * @param alpha   误受风险（0.05 → 95% 置信）
   */
  function stringerUpperBound(taints, n, bookCents, alpha) {
    if (!n || n <= 0 || !(bookCents > 0)) return null;
    var sorted = (taints || []).slice().sort(function (a, b) { return b - a; });
    var prev = poissonUpperFactor(0, alpha);
    var acc = prev;
    var factors = [prev];
    for (var j = 1; j <= sorted.length; j++) {
      var r = poissonUpperFactor(j, alpha);
      factors.push(r);
      acc += sorted[j - 1] * (r - prev);
      prev = r;
    }
    return {
      upperCents: Math.min(Math.round(bookCents / n * acc), bookCents),
      rawCents: Math.round(bookCents / n * acc),
      capped: Math.round(bookCents / n * acc) > bookCents,
      factorTotal: acc,
      factors: factors,
      perItemCents: Math.round(bookCents / n)
    };
  }

  var CONCLUSION_WORDS = {
    ok: ['相符', '符合', '√', '✓', '是', '对', '通过', '无误', '正常', 'ok', 'yes', 'y',
         '未见异常', '无异常', '未发现异常', '没有异常', '无差异', '未发现差异', '没发现问题'],
    exception: ['不符', '不符合', '不相符', '未相符', '不匹配', '×', '✗', '差异', '有误', '例外',
                '错误', '不一致', '异常', '否', 'no', 'n'],
    missing: ['未取得', '未提供', '未获取', '未回函', '未收到', '未回复', '无回函', '没拿到', '无法取得', '无资料', '缺失'],
    na: ['不适用', '本期无发生', '无发生', '无业务', 'na', 'n/a', '-', '—']
  };

  var NEG_CHARS = ['不', '未', '无', '没', '非', '否'];
  // 长度相同时的优先级：把例外误判成相符会让它从评估里彻底消失，
  // 反过来只是多报一笔让人复核，所以例外排前面。
  var CONCLUS_RANK = { exception: 0, missing: 1, na: 2, ok: 3 };

  /** 把自由填写的结论归一。认不出来的不猜，返回 'other' 让上层报出来。 */
  function classifyConclusion(raw) {
    var t = norm(raw);
    if (!t) return '';

    // 整格就是这个词：先精确判，避免「未见异常」被里面的「异常」抢走
    for (var key in CONCLUSION_WORDS) {
      if (!CONCLUSION_WORDS.hasOwnProperty(key)) continue;
      var list = CONCLUSION_WORDS[key];
      for (var i = 0; i < list.length; i++) {
        if (t === norm(list[i])) return key;
      }
    }

    // 按最长命中匹配：「不符合规定」必须落到「不符合」，不能被「符合」抢走
    var best = null;
    for (var k2 in CONCLUSION_WORDS) {
      if (!CONCLUSION_WORDS.hasOwnProperty(k2)) continue;
      var l2 = CONCLUSION_WORDS[k2];
      for (var j = 0; j < l2.length; j++) {
        var w = norm(l2[j]);
        if (w.length < 2) continue;
        var at = t.indexOf(w);
        if (at < 0) continue;

        var key2 = k2;
        // 「不相符」「未相符」：命中肯定词但前面有否定字 → 例外
        // 「未见异常」「无差异」：命中否定词但整个词是肯定意思 → 相符
        if (at > 0 && NEG_CHARS.indexOf(t.charAt(at - 1)) >= 0) {
          if (key2 === 'ok') key2 = 'exception';
          else if (key2 === 'exception') key2 = 'ok';
        }

        var better = !best || w.length > best.len ||
          (w.length === best.len && CONCLUS_RANK[key2] < CONCLUS_RANK[best.key]);
        if (better) best = { key: key2, len: w.length };
      }
    }
    return best ? best.key : 'other';
  }

  /** 从抽凭表里读回现场填的结果 */
  function parseVouchingRows(rows) {
    var warnings = [];
    var headIdx = -1;
    for (var r = 0; r < Math.min(rows.length, 40); r++) {
      var cells = (rows[r] || []).map(norm);
      if (cells.indexOf('序号') >= 0 && cells.indexOf('凭证号') >= 0) { headIdx = r; break; }
    }
    if (headIdx < 0) {
      return { records: [], warnings: ['没找到抽凭表的表头：需要一行里有「序号」和「凭证号」。确认选的是本工具生成的抽凭表。'], headerRowIndex: -1 };
    }

    var head = (rows[headIdx] || []).map(function (c) { return str(c).trim(); });
    function col(name) {
      for (var i = 0; i < head.length; i++) if (norm(head[i]) === norm(name)) return i;
      return -1;
    }

    var cSeq = col('序号'), cRow = col('原始行号'), cDate = col('凭证日期'), cVou = col('凭证号');
    var cAcct = col('科目'), cSum = col('摘要'), cAbs = col('金额（绝对值）');
    var cLayer = col('抽样类别'), cConc = col('核对结论'), cMis = col('错报金额（元）'), cNote = col('核对说明');

    if (cConc < 0) warnings.push('这份表没有「核对结论」列：请用 v0.2 之后生成的抽凭表回录，或手工加一列。');
    if (cMis < 0) warnings.push('这份表没有「错报金额（元）」列，错报一律按 0 计。');

    var records = [];
    var misUnreadable = [];       // 错报金额写了但读不出来：不能当 0，否则错误金额会凭空消失
    for (var i = headIdx + 1; i < rows.length; i++) {
      var row = rows[i] || [];
      var seqText = str(row[cSeq]).trim();
      if (seqText === '合计') break;
      if (!seqText) continue;
      var absParsed = cAbs >= 0 ? toCents(row[cAbs]) : { cents: null };
      var misParsed = cMis >= 0 ? toCents(row[cMis]) : { cents: null };
      var rawConc = cConc >= 0 ? str(row[cConc]).trim() : '';
      var hasMisText = cMis >= 0 && str(row[cMis]).trim() !== '';
      var misCents = misParsed.cents;
      if (hasMisText && misCents === null) {
        // 写了「约500元」这种：记 null 并报出来，绝不默默当 0
        misCents = null;
        misUnreadable.push(str(row[cSeq]).trim() + '（原值「' + str(row[cMis]).trim() + '」）');
      }
      if (misCents === null && !hasMisText) misCents = 0;
      records.push({
        seq: parseInt(seqText, 10),
        srcRow: cRow >= 0 ? str(row[cRow]).trim() : '',
        dateText: cDate >= 0 ? str(row[cDate]).trim() : '',
        voucher: cVou >= 0 ? str(row[cVou]).trim() : '',
        account: cAcct >= 0 ? str(row[cAcct]).trim() : '',
        summary: cSum >= 0 ? str(row[cSum]).trim() : '',
        absCents: absParsed.cents,
        layer: cLayer >= 0 ? str(row[cLayer]).trim() : '',
        rawConclusion: rawConc,
        conclusion: classifyConclusion(rawConc),
        misCents: misCents,
        misUnreadable: hasMisText && misParsed.cents === null,
        note: cNote >= 0 ? str(row[cNote]).trim() : ''
      });
    }

    if (misUnreadable.length) {
      warnings.push('这些行的「错报金额」读不出来，未计入评估：' + misUnreadable.slice(0, 5).join('、') +
        (misUnreadable.length > 5 ? ' 等 ' + misUnreadable.length + ' 行' : '') + '。请改成纯数字后重试。');
    }

    if (!records.length) warnings.push('表头下面没有读到数据行。');
    return { records: records, warnings: warnings, headerRowIndex: headIdx };
  }

  /** 从「抽样信息」页读回总体参数（重新导入时不用再手填一遍） */
  function parseInfoRows(rows) {
    var out = {};
    for (var i = 0; i < rows.length; i++) {
      var k = str((rows[i] || [])[0]).trim();
      var v = (rows[i] || [])[1];
      if (!k) continue;
      if (k === '总体笔数') out.popCount = parseInt(v, 10);
      else if (k === '总体金额（绝对值合计）') out.popCentsText = str(v);
      else if (k === '样本笔数') out.sampleCount = parseInt(v, 10);
      else if (k === '随机种子') out.seed = parseInt(v, 10);
      // 评估要知道样本是怎么抽的：等距/分层抽出来的不能套 MUS 上限
      else if (k === '抽样方法') out.sampleMethodLabel = str(v).trim();
      else if (k.indexOf('全查层笔数') === 0) out.fullCount = parseInt(str(v), 10) || 0;
      else if (k === '抽样单元') out.unit = str(v).trim();
    }
    if (out.popCentsText) out.popCents = toCents(out.popCentsText).cents;
    // 抽样方法的中文标签 → 内部代号
    var lbl = out.sampleMethodLabel || '';
    if (/金额加权/.test(lbl)) out.sampleMethod = 'pps';
    else if (/等距/.test(lbl)) out.sampleMethod = 'systematic';
    else if (/随机/.test(lbl)) out.sampleMethod = 'random';
    else out.sampleMethod = '';
    return out;
  }

  /**
   * 评估：计数 + 例外汇总 + MUS 上界。
   * @param records  parseVouchingRows 的结果
   * @param opts { alpha, tolerableCents, populationCents }
   */
  function evaluateVouching(records, opts) {
    opts = opts || {};
    var alpha = opts.alpha || 0.05;
    var counts = { ok: 0, exception: 0, missing: 0, na: 0, blank: 0, other: 0 };
    var others = [];

    records.forEach(function (r) {
      var c = r.conclusion;
      if (!c) counts.blank++;
      else if (counts[c] === undefined) counts.other++;
      else counts[c]++;
      if (c === 'other') others.push({ seq: r.seq, text: r.rawConclusion });
    });

    var examined = counts.ok + counts.exception + counts.missing;
    var exceptions = records.filter(function (r) { return r.conclusion === 'exception'; });
    var missings = records.filter(function (r) { return r.conclusion === 'missing'; });

    var grossOver = 0, grossUnder = 0;
    exceptions.forEach(function (r) {
      if (r.misCents > 0) grossOver += r.misCents;
      else if (r.misCents < 0) grossUnder += -r.misCents;
    });

    // MUS 的 taint：错报占该笔账面的比例。超过 100% 的按 100% 处理。
    var taints = [];
    exceptions.forEach(function (r) {
      if (!r.absCents) { taints.push(1); return; }
      taints.push(Math.min(1, Math.abs(r.misCents) / r.absCents));
    });
    // 未取得的按 100% 错报计入，这是保守做法，单独报出来让审计人员知道
    missings.forEach(function () { taints.push(1); });

    var um = null;
    var umBlocked = '';
    var popCents = opts.populationCents || 0;

    // 货币单元抽样（MUS）上限只对「按金额加权」的样本成立：
    // 等距/分层抽出来的样本，每笔被抽中的概率不相等，套 Stringer 公式算出来的数没有意义。
    // 宁可不给，也不能给一个看着像结论的数。
    var blockers = [];
    var method = opts.sampleMethod || '';
    if (!method) {
      blockers.push('读不到这份抽凭表的抽样方法（「抽样信息」页可能被删过），无法确认能否计算上限');
    } else if (method !== 'pps') {
      blockers.push('这份样本是按「' + (opts.sampleMethodLabel || method) +
        '」抽的，不是「按金额加权」，套不了货币单元抽样上限');
    }
    if (opts.fullCount > 0) {
      blockers.push('样本里有 ' + opts.fullCount + ' 笔是大额全查项，全查项不参与 MUS 上界计算');
    }
    if (counts.blank > 0) {
      blockers.push('还有 ' + counts.blank + ' 行没填核对结论');
    }
    var misBad = records.filter(function (r) { return r.misUnreadable; }).length;
    if (misBad > 0) {
      blockers.push('有 ' + misBad + ' 行「错报金额」读不出来');
    }

    if (blockers.length) {
      umBlocked = blockers.join('；') + '。';
    } else if (examined > 0 && popCents > 0) {
      um = stringerUpperBound(taints, examined, popCents, alpha);
    }

    // 判断抽样（非统计）的推断错报：比率法按金额比例，差异法按笔数比例。
    // 两种都给出来，用哪种由审计人员判断。
    var sampleBook = 0;
    records.forEach(function (r) {
      if (r.conclusion === 'na') return;
      if (r.absCents) sampleBook += Math.abs(r.absCents);
    });
    var sampleMisAbs = 0;
    exceptions.forEach(function (r) { sampleMisAbs += Math.abs(r.misCents || 0); });

    var projection = {
      sampleBookCents: sampleBook,
      sampleMisAbsCents: sampleMisAbs,
      ratioCents: sampleBook > 0 ? Math.round(sampleMisAbs / sampleBook * popCents) : 0,
      differenceCents: (examined > 0 && opts.populationCount)
        ? Math.round(sampleMisAbs / examined * opts.populationCount) : 0
    };

    return {
      counts: counts,
      examined: examined,
      others: others,
      exceptions: exceptions,
      missings: missings,
      grossOverCents: grossOver,
      grossUnderCents: grossUnder,
      netCents: grossOver - grossUnder,
      alpha: alpha,
      taintCount: taints.length,
      upper: um,
      umBlocked: umBlocked,
      sampleMethod: method,
      tolerableCents: opts.tolerableCents || null,
      populationCents: popCents,
      projection: projection,
      // 只给对照，不下结论
      verdictHint: (um && opts.tolerableCents)
        ? (um.upperCents > opts.tolerableCents ? '超过' : '未超过')
        : ''
    };
  }

  /** 评估表：把回录结果与错报上限排成一页 */
  function buildEvaluationSheet(ev, box) {
    // 金额一律写数值（元），显示格式交给 applyMoneyFormat；写成文本就没法求和、也贴不进底稿
    function yuan(c) { return (c === null || c === undefined) ? '' : c / 100; }
    var rows = [];
    var c = ev.counts;
    rows.push(['抽查结果评估', '', '索引号：' + ((box && box.indexNo) || '')]);
    rows.push(['工具版本', '抽凭助手 v' + VERSION]);
    rows.push([]);
    // 计数写成文本：否则会被金额格式套成 9.00，看着像金额
    rows.push(['样本笔数', String(c.ok + c.exception + c.missing + c.na + c.blank + c.other)]);
    rows.push(['已检查（参与评估）', String(ev.examined), '不含「不适用」与未填写的']);
    rows.push(['相符', String(c.ok)]);
    rows.push(['不符', String(c.exception)]);
    rows.push(['未取得', String(c.missing), '按 100% 错报计入上限（保守做法）']);
    rows.push(['不适用', String(c.na)]);
    rows.push(['尚未填写', String(c.blank), '这些行没有结论，不计入评估']);
    if (c.other) rows.push(['结论写法认不出', c.other, ev.others.map(function (o) { return '第' + o.seq + '行「' + o.text + '」'; }).join('；')]);
    rows.push([]);
    rows.push(['例外汇总']);
    rows.push(['例外笔数', String(ev.exceptions.length)]);
    rows.push(['高估金额（账面多计）', yuan(ev.grossOverCents)]);
    rows.push(['低估金额（账面少计）', yuan(ev.grossUnderCents)]);
    rows.push(['错报净额', yuan(ev.netCents)]);
    rows.push([]);

    if (ev.upper) {
      rows.push(['货币单元抽样错报上限']);
      rows.push(['置信水平', ((1 - ev.alpha) * 100).toFixed(0) + '%']);
      rows.push(['每笔平均账面金额', yuan(ev.upper.perItemCents)]);
      rows.push(['总体账面金额', yuan(ev.populationCents)]);
      rows.push(['上界系数', ev.upper.factorTotal.toFixed(4)]);
      rows.push(['错报上限', yuan(ev.upper.upperCents), '统计推断值，不是审计结论']);
      if (ev.upper.capped) {
        rows.push(['封顶说明', '按总体账面金额封顶',
          '未封顶的理论值 ' + moneyText(ev.upper.rawCents) + '。样本量相对总体太小时会出现这种情况，' +
          '错报不可能超过总体账面，故取总体账面为上限。']);
      }
      if (ev.tolerableCents) {
        rows.push(['可容忍错报', yuan(ev.tolerableCents)]);
        rows.push(['对照（仅供判断参考）', ev.verdictHint, '是否可接受，由审计人员判断']);
      }
    } else {
      rows.push(['货币单元抽样错报上限', '未计算', ev.umBlocked || '缺总体账面金额或没有已检查的样本']);
    }

    rows.push([]);
    rows.push(['判断抽样推断（非统计抽样用）']);
    rows.push(['样本账面合计', yuan(ev.projection.sampleBookCents)]);
    rows.push(['样本错报合计', yuan(ev.projection.sampleMisAbsCents)]);
    rows.push(['推断错报（比率法）', yuan(ev.projection.ratioCents), '样本错报 ÷ 样本账面 × 总体账面']);
    rows.push(['推断错报（差异法）', yuan(ev.projection.differenceCents), '样本错报 ÷ 样本笔数 × 总体笔数']);
    rows.push(['说明', '两种方法结果不同是正常的；用哪种、要不要再加程序，由审计人员判断。']);

    rows.push([]);
    rows.push(['说明']);
    rows.push(['1', '上界按 Stringer 方法计算：UM = (总体账面 / 样本量) × [R0 + Σ t_j × (Rj − Rj−1)]，' +
                    't_j 是第 j 笔错报占该笔账面的比例（降序），Rj 是 j 笔错报时的泊松上界系数。']);
    rows.push(['2', '「未取得」按 100% 错报计入上限。如果你认为应当单独处理，请调整结论列后重新评估。']);
    rows.push(['3', '上界超过可容忍错报，只说明统计上不能排除总体存在重大错报；能不能接受，要结合性质、原因、其他程序一起判断。']);
    rows.push(['4', '本页不代替判断。结论、评价、是否追加程序，由审计人员决定并签字。']);

    rows.push([]);
    rows.push(['例外明细']);
    rows.push(['序号', '原始行号', '凭证号', '科目', '账面金额', '错报金额', '错报率']);
    ev.exceptions.forEach(function (r) {
      var abs = r.absCents === null ? '' : r.absCents / 100;
      var mis = r.misCents / 100;
      var rate = r.absCents ? (Math.abs(r.misCents) / r.absCents) : 1;
      rows.push([r.seq, r.srcRow, r.voucher, r.account, abs, mis, (rate * 100).toFixed(2) + '%']);
    });
    if (!ev.exceptions.length) rows.push(['（无）']);

    return { name: '评估表', rows: rows, moneyCols: [1, 4, 5] };
  }

  /** 评估表文件名：抽查评估_客户_YYYYMMDD.xlsx */
  function evalFileName(box) {
    var parts = ['抽查评估'];
    if (box && box.client) parts.push(box.client);
    var d = new Date();
    parts.push('' + d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate()));
    return parts.join('_').replace(/[\\/:*?"<>|]/g, '_') + '.xlsx';
  }

  /** 评估表装配：与 buildWorkbook 同一套写法，测试和界面走同一条路 */
  function buildEvaluationWorkbook(XLSX, ev, box) {
    var sheet = buildEvaluationSheet(ev, box);
    var wb = XLSX.utils.book_new();
    var ws = XLSX.utils.aoa_to_sheet(sheet.rows);
    ws['!cols'] = colWidths(sheet.rows);
    if (sheet.moneyCols) applyMoneyFormat(XLSX, ws, sheet.rows.length, sheet.moneyCols);
    XLSX.utils.book_append_sheet(wb, ws, sheet.name);
    return { wb: wb, fileName: evalFileName(box), sheetNames: [sheet.name] };
  }

  /* ---------- 工作簿装配（需要传入 XLSX 库，便于在 node 里复用同一段代码） ---------- */

  function colWidths(rows) {
    var w = [];
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r] || [];
      for (var c = 0; c < row.length; c++) {
        var v = row[c];
        var s = (v === null || v === undefined) ? '' : String(v);
        var len = 0;
        for (var k = 0; k < s.length; k++) len += s.charCodeAt(k) > 255 ? 2 : 1;
        w[c] = Math.max(w[c] || 0, len);
      }
    }
    return w.map(function (x) { return { wch: Math.min(46, Math.max(8, (x || 0) + 2)) }; });
  }

  function applyMoneyFormat(XLSX, ws, rowCount, cols) {
    for (var r = 0; r < rowCount; r++) {
      for (var i = 0; i < cols.length; i++) {
        var cell = ws[XLSX.utils.encode_cell({ r: r, c: cols[i] })];
        if (cell && cell.t === 'n') cell.z = '#,##0.00';
      }
    }
  }

  /**
   * 装配成真正可下载的工作簿。
   * 界面层与测试都走这一个入口，避免「测的不是发的那个」。
   * @param XLSX  SheetJS 命名空间
   * @param box   buildSheets 需要的同一个数据包
   * @returns {{wb, fileName, sheetNames}}
   */
  function buildWorkbook(XLSX, box) {
    var sheets = buildSheets(box);
    var wb = XLSX.utils.book_new();
    var names = [];
    sheets.forEach(function (s) {
      var ws = XLSX.utils.aoa_to_sheet(s.rows);
      ws['!cols'] = colWidths(s.rows);
      if (s.moneyCols) applyMoneyFormat(XLSX, ws, s.rows.length, s.moneyCols);
      XLSX.utils.book_append_sheet(wb, ws, s.name);
      names.push(s.name);
    });
    return { wb: wb, fileName: fileName(box), sheetNames: names };
  }

  /* ============================================================
   * 导出
   * ============================================================ */

  return {
    VERSION: VERSION,
    // 基础
    norm: norm, toCents: toCents, moneyText: moneyText, toDate: toDate, dateText: dateText,
    dnum: dnum, mkDate: mkDate, pad2: pad2,
    // 识别与解析
    FIELDS: FIELDS, LABEL_OF: LABEL_OF, ALIASES: ALIASES,
    matchField: matchField, mapRow: mapRow, detectHeader: detectHeader, parseRows: parseRows,
    isSubtotalRow: isSubtotalRow, isBlankRow: isBlankRow, sheetVerdict: sheetVerdict,
    transposeRows: transposeRows, maybeTranspose: maybeTranspose,
    learnAlias: learnAlias, forgetAliases: forgetAliases,
    learnedAliases: learnedAliases, setLearnedAliases: setLearnedAliases,
    // 总体与抽样
    buildPopulation: buildPopulation, buildSample: buildSample,
    sampleSystematic: sampleSystematic, sampleRandom: sampleRandom, samplePPS: samplePPS, mulberry32: mulberry32,
    sumAbs: sumAbs, sumSigned: sumSigned, cmpForSampling: cmpForSampling,
    // 输出
    EVIDENCE_PRESETS: EVIDENCE_PRESETS, presetColumns: presetColumns,
    buildSheets: buildSheets, buildWorkbook: buildWorkbook, colWidths: colWidths,
    fileName: fileName,
    // 回录与评估
    poissonCdf: poissonCdf, poissonUpperFactor: poissonUpperFactor,
    stringerUpperBound: stringerUpperBound, classifyConclusion: classifyConclusion,
    parseVouchingRows: parseVouchingRows, parseInfoRows: parseInfoRows,
    evaluateVouching: evaluateVouching, buildEvaluationSheet: buildEvaluationSheet,
    buildEvaluationWorkbook: buildEvaluationWorkbook, evalFileName: evalFileName
  };
});
