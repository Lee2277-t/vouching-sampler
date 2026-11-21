/* 抽凭助手 · 界面层
 * 只做「读文件 → 填表 → 调 core → 下载」，所有金额与抽样逻辑都在 core.js。
 */
(function () {
  'use strict';

  var Core = window.AuditCore;
  var XLSX = window.XLSX;

  function $(id) { return document.getElementById(id); }

  function esc(s) {
    return String(s === null || s === undefined ? '' : s).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  }

  function pad2(n) { return String(n).padStart(2, '0'); }

  function money(cents) {
    return (cents === null || cents === undefined) ? '' : Core.moneyText(cents);
  }

  /* 界面上可调的字段（顺序即显示顺序） */
  var UI_FIELDS = ['date', 'voucher', 'voucherWord', 'summary', 'account', 'accountCode',
                   'entity', 'debit', 'credit', 'amount', 'direction'];

  var state = {
    books: [],        // [{fileName, wb, ok}]
    fileErrors: [],   // 读不出来的文件
    multiFile: false,
    sheetInfos: [],   // 所有文件的所有工作表：{key, bookIndex, fileName, sheetName, verdict, rows}
    sheets: [],       // 选中的工作表：{key, fileName, name, label, rows, firstExcelRow, header, map, verdict}
    dropped: [],      // 多选时因不可用被跳过的工作表
    multi: false,
    thresholdTouched: false,
    lastFileTag: '',
    parsed: null,
    pop: null,
    verdict: null
  };

  /* 记住用户教过的列名。只存表头文字，不存任何金额、摘要、客户名。
   * file:// 下个别浏览器禁 localStorage，存不了就算了，不影响主流程。 */
  var LEARN_KEY = 'chouping.learned.v1';

  function loadLearned() {
    try {
      var raw = window.localStorage.getItem(LEARN_KEY);
      if (raw) Core.setLearnedAliases(JSON.parse(raw));
    } catch (e) { /* 忽略 */ }
  }

  function saveLearned() {
    try {
      window.localStorage.setItem(LEARN_KEY, JSON.stringify(Core.learnedAliases()));
    } catch (e) { /* 忽略 */ }
  }

  function renderLearnHint() {
    var n = Object.keys(Core.learnedAliases()).length;
    var el = $('learnHint');
    if (el) {
      el.innerHTML = n
        ? '已记住 <b>' + n + '</b> 个列名（只记表头文字，不记数据）。'
        : '还没有记住任何列名。在上面改一次列，下次遇到同样的写法就会自动认。';
    }
    var btn = $('forgetBtn');
    if (btn) btn.classList.toggle('hidden', !n);
  }

  $('forgetBtn').addEventListener('click', function () {
    Core.forgetAliases();
    saveLearned();
    renderLearnHint();
  });

  loadLearned();

  /* ============================================================
   * 简易 / 高级
   * 默认简易：只留 文件、科目、期间、样本量。其余全收进「高级」。
   * ============================================================ */

  var MODE_KEY = 'chouping.mode.v1';

  function isSimple() {
    return document.body.classList.contains('mode-simple');
  }

  function setMode(simple) {
    document.body.classList.toggle('mode-simple', simple);
    $('modeSimple').classList.toggle('on', simple);
    $('modeExpert').classList.toggle('on', !simple);
    $('modeHint').textContent = simple
      ? '选文件 → 填科目 → 点生成。其余都不用管。'
      : '所有参数都露出来了。改哪项都会立刻重算。';
    try { window.localStorage.setItem(MODE_KEY, simple ? 'simple' : 'expert'); } catch (e) { /* 忽略 */ }
    renderSimpleSummary();
  }

  $('modeSimple').addEventListener('click', function () { setMode(true); });
  $('modeExpert').addEventListener('click', function () { setMode(false); });

  /** 简易模式下那一句话：认了多少笔、多少钱、什么期间 */
  function renderSimpleSummary() {
    var el = $('simpleSummary');
    if (!el) return;

    if (!isSimple()) {
      el.classList.add('hidden');
      return;
    }
    el.classList.remove('hidden');

    if (!state.books.length) {
      el.className = 'note ok';
      el.innerHTML = '第一次用：点上面「选择 Excel 文件」导入明细账；手上没有文件的话，点「用示例数据试一下」。';
      return;
    }

    if (state.verdict && !state.verdict.usable) {
      el.className = 'note bad';
      el.innerHTML = '<b>这个文件读不出明细账</b>：' + esc(state.verdict.problems.join('；')) +
        '。如果是科目余额表或报表，请让客户另导一份序时账／明细账。';
      return;
    }

    var pop = state.pop;
    if (!pop) return;

    if (!pop.stats.count) {
      el.className = 'note warn';
      el.innerHTML = '按当前条件没有符合条件的记录（0 笔）。清空「科目关键词」或放宽期间再试。';
      return;
    }

    var labels = state.sheets.map(function (s) {
      return state.multiFile ? s.label : (s.fileName + ' · ' + s.name);
    });
    var src = labels.slice(0, 2).join('、') + (labels.length > 2 ? ' 等 ' + labels.length + ' 张' : '');

    el.className = 'note ok';
    el.innerHTML = '已认出 <b>' + pop.stats.count + '</b> 笔' +
      (src ? '（' + esc(src) + '）' : '') +
      '，合计 <b>' + money(pop.stats.totalAbsCents) + '</b>' +
      (state.dateSpan ? '，日期 ' + esc(state.dateSpan) : '') +
      '。没问题就点下面第 3 步的「生成并下载抽凭表」。';
  }

  /** 没有文件也能跑通一遍：造 60 笔假账，走完全相同的识别与抽样路径 */
  $('demoBtn').addEventListener('click', function () {
    var rows = [['日期', '凭证号', '摘要', '科目', '借方金额', '贷方金额']];
    for (var i = 1; i <= 60; i++) {
      var dt = new Date(Date.UTC(2026, 0, 1) + (i - 1) * 2 * 86400000);
      var amt = 500 + ((i * 7919) % 40000);
      var isCredit = i % 5 === 0;
      rows.push([
        dt.getUTCFullYear() + '-' + pad2(dt.getUTCMonth() + 1) + '-' + pad2(dt.getUTCDate()),
        '记-' + String(i).padStart(4, '0'),
        '示例业务' + i + '（假数据）',
        isCredit ? '6001 主营业务收入' : '1122 应收账款',
        isCredit ? null : amt,
        isCredit ? amt : null
      ]);
    }
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), '示例序时账');

    state.books = [{ fileName: '示例数据.xlsx', wb: wb, ok: true }];
    state.fileErrors = [];
    state.multiFile = false;
    $('accountText').value = '应收账款';
    analyzeBooks();
    renderSheetOptions();
  });

  /* ============================================================
   * 读文件
   * ============================================================ */

  /** 按「表的走向」决定要不要转置：自动就是横排读不出来时试竖排 */
  function pickOrientation(rows) {
    var mode = $('orientSelect').value;
    if (mode === 'normal') {
      var h = Core.detectHeader(rows);
      return { rows: rows, header: h, verdict: Core.sheetVerdict(h), transposed: false };
    }
    return Core.maybeTranspose(rows, mode === 'transposed');
  }

  function decodeText(buf) {
    try { return new TextDecoder('utf-8', { fatal: true }).decode(buf); }
    catch (e) {
      try { return new TextDecoder('gbk').decode(buf); }
      catch (e2) { return new TextDecoder('utf-8').decode(buf); }
    }
  }

  /** 把一个 File 读成工作簿。失败不抛，返回 ok:false 交给上层汇总。 */
  function readWorkbook(file) {
    return new Promise(function (resolve) {
      var reader = new FileReader();
      reader.onload = function () {
        try {
          var buf = new Uint8Array(reader.result);
          // cellDates:false → 日期读成 Excel 序列号，由 core 统一解析，避开时区偏移
          var wb = /\.csv$/i.test(file.name)
            ? XLSX.read(decodeText(buf), { type: 'string', cellDates: false })
            : XLSX.read(buf, { type: 'array', cellDates: false });
          resolve({ fileName: file.name, wb: wb, ok: true });
        } catch (err) {
          resolve({ fileName: file.name, ok: false, error: err.message });
        }
      };
      reader.onerror = function () {
        resolve({ fileName: file.name, ok: false, error: '文件读不出来' });
      };
      reader.readAsArrayBuffer(file);
    });
  }

  $('fileInput').addEventListener('change', function (e) {
    var files = Array.prototype.slice.call(e.target.files || []);
    if (!files.length) return;

    $('fileInfo').textContent = '正在读取 ' + files.length + ' 个文件…';

    Promise.all(files.map(readWorkbook)).then(function (books) {
      var good = books.filter(function (b) { return b.ok; });
      var bad = books.filter(function (b) { return !b.ok; });

      if (!good.length) {
        $('fileInfo').textContent = '这些文件都读不了：' +
          bad.map(function (b) { return b.fileName + '（' + b.error + '）'; }).join('；');
        return;
      }

      state.books = good;
      state.fileErrors = bad;
      state.multiFile = good.length > 1;
      analyzeBooks();
      renderSheetOptions();
    });
  });

  $('sheetSelect').addEventListener('change', function () {
    loadSelected();
  });

  $('orientSelect').addEventListener('change', function () {
    if (!state.books.length) return;
    analyzeBooks();
    renderSheetOptions();
  });

  /** 只读前 n 行做判定，别为了挑工作表把几十万行全读一遍 */
  function previewRows(ws, n) {
    var ref = ws && ws['!ref'];
    if (!ref) return [];
    var r = XLSX.utils.decode_range(ref);
    return XLSX.utils.sheet_to_json(ws, {
      header: 1, raw: true, defval: null, blankrows: true,
      range: { s: { r: r.s.r, c: r.s.c }, e: { r: Math.min(r.e.r, r.s.r + n - 1), c: r.e.c } }
    });
  }

  /** 每本书的每张表各打一次分：能不能当明细账用，为什么不能 */
  function analyzeBooks() {
    state.sheetInfos = [];
    state.books.forEach(function (b, bi) {
      (b.wb.SheetNames || []).forEach(function (name) {
        var ws = b.wb.Sheets[name];
        var picked = pickOrientation(previewRows(ws, 30));
        var rows = 0;
        if (ws && ws['!ref']) {
          var rr = XLSX.utils.decode_range(ws['!ref']);
          rows = rr.e.r - rr.s.r + 1;
        }
        state.sheetInfos.push({
          key: bi + '\u0001' + name,
          bookIndex: bi,
          fileName: b.fileName,
          sheetName: name,
          verdict: picked.verdict,
          transposed: picked.transposed,
          rows: rows
        });
      });
    });
  }

  /** 把所有文件的所有工作表铺进下拉，并自动选中得分最高的那张 */
  function renderSheetOptions() {
    // 换文件才重置阈值；只换工作表不该把用户手填的阈值冲掉
    var fileTag = state.books.map(function (b) { return b.fileName; }).join('|');
    if (fileTag !== state.lastFileTag) {
      state.thresholdTouched = false;
      state.lastFileTag = fileTag;
    }
    if (!state.sheetInfos.length) {
      $('fileInfo').textContent = '这些文件里没有工作表。';
      $('loadPanel').classList.add('hidden');
      return;
    }

    var multiFile = state.books.length > 1;

    $('sheetSelect').innerHTML = state.sheetInfos.map(function (i) {
      var mark = i.verdict.usable ? '✓' : '✗';
      var why = i.verdict.problems.length ? ' · ' + i.verdict.problems.join('，') : '';
      var turn = i.transposed ? ' · 竖排已转置' : '';
      var label = (multiFile ? i.fileName + ' · ' : '') + i.sheetName + ' · ' + i.rows + ' 行' + turn + why;
      return '<option value="' + esc(i.key) + '">' + mark + ' ' + esc(label) + '</option>';
    }).join('');
    $('loadPanel').classList.remove('hidden');

    var anyUsable = state.sheetInfos.some(function (i) { return i.verdict.usable; });
    var head = multiFile
      ? '<b>' + state.books.length + ' 个文件</b> · 共 ' + state.sheetInfos.length + ' 个工作表'
      : '<b>' + esc(state.books[0].fileName) + '</b> · 共 ' + state.sheetInfos.length + ' 个工作表';
    if (!anyUsable) head += ' · <b>没有找到可用的明细账</b>';
    if (state.fileErrors.length) {
      head += ' · <b>' + state.fileErrors.length + ' 个文件读不了</b>（' +
        state.fileErrors.map(function (b) { return esc(b.fileName); }).join('、') + '）';
    }
    $('fileInfo').innerHTML = head;

    // 自动选最像明细账的那张，而不是第一个文件的第一张
    var best = state.sheetInfos.filter(function (i) { return i.verdict.usable; })
      .sort(function (a, b) { return b.verdict.score - a.verdict.score; })[0] || state.sheetInfos[0];
    Array.prototype.forEach.call($('sheetSelect').options, function (o) {
      o.selected = (o.value === best.key);
    });
    loadSelected();
  }

  /** 读取选中的工作表（可多张）。每张各自识别表头，不可用的直接跳过并说明。 */
  function loadSelected() {
    var pick = [];
    Array.prototype.forEach.call($('sheetSelect').options, function (o) {
      if (o.selected) pick.push(o.value);
    });
    if (!pick.length) {
      var first = $('sheetSelect').options[0];
      if (!first) return;
      first.selected = true;
      pick = [first.value];
    }

    var loaded = pick.map(function (key) {
      var cut = key.indexOf('\u0001');
      var bi = parseInt(key.slice(0, cut), 10);
      var sheetName = key.slice(cut + 1);
      var book = state.books[bi];
      var ws = book.wb.Sheets[sheetName];
      var ref = (ws && ws['!ref']) || 'A1';
      var range = XLSX.utils.decode_range(ref);
      var rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
      var picked = pickOrientation(rawRows);
      var rows = picked.rows;
      var h = picked.header;
      var info = state.sheetInfos.filter(function (i) { return i.key === key; })[0];
      return {
        key: key,
        fileName: book.fileName,
        name: sheetName,
        label: (state.multiFile ? book.fileName + ' · ' : '') + sheetName,
        rows: rows,
        transposed: picked.transposed,
        firstExcelRow: picked.transposed ? 1 : range.s.r + 1,
        header: { rowIndex: h.rowIndex >= 0 ? h.rowIndex : 0, span: h.span || 1, map: h.map, labels: h.labels },
        map: h.map,
        verdict: info ? info.verdict : picked.verdict
      };
    });

    state.multi = loaded.length > 1;
    state.dropped = [];

    if (state.multi) {
      // 多选时把不可用的表剔掉：选了一张余额表进来，不该拖垮整批
      var usable = loaded.filter(function (s) { return s.verdict.usable; });
      loaded.forEach(function (s) {
        if (!s.verdict.usable) state.dropped.push(s.label + '（' + s.verdict.problems.join('，') + '）');
      });
      state.sheets = usable.length ? usable : loaded;
      state.verdict = usable.length
        ? { usable: true, problems: [], hasDate: true, hasAmount: true, hasVoucher: true, score: 1 }
        : { usable: false, problems: loaded[0].verdict.problems, hasDate: false, hasAmount: false, hasVoucher: false, score: 0 };
    } else {
      state.sheets = loaded;
      state.verdict = loaded[0].verdict;
    }

    $('headerRowInput').value = state.sheets[0].header.rowIndex + state.sheets[0].firstExcelRow;
    renderMap();
    refresh();
  }

  $('redetectBtn').addEventListener('click', function () {
    if (!state.sheets.length) return;
    var s = state.sheets[0];
    var v = parseInt($('headerRowInput').value, 10);
    if (!isFinite(v) || v < 1) return;
    var idx = v - s.firstExcelRow;
    if (idx < 0 || idx >= s.rows.length) {
      $('fileInfo').textContent = '表头行号超出这张表的范围。';
      return;
    }
    s.header.rowIndex = idx;
    s.header.span = 1;
    s.map = Core.mapRow(s.rows[idx] || []).map;
    renderMap();
    refresh();
  });

  /* ============================================================
   * 列识别
   * ============================================================ */

  function renderMap() {
    if (!state.sheets.length) return;

    if (state.multi) {
      var html = '<thead><tr><th>工作表</th><th>数据行数</th><th>识别到的列</th></tr></thead><tbody>';
      state.sheets.forEach(function (s) {
        var names = UI_FIELDS.filter(function (f) { return s.map[f] !== undefined; })
          .map(function (f) { return Core.LABEL_OF[f]; });
        var dataRows = Math.max(0, s.rows.length - s.header.rowIndex - (s.header.span || 1));
        html += '<tr><td>' + esc(s.label) + '</td>' +
          '<td class="num">' + dataRows + '</td>' +
          '<td class="sample">' + esc(names.join('、') || '（未识别）') + '</td></tr>';
      });
      html += '</tbody>';
      $('mapTable').innerHTML = html;
      $('mapPanel').classList.remove('hidden');
      $('headerRowInput').disabled = true;
      $('redetectBtn').disabled = true;
      return;
    }

    $('headerRowInput').disabled = false;
    $('redetectBtn').disabled = false;

    var s = state.sheets[0];
    var rawHead = s.rows[s.header.rowIndex] || [];
    var labels = (s.header.labels && s.header.labels.length)
      ? s.header.labels
      : rawHead.map(function (c) { return String(c === null || c === undefined ? '' : c).trim(); });

    function option(value, label, selected) {
      return '<option value="' + value + '"' + (selected ? ' selected' : '') + '>' + label + '</option>';
    }

    var html = '<thead><tr><th>含义</th><th>对应列</th><th>该列第一个数据</th></tr></thead><tbody>';
    UI_FIELDS.forEach(function (f) {
      var cur = s.map[f];
      var cols = option('', '（不映射）', cur === undefined) +
        labels.map(function (t, i) {
          return option(i, '第 ' + (i + 1) + ' 列：' + esc(t || '（空表头）'), cur === i);
        }).join('');

      html += '<tr class="' + (cur === undefined ? 'off' : '') + '">' +
        '<td>' + esc(Core.LABEL_OF[f]) + '</td>' +
        '<td><select data-field="' + f + '">' + cols + '</select></td>' +
        '<td class="sample">' + esc(sampleOfColumn(s, cur)) + '</td>' +
      '</tr>';
    });
    html += '</tbody>';

    $('mapTable').innerHTML = html;
    $('mapPanel').classList.remove('hidden');
    renderLearnHint();

    Array.prototype.forEach.call($('mapTable').querySelectorAll('select'), function (sel) {
      sel.addEventListener('change', function () {
        var f = this.getAttribute('data-field');
        if (this.value === '') {
          delete s.map[f];
        } else {
          var col = parseInt(this.value, 10);
          s.map[f] = col;
          // 手工指的列名记下来，下次同样的写法自动认（只记表头文字）
          var label = labels[col] || (s.rows[s.header.rowIndex] || [])[col];
          if (Core.learnAlias(label, f)) saveLearned();
        }
        renderLearnHint();
        renderMap();
        refresh();
      });
    });
  }

  function sampleOfColumn(s, c) {
    if (c === undefined) return '—';
    var start = s.header.rowIndex + (s.header.span || 1);
    for (var r = start; r < s.rows.length; r++) {
      var v = (s.rows[r] || [])[c];
      if (v !== null && v !== undefined && String(v).trim() !== '') return String(v);
    }
    return '（整列为空）';
  }

  /* ============================================================
   * 筛选 / 抽样参数
   * ============================================================ */

  function yuanToCents(v) {
    var s = String(v || '').trim();
    if (!s) return null;
    var r = Core.toCents(s);
    return r.cents;
  }

  /** 填了但读不出来的数字（如「10万」「五千」）：这些不能当成「没填」 */
  function badNumberFields() {
    var bad = [];
    [['thresholdYuan', '大额全查阈值'], ['minYuan', '金额下限'], ['maxYuan', '金额上限'],
     ['tolerableYuan', '可容忍错报'], ['popYuan', '总体账面金额']].forEach(function (p) {
      var el = $(p[0]);
      if (!el) return;
      var raw = String(el.value || '').trim();
      if (raw && Core.toCents(raw).cents === null) bad.push(p[1] + '「' + raw + '」');
    });
    return bad;
  }

  function readFilter() {
    return {
      dateFrom: Core.toDate($('dateFrom').value),
      dateTo: Core.toDate($('dateTo').value),
      accountText: $('accountText').value.trim(),
      entity: $('entitySelect').value,
      direction: $('direction').value,
      unit: $('unitSelect').value,
      minCents: yuanToCents($('minYuan').value),
      maxCents: yuanToCents($('maxYuan').value)
    };
  }

  function readParams() {
    return {
      thresholdCents: yuanToCents($('thresholdYuan').value) || 0,
      n: Math.max(0, parseInt($('sampleN').value, 10) || 0),
      seed: (parseInt($('seedInput').value, 10) || 20260101) >>> 0,
      method: $('methodSelect').value
    };
  }

  /* ============================================================
   * 刷新
   * ============================================================ */

  /** 大额全查阈值：给个按总体自动算的建议值（总体金额 ÷ 抽几笔） */
  function suggestThresholdCents(pop) {
    var n = parseInt($('sampleN').value, 10) || 25;
    if (!pop || !pop.stats.count || n <= 0) return 0;
    var yuan = pop.stats.totalAbsCents / 100 / n;
    if (yuan <= 0) return 0;
    var mag = Math.pow(10, Math.floor(Math.log10(yuan)));
    var step = mag / 10;
    return Math.round(yuan / step) * step * 100;
  }

  /** 「金额加权」不需要阈值，选了就把它藏起来，免得两套口径打架 */
  function updateMethodUI() {
    var pps = $('methodSelect').value === 'pps';
    var f = $('thresholdField');
    if (f) f.classList.toggle('hidden', pps);
  }

  function refresh() {
    if (!state.sheets.length) return;

    var records = [], problems = [], skipped = [];
    var warnings = { extraDecimals: 0, blankRows: 0, subtotalRows: 0, emptyAmountRows: 0 };
    var anyVoucher = false;

    state.sheets.forEach(function (s) {
      var parsed = Core.parseRows(
        s.rows,
        { rowIndex: s.header.rowIndex, map: s.map, span: s.header.span || 1 },
        { map: s.map, firstExcelRow: s.firstExcelRow }
      );
      parsed.records.forEach(function (r) { r.sheet = s.label; });
      records = records.concat(parsed.records);
      problems = problems.concat(parsed.problems);
      skipped = skipped.concat(parsed.skipped);
      warnings.extraDecimals += parsed.warnings.extraDecimals || 0;
      warnings.blankRows += parsed.warnings.blankRows || 0;
      warnings.subtotalRows += parsed.warnings.subtotalRows || 0;
      warnings.emptyAmountRows += parsed.warnings.emptyAmountRows || 0;
      if (parsed.records.some(function (r) { return !!r.voucher; })) anyVoucher = true;
    });

    warnings.noVoucher = records.length > 0 && !anyVoucher;

    // 数据自己带出来的期间，用于简易模式那句话
    var lo = null, hi = null;
    records.forEach(function (r) {
      if (!r.date) return;
      var k = Core.dnum(r.date);
      if (!lo || k < lo.k) lo = { k: k, t: r.dateText };
      if (!hi || k > hi.k) hi = { k: k, t: r.dateText };
    });
    state.dateSpan = (lo && hi) ? (lo.t + ' ~ ' + hi.t) : '';

    state.parsed = { records: records, problems: problems, skipped: skipped, warnings: warnings };
    state.pop = Core.buildPopulation(records, readFilter());

    // 阈值没被手工改过，就按总体自动建议一个（改过之后不再覆盖）
    if (!state.thresholdTouched) {
      var sug = suggestThresholdCents(state.pop);
      $('thresholdYuan').value = String(Math.round(sug / 100));
    }
    updateMethodUI();

    // 用户手工改了表头行/列识别之后，可用性要用「现在的映射」重算，
    // 否则程序自己给的出路走不通：统计已经对了，红条还拦着不让生成
    if (state.sheets.length === 1) {
      var s0 = state.sheets[0];
      state.verdict = Core.sheetVerdict({ map: s0.map, labels: s0.header.labels });
    }

    renderEntitySelect();
    renderPreview();
    renderStats();
    renderWarnings();
    renderSimpleSummary();
  }

  /** 主体下拉：从数据里现取，没有主体列就禁用 */
  function renderEntitySelect() {
    var sel = $('entitySelect');
    var counts = {};
    var total = 0;
    (state.parsed && state.parsed.records || []).forEach(function (r) {
      if (!r.entity) return;
      total++;
      counts[r.entity] = (counts[r.entity] || 0) + 1;
    });

    var keys = Object.keys(counts).sort();
    var cur = sel.value;

    if (!keys.length) {
      sel.innerHTML = '<option value="">全部主体</option>';
      sel.disabled = true;
      sel.value = '';
      return;
    }

    sel.disabled = false;
    sel.innerHTML = '<option value="">全部主体（' + total + ' 行）</option>' +
      keys.map(function (k) {
        return '<option value="' + esc(k) + '">' + esc(k) + '（' + counts[k] + ' 行）</option>';
      }).join('');
    if (keys.indexOf(cur) >= 0) sel.value = cur;
  }

  function renderPreview() {
    var recs = (state.parsed.records || []).slice(0, 5);
    var cols = state.multi ? 8 : 7;
    var html = '<thead><tr>' + (state.multi ? '<th>工作表</th>' : '') +
               '<th>Excel 行</th><th>日期</th><th>凭证号</th><th>科目</th><th>摘要</th>' +
               '<th>借方</th><th>贷方</th></tr></thead><tbody>';
    if (!recs.length) {
      html += '<tr><td colspan="' + cols + '" class="muted">没有解析出明细行。</td></tr>';
    }
    recs.forEach(function (r) {
      html += '<tr>' +
        (state.multi ? '<td>' + esc(r.sheet || '') + '</td>' : '') +
        '<td class="num">' + r.srcRow + '</td>' +
        '<td>' + esc(r.dateText) + '</td>' +
        '<td>' + esc(r.voucher) + '</td>' +
        '<td>' + esc(r.account) + '</td>' +
        '<td>' + esc(r.summary) + '</td>' +
        '<td class="num">' + money(r.debitCents) + '</td>' +
        '<td class="num">' + money(r.creditCents) + '</td>' +
      '</tr>';
    });
    html += '</tbody>';

    $('previewTable').innerHTML = html;
    $('previewPanel').classList.remove('hidden');
  }

  function renderStats() {
    var pop = state.pop;
    var box = $('statsBox');

    if (!pop || !state.parsed) {
      box.innerHTML = '<div class="empty">导入明细账后，这里会实时显示总体、分层与预计样本。</div>';
      return;
    }

    var params = readParams();
    var smp = Core.buildSample(pop.included, params);
    var m = smp.meta;
    var u = pop.unit === 'voucher' ? '张' : '笔';

    var exclCount = {};
    pop.excluded.forEach(function (e) { exclCount[e.reason] = (exclCount[e.reason] || 0) + 1; });
    var exclText = Object.keys(exclCount).map(function (k) {
      return k + ' ' + exclCount[k] + ' 笔';
    }).join('、') || '无';

    box.innerHTML =
      '<div class="head">当前结果（改动上面任何一项都会立刻重算）</div>' +
      '<div class="body">' +
        cell('总体' + (pop.unit === 'voucher' ? '凭证数' : '笔数'), pop.stats.count, u, 'accent') +
        cell('总体金额', money(pop.stats.totalAbsCents), '', '') +
        (m.method === 'pps'
          ? cell('金额间隔', money(Math.round(m.intervalCents || 0)), '单笔超过必中', 'warn')
          : cell('大额全查', m.fullCount, u + ' / ' + money(m.fullAbsCents), 'warn')) +
        (m.method === 'pps'
          ? cell('实抽笔数', m.sampleCount, '大额占多个间隔', '')
          : cell('其余可抽', m.restCount, u + ' / ' + money(m.restAbsCents), '')) +
      '</div>' +
      '<div class="body">' +
        cell('预计样本', m.sampleCount, u, 'ok') +
        cell('覆盖率', (m.coverage * 100).toFixed(2) + '%', '', 'ok') +
        cell('样本金额', money(m.sampleAbsCents), '', '') +
        cell('排除记录', pop.stats.excludedCount, u, pop.stats.excludedCount ? 'bad' : '') +
      '</div>' +
      '<div class="foot">' +
        '排除原因：' + esc(exclText) +
        ' ｜ 涉及凭证 <b>' + m.voucherCount + '</b> 张' +
        (m.voucherCount && m.voucherCount < m.sampleCount
          ? '（有 ' + (m.sampleCount - m.voucherCount) + ' 行与别的样本同凭证，现场只需翻一次）'
          : '') +
        (params.method === 'systematic' && m.k
          ? ' ｜ 等距：间距 <b>' + m.k + '</b>，起点 <b>' + m.start + '</b>（复核时照这个数对）'
          : '') +
        ' ｜ 净额 ' + money(pop.stats.netCents) +
      '</div>';

    if (pop.stats.count === 0) {
      box.innerHTML += '<div class="note bad" style="margin:12px 16px 14px">总体是 0 笔：先看下面的「数据检查」，' +
        '常见原因是日期列没认对，或期间条件把数据全排除了。</div>';
    }
  }

  function cell(k, v, unit, cls) {
    return '<div class="cell ' + (cls || '') + '"><div class="k">' + esc(k) + '</div>' +
      '<div class="v">' + esc(v) + (unit ? '<small>' + esc(unit) + '</small>' : '') + '</div></div>';
  }

  function renderWarnings() {
    var p = state.parsed;
    var box = $('warnBox');
    var items = [];
    var level = '';

    if (!p) { box.classList.add('hidden'); return; }

    if (state.verdict && !state.verdict.usable) {
      items.push('<b>这张表不能当抽凭总体</b>：' + esc(state.verdict.problems.join('；')) +
        '。在「工作表」里选带 ✓ 的那张；如果整个文件都没 ✓，说明客户给的是科目余额表或报表，' +
        '需要请他另导一份明细账 / 序时账。');
      level = 'bad';
    }

    var hasAmountCol = state.sheets.some(function (s) {
      return s.map.debit !== undefined || s.map.credit !== undefined || s.map.amount !== undefined;
    });
    if (!hasAmountCol) {
      items.push('<b>没有识别到金额列</b>：在「列识别」里手工指定借方/贷方（或单列金额）后再看结果。');
      level = 'bad';
    }
    if (state.sheets.every(function (s) { return s.map.date === undefined; })) {
      items.push('<b>没有识别到日期列</b>：一旦设了期间，所有记录都会因「日期缺失」被排除。');
      level = level || 'warn';
    }
    if (state.dropped && state.dropped.length) {
      items.push('多选时跳过了不可用的工作表：' + esc(state.dropped.join('；')));
      level = level || 'warn';
    }
    if (p.warnings.noVoucher) {
      items.push('<b>没有识别到凭证号列</b>：抽凭表会空白，现场只能按「原始行号」回原表翻。' +
        '如果表里是「字」+「号」两列，在下面「列识别」里把它们指给「凭证号」和「凭证字」。');
      level = level || 'warn';
    }
    if (p.problems.length) {
      items.push('<b>' + p.problems.length + ' 行的金额读不出来</b>（不在总体也不在排除清单，需回 Excel 核对）：' +
        p.problems.slice(0, 3).map(function (x) { return '第 ' + x.srcRow + ' 行「' + esc(x.text) + '」'; }).join('；') +
        (p.problems.length > 3 ? ' 等' : ''));
      level = 'bad';
    }
    if (p.warnings.extraDecimals) {
      items.push(p.warnings.extraDecimals + ' 行金额超过两位小数，已按四舍五入到分。建议回原始数据确认。');
      level = level || 'warn';
    }
    if (p.warnings.blankRows || p.warnings.subtotalRows || p.warnings.emptyAmountRows) {
      var bits = [];
      if (p.warnings.blankRows) bits.push('空行 ' + p.warnings.blankRows);
      if (p.warnings.subtotalRows) bits.push('小计/合计行 ' + p.warnings.subtotalRows);
      if (p.warnings.emptyAmountRows) bits.push('无金额行 ' + p.warnings.emptyAmountRows);
      items.push('跳过的非明细行：' + bits.join('、') + ' 行（明细行 ' + p.records.length + ' 条）。明细清单见导出文件「数据检查」页。');
    }

    if (!items.length) {
      box.className = 'note ok';
      box.innerHTML = '<b>数据检查通过</b>：' + p.records.length + ' 条明细行全部解析成功，没有读不出来的金额。';
      box.classList.remove('hidden');
      return;
    }

    box.className = 'note ' + (level || 'warn');
    box.innerHTML = '<b>先看这里</b><ul>' + items.map(function (t) { return '<li>' + t + '</li>'; }).join('') + '</ul>';
    box.classList.remove('hidden');
  }

  /* ============================================================
   * 生成
   * ============================================================ */

  function mapNotes() {
    var notes = [];
    state.sheets.forEach(function (s) {
      UI_FIELDS.forEach(function (f) {
        var c = s.map[f];
        if (c === undefined) return;
        var label = (s.header.labels && s.header.labels[c]) || '';
        notes.push((state.multi ? s.label + ' · ' : '') + Core.LABEL_OF[f] +
          ' → 第 ' + (c + 1) + ' 列「' + label + '」');
      });
      notes.push('工作表：' + s.label + '；表头位于第 ' + (s.header.rowIndex + s.firstExcelRow) + ' 行' +
        (s.header.span > 1 ? '（两级表头，占 ' + s.header.span + ' 行）' : ''));
    });
    return notes;
  }

  function setResult(kind, html) {
    $('resultBox').innerHTML = '<div class="note ' + kind + '">' + html + '</div>';
  }

  $('genBtn').addEventListener('click', function () {
    // 宁可不出，也不能出一份假样本
    var badNums = badNumberFields();
    if (badNums.length) {
      setResult('bad', '这些不是能识别的金额：' + esc(badNums.join('、')) +
        '。请填纯数字（如 100000），不要写「10万」「五千」。条件读不出来时程序不会当成「没填」，以免抽错范围。');
      return;
    }
    if (state.verdict && !state.verdict.usable) {
      setResult('bad', '这张表不能当抽凭总体（' + esc(state.verdict.problems.join('；')) +
        '）。换一张工作表，或请客户另导明细账 / 序时账。');
      return;
    }
    if (!state.pop || !state.pop.stats.count) {
      setResult('bad', '总体是空的，没法抽。先导入明细账，或者放宽上面的期间/科目/金额条件。');
      return;
    }

    var params = readParams();
    var smp = Core.buildSample(state.pop.included, params);

    if (!smp.meta.sampleCount) {
      setResult('bad', '样本量是 0：把「抽样层样本量」调大，或者把全查阈值调低。');
      return;
    }

    var today = new Date();
    var box = {
      client: $('client').value.trim(),
      indexNo: $('indexNo').value.trim(),
      preparer: $('preparer').value.trim(),
      reviewer: $('reviewer').value.trim(),
      dateText: today.getFullYear() + '-' + pad2(today.getMonth() + 1) + '-' + pad2(today.getDate()),
      periodText: periodText(),
      accountText: $('accountText').value.trim(),
      presetName: $('presetSelect').value,
      sample: smp.sample,
      meta: Object.assign({}, smp.meta, { unit: state.pop.unit, multiSheet: state.multi }),
      sheetNames: state.sheets.map(function (s) { return s.label; }),
      transposed: state.sheets.some(function (s) { return s.transposed; }),
      populationStats: state.pop.stats,
      excluded: state.pop.excluded,
      problems: state.parsed.problems,
      skipped: state.parsed.skipped,
      warnings: Object.assign({ parsedRows: state.parsed.records.length }, state.parsed.warnings),
      mapNotes: mapNotes(),
      filter: readFilter(),
      headerRow: state.sheets.length === 1 ? (state.sheets[0].firstExcelRow + state.sheets[0].header.rowIndex) : null
    };

    var name = Core.fileName(box);
    var u = state.pop.unit === 'voucher' ? '张' : '笔';

    try {
      // 装配与测试走同一个入口，避免「测的不是发的那个」
      var out = Core.buildWorkbook(XLSX, box);
      XLSX.writeFile(out.wb, out.fileName);
    } catch (err) {
      setResult('bad', '生成失败：' + esc(err.message) + '。把这条信息发给做工具的人。');
      return;
    }

    setResult('ok',
      '<b>已生成</b>：' + esc(name) + '<ul>' +
      '<li>样本 ' + smp.meta.sampleCount + ' ' + u + '（全查 ' + smp.meta.fullCount + ' + 抽样 ' + smp.meta.n + '），' +
        '覆盖金额 ' + money(smp.meta.sampleAbsCents) + '，占总体 ' + (smp.meta.coverage * 100).toFixed(2) + '%</li>' +
      '<li>排除 ' + state.pop.stats.excludedCount + ' 笔，逐条列在「排除清单」页并带原始行号</li>' +
      '<li>文件共 5 页：抽凭表 / 抽样信息 / 排除清单 / 数据检查 / 说明</li>' +
      '<li>复核用参数：种子 ' + smp.meta.seed +
        (smp.meta.k ? '；间距 ' + smp.meta.k + '；起点 ' + smp.meta.start : '') + '</li>' +
      '</ul>' +
      '下载没弹出的话，检查浏览器是否拦了下载。');

    if (state.parsed.problems.length) {
      setResult('warn',
        '<b>文件已生成，但有 ' + state.parsed.problems.length + ' 行金额没读出来</b>，这些行既不在样本里也不在排除清单里。' +
        '请在「数据检查」页核对后再定稿。');
    }
  });

  /* ============================================================
   * 回录与评估
   * ============================================================ */

  var evState = { records: null, info: null, warnings: [] };

  function rowsOf(wb, name) {
    var ws = wb.Sheets[name];
    if (!ws) return null;
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  }

  $('evalFile').addEventListener('change', function (e) {
    var f = e.target.files && e.target.files[0];
    if (!f) return;

    var reader = new FileReader();
    reader.onload = function () {
      try {
        var wb = XLSX.read(new Uint8Array(reader.result), { type: 'array', cellDates: false });
        var mainRows = rowsOf(wb, '抽凭表');
        var mainName = '抽凭表';

        if (!mainRows) {
          for (var i = 0; i < wb.SheetNames.length; i++) {
            var rs = rowsOf(wb, wb.SheetNames[i]);
            if (rs && Core.parseVouchingRows(rs).headerRowIndex >= 0) {
              mainRows = rs; mainName = wb.SheetNames[i]; break;
            }
          }
        }
        if (!mainRows) {
          $('evalInfo').textContent = '这个文件里没找到抽凭表（需要一行里同时有「序号」和「凭证号」）。';
          return;
        }

        var pv = Core.parseVouchingRows(mainRows);
        var infoRows = rowsOf(wb, '抽样信息');

        evState.records = pv.records;
        evState.warnings = pv.warnings.slice();
        evState.info = infoRows ? Core.parseInfoRows(infoRows) : {};

        if (evState.info.popCents) $('popYuan').value = Core.moneyText(evState.info.popCents).replace(/,/g, '');

        $('evalParams').classList.remove('hidden');
        $('evalActions').classList.remove('hidden');
        $('evalInfo').innerHTML = '<b>' + esc(f.name) + '</b> · 读到 ' + pv.records.length + ' 笔（' + esc(mainName) + '）';
        renderEval();
      } catch (err) {
        // 读失败必须把上一份的结果清掉，否则旧数字还挂在屏幕上、还能导出
        evState.records = null;
        evState.info = null;
        $('evalResult').innerHTML = '';
        $('evalParams').classList.add('hidden');
        $('evalActions').classList.add('hidden');
        $('evalInfo').textContent = '读不了这个文件：' + err.message;
      }
    };
    reader.readAsArrayBuffer(f);
  });

  /** 用当前参数算一遍评估结果（页面渲染与导出共用） */
  function currentEval() {
    return Core.evaluateVouching(evState.records, {
      alpha: parseFloat($('alphaSelect').value),
      tolerableCents: yuanToCents($('tolerableYuan').value),
      populationCents: yuanToCents($('popYuan').value),
      populationCount: (evState.info && evState.info.popCount) || null,
      // 抽样方法/全查层决定能不能算 MUS 上限；读不到就不算
      sampleMethod: (evState.info && evState.info.sampleMethod) || '',
      sampleMethodLabel: (evState.info && evState.info.sampleMethodLabel) || '',
      fullCount: (evState.info && evState.info.fullCount) || 0
    });
  }

  function renderEval() {
    var box = $('evalResult');
    if (!evState.records) { box.innerHTML = ''; return; }

    var ev = currentEval();

    var html = '<div class="stats">' +
      '<div class="head">抽查结果（改上面参数会立刻重算）</div>' +
      '<div class="body">' +
        cell('相符', ev.counts.ok, '笔', 'ok') +
        cell('不符', ev.counts.exception, '笔', ev.counts.exception ? 'bad' : '') +
        cell('未取得', ev.counts.missing, '笔', ev.counts.missing ? 'warn' : '') +
        cell('尚未填写', ev.counts.blank, '笔', ev.counts.blank ? 'warn' : '') +
      '</div>' +
      '<div class="body">' +
        cell('参与评估', ev.examined, '笔', '') +
        cell('高估金额', Core.moneyText(ev.grossOverCents), '', '') +
        cell('低估金额', Core.moneyText(ev.grossUnderCents), '', '') +
        cell('错报净额', Core.moneyText(ev.netCents), '', '') +
      '</div>' +
      '<div class="body">' +
        cell('推断错报（比率法）', Core.moneyText(ev.projection.ratioCents), '', 'warn') +
        cell('推断错报（差异法）', Core.moneyText(ev.projection.differenceCents), '', 'warn') +
        cell('样本账面合计', Core.moneyText(ev.projection.sampleBookCents), '', '') +
        cell('样本错报合计', Core.moneyText(ev.projection.sampleMisAbsCents), '', '') +
      '</div>' +
      '<div class="foot">共 ' + evState.records.length + ' 笔样本；「不适用」' + ev.counts.na +
        ' 笔不计入评估分母，「未取得」按 100% 错报计入上限。</div>' +
      '</div>';

    if (ev.upper) {
      var better = ev.verdictHint === '未超过';
      html += '<div class="note ' + (ev.verdictHint ? (better ? 'ok' : 'bad') : '') + '">' +
        '<b>错报上限 ' + Core.moneyText(ev.upper.upperCents) + '</b>' +
        '（' + (ev.alpha * 100).toFixed(0) + '% 误受风险；每笔平均账面 ' + Core.moneyText(ev.upper.perItemCents) +
        ' × 系数 ' + ev.upper.factorTotal.toFixed(3) + '）' +
        (ev.verdictHint
          ? '<br>对照可容忍错报 ' + Core.moneyText(ev.tolerableCents) + '：<b>' + ev.verdictHint + '</b>。' +
            '这只是统计上的对照 —— 能不能接受，要结合错报性质、原因和其他程序一起判断。'
          : '<br>填上「可容忍错报」后可以看对照。') +
        '</div>';
    } else {
      html += '<div class="note warn"><b>这份样本算不了「错报上限」</b>' +
        (ev.umBlocked ? '：' + esc(ev.umBlocked) : '：需要「总体账面金额」，并且至少有一笔已检查的样本。') +
        '<br>上限是货币单元抽样（按金额加权）才成立的统计推断。要它，请回第 2 步把「怎么抽」改成' +
        '「按金额加权」，重新生成抽凭表、填完再回录。下面的推断错报仍然可以看。' +
        '</div>';
    }

    if (ev.counts.other) {
      html += '<div class="note warn"><b>有 ' + ev.counts.other + ' 行的结论认不出</b>：' +
        ev.others.map(function (o) { return '第 ' + o.seq + ' 行「' + esc(o.text) + '」'; }).join('；') +
        '。请改成 相符／不符／未取得／不适用 之一。</div>';
    }

    if (ev.exceptions.length) {
      html += '<h3>例外明细</h3><div class="scroll"><table class="data"><thead><tr>' +
        '<th>序号</th><th>原始行号</th><th>凭证号</th><th>科目</th><th>账面金额</th><th>错报金额</th><th>错报率</th>' +
        '</tr></thead><tbody>';
      ev.exceptions.forEach(function (r) {
        var rate = r.absCents ? Math.abs(r.misCents) / r.absCents : 1;
        html += '<tr><td class="num">' + r.seq + '</td><td class="num">' + esc(r.srcRow) + '</td>' +
          '<td>' + esc(r.voucher) + '</td><td>' + esc(r.account) + '</td>' +
          '<td class="num">' + Core.moneyText(r.absCents) + '</td>' +
          '<td class="num">' + Core.moneyText(r.misCents) + '</td>' +
          '<td class="num">' + (rate * 100).toFixed(2) + '%</td></tr>';
      });
      html += '</tbody></table></div>';
    }

    if (evState.warnings.length) {
      html += '<div class="note warn"><ul>' +
        evState.warnings.map(function (t) { return '<li>' + esc(t) + '</li>'; }).join('') + '</ul></div>';
    }

    box.innerHTML = html;
  }

  ['alphaSelect', 'tolerableYuan', 'popYuan'].forEach(function (id) {
    $(id).addEventListener('input', renderEval);
    $(id).addEventListener('change', renderEval);
  });

  $('evalExportBtn').addEventListener('click', function () {
    if (!evState.records) {
      $('evalInfo').textContent = '先选一份已填写的抽凭表，再导出评估表。';
      return;
    }
    var ev = currentEval();
    var name = Core.evalFileName({ client: $('client').value.trim() });
    try {
      var out = Core.buildEvaluationWorkbook(XLSX, ev, {
        indexNo: $('indexNo').value.trim(),
        client: $('client').value.trim()
      });
      XLSX.writeFile(out.wb, out.fileName);
    } catch (err) {
      $('evalInfo').textContent = '导出失败：' + err.message;
      return;
    }

    $('evalInfo').innerHTML = '<b>评估表已导出</b>：' + esc(name) +
      '。错报上界与推断错报都是统计值，结论仍要人工判断并签字。';
  });

  function periodText() {
    var a = $('dateFrom').value, b = $('dateTo').value;
    if (!a && !b) return '全部日期';
    return (a || '最早') + ' 至 ' + (b || '最晚');
  }

  /* ============================================================
   * 初始化
   * ============================================================ */

  $('presetSelect').innerHTML = Object.keys(Core.EVIDENCE_PRESETS).map(function (k) {
    return '<option value="' + esc(k) + '">' + esc(k) + '</option>';
  }).join('');

  $('seedBtn').addEventListener('click', function () {
    $('seedInput').value = Math.floor(Math.random() * 90000000) + 10000000;
    refresh();
  });

  /** 输入框防抖：真底稿动辄上万行，每敲一个字就重算会卡死 */
  function debounce(fn, ms) {
    var timer = null;
    return function () {
      if (timer) clearTimeout(timer);
      timer = setTimeout(fn, ms);
    };
  }

  var refreshSoon = debounce(refresh, 180);

  ['dateFrom', 'dateTo', 'accountText', 'direction', 'minYuan', 'maxYuan',
   'thresholdYuan', 'sampleN', 'methodSelect', 'seedInput'].forEach(function (id) {
    $(id).addEventListener('input', refreshSoon);
    $(id).addEventListener('change', refreshSoon);
  });

  // 手工改过阈值就不再自动建议（否则用户改一次就被覆盖一次）
  ['thresholdYuan', 'sampleN'].forEach(function (id) {
    $(id).addEventListener('input', function () { state.thresholdTouched = true; });
  });

  ['entitySelect', 'unitSelect'].forEach(function (id) {
    $(id).addEventListener('change', refreshSoon);
  });

  /* 默认值：期间留空由用户填；证据列按销售/应收起步 */
  $('presetSelect').value = '销售/应收';
  $('verText').textContent = 'v' + Core.VERSION;

  (function initMode() {
    var saved = null;
    try { saved = window.localStorage.getItem(MODE_KEY); } catch (e) { /* 忽略 */ }
    setMode(saved ? saved === 'simple' : true);
  })();
  renderStats();
})();
