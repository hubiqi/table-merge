/* 表格合并工具 —— 纯前端，基于 SheetJS */
(function () {
  'use strict';

  var $ = function (s) { return document.querySelector(s); };
  var uid = 1;

  var state = {
    tables: [],   // {id, file, sheet, rawRows, firstRow, hdrRow, enabled}
    result: null
  };

  /* ================= 基础工具 ================= */
  var utf8Strict = new TextDecoder('utf-8', { fatal: true });
  var gbk = (function () { try { return new TextDecoder('gbk'); } catch (e) { return null; } })();

  function isText(name) { return /\.(csv|tsv|txt)$/i.test(name); }

  function fmtDate(d) {
    var p = function (n) { return (n < 10 ? '0' : '') + n; };
    var s = d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
    if (d.getHours() || d.getMinutes() || d.getSeconds())
      s += ' ' + p(d.getHours()) + ':' + p(d.getMinutes()) + ':' + p(d.getSeconds());
    return s;
  }
  function str(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return fmtDate(v);
    return String(v);
  }
  function keyOf(s) { return str(s).replace(/[\s\u3000]+/g, '').toLowerCase(); }
  function hasVal(v) { return !(v === null || v === undefined || String(v).trim() === ''); }
  function rowHasVal(r) { return !!r && r.some(hasVal); }
  function countVal(r) { return r ? r.filter(hasVal).length : 0; }
  function widthOf(rows) {
    var w = 0;
    for (var i = 0; i < rows.length; i++) if (rows[i] && rows[i].length > w) w = rows[i].length;
    return w;
  }

  /* ================= 解析 ================= */
  function readWorkbook(file) {
    return file.arrayBuffer().then(function (buf) {
      if (isText(file.name)) {
        var text = null;
        try { text = utf8Strict.decode(buf); }
        catch (e) { if (gbk) text = gbk.decode(buf); }
        if (text === null) text = new TextDecoder('utf-8').decode(buf);
        if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
        return XLSX.read(text, { type: 'string', cellDates: true, raw: true });
      }
      return XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true, raw: true });
    });
  }

  /* 按真实单元格算范围。
     很多导出的 xlsx 里 <dimension> 是错的（比如只写到表头行），
     SheetJS 会照着这个范围读，导致数据行被整段丢掉 —— 这里绕开它。 */
  function trueRange(ws) {
    var s = { r: Infinity, c: Infinity }, e = { r: -1, c: -1 }, any = false;
    for (var k in ws) {
      if (!Object.prototype.hasOwnProperty.call(ws, k)) continue;
      if (k.charCodeAt(0) === 33 /* ! */ || k.charCodeAt(0) === 36 /* $ */) continue;
      var a;
      try { a = XLSX.utils.decode_cell(k); } catch (err) { continue; }
      if (!a || a.r == null || a.c == null) continue;
      any = true;
      if (a.r < s.r) s.r = a.r;
      if (a.c < s.c) s.c = a.c;
      if (a.r > e.r) e.r = a.r;
      if (a.c > e.c) e.c = a.c;
    }
    if (!any || e.r < 0) return null;
    return { s: s, e: e };
  }

  function sheetToRows(ws) {
    var rng = trueRange(ws);
    if (!rng) return { rows: [], firstRow: 0, declared: null, rng: null };
    var opt = { header: 1, blankrows: true, defval: null, raw: true, range: rng };
    var rows = XLSX.utils.sheet_to_json(ws, opt);
    var declared = null;
    try { if (ws['!ref']) declared = ws['!ref']; } catch (e) { declared = null; }
    return { rows: rows, firstRow: rng.s.r, declared: declared, rng: rng };
  }

  function firstNonBlank(rows) {
    for (var i = 0; i < rows.length; i++) if (rowHasVal(rows[i])) return i;
    return 0;
  }
  function detectHeaderRow(rows) {
    var i = firstNonBlank(rows);
    if (rows.length && countVal(rows[i]) < 2) {
      for (var j = i + 1; j < Math.min(rows.length, i + 8); j++) {
        if (countVal(rows[j]) > countVal(rows[i])) return j;
      }
    }
    return i;
  }

  /* 表头解析：names 显示名 / keys 匹配键 / startRow 数据起始下标 */
  function headerInfo(t, headerMode, width) {
    var names = [], keys = [];
    if (headerMode === 'none') {
      for (var c0 = 0; c0 < width; c0++) { names.push('列' + (c0 + 1)); keys.push('列' + (c0 + 1)); }
      return { names: names, keys: keys, startRow: 0, row: -1 };
    }
    var hr;
    if (t.hdrRow !== null && t.hdrRow !== undefined && t.hdrRow >= 0) hr = t.hdrRow;
    else hr = headerMode === 'first' ? firstNonBlank(t.rawRows) : detectHeaderRow(t.rawRows);
    if (hr >= t.rawRows.length) hr = firstNonBlank(t.rawRows);
    /* 兜底：表头行以下必须还有数据，否则退回第一行非空行、再不行退回第 0 行 */
    if (t.rawRows.length > 1 && hr >= t.rawRows.length - 1) {
      var alt = firstNonBlank(t.rawRows);
      if (alt < t.rawRows.length - 1) hr = alt; else hr = t.rawRows.length - 2;
    }
    var row = t.rawRows[hr] || [];
    var seen = {};
    for (var i = 0; i < width; i++) {
      var raw = hasVal(row[i]) ? str(row[i]).trim() : '列' + (i + 1);
      var name = raw, k = keyOf(raw), n = 1;
      while (seen[k] !== undefined) { n++; name = raw + ' (' + n + ')'; k = keyOf(name); }
      seen[k] = i;
      names.push(name); keys.push(k);
    }
    return { names: names, keys: keys, startRow: hr + 1, row: hr };
  }

  /* ================= 合并（纯函数，便于测试） ================= */
  function mergeTables(active, o) {
    var srcCols = o.withSource ? ['来源文件', '来源工作表'] : [];
    var header = null, rows = [], totalRows = 0, seen = {};

    function push(row) {
      if (o.dedupe) {
        var sig = JSON.stringify(row.map(str));
        if (seen[sig]) return;
        seen[sig] = 1;
      }
      rows.push(row);
    }

    if (o.mode === 'align') {
      var infos = active.map(function (t) {
        return { t: t, info: headerInfo(t, o.headerMode, widthOf(t.rawRows)) };
      });
      var colIndex = {}, cols = [];
      infos.forEach(function (x) {
        x.map = {};
        x.info.keys.forEach(function (k, i) {
          x.map[k] = i;
          if (colIndex[k] === undefined) { colIndex[k] = cols.length; cols.push({ key: k, name: x.info.names[i] }); }
        });
      });
      header = srcCols.concat(cols.map(function (c) { return c.name; }));

      infos.forEach(function (x) {
        var data = x.t.rawRows;
        for (var r = x.info.startRow; r < data.length; r++) {
          var row = data[r] || [];
          if (!o.keepBlank && !rowHasVal(row)) continue;
          var out = [];
          if (o.withSource) out.push(x.t.file, x.t.sheet);
          for (var c = 0; c < cols.length; c++) {
            var idx = x.map[cols[c].key];
            out.push(idx === undefined ? null : (row[idx] === undefined ? null : row[idx]));
          }
          totalRows++;
          push(out);
        }
      });
    } else {
      var w = 0;
      active.forEach(function (t) { w = Math.max(w, widthOf(t.rawRows)); });
      active.forEach(function (t) {
        var info = headerInfo(t, o.headerMode, w);
        if (info.row >= 0) {
          var h = [];
          if (o.withSource) h.push(t.file, t.sheet);
          for (var c1 = 0; c1 < w; c1++) h.push(info.names[c1]);
          push(h);
        }
        var data = t.rawRows;
        for (var r = info.startRow; r < data.length; r++) {
          var row = data[r] || [];
          if (!o.keepBlank && !rowHasVal(row)) continue;
          var out = [];
          if (o.withSource) out.push(t.file, t.sheet);
          for (var c2 = 0; c2 < w; c2++) out.push(row[c2] === undefined ? null : row[c2]);
          totalRows++;
          push(out);
        }
      });
    }
    return { header: header, rows: rows, srcCols: srcCols, totalRows: totalRows, mode: o.mode, tables: active.length };
  }

  function currentOpts() {
    return {
      mode: $('#mode').value,
      headerMode: $('#headerMode').value,
      withSource: $('#withSource').checked,
      dedupe: $('#dedupe').checked,
      keepBlank: $('#keepBlank').checked
    };
  }

  function buildResult() {
    var active = state.tables.filter(function (t) { return t.enabled; });
    state.result = active.length ? mergeTables(active, currentOpts()) : null;
    return state.result;
  }

  window.__mt = {
    state: state, mergeTables: mergeTables, sheetToRows: sheetToRows, readWorkbook: readWorkbook,
    headerInfo: headerInfo, trueRange: trueRange, currentOpts: currentOpts
  };
})();

/* ================= 界面 ================= */
(function () {
  'use strict';
  var MT = window.__mt;
  var state = MT.state;
  var $ = function (s) { return document.querySelector(s); };
  var uid = 1;

  function esc(s) {
    return String(s === null || s === undefined ? '' : s).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  }

  /* ---------- 数据源列表 ---------- */
  function tableStats(t) {
    var w = 0, i;
    for (i = 0; i < t.rawRows.length; i++) if (t.rawRows[i] && t.rawRows[i].length > w) w = t.rawRows[i].length;
    var info = MT.headerInfo(t, $('#headerMode').value, w);
    var dataRows = 0;
    for (var r = info.startRow; r < t.rawRows.length; r++) {
      var row = t.rawRows[r] || [];
      if (row.some(function (v) { return !(v === null || v === undefined || String(v).trim() === ''); })) dataRows++;
    }
    return { width: w, info: info, dataRows: dataRows, totalRows: t.rawRows.length };
  }

  function renderList() {
    var box = $('#tableList');
    box.innerHTML = '';
    if (!state.tables.length) {
      box.innerHTML = '<div style="color:var(--sub);font-size:13px">还没有数据，请先选择文件。</div>';
      return;
    }
    state.tables.forEach(function (t) {
      var s = tableStats(t);
      var el = document.createElement('div');
      el.className = 'src' + (t.enabled ? '' : ' off');
      var warn = '';
      if (!s.dataRows) {
        warn = '<div class="warnline">⚠ 表头行以下没有数据行（共 ' + s.totalRows + ' 行）。可能是表头行设置不对，或该文件本身只有表头。</div>';
      }
      el.innerHTML =
        '<input type="checkbox" ' + (t.enabled ? 'checked' : '') + '>' +
        '<div class="grow"><div class="name">' + esc(t.file) + '</div>' +
        '<div class="meta">工作表：' + esc(t.sheet) + ' · 数据 ' + s.dataRows + ' 行 · ' + s.width + ' 列' +
        (t.note ? ' · <span style="color:var(--warn)">' + esc(t.note) + '</span>' : '') + '</div>' + warn + '</div>' +
        '<label class="hdr">表头行 <input type="number" min="1" value="' +
        (t.firstRow + (s.info.row < 0 ? 1 : s.info.row) + 1) + '"' + (s.info.row < 0 ? ' disabled' : '') + '></label>';

      el.querySelector('input[type=checkbox]').onchange = function (e) {
        t.enabled = e.target.checked; update();
      };
      var num = el.querySelector('input[type=number]');
      if (num) num.onchange = function (e) {
        var v = parseInt(e.target.value, 10);
        if (!isNaN(v)) t.hdrRow = Math.max(0, v - 1 - t.firstRow);
        update();
      };
      box.appendChild(el);
    });
  }

  /* ---------- 结果预览 ---------- */
  function renderResult() {
    var res = state.result;
    if (!res) {
      $('#resCard').hidden = !state.tables.length;
      $('#stats').innerHTML = '';
      $('#preview').innerHTML = '';
      $('#previewNote').textContent = '';
      return;
    }
    $('#resCard').hidden = false;
    var colCount = res.header ? res.header.length
      : res.rows.reduce(function (a, r) { return Math.max(a, r.length); }, 0);
    $('#stats').innerHTML = '<b>' + res.tables + '</b> 个数据源 · <b>' + res.rows.length + '</b> 行 × <b>' +
      colCount + '</b> 列' + (res.mode === 'align' ? ' · 表头并集对齐' : ' · 原样堆叠');

    var hasData = res.mode !== 'align' ? res.rows.length > (res.header ? 0 : 1) : res.rows.length > 0;
    if (!hasData) {
      $('#previewNote').innerHTML = '<span style="color:var(--danger)">⚠ 只读到了表头，没有数据行。</span>' +
        '请检查上方的「表头行」是否指到了最后一行；若文件里确实有数据，请点开下面的「诊断信息」把内容发我。';
    }

    var LIMIT = 60;
    var show = res.rows.slice(0, LIMIT);
    var html = '<thead><tr>';
    if (res.header) {
      res.header.forEach(function (h) { html += '<th>' + esc(h) + '</th>'; });
    } else {
      var n = Math.max(colCount, (show[0] || []).length);
      for (var i = 0; i < n; i++) html += '<th>' + esc((show[0] || [])[i]) + '</th>';
    }
    html += '</tr></thead><tbody>';
    var startAt = res.header ? 0 : 1;
    for (var r = startAt; r < show.length; r++) {
      html += '<tr>';
      for (var c = 0; c < colCount; c++) {
        var cls = (res.srcCols.length && c < res.srcCols.length) ? ' class="srcCol"' : '';
        html += '<td' + cls + '>' + esc(show[r][c]) + '</td>';
      }
      html += '</tr>';
    }
    html += '</tbody>';
    $('#preview').innerHTML = html;
    if (hasData) {
      $('#previewNote').textContent = res.rows.length > LIMIT
        ? '预览仅显示前 ' + LIMIT + ' 行，导出的文件包含全部 ' + res.rows.length + ' 行。'
        : '已显示全部 ' + res.rows.length + ' 行。';
    }
    renderDiag();
  }

  /* ---------- 诊断信息 ---------- */
  function renderDiag() {
    var d = $('#diag');
    if (!d) return;
    var lines = [];
    state.tables.forEach(function (t) {
      var s = tableStats(t);
      lines.push('【' + t.file + ' / ' + t.sheet + '】 单元格范围 ' + (t.rngTxt || '?') +
        ' · 已读 ' + s.totalRows + ' 行 · 表头行 ' + (t.firstRow + (s.info.row < 0 ? 1 : s.info.row) + 1) +
        ' · 数据 ' + s.dataRows + ' 行');
      for (var i = 0; i < Math.min(t.rawRows.length, 4); i++) {
        lines.push('  ' + (t.firstRow + i + 1) + ': ' +
          JSON.stringify((t.rawRows[i] || []).slice(0, 12).map(function (v) {
            return v === null ? null : (v instanceof Date ? v.toISOString().slice(0, 10) : String(v).slice(0, 40));
          })));
      }
      lines.push('');
    });
    d.textContent = lines.join('\n') || '（无）';
  }

  /* ---------- 导出 ---------- */
  function buildAOA() {
    var res = state.result;
    if (!res) return null;
    return res.header ? [res.header].concat(res.rows) : res.rows;
  }
  function stamp(pre) {
    var d = new Date(), p = function (n) { return (n < 10 ? '0' : '') + n; };
    return pre + '_' + d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '_' + p(d.getHours()) + p(d.getMinutes());
  }
  function download(blob, name) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  }
  function exportXlsx() {
    var aoa = buildAOA();
    if (!aoa) return;
    var ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
    var widths = [], sample = Math.min(aoa.length, 200);
    for (var r = 0; r < sample; r++) {
      aoa[r].forEach(function (v, i) {
        var len = 0;
        var s = String(v === null || v === undefined ? '' : (v instanceof Date ? '2000-00-00 00:00:00' : v));
        len = s.length + (s.match(/[^\x00-\xff]/g) || []).length;
        widths[i] = Math.max(widths[i] || 6, Math.min(len + 2, 42));
      });
    }
    ws['!cols'] = widths.map(function (w) { return { wch: w || 10 }; });
    if (aoa.length > 1 && aoa[0] && aoa[0].length) {
      ws['!autofilter'] = {
        ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: aoa[0].length - 1 } })
      };
    }
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '合并结果');
    XLSX.writeFile(wb, stamp('合并结果') + '.xlsx');
  }
  function exportCsv() {
    var aoa = buildAOA();
    if (!aoa) return;
    var q = function (v) {
      var s = String(v === null || v === undefined ? '' : v);
      return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    var lines = aoa.map(function (row) { return row.map(q).join(','); });
    download(new Blob(['\uFEFF' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }), stamp('合并结果') + '.csv');
  }

  /* ---------- 主流程 ---------- */
  var inited = false;
  function update() {
    $('#optCard').hidden = !state.tables.length;
    renderList();
    var active = state.tables.filter(function (t) { return t.enabled; });
    state.result = active.length ? MT.mergeTables(active, MT.currentOpts()) : null;
    renderResult();
  }

  function addFiles(files) {
    var list = Array.prototype.slice.call(files);
    if (!list.length) return;
    $('#parsing').hidden = false;
    var i = 0, errors = [];
    var step = function () {
      if (i >= list.length) {
        $('#parsing').hidden = true;
        if (errors.length) {
          var old = $('#drop').querySelector('.err');
          if (old) old.remove();
          var p = document.createElement('div');
          p.className = 'err';
          p.textContent = '以下文件读取失败：' + errors.join('；');
          $('#drop').appendChild(p);
        }
        update();
        return;
      }
      var f = list[i];
      $('#parsingText').textContent = '正在解析 ' + f.name + '（' + (i + 1) + '/' + list.length + '）…';
      MT.readWorkbook(f).then(function (wb) {
        var sheets = $('#allSheets').checked ? wb.SheetNames : wb.SheetNames.slice(0, 1);
        sheets.forEach(function (name) {
          var sr = MT.sheetToRows(wb.Sheets[name]);
          if (!sr.rows.length) return;
          state.tables.push({
            id: uid++, file: f.name, sheet: name,
            rawRows: sr.rows, firstRow: sr.firstRow,
            rngTxt: sr.rng ? XLSX.utils.encode_range(sr.rng) + (sr.declared && sr.declared !== XLSX.utils.encode_range(sr.rng) ? '（文件声明 ' + sr.declared + '，已按真实范围读取）' : '') : '',
            hdrRow: null, enabled: true
          });
        });
      }).catch(function (err) {
        errors.push(f.name + '（' + (err && err.message ? err.message : '无法识别') + '）');
      }).then(function () { i++; setTimeout(step, 0); });
    };
    step();
  }

  function init() {
    if (inited) return; inited = true;
    $('#pick').onclick = function () { $('#file').click(); };
    $('#file').onchange = function (e) { addFiles(e.target.files); e.target.value = ''; };
    var drop = $('#drop');
    ['dragenter', 'dragover'].forEach(function (ev) {
      drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('over'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('over'); });
    });
    drop.addEventListener('drop', function (e) {
      if (e.dataTransfer && e.dataTransfer.files) addFiles(e.dataTransfer.files);
    });
    ['mode', 'headerMode', 'withSource', 'dedupe', 'keepBlank', 'allSheets'].forEach(function (id) {
      var el = document.getElementById(id);
      if (el) el.onchange = update;
    });
    $('#exportX').onclick = exportXlsx;
    $('#exportC').onclick = exportCsv;
    $('#reset').onclick = function () {
      state.tables = []; state.result = null;
      var old = $('#drop').querySelector('.err');
      if (old) old.remove();
      $('#resCard').hidden = true;
      update();
    };
  }

  /* 扩展调试/测试钩子 */
  MT.addFiles = addFiles;
  MT.update = update;
  MT.buildAOA = buildAOA;
  MT.exportXlsx = exportXlsx;

  if (typeof document !== 'undefined') {
    document.addEventListener('DOMContentLoaded', init);
    if (document.readyState !== 'loading') init();
  }
})();
