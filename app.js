/* 表格合并工具 —— 纯前端，基于 SheetJS */
(function () {
  'use strict';

  var $ = function (s) { return document.querySelector(s); };
  var uid = 1;

  var state = {
    tables: [],   // {id, file, sheet, rawRows, firstRow, headerRow, enabled}
    result: null
  };

  /* ================= 工具 ================= */
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

  function sheetToRows(ws) {
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: true, defval: null, raw: true });
    var firstRow = 0;
    if (ws['!ref']) {
      try { firstRow = XLSX.utils.decode_range(ws['!ref']).s.r; } catch (e) { firstRow = 0; }
    }
    return { rows: rows, firstRow: firstRow };
  }

  function firstNonBlank(rows) {
    for (var i = 0; i < rows.length; i++) if (rowHasVal(rows[i])) return i;
    return 0;
  }
  function detectHeaderRow(rows) {
    var i = firstNonBlank(rows);
    var cnt = function (r) { return r ? r.filter(hasVal).length : 0; };
    if (rows.length && cnt(rows[i]) < 2) {
      for (var j = i + 1; j < Math.min(rows.length, i + 8); j++) {
        if (cnt(rows[j]) > cnt(rows[i])) return j;
      }
    }
    return i;
  }

  /* 表头解析：names 显示名 / keys 匹配键 / startRow 数据起始下标 */
  function headerInfo(t, headerMode, width) {
    var names = [], keys = [];
    if (headerMode === 'none') {
      for (var c = 0; c < width; c++) { names.push('列' + (c + 1)); keys.push('列' + (c + 1)); }
      return { names: names, keys: keys, startRow: 0, row: -1 };
    }
    var hr;
    if (t.hdrRow !== null && t.hdrRow !== undefined && t.hdrRow >= 0) hr = t.hdrRow;
    else hr = headerMode === 'first' ? firstNonBlank(t.rawRows) : detectHeaderRow(t.rawRows);
    if (hr >= t.rawRows.length) hr = firstNonBlank(t.rawRows);
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

  /* ================= 合并 ================= */
  function currentOpts() {
    return {
      mode: $('#mode').value,
      headerMode: $('#headerMode').value,
      withSource: $('#withSource').checked,
      dedupe: $('#dedupe').checked,
      keepBlank: $('#keepBlank').checked
    };
  }

  function pushRow(rows, seen, row, o) {
    if (o.dedupe) {
      var sig = JSON.stringify(row.map(str));
      if (seen[sig]) return;
      seen[sig] = 1;
    }
    rows.push(row);
  }

  function buildResult() {
    var o = currentOpts();
    var active = state.tables.filter(function (t) { return t.enabled; });
    if (!active.length) { state.result = null; return null; }

    var srcCols = o.withSource ? ['来源文件', '来源工作表'] : [];
    var header = null, rows = [], totalRows = 0, seen = {};

    if (o.mode === 'align') {
      var infos = active.map(function (t) {
        var w = widthOf(t.rawRows);
        return { t: t, info: headerInfo(t, o.headerMode, w) };
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
          pushRow(rows, seen, out, o);
        }
      });
    } else {
      /* 直接堆叠：每张表以自己的表头行开头，原样上下拼接 */
      var w = 0;
      active.forEach(function (t) { w = Math.max(w, widthOf(t.rawRows)); });
      active.forEach(function (t) {
        var info = headerInfo(t, o.headerMode, w);
        if (info.row >= 0) {
          var h = [];
          if (o.withSource) h.push(t.file, t.sheet);
          for (var c = 0; c < w; c++) h.push(info.names[c]);
          pushRow(rows, seen, h, o);
        }
        var data = t.rawRows;
        for (var r = info.startRow; r < data.length; r++) {
          var row = data[r] || [];
          if (!o.keepBlank && !rowHasVal(row)) continue;
          var out = [];
          if (o.withSource) out.push(t.file, t.sheet);
          for (var c2 = 0; c2 < w; c2++) out.push(row[c2] === undefined ? null : row[c2]);
          totalRows++;
          pushRow(rows, seen, out, o);
        }
      });
    }

    state.result = {
      header: header, rows: rows, srcCols: srcCols,
      totalRows: totalRows, mode: o.mode, tables: active.length,
      missing: active.reduce(function (a, t) { return a + (t.rawRows.length - (headerInfo(t, o.headerMode, widthOf(t.rawRows)).startRow)); }, 0) - totalRows
    };
    return state.result;
  }

  /* ================= 渲染：数据源列表 ================= */
  function renderList() {
    var box = $('#tableList');
    box.innerHTML = '';
    if (!state.tables.length) {
      box.innerHTML = '<div class="meta" style="color:var(--sub);font-size:13px">还没有数据，请先选择文件。</div>';
      return;
    }
    state.tables.forEach(function (t) {
      var w = widthOf(t.rawRows), hm = $('#headerMode').value;
      var info = headerInfo(t, hm, w);
      var dataRows = 0;
      for (var r = info.startRow; r < t.rawRows.length; r++) if (rowHasVal(t.rawRows[r])) dataRows++;

      var el = document.createElement('div');
      el.className = 'src' + (t.enabled ? '' : ' off');
      el.innerHTML =
        '<input type="checkbox" ' + (t.enabled ? 'checked' : '') + '>' +
        '<div class="grow"><div class="name">' + esc(t.file) + '</div>' +
        '<div class="meta">工作表：' + esc(t.sheet) + ' · ' + dataRows + ' 行 · ' + w + ' 列</div></div>' +
        '<label class="hdr">表头行 <input type="number" min="1" value="' + (t.firstRow + (info.row < 0 ? 1 : info.row) + 1) + '"' + (info.row < 0 ? ' disabled' : '') + '></label>';

      el.querySelector('input[type=checkbox]').onchange = function (e) {
        t.enabled = e.target.checked;
        update();
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
  function esc(s) {
    return str(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
  }

  /* ================= 渲染：预览 ================= */
  function renderResult() {
    var res = state.result;
    if (!res || !res.rows.length) {
      $('#resCard').hidden = !state.tables.length;
      $('#stats').innerHTML = '<span style="color:var(--warn)">无可合并的数据（可能全部被取消勾选）</span>';
      $('#preview').innerHTML = '';
      $('#previewNote').textContent = '';
      return;
    }
    $('#resCard').hidden = false;
    var colCount = res.header ? res.header.length : Math.max.apply(null, res.rows.map(function (r) { return r.length; }));
    $('#stats').innerHTML = '<b>' + res.tables + '</b> 个数据源 · <b>' + res.rows.length + '</b> 行 × <b>' + colCount + '</b> 列' +
      (res.mode === 'align' ? ' · 表头并集对齐' : ' · 原样堆叠');

    var LIMIT = 60;
    var show = res.rows.slice(0, LIMIT);
    var html = '<thead><tr>';
    if (res.header) {
      res.header.forEach(function (h, i) { html += '<th>' + esc(h) + '</th>'; });
    } else {
      var r0 = show[0] || [];
      var n = Math.max(colCount, r0.length);
      for (var i = 0; i < n; i++) html += '<th>' + esc(r0[i] === undefined ? '' : r0[i]) + '</th>';
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
    $('#previewNote').textContent = res.rows.length > LIMIT
      ? '预览仅显示前 ' + LIMIT + ' 行，导出的文件包含全部 ' + res.rows.length + ' 行。'
      : '已显示全部 ' + res.rows.length + ' 行。';
  }

  /* ================= 导出 ================= */
  function buildAOA() {
    var res = state.result;
    if (!res) return null;
    return res.header ? [res.header].concat(res.rows) : res.rows;
  }

  function exportXlsx() {
    var aoa = buildAOA();
    if (!aoa) return;
    var ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
    /* 列宽自适应 */
    var widths = [];
    var sample = Math.min(aoa.length, 200);
    for (var r = 0; r < sample; r++) {
      aoa[r].forEach(function (v, i) {
        var len = str(v).length + (str(v).match(/[^\x00-\xff]/g) || []).length;
        widths[i] = Math.max(widths[i] || 6, Math.min(len + 2, 42));
      });
    }
    ws['!cols'] = widths.map(function (w) { return { wch: w || 10 }; });
    if (aoa.length > 1) {
      ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: aoa[0].length - 1 } }) };
    }
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '合并结果');
    XLSX.writeFile(wb, stamp('合并结果') + '.xlsx');
  }

  function exportCsv() {
    var res = state.result;
    if (!res) return;
    var lines = [];
    var q = function (v) {
      var s = str(v);
      return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    var aoa = buildAOA();
    aoa.forEach(function (row) { lines.push(row.map(q).join(',')); });
    var blob = new Blob(['\uFEFF' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    download(blob, stamp('合并结果') + '.csv');
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

  /* ================= 主流程 ================= */
  function update() {
    $('#optCard').hidden = !state.tables.length;
    renderList();
    buildResult();
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
      readWorkbook(f).then(function (wb) {
        var sheets = $('#allSheets').checked ? wb.SheetNames : wb.SheetNames.slice(0, 1);
        sheets.forEach(function (name) {
          var sr = sheetToRows(wb.Sheets[name]);
          if (!sr.rows.length) return;
          state.tables.push({
            id: uid++, file: f.name, sheet: name,
            rawRows: sr.rows, firstRow: sr.firstRow,
            hdrRow: null, enabled: true
          });
        });
      }).catch(function (err) {
        errors.push(f.name + '（' + (err && err.message ? err.message : '无法识别') + '）');
      }).then(function () { i++; setTimeout(step, 0); });
    };
    step();
  }

  /* ================= 事件 ================= */
  var inited = false;
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

  /* 供自动化测试调用的钩子（不影响正常使用） */
  window.__mt = { addFiles: addFiles, update: update, buildAOA: buildAOA, state: state };

  document.addEventListener('DOMContentLoaded', init);
  if (document.readyState !== 'loading') init();
})();
