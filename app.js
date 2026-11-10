/* FormulaKit — Excel add-in logic (Office.js, shared runtime). */
/* global Office, Excel, FK */
(function () {
  "use strict";

  var S = {
    list: [],          // current trace list: [{sheet, addr, label, value}]
    idx: -1,
    originPushed: false,
    back: [],          // navigation history
    explore: null,     // last explored cell {a, refs}
    nf: null,          // number-format cycle memory
    border: null,      // border cycle memory
    mapIssues: []
  };

  var COLORS = ["#2f6fd6", "#c0392b", "#1e8449", "#8e44ad", "#d35400", "#16a085", "#b7950b", "#c2185b"];
  var BLUE = "#0000FF", GREEN = "#008000", BLACK = "#000000";
  var NUM_FORMATS = [
    '#,##0_);(#,##0);"-"_)',
    '#,##0.0_);(#,##0.0);"-"_)',
    '#,##0.00_);(#,##0.00);"-"_)',
    '0.0%_);(0.0%);"-"_)',
    '0.0"x"_);(0.0"x");"-"_)',
    "dd-mmm-yy",
    "General"
  ];
  var KEYS = [
    ["Trace precedents", "P"], ["Trace dependents", "D"], ["Jump to next in list", "J"],
    ["Go back", "B"], ["Formula explorer", "E"], ["Model map", "M"],
    ["Cycle number format", "N"], ["Cycle font colour", "C"], ["Auto-colour", "A"], ["Cycle borders", "K"]
  ];

  // ------------------------------------------------------------------ UI helpers

  function $(id) { return document.getElementById(id); }

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  }

  function status(msg, isErr) {
    var el = $("status");
    if (!el) return;
    el.textContent = msg || "";
    el.className = isErr ? "status err" : "status";
  }

  function friendly(e) {
    var m = (e && (e.message || e.code)) || String(e);
    if (e && e.code === "ApiNotFound") return "This needs a newer version of Excel for Mac. Update Excel and try again.";
    return m;
  }

  function guard(fn) {
    return function () {
      var args = arguments;
      status("");
      return Promise.resolve()
        .then(function () { return fn.apply(null, args); })
        .catch(function (e) { console.error(e); status(friendly(e), true); });
    };
  }

  function showTab(name) {
    var navs = document.querySelectorAll("nav button");
    for (var i = 0; i < navs.length; i++) navs[i].classList.toggle("on", navs[i].getAttribute("data-tab") === name);
    var secs = document.querySelectorAll("main section");
    for (var j = 0; j < secs.length; j++) secs[j].classList.toggle("on", secs[j].id === "tab-" + name);
  }

  function showPane() {
    try { return Office.addin.showAsTaskpane().catch(function () {}); }
    catch (e) { return Promise.resolve(); }
  }

  function label(sheet, addr, home) {
    return sheet && sheet !== home ? sheet + "!" + addr : addr;
  }

  function fmtValue(v, type, text) {
    if (type === "Empty") return "(empty)";
    if (type === "Error") return String(v);
    if (text && text.indexOf("##") < 0) return text;
    if (typeof v === "number") return v.toLocaleString(undefined, { maximumFractionDigits: 4 });
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    return String(v);
  }

  function fmtNum(n) {
    return n.toLocaleString(undefined, { maximumFractionDigits: 2 });
  }

  // ------------------------------------------------------------------ Excel helpers

  async function getActive(ctx) {
    var c = ctx.workbook.getActiveCell();
    c.load("address,formulas,values,valueTypes,text,rowIndex,columnIndex");
    c.worksheet.load("name");
    await ctx.sync();
    var p = FK.splitSheetAddress(c.address);
    return {
      range: c,
      sheet: c.worksheet.name,
      addr: p.addr,
      formula: String(c.formulas[0][0]),
      value: c.values[0][0],
      type: c.valueTypes[0][0],
      text: c.text[0][0],
      row: c.rowIndex + 1,
      col: c.columnIndex + 1
    };
  }

  // Turn parsed reference tokens into loaded ranges. Returns an array aligned
  // with refs; unresolved entries are null.
  async function resolveRefs(ctx, home, refs) {
    var wb = ctx.workbook;
    var homeWs = wb.worksheets.getItem(home.sheet);
    var sheets = {};
    var i, r;

    // phase 1: look up sheets, names, tables
    for (i = 0; i < refs.length; i++) {
      r = refs[i];
      if (r.kind === "address" && r.sheet && !sheets[r.sheet]) {
        sheets[r.sheet] = wb.worksheets.getItemOrNullObject(r.sheet);
      } else if (r.kind === "name") {
        r._wb = wb.names.getItemOrNullObject(r.name);
        r._wb.load("type");
        r._ws = homeWs.names.getItemOrNullObject(r.name);
        r._ws.load("type");
      } else if (r.kind === "table") {
        r._t = wb.tables.getItemOrNullObject(r.table);
        r._t.load("name");
      }
    }
    await ctx.sync();

    // phase 1b: column names of found tables
    var needCols = false;
    for (i = 0; i < refs.length; i++) {
      r = refs[i];
      if (r.kind === "table" && !r._t.isNullObject && r.column) {
        r._t.columns.load("items/name");
        needCols = true;
      }
    }
    if (needCols) await ctx.sync();

    // phase 2: ranges
    var rngs = refs.map(function (r) {
      var rng = null;
      if (r.kind === "address") {
        var w = r.sheet ? sheets[r.sheet] : homeWs;
        if (r.sheet && w.isNullObject) return null;
        rng = w.getRange(r.addr);
      } else if (r.kind === "name") {
        var nm = null;
        if (!r._ws.isNullObject && r._ws.type === "Range") nm = r._ws;
        else if (!r._wb.isNullObject && r._wb.type === "Range") nm = r._wb;
        if (!nm) return null;
        rng = nm.getRangeOrNullObject();
      } else if (r.kind === "table") {
        if (r._t.isNullObject) return null;
        if (r.column) {
          var idx = -1, items = r._t.columns.items;
          for (var k = 0; k < items.length; k++) {
            if (items[k].name.toLowerCase() === r.column.toLowerCase()) { idx = k; break; }
          }
          if (idx < 0) return null;
          rng = r._t.columns.getItemAt(idx).getDataBodyRange();
          if (r.thisRow) rng = rng.getIntersectionOrNullObject(homeWs.getRange(home.row + ":" + home.row));
        } else {
          rng = r._t.getDataBodyRange();
        }
      }
      if (!rng) return null;
      rng.load("address,cellCount");
      return rng;
    });
    await ctx.sync();

    // phase 3: values for small ranges
    rngs.forEach(function (rng) {
      if (rng && !rng.isNullObject && rng.cellCount <= 5000) rng.load("values,valueTypes,text");
    });
    await ctx.sync();

    return rngs.map(function (rng, i) {
      if (!rng || rng.isNullObject) return null;
      var p = FK.splitSheetAddress(rng.address);
      var out = { ref: refs[i], sheet: p.sheet || home.sheet, addr: p.addr, cellCount: rng.cellCount };
      if (rng.cellCount === 1) {
        out.value = rng.values[0][0];
        out.type = rng.valueTypes[0][0];
        out.text = rng.text[0][0];
        out.display = fmtValue(out.value, out.type, out.text);
      } else if (rng.cellCount <= 5000) {
        var sum = 0, nums = 0;
        rng.values.forEach(function (row) {
          row.forEach(function (v) { if (typeof v === "number") { sum += v; nums++; } });
        });
        out.display = rng.values.length + "×" + rng.values[0].length + (nums ? "  Σ " + fmtNum(sum) : "");
      } else {
        out.display = rng.cellCount.toLocaleString() + " cells";
      }
      return out;
    });
  }

  async function navigate(item, pushOrigin) {
    await Excel.run(async function (ctx) {
      if (pushOrigin) {
        var c = ctx.workbook.getActiveCell();
        c.load("address");
        await ctx.sync();
        var p = FK.splitSheetAddress(c.address);
        S.back.push({ sheet: p.sheet, addr: p.addr });
        if (S.back.length > 50) S.back.shift();
      }
      var ws = ctx.workbook.worksheets.getItem(item.sheet);
      ws.activate();
      ws.getRange(item.addr).select();
      await ctx.sync();
    }).catch(function (e) {
      throw new Error("Couldn't go to " + item.sheet + "!" + item.addr + " — the sheet may be hidden or protected. (" + friendly(e) + ")");
    });
  }

  // ------------------------------------------------------------------ Trace

  function renderTrace(title, sub, list, emptyMsg) {
    S.list = list;
    S.idx = -1;
    S.originPushed = false;
    var html = '<div class="title">' + esc(title) + "</div>";
    if (sub) html += '<div class="sub">' + esc(sub) + "</div>";
    if (!list.length) {
      html += '<div class="empty">' + esc(emptyMsg) + "</div>";
    } else {
      html += '<div class="list">';
      list.forEach(function (it, i) {
        html += '<button class="item" data-i="' + i + '"><span class="a">' +
          (it.color ? '<span class="sw" style="background:' + it.color + '"></span>' : "") +
          esc(it.label) + '</span><span class="v">' + esc(it.value || "") + "</span></button>";
      });
      html += "</div>";
    }
    $("traceOut").innerHTML = html;
    bindList($("traceOut"));
  }

  function bindList(root) {
    var btns = root.querySelectorAll(".item");
    for (var i = 0; i < btns.length; i++) {
      btns[i].addEventListener("click", guard(onItemClick));
      btns[i].addEventListener("keydown", onItemKey);
    }
  }

  async function onItemClick(ev) {
    var el = ev && ev.currentTarget;
    var i = +el.getAttribute("data-i");
    var list = el.closest("#mapOut") ? S.mapIssues : S.list;
    await jumpTo(list, i);
  }

  function onItemKey(ev) {
    var el = ev.currentTarget;
    if (ev.key === "ArrowDown" && el.nextElementSibling) { el.nextElementSibling.focus(); ev.preventDefault(); }
    if (ev.key === "ArrowUp" && el.previousElementSibling) { el.previousElementSibling.focus(); ev.preventDefault(); }
  }

  async function jumpTo(list, i) {
    var item = list[i];
    if (!item) return;
    await navigate(item, !S.originPushed);
    S.originPushed = true;
    if (list === S.list) {
      S.idx = i;
      var btns = $("traceOut").querySelectorAll(".item");
      for (var k = 0; k < btns.length; k++) btns[k].classList.toggle("cur", k === i);
    }
  }

  async function tracePrecedents() {
    showTab("trace");
    await showPane();
    await Excel.run(async function (ctx) {
      var a = await getActive(ctx);
      if (a.formula.charAt(0) !== "=") {
        renderTrace("Precedents of " + a.addr, "", [], "This cell has no formula.");
        return;
      }
      var refs = FK.findRefs(a.formula);
      var items = await resolveRefs(ctx, a, refs);
      var seen = {}, list = [], colorOf = {};
      items.forEach(function (it, i) {
        if (!it) return;
        var key = (it.sheet + "!" + it.addr).toUpperCase();
        if (!colorOf[key]) colorOf[key] = COLORS[Object.keys(colorOf).length % COLORS.length];
        if (seen[key]) return;
        seen[key] = true;
        list.push({ sheet: it.sheet, addr: it.addr, label: label(it.sheet, it.addr, a.sheet), value: it.display, color: colorOf[key] });
      });
      renderTrace("Precedents of " + label(a.sheet, a.addr, ""), a.formula, list,
        "No references to cells in this workbook (links to other files are skipped).");
    });
  }

  async function scanDependents(ctx, a) {
    var wb = ctx.workbook;
    var sheets = wb.worksheets;
    sheets.load("items/name");
    var names = wb.names;
    names.load("items/name,items/type,items/formula");
    await ctx.sync();

    var nameMap = {};
    names.items.forEach(function (n) {
      if (n.type !== "Range") return;
      var f = String(n.formula).replace(/^=/, "");
      if (f.indexOf(",") >= 0) return;
      var p = FK.splitSheetAddress(f);
      var b = FK.parseA1(p.addr);
      if (b && p.sheet) nameMap[n.name.toUpperCase()] = { sheet: p.sheet, bounds: b };
    });
    var hasNames = Object.keys(nameMap).length > 0;
    var target = { r1: a.row, r2: a.row, c1: a.col, c2: a.col };
    var tgtU = a.sheet.toUpperCase();
    var hits = [];

    for (var s = 0; s < sheets.items.length; s++) {
      var ws = sheets.items[s];
      status("Scanning " + ws.name + "…");
      var ur = ws.getUsedRangeOrNullObject();
      ur.load("rowIndex,columnIndex,rowCount,columnCount");
      await ctx.sync();
      if (ur.isNullObject) continue;
      var step = Math.max(1, Math.floor(40000 / ur.columnCount));
      for (var r0 = 0; r0 < ur.rowCount; r0 += step) {
        var n = Math.min(step, ur.rowCount - r0);
        var blk = ws.getRangeByIndexes(ur.rowIndex + r0, ur.columnIndex, n, ur.columnCount);
        blk.load("formulas");
        await ctx.sync();
        for (var i = 0; i < n; i++) {
          for (var j = 0; j < ur.columnCount; j++) {
            var f = blk.formulas[i][j];
            if (typeof f !== "string" || f.charAt(0) !== "=") continue;
            if (ws.name !== a.sheet && !hasNames && f.toUpperCase().indexOf(tgtU) < 0) continue;
            var refs = FK.findRefs(f);
            for (var k = 0; k < refs.length; k++) {
              var t = refs[k], sh = null, b = null;
              if (t.kind === "address") { sh = t.sheet || ws.name; b = t.bounds; }
              else if (t.kind === "name" && nameMap[t.name.toUpperCase()]) {
                sh = nameMap[t.name.toUpperCase()].sheet; b = nameMap[t.name.toUpperCase()].bounds;
              }
              if (sh && sh.toUpperCase() === tgtU && FK.overlaps(b, target)) {
                hits.push({ sheet: ws.name, addr: FK.numToCol(ur.columnIndex + j + 1) + (ur.rowIndex + r0 + i + 1) });
                break;
              }
            }
          }
        }
      }
    }
    status("");
    return hits;
  }

  async function traceDependents() {
    showTab("trace");
    await showPane();
    await Excel.run(async function (ctx) {
      var a = await getActive(ctx);
      var list = null;
      try {
        var deps = a.range.getDirectDependents();
        deps.load("addresses");
        await ctx.sync();
        list = [];
        deps.addresses.forEach(function (s) {
          FK.splitAddressList(s).forEach(function (full) {
            var p = FK.splitSheetAddress(full);
            list.push({ sheet: p.sheet || a.sheet, addr: p.addr });
          });
        });
      } catch (e) {
        list = (e && e.code === "ItemNotFound") ? [] : null;
      }
      if (list === null) list = await scanDependents(ctx, a);

      list = list.slice(0, 500);
      var rngs = list.map(function (it) {
        var r = ctx.workbook.worksheets.getItem(it.sheet).getRange(it.addr).getCell(0, 0);
        r.load("text,values,valueTypes");
        return r;
      });
      await ctx.sync();
      list.forEach(function (it, i) {
        it.label = label(it.sheet, it.addr, a.sheet);
        it.value = fmtValue(rngs[i].values[0][0], rngs[i].valueTypes[0][0], rngs[i].text[0][0]);
      });
      renderTrace("Dependents of " + label(a.sheet, a.addr, ""), list.length ? list.length + " formula cell(s) use this cell." : "",
        list, "No formulas in this workbook use this cell.");
    });
  }

  async function jumpNext() {
    if (!S.list.length) { await tracePrecedents(); return; }
    await jumpTo(S.list, (S.idx + 1) % S.list.length);
  }

  async function goBack() {
    var item = S.back.pop();
    if (!item) { status("Nothing to go back to."); return; }
    await navigate(item, false);
    S.originPushed = false;
  }

  // ------------------------------------------------------------------ Explorer

  function highlight(formula, refs, colorOf) {
    var html = "", pos = 0;
    refs.forEach(function (r) {
      html += esc(formula.slice(pos, r.start));
      var c = colorOf[r.start];
      html += c ? '<b style="color:' + c + '">' + esc(r.text) + "</b>" : esc(r.text);
      pos = r.end;
    });
    return html + esc(formula.slice(pos));
  }

  async function explore() {
    showTab("explore");
    await showPane();
    await Excel.run(async function (ctx) {
      var a = await getActive(ctx);
      if (a.formula.charAt(0) !== "=") {
        $("exploreOut").innerHTML = '<div class="title">' + esc(label(a.sheet, a.addr, "")) +
          '</div><div class="empty">No formula. Value: ' + esc(fmtValue(a.value, a.type, a.text)) + "</div>";
        return;
      }
      var refs = FK.findRefs(a.formula);
      var items = await resolveRefs(ctx, a, refs);

      var colorOf = {}, keyColor = {}, list = [], seen = {};
      var sub = a.formula;
      for (var i = refs.length - 1; i >= 0; i--) {
        if (items[i] && items[i].cellCount === 1) {
          sub = sub.slice(0, refs[i].start) + FK.literal(items[i].value, items[i].type) + sub.slice(refs[i].end);
        }
      }
      items.forEach(function (it, i) {
        if (!it) return;
        var key = (it.sheet + "!" + it.addr).toUpperCase();
        if (!keyColor[key]) keyColor[key] = COLORS[Object.keys(keyColor).length % COLORS.length];
        colorOf[refs[i].start] = keyColor[key];
        if (!seen[key]) {
          seen[key] = true;
          list.push({ sheet: it.sheet, addr: it.addr, label: label(it.sheet, it.addr, a.sheet), value: it.display, color: keyColor[key] });
        }
      });

      S.explore = { a: a, refs: refs };
      var html = '<div class="title">' + esc(label(a.sheet, a.addr, "")) + "</div>" +
        '<div class="big">' + esc(fmtValue(a.value, a.type, a.text)) + "</div>" +
        '<div class="label">Formula</div><div class="code">' + highlight(a.formula, refs, colorOf) + "</div>" +
        '<div class="label">With values</div><div class="code">' + esc(sub) + "</div>";
      if (FK.functionParts(a.formula).length) {
        html += '<div class="row"><button class="btn" id="bParts">Evaluate each function</button></div><div id="partsOut"></div>';
      }
      html += '<div class="label">Inputs</div><div id="exploreList"></div>';
      $("exploreOut").innerHTML = html;

      // reuse the trace list so clicks and the jump shortcut work
      renderTrace("Precedents of " + label(a.sheet, a.addr, ""), a.formula, list, "No inputs in this workbook.");
      var lst = $("exploreList");
      lst.innerHTML = list.length ? $("traceOut").querySelector(".list").outerHTML : '<div class="empty">No inputs in this workbook.</div>';
      bindList(lst);
      if ($("bParts")) $("bParts").addEventListener("click", guard(evaluateParts));
    });
  }

  async function evaluateParts() {
    var ex = S.explore;
    if (!ex) return;
    var q = FK.qualify(ex.a.formula, ex.refs, ex.a.sheet);
    var parts = FK.functionParts(q).slice(0, 25);
    if (!parts.length) return;
    var prefix = FK.quoteSheet(ex.a.sheet) + "!";
    var results = [];

    await Excel.run(async function (ctx) {
      var tmp = ctx.workbook.worksheets.add("FK_tmp_" + Math.floor(Math.random() * 1e6));
      tmp.visibility = Excel.SheetVisibility.hidden;
      try {
        var cells = parts.map(function (p, i) {
          var c = tmp.getRange("A" + (i * 2000 + 1));
          c.formulas = [["=" + p.text]];
          c.load("values,valueTypes");
          return c;
        });
        await ctx.sync();
        results = cells.map(function (c) { return fmtValue(c.values[0][0], c.valueTypes[0][0], ""); });
      } finally {
        tmp.delete();
        await ctx.sync();
      }
    });

    var html = '<div class="label">Function results</div>';
    parts.forEach(function (p, i) {
      var t = p.text.split(prefix).join("").replace(/_xlfn\./gi, "");
      if (t.length > 70) t = t.slice(0, 67) + "…";
      html += '<div class="step" style="padding-left:' + (p.depth * 12) + 'px"><span>' + esc(t) +
        '</span><span class="r">' + esc(results[i]) + "</span></div>";
    });
    html += '<p class="note">Calculated on a temporary hidden sheet, which clears Excel\'s Undo history.</p>';
    $("partsOut").innerHTML = html;
  }

  // ------------------------------------------------------------------ Model map

  var LEGEND = [
    ["N", "Typed-in number (input)", "#DDEBF7", "#0000C0"],
    ["F", "Formula, start of a block", "#FFFFFF", "#000000"],
    ["<", "Same formula as the cell to the left", "#F2F2F2", "#808080"],
    ["^", "Same formula as the cell above", "#F2F2F2", "#808080"],
    ["L", "Formula linking to another sheet", "#E2EFDA", "#008000"],
    ["H", "Formula with a typed-in number", "#FFE699", "#9C5700"],
    ["X", "Breaks the pattern of its neighbours", "#FFC7CE", "#9C0006"],
    ["E", "Error value", "#C00000", "#FFFFFF"],
    ["T", "Text / labels", "#FFFFFF", "#A6A6A6"]
  ];

  function mapSheetName(src) {
    return ("Map " + src).replace(/[\\\/\?\*\[\]:]/g, "").slice(0, 31);
  }

  async function modelMap() {
    showTab("map");
    await showPane();
    await Excel.run(async function (ctx) {
      var src = ctx.workbook.worksheets.getActiveWorksheet();
      src.load("name");
      var ur = src.getUsedRangeOrNullObject();
      ur.load("rowIndex,columnIndex,rowCount,columnCount");
      await ctx.sync();
      if (ur.isNullObject) { status("This sheet is empty."); return; }
      if (/^Map /.test(src.name)) { status("This is already a map sheet. Go to the model sheet first."); return; }

      var nR = ur.rowCount, nC = ur.columnCount;
      if (nR * nC > 2000000) { status("This sheet's used range is too large to map (" + (nR * nC).toLocaleString() + " cells).", true); return; }

      var F = [], T = [];
      var step = Math.max(1, Math.floor(30000 / nC));
      for (var r0 = 0; r0 < nR; r0 += step) {
        status("Reading " + src.name + "… " + Math.round((r0 / nR) * 100) + "%");
        var n = Math.min(step, nR - r0);
        var blk = src.getRangeByIndexes(ur.rowIndex + r0, ur.columnIndex, n, nC);
        blk.load("formulasR1C1,valueTypes");
        await ctx.sync();
        Array.prototype.push.apply(F, blk.formulasR1C1);
        Array.prototype.push.apply(T, blk.valueTypes);
      }
      var res = FK.classifyGrid(F, T);

      var name = mapSheetName(src.name);
      var old = ctx.workbook.worksheets.getItemOrNullObject(name);
      await ctx.sync();
      if (!old.isNullObject) old.delete();
      var m = ctx.workbook.worksheets.add(name);
      await ctx.sync();

      for (r0 = 0; r0 < nR; r0 += step) {
        status("Writing map… " + Math.round((r0 / nR) * 100) + "%");
        var rows = res.codes.slice(r0, r0 + step);
        var out = m.getRangeByIndexes(ur.rowIndex + r0, ur.columnIndex, rows.length, nC);
        out.numberFormat = rows.map(function (row) { return row.map(function () { return "@"; }); });
        out.values = rows;
        await ctx.sync();
      }

      var mapRng = m.getRangeByIndexes(ur.rowIndex, ur.columnIndex, nR, nC);
      mapRng.format.horizontalAlignment = "Center";
      mapRng.format.font.size = 9;
      m.getRangeByIndexes(0, 0, 1, ur.columnIndex + nC).getEntireColumn().format.columnWidth = 16;
      LEGEND.forEach(function (l) {
        var cf = mapRng.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
        cf.cellValue.format.fill.color = l[2];
        cf.cellValue.format.font.color = l[3];
        cf.cellValue.rule = { formula1: '="' + l[0] + '"', operator: "EqualTo" };
      });
      await ctx.sync();
      status("");

      // review list
      var issues = [];
      for (var r = 0; r < nR && issues.length < 300; r++) {
        for (var c = 0; c < nC && issues.length < 300; c++) {
          var code = res.codes[r][c];
          if (code === "X" || code === "H" || code === "E") {
            issues.push({
              sheet: src.name,
              addr: FK.numToCol(ur.columnIndex + c + 1) + (ur.rowIndex + r + 1),
              code: code,
              formula: String(F[r][c])
            });
          }
        }
      }
      if (issues.length) {
        var fr = issues.map(function (it) { var c = src.getRange(it.addr); c.load("formulas"); return c; });
        await ctx.sync();
        issues.forEach(function (it, i) { it.formula = String(fr[i].formulas[0][0]); });
      }
      S.mapIssues = issues;

      var html = '<div class="title">' + esc(name) + '</div><div class="sub">Each cell sits at the same address as on ' + esc(src.name) + ".</div>" +
        '<table class="legend">';
      LEGEND.forEach(function (l) {
        html += '<tr><td><span class="chip" style="background:' + l[2] + ";color:" + l[3] + ';border:1px solid #ccc">' + esc(l[0]) +
          "</span></td><td>" + esc(l[1]) + '</td><td class="n">' + (res.counts[l[0]] || 0).toLocaleString() + "</td></tr>";
      });
      html += "</table>";
      html += '<div class="label">Cells to review (' + issues.length + (issues.length >= 300 ? "+" : "") + ")</div>";
      if (!issues.length) {
        html += '<div class="empty">No pattern breaks, hardcoded numbers in formulas, or errors found.</div>';
      } else {
        html += '<div class="list">';
        issues.forEach(function (it, i) {
          var l = LEGEND.filter(function (x) { return x[0] === it.code; })[0];
          html += '<button class="item" data-i="' + i + '"><span class="a"><span class="chip" style="background:' + l[2] + ";color:" + l[3] + '">' +
            it.code + "</span> " + esc(it.addr) + '</span><span class="v">' + esc(it.formula) + "</span></button>";
        });
        html += "</div>";
      }
      html += '<p class="note">Hardcode check ignores 0 and 1. X is a prompt to check, not proof of an error.</p>';
      $("mapOut").innerHTML = html;
      bindList($("mapOut"));
    });
  }

  async function removeMaps() {
    await Excel.run(async function (ctx) {
      var ws = ctx.workbook.worksheets;
      ws.load("items/name");
      await ctx.sync();
      var n = 0;
      ws.items.forEach(function (s) { if (/^Map /.test(s.name)) { s.delete(); n++; } });
      await ctx.sync();
      status(n ? "Removed " + n + " map sheet(s)." : "No map sheets found.");
    });
  }

  // ------------------------------------------------------------------ Formatting

  // Selection limited to the used range (so whole-column selections stay fast).
  async function workingRange(ctx) {
    var sel = ctx.workbook.getSelectedRange();
    sel.load("address,cellCount,rowCount,columnCount");
    await ctx.sync();
    if (sel.cellCount <= 50000) return sel;
    var used = sel.worksheet.getUsedRange();
    var r = sel.getIntersectionOrNullObject(used);
    r.load("address,cellCount,rowCount,columnCount");
    await ctx.sync();
    return r.isNullObject ? null : r;
  }

  async function cycleNumberFormat() {
    await Excel.run(async function (ctx) {
      var r = await workingRange(ctx);
      if (!r) return;
      var first = r.getCell(0, 0);
      first.load("numberFormat");
      await ctx.sync();
      var cur = String(first.numberFormat[0][0]);
      var idx;
      if (S.nf && S.nf.applied === cur) idx = S.nf.idx + 1;
      else idx = NUM_FORMATS.indexOf(cur) + 1;
      if (idx >= NUM_FORMATS.length) idx = 0;
      var fmt = NUM_FORMATS[idx];
      var arr = [];
      for (var i = 0; i < r.rowCount; i++) {
        var row = [];
        for (var j = 0; j < r.columnCount; j++) row.push(fmt);
        arr.push(row);
      }
      r.numberFormat = arr;
      first.load("numberFormat");
      await ctx.sync();
      S.nf = { idx: idx, applied: String(first.numberFormat[0][0]) };
      status("Number format: " + fmt);
    });
  }

  async function cycleFontColour() {
    await Excel.run(async function (ctx) {
      var sel = ctx.workbook.getSelectedRange();
      var first = sel.getCell(0, 0);
      first.format.font.load("color");
      await ctx.sync();
      var cur = String(first.format.font.color || "").toUpperCase();
      var nxt = cur === BLUE ? GREEN : cur === GREEN ? BLACK : BLUE;
      sel.format.font.color = nxt;
      await ctx.sync();
      status("Font colour: " + (nxt === BLUE ? "blue" : nxt === GREEN ? "green" : "black"));
    });
  }

  async function autoColour() {
    await Excel.run(async function (ctx) {
      var r = await workingRange(ctx);
      if (!r) return;
      r.worksheet.load("name");
      if (r.cellCount === 1) {
        r.load("formulas,valueTypes");
        await ctx.sync();
        var f = r.formulas[0][0], t = r.valueTypes[0][0];
        if (typeof f === "string" && f.charAt(0) === "=") r.format.font.color = f.indexOf("!") >= 0 ? GREEN : BLACK;
        else if (t === "Double" || t === "Integer") r.format.font.color = BLUE;
        await ctx.sync();
        return;
      }
      var nums = r.getSpecialCellsOrNullObject("Constants", "Numbers");
      var fs = r.getSpecialCellsOrNullObject("Formulas");
      nums.load("cellCount");
      fs.load("cellCount");
      await ctx.sync();
      if (!nums.isNullObject) nums.format.font.color = BLUE;
      var links = [];
      if (!fs.isNullObject) {
        fs.format.font.color = BLACK;
        fs.areas.load("items/formulas,items/rowIndex,items/columnIndex");
        await ctx.sync();
        fs.areas.items.forEach(function (area) {
          area.formulas.forEach(function (row, i) {
            row.forEach(function (f, j) {
              if (typeof f === "string" && f.indexOf("!") >= 0) {
                links.push(FK.numToCol(area.columnIndex + j + 1) + (area.rowIndex + i + 1));
              }
            });
          });
        });
      }
      var ws = ctx.workbook.worksheets.getItem(r.worksheet.name);
      for (var k = 0; k < links.length; k += 100) {
        ws.getRanges(links.slice(k, k + 100).join(",")).format.font.color = GREEN;
      }
      await ctx.sync();
      status("Coloured " + (nums.isNullObject ? 0 : nums.cellCount) + " inputs, " + links.length + " links, " +
        ((fs.isNullObject ? 0 : fs.cellCount) - links.length) + " formulas.");
    });
  }

  async function cycleBorders() {
    await Excel.run(async function (ctx) {
      var sel = ctx.workbook.getSelectedRange();
      sel.load("address,rowCount,columnCount");
      await ctx.sync();
      var state = (S.border && S.border.addr === sel.address) ? (S.border.state + 1) % 4 : 1;
      S.border = { addr: sel.address, state: state };
      var b = sel.format.borders;
      var edges = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
      if (sel.rowCount > 1) edges.push("InsideHorizontal");
      if (sel.columnCount > 1) edges.push("InsideVertical");
      edges.forEach(function (e) { b.getItem(e).style = "None"; });
      function line(e, style) { var it = b.getItem(e); it.style = style; if (style === "Continuous") it.weight = "Thin"; }
      if (state === 1) line("EdgeBottom", "Continuous");
      if (state === 2) { line("EdgeTop", "Continuous"); line("EdgeBottom", "Double"); }
      if (state === 3) ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"].forEach(function (e) { line(e, "Continuous"); });
      await ctx.sync();
      status(["Borders cleared", "Bottom line", "Total line", "Box"][state]);
    });
  }

  // ------------------------------------------------------------------ Wiring

  var ACTIONS = {
    FK_PRECEDENTS: tracePrecedents,
    FK_DEPENDENTS: traceDependents,
    FK_NEXT: jumpNext,
    FK_BACK: goBack,
    FK_EXPLORE: explore,
    FK_MAP: modelMap,
    FK_NUMFMT: cycleNumberFormat,
    FK_FONT: cycleFontColour,
    FK_AUTOCOLOUR: autoColour,
    FK_BORDERS: cycleBorders
  };

  function wireUi() {
    var navs = document.querySelectorAll("nav button");
    for (var i = 0; i < navs.length; i++) {
      navs[i].addEventListener("click", function (ev) { showTab(ev.currentTarget.getAttribute("data-tab")); });
    }
    var map = {
      bPrec: tracePrecedents, bDep: traceDependents, bBack: goBack, bExplore: explore,
      bMap: modelMap, bMapDel: removeMaps, bNum: cycleNumberFormat, bFont: cycleFontColour,
      bAuto: autoColour, bBorder: cycleBorders
    };
    Object.keys(map).forEach(function (id) { if ($(id)) $(id).addEventListener("click", guard(map[id])); });

    var html = "";
    KEYS.forEach(function (k) {
      html += "<tr><td>" + esc(k[0]) + "</td><td><kbd>⌘ ⌥ " + k[1] + "</kbd></td></tr>";
    });
    if ($("keysTable")) $("keysTable").innerHTML = html;
  }

  Office.onReady(function () {
    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", wireUi);
    else wireUi();

    Object.keys(ACTIONS).forEach(function (id) {
      try {
        Office.actions.associate(id, function () { return guard(ACTIONS[id])(); });
      } catch (e) { console.warn("associate failed", id, e); }
    });
    try { Office.addin.setStartupBehavior(Office.StartupBehavior.load); } catch (e) { /* older Excel */ }
  });

  // exposed for testing
  if (typeof window !== "undefined") window.FKApp = { S: S, showTab: showTab, renderTrace: renderTrace };
})();
