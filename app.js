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
    ["Trace precedents", "P"], ["Trace dependents", "U"], ["Jump to next in list", "J"],
    ["Go back", "B"], ["Formula explorer", "E"], ["Model map", "G"],
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
        out.values = rng.values;
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

  function listHtml(list, home) {
    var html = '<div class="grid">', last = null;
    var multi = list.some(function (x) { return x.sheet !== home; });
    list.forEach(function (it, i) {
      if (multi && it.sheet !== last) { html += '<div class="gh">' + esc(it.sheet) + "</div>"; last = it.sheet; }
      html += '<button class="item gr" data-i="' + i + '">' +
        '<span class="sq" style="background:' + (it.color || "var(--muted)") + '"></span>' +
        '<span class="gl">' + (it.name ? '<span class="lbl">' + esc(it.name) + '</span><span class="adr">' + esc(cleanAddr(it.addr)) + "</span>"
                                       : '<span class="lbl mono">' + esc(cleanAddr(it.addr)) + "</span>") + "</span>" +
        '<span class="gv">' + esc(it.value || "") + "</span></button>";
    });
    return html + "</div>";
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
      var btns = document.querySelectorAll("#" + S.listBox + " .item");
      for (var k = 0; k < btns.length; k++) btns[k].classList.toggle("cur", k === i);
    }
  }

  async function tracePrecedents() {
    await analyze();
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
        it.value = fmtValue(rngs[i].values[0][0], rngs[i].valueTypes[0][0], rngs[i].text[0][0]);
      });
      await loadLabels(ctx, [a].concat(list.slice(0, 150)));
      a.name = nameFor(a, null);
      list.forEach(function (it) { it.name = nameFor(it, null); });
      $("traceOut").innerHTML = cellCard(a) +
        '<div class="sec"><span>Used by (' + list.length + ')</span><span class="hint">click to jump · ⌘⌥J next · ⌘⌥B back</span></div>' +
        '<div id="fkList"></div>';
      setList(sortBySheet(list, a.sheet), "fkList", a.sheet);
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

  // ------------------------------------------------------------------ Explorer (precedents + breakdown)

  var LETTER = /[A-Za-zÀ-ɏЀ-ӿ]/;
  var DATEISH = /^(19|20)\d{2}[A-Z]?$|^\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}$|^[A-Za-zЀ-ӿ]{3,9}[\s\-']\d{2,4}$|^(FY|Q[1-4]|H[12])/i;

  function shorten(s, n) {
    s = String(s).replace(/\s+/g, " ").trim();
    return s.length > n ? s.slice(0, n - 1) + "…" : s;
  }

  // Row label = the text cells to the left of the cell (up to 3, joined),
  // column header = nearest text or date-like cell above.
  async function loadLabels(ctx, cells) {
    var jobs = cells.map(function (c) {
      var b = FK.parseA1(c.addr);
      if (!b || !c.sheet) return null;
      var ws = ctx.workbook.worksheets.getItem(c.sheet);
      var j = { c: c, b: b, row: null, col: null };
      if (b.c1 > 1) {
        j.row = ws.getRangeByIndexes(b.r1 - 1, 0, 1, Math.min(b.c1 - 1, 30));
        j.row.load("values,valueTypes");
      }
      if (b.r1 > 1) {
        var top = Math.max(0, b.r1 - 1 - 40);
        j.col = ws.getRangeByIndexes(top, b.c1 - 1, b.r1 - 1 - top, 1);
        j.col.load("text,valueTypes");
      }
      return j;
    });
    await ctx.sync();
    jobs.forEach(function (j) {
      if (!j) return;
      j.c.rowLabel = "";
      j.c.colHeader = "";
      if (j.row) {
        var vals = j.row.values[0], types = j.row.valueTypes[0], texts = [];
        for (var i = 0; i < vals.length; i++) {
          var t = String(vals[i]).replace(/\s+/g, " ").trim();
          if (types[i] === "String" && LETTER.test(t) && texts.indexOf(t) < 0) texts.push(shorten(t, 45));
        }
        if (texts.length > 3) texts = [texts[0]].concat(texts.slice(-2));
        j.c.rowLabel = texts.join(" › ");
      }
      if (j.col) {
        for (var r = j.col.text.length - 1; r >= 0; r--) {
          var h = String(j.col.text[r][0]).trim(), ty = j.col.valueTypes[r][0];
          if (!h) continue;
          if ((ty === "String" && h.length <= 40) || ((ty === "Double" || ty === "Integer") && DATEISH.test(h))) {
            j.c.colHeader = shorten(h, 25);
            break;
          }
        }
      }
    });
  }

  function nameFor(it, home) {
    var parts = [];
    var rb = FK.parseA1(it.addr);
    if (rb && rb.r2 > rb.r1) {
      if (rb.c2 === rb.c1 && it.colHeader) parts.push(it.colHeader);
      parts.push("rows " + rb.r1 + "–" + rb.r2);
      return parts.join(" · ");
    }
    var b = FK.parseA1(it.addr), hb = home ? FK.parseA1(home.addr) : null;
    var sameCol = b && hb && b.c1 === hb.c1 && it.sheet === home.sheet;
    if (it.rowLabel) parts.push(it.rowLabel);
    if (it.colHeader && (!sameCol || !it.rowLabel)) parts.push(it.colHeader);
    return parts.join(" · ");
  }

  // Numbers for computed steps: sensible precision by size.
  function fmtComp(v) {
    if (v === undefined) return "?";
    if (FK.isErr(v)) return v.err;
    if (FK.isRange(v)) return v.range.length + "×" + v.range[0].length + " range";
    if (v && v.missing) return "—";
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    if (typeof v === "number") {
      var a = Math.abs(v);
      var d = a >= 1000 ? 0 : a >= 1 ? 2 : a === 0 ? 0 : 4;
      return v.toLocaleString(undefined, { maximumFractionDigits: d, minimumFractionDigits: 0 });
    }
    return '"' + shorten(v, 30) + '"';
  }

  function cleanAddr(a) { return String(a).replace(/\$/g, ""); }

  // Compact expression: refs become coloured chips (address or label), * / become × ÷.
  function exprHtml(ex, s, e, mode) {
    var f = ex.a.formula, html = "", pos = s;
    function plain(t) {
      return esc(t).replace(/\*/g, " × ").replace(/\//g, " ÷ ").replace(/\+/g, " + ").replace(/,/g, ", ")
        .replace(/_xl(fn|ws)\./gi, "").replace(/\s{2,}/g, " ");
    }
    ex.refs.forEach(function (r, i) {
      if (r.start < s || r.end > e) return;
      html += plain(f.slice(pos, r.start));
      var it = ex.items[i], color = ex.colorOf[r.start] || "var(--muted)";
      var addr = it ? cleanAddr(it.addr) : cleanAddr(r.text);
      var other = it && it.sheet !== ex.a.sheet;
      var txt = mode === "labels" && it && it.name ? it.name : (other ? it.sheet + "!" : "") + addr;
      html += '<span class="rc" style="--c:' + color + '" title="' + esc((it ? it.sheet + "!" : "") + addr + (it && it.name ? " — " + it.name : "")) + '">' + esc(txt) + "</span>";
      pos = r.end;
    });
    return html + plain(f.slice(pos, e));
  }

  var OP_LABEL = { "+": "+", "-": "−", "*": "×", "/": "÷", "^": "^", "&": "&", "=": "=", "<>": "≠", "<": "<", ">": ">", "<=": "≤", ">=": "≥" };
  var ARG_NAMES = {
    IF: ["if", "then", "else"], IFERROR: ["value", "if error"], ROUND: ["", "digits"], ROUNDUP: ["", "digits"],
    ROUNDDOWN: ["", "digits"], VLOOKUP: ["look up", "in", "column", "exact?"], XLOOKUP: ["look up", "in", "return"],
    INDEX: ["from", "row", "column"], MATCH: ["look up", "in", "type"], SUMIF: ["range", "criteria", "sum"],
    SUMIFS: ["sum", "range", "criteria"], MIN: [], MAX: []
  };
  function group(op) { return op === "+" || op === "-" ? "add" : op === "*" || op === "/" ? "mul" : null; }
  function core(n) { while (n.k === "paren") n = n.x; return n; }

  function childrenOf(n) {
    var c = core(n), out = [];
    if (c.k === "bin" && group(c.op)) {
      var g = group(c.op);
      (function collect(m, op) {
        if (m.k === "bin" && group(m.op) === g) { collect(m.l, op); out.push({ tag: OP_LABEL[m.op], n: m.r, isOp: true }); }
        else out.push({ tag: op, n: m, isOp: true });
      })(c, "");
    } else if (c.k === "bin") {
      out.push({ tag: "", n: c.l, isOp: true });
      out.push({ tag: OP_LABEL[c.op] || c.op, n: c.r, isOp: true });
    } else if (c.k === "fn") {
      var names = ARG_NAMES[c.name] || [];
      c.args.forEach(function (arg, i) { if (arg.k !== "missing") out.push({ tag: names[i] || "", n: arg }); });
    } else if (c.k === "unary") out.push({ tag: c.op === "-" ? "−" : "+", n: c.x, isOp: true });
    else if (c.k === "pct") out.push({ tag: "%", n: c.x, isOp: true });
    return out;
  }

  // One level of the calculation: leaves shown as cell chips / numbers, deeper parts as (…) or NAME(…).
  function stepHtml(ex, n) {
    var c = core(n), kids = childrenOf(n);
    function piece(m) {
      var k = core(m);
      if (k.k === "ref" || k.k === "num" || k.k === "str" || k.k === "bool" || k.k === "err") return exprHtml(ex, m.s, m.e, "cells");
      if (k.k === "fn") return "<b>" + esc(k.name) + "</b>(…)";
      return "(…)";
    }
    var html;
    if (c.k === "fn") {
      html = "<b>" + esc(c.name) + "</b>(" + kids.map(function (k) { return piece(k.n); }).join(", ") + ")";
    } else {
      html = kids.map(function (k) { return (k.tag ? " " + esc(k.tag) + " " : "") + piece(k.n); }).join("");
    }
    return n.k === "paren" ? "(" + html + ")" : html;
  }

  function renderBreakdown() {
    var ex = S.explore, out = $("fkTree");
    if (!out) return;
    if (!ex.root) {
      out.innerHTML = '<div class="empty pad">This formula is too complex to break down. The inputs are listed below.</div>';
      return;
    }
    var html = '<div class="th"><span>Calculation</span><span>Value</span></div>', unknown = 0, count = 0;
    function walk(n, depth, ch) {
      if (count++ > 200) return;
      var c = core(n), kids = childrenOf(n), open = !ex.collapsed.has(n), val = n.val, body;
      if (val === undefined && c.k !== "unk") unknown++;
      var shown = fmtComp(val);
      if (c.k === "ref") {
        var it = ex.items[c.ref], color = ex.colorOf[ex.refs[c.ref].start] || "var(--muted)";
        var addr = it ? (it.sheet !== ex.a.sheet ? it.sheet + "!" : "") + cleanAddr(it.addr) : cleanAddr(ex.refs[c.ref].text);
        body = '<span class="sq" style="background:' + color + '"></span><span class="gl">' +
          (it && it.name ? '<span class="lbl">' + esc(it.name) + '</span><span class="adr">' + esc(addr) + "</span>"
                         : '<span class="lbl mono">' + esc(addr) + "</span>") + "</span>";
        if (it && it.cellCount === 1 && it.text && it.text.indexOf("##") < 0) shown = it.text;
      } else if (c.k === "num") {
        body = '<span class="mono">' + esc(String(c.v)) + "</span>" + (c.v !== 0 && c.v !== 1 ? ' <span class="hc">typed-in</span>' : "");
        shown = "";
      } else if (c.k === "str" || c.k === "bool" || c.k === "err") {
        body = '<span class="mono">' + esc(ex.a.formula.slice(c.s, c.e)) + "</span>";
        shown = "";
      } else {
        body = '<span class="ex">' + (open && kids.length ? stepHtml(ex, n) : exprHtml(ex, n.s, n.e, "cells")) + "</span>";
      }
      var tag = ch && ch.tag ? '<span class="' + (ch.isOp ? "opx" : "tag") + '">' + esc(ch.tag) + "</span>" : "";
      html += '<div class="tr' + (depth === 0 ? " root" : "") + '" style="--d:' + Math.min(depth, 8) + '">' +
        '<span class="tog" data-n="' + ex.nodes.length + '">' + (kids.length ? (open ? "▾" : "▸") : "") + "</span>" +
        tag + '<span class="tb">' + body + "</span>" +
        '<span class="tv' + (FK.isErr(val) ? " bad" : "") + '">' + esc(shown) + "</span></div>";
      ex.nodes.push(n);
      if (open) kids.forEach(function (k) { walk(k.n, depth + 1, k); });
    }
    ex.nodes = [];
    walk(ex.root, 0, null);
    if (unknown) {
      html += '<div class="foot"><button class="btn sm" id="fkCalc">Calculate the “?” steps in Excel</button>' +
        '<span class="note">Uses a temporary hidden sheet (clears Undo).</span></div>';
    }
    out.innerHTML = html;
    var togs = out.querySelectorAll(".tog");
    for (var i = 0; i < togs.length; i++) {
      togs[i].addEventListener("click", function (ev) {
        var node = S.explore.nodes[+ev.currentTarget.getAttribute("data-n")];
        if (!node || !childrenOf(node).length) return;
        if (S.explore.collapsed.has(node)) S.explore.collapsed.delete(node); else S.explore.collapsed.add(node);
        renderBreakdown();
      });
    }
    if ($("fkCalc")) $("fkCalc").addEventListener("click", guard(fillUnknown));
  }

  function renderFormula() {
    var ex = S.explore;
    $("fkFormula").innerHTML = "= " + exprHtml(ex, 1, ex.a.formula.length, ex.mode);
    var btns = document.querySelectorAll("#fkMode button");
    for (var i = 0; i < btns.length; i++) btns[i].classList.toggle("on", btns[i].getAttribute("data-m") === ex.mode);
  }

  function allNodes(root) {
    var out = [];
    (function walk(n) { out.push(n); childrenOf(n).forEach(function (k) { walk(k.n); }); })(root);
    return out;
  }

  async function fillUnknown() {
    var ex = S.explore;
    var nodes = allNodes(ex.root).filter(function (n) { return n.val === undefined && n.k !== "unk" && n.k !== "missing"; }).slice(0, 40);
    if (!nodes.length) return;
    await Excel.run(async function (ctx) {
      var tmp = ctx.workbook.worksheets.add("FK_tmp_" + Math.floor(Math.random() * 1e6));
      tmp.visibility = Excel.SheetVisibility.hidden;
      try {
        var cells = nodes.map(function (n, i) {
          var c = tmp.getRange("A" + (i * 1000 + 1));
          c.formulas = [["=" + FK.qualifySlice(ex.a.formula, ex.refs, n.s, n.e, ex.a.sheet)]];
          c.load("values,valueTypes");
          return c;
        });
        await ctx.sync();
        nodes.forEach(function (n, i) {
          var v = cells[i].values[0][0], t = cells[i].valueTypes[0][0];
          n.val = t === "Error" ? { err: String(v) } : t === "Empty" ? 0 : v;
        });
      } finally {
        tmp.delete();
        await ctx.sync();
      }
    });
    renderBreakdown();
  }

  function cellCard(a) {
    return '<div class="card"><div class="cl"><div class="ca"><span class="addr">' + esc(cleanAddr(a.addr)) + '</span><span class="shn">' +
      esc(a.sheet) + "</span></div>" + (a.name ? '<div class="cn">' + esc(a.name) + "</div>" : "") +
      '</div><div class="cv">' + esc(fmtValue(a.value, a.type, a.text)) + "</div></div>";
  }

  function sortBySheet(list, home) {
    return list.map(function (it, i) { return { it: it, i: i }; }).sort(function (x, y) {
      var hx = x.it.sheet === home ? 0 : 1, hy = y.it.sheet === home ? 0 : 1;
      if (hx !== hy) return hx - hy;
      if (x.it.sheet !== y.it.sheet) return x.it.sheet < y.it.sheet ? -1 : 1;
      return x.i - y.i;
    }).map(function (x) { return x.it; });
  }

  function setList(list, boxId, home) {
    S.list = list;
    S.listBox = boxId;
    S.idx = -1;
    S.originPushed = false;
    var el = $(boxId);
    el.innerHTML = list.length ? listHtml(list, home) : '<div class="empty pad">None in this workbook (links to other files are skipped).</div>';
    bindList(el);
  }

  async function analyze() {
    showTab("trace");
    await showPane();
    await Excel.run(async function (ctx) {
      var a = await getActive(ctx);
      var outEl = $("traceOut");
      await loadLabels(ctx, [a]);
      a.name = nameFor(a, null);
      if (a.formula.charAt(0) !== "=") {
        outEl.innerHTML = cellCard(a) + '<div class="empty pad">No formula here — this is a typed-in value. Use Dependents (⌘⌥U) to see where it is used.</div>';
        S.list = [];
        return;
      }
      var refs = FK.findRefs(a.formula);
      var items = await resolveRefs(ctx, a, refs);
      var uniq = [], byKey = {};
      items.forEach(function (it) {
        if (!it) return;
        var key = (it.sheet + "!" + it.addr).toUpperCase();
        if (!byKey[key]) { byKey[key] = it; uniq.push(it); }
      });
      await loadLabels(ctx, uniq);

      var colorOf = {}, keyColor = {}, list = [];
      items.forEach(function (it, i) {
        if (!it) return;
        var key = (it.sheet + "!" + it.addr).toUpperCase(), src = byKey[key];
        it.rowLabel = src.rowLabel; it.colHeader = src.colHeader;
        it.name = nameFor(it, a);
        if (!keyColor[key]) {
          keyColor[key] = COLORS[Object.keys(keyColor).length % COLORS.length];
          list.push({ sheet: it.sheet, addr: it.addr, name: it.name, value: it.display, color: keyColor[key] });
        }
        colorOf[refs[i].start] = keyColor[key];
      });

      var root = FK.parseFormula(a.formula, refs);
      if (root) {
        FK.evalTree(root, function (i) {
          var it = items[i];
          if (!it) return undefined;
          if (it.cellCount === 1) {
            if (it.type === "Empty") return 0;
            if (it.type === "Error") return { err: String(it.value) };
            return it.value;
          }
          return it.values ? { range: it.values } : undefined;
        });
      }
      var prevMode = S.explore && S.explore.mode;
      S.explore = { a: a, refs: refs, items: items, root: root, colorOf: colorOf, mode: prevMode || "labels", collapsed: new Set(), nodes: [] };

      outEl.innerHTML = cellCard(a) +
        '<div class="sec"><span>Formula</span><span class="seg" id="fkMode"><button data-m="labels">Labels</button><button data-m="cells">Cells</button></span></div>' +
        '<div class="fbox" id="fkFormula"></div>' +
        '<div class="sec"><span>Breakdown</span><span class="hint">▸ click to fold</span></div>' +
        '<div class="tree" id="fkTree"></div>' +
        '<div class="sec"><span>Inputs (' + list.length + ')</span><span class="hint">click to jump · ⌘⌥J next · ⌘⌥B back</span></div>' +
        '<div id="fkList"></div>';
      renderFormula();
      renderBreakdown();
      var mb = document.querySelectorAll("#fkMode button");
      for (var k = 0; k < mb.length; k++) {
        mb[k].addEventListener("click", function (ev) {
          S.explore.mode = ev.currentTarget.getAttribute("data-m");
          renderFormula();
          renderBreakdown();
        });
      }
      setList(sortBySheet(list, a.sheet), "fkList", a.sheet);
    });
  }

  async function explore() {
    await analyze();
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
  if (typeof window !== "undefined") window.FKApp = { S: S, showTab: showTab };
})();
