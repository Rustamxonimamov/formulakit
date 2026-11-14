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
    ["Explore precedents", "P"], ["Explore dependents", "U"], ["Jump to next in list", "J"],
    ["Go back", "B"], ["Explore (same as P)", "E"], ["Model map", "G"],
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
      if (rng && !rng.isNullObject && rng.cellCount <= 5000) rng.load("values,valueTypes,text,formulas");
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
        var f0 = rng.formulas[0][0];
        out.formula = typeof f0 === "string" && f0.charAt(0) === "=" ? f0 : "";
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
      if (String(item.addr).indexOf(",") >= 0) ws.getRanges(item.addr).select();   // several separate cells
      else ws.getRange(item.addr).select();
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

  // ------------------------------------------------------------------ Explorer (Arixcel-style logical tree)
  //
  // Root = the selected cell. Below it, the formula's logical tree: functions with their
  // Excel argument names (Info), every element's value, and where it points (Location).
  // Moving the selection selects that location in the sheet.

  var X = { mode: "prec", root: null, cur: null, rows: [], pushed: false };

  function cleanAddr(a) { return String(a).replace(/\$/g, ""); }

  var ARGSPEC = {
    IF: "logical_test,value_if_true,value_if_false", IFERROR: "value,value_if_error", IFNA: "value,value_if_na",
    OFFSET: "reference,rows,cols,height,width", INDEX: "array,row_num,column_num", MATCH: "lookup_value,lookup_array,match_type",
    VLOOKUP: "lookup_value,table_array,col_index_num,range_lookup", HLOOKUP: "lookup_value,table_array,row_index_num,range_lookup",
    XLOOKUP: "lookup_value,lookup_array,return_array,if_not_found,match_mode,search_mode",
    SUMIF: "range,criteria,sum_range", COUNTIF: "range,criteria", AVERAGEIF: "range,criteria,average_range",
    SUMIFS: "sum_range,(criteria_range,criteria)*", COUNTIFS: "(criteria_range,criteria)*", AVERAGEIFS: "average_range,(criteria_range,criteria)*",
    MAXIFS: "max_range,(criteria_range,criteria)*", MINIFS: "min_range,(criteria_range,criteria)*",
    ROUND: "number,num_digits", ROUNDUP: "number,num_digits", ROUNDDOWN: "number,num_digits", MOD: "number,divisor",
    POWER: "number,power", ABS: "number", INT: "number", SQRT: "number", SIGN: "number", EXP: "number", LN: "number",
    EOMONTH: "start_date,months", EDATE: "start_date,months", DATE: "year,month,day", YEAR: "serial_number",
    MONTH: "serial_number", DAY: "serial_number", LEFT: "text,num_chars", RIGHT: "text,num_chars", MID: "text,start_num,num_chars",
    LEN: "text", UPPER: "text", LOWER: "text", TRIM: "text", NOT: "logical", TEXT: "value,format_text", NPV: "rate,value*",
    PMT: "rate,nper,pv,fv,type", CHOOSE: "index_num,value*", SUM: "number*", PRODUCT: "number*", AVERAGE: "number*",
    MIN: "number*", MAX: "number*", MEDIAN: "number*", COUNT: "value*", COUNTA: "value*", SUMPRODUCT: "array*",
    AND: "logical*", OR: "logical*", CONCAT: "text*", CONCATENATE: "text*", ISNUMBER: "value", ISERROR: "value",
    ISTEXT: "value", ISNA: "value", ROWS: "array", COLUMNS: "array", IFS: "(logical_test,value_if_true)*"
  };

  function argName(fn, i) {
    var spec = ARGSPEC[fn];
    if (!spec) return "";
    var parts = spec.split(/,(?![^(]*\))/), fixed = [];
    for (var k = 0; k < parts.length; k++) {
      var p = parts[k];
      if (/\*$/.test(p)) {
        var rest = i - fixed.length;
        if (rest < 0) return fixed[i];
        if (p.charAt(0) === "(") {
          var pair = p.slice(1, -2).split(",");
          return pair[rest % 2] + (Math.floor(rest / 2) + 1);
        }
        return p.slice(0, -1) + (rest + 1);
      }
      fixed.push(p);
    }
    return fixed[i] || "";
  }

  function core(n) { while (n && n.k === "paren") n = n.x; return n; }
  function group(op) { return op === "+" || op === "-" ? "add" : op === "*" || op === "/" ? "mul" : null; }
  function isLeafExpr(n) { var c = core(n); return c.k === "num" || c.k === "str" || c.k === "bool" || c.k === "err" || c.k === "unk" || c.k === "missing"; }

  // ---- cell contexts (one per formula cell shown in the tree)

  function itemVal(it) {
    if (!it) return undefined;
    if (it.cellCount === 1) {
      if (it.type === "Empty") return 0;
      if (it.type === "Error") return { err: String(it.value) };
      return it.value;
    }
    return it.values ? { range: it.values } : undefined;
  }

  async function makeCx(ctx, sheet, addr) {
    var c = ctx.workbook.worksheets.getItem(sheet).getRange(addr);
    c.load("formulas,values,valueTypes,text,rowIndex");
    await ctx.sync();
    var cx = { sheet: sheet, addr: cleanAddr(addr), formula: "", raw: c.formulas[0][0], text: fmtValue(c.values[0][0], c.valueTypes[0][0], c.text[0][0]),
               refs: [], items: [], root: null, cache: {} };
    var f = c.formulas[0][0];
    if (typeof f === "string" && f.charAt(0) === "=") {
      cx.formula = f;
      cx.refs = FK.findRefs(f);
      cx.items = await resolveRefs(ctx, { sheet: sheet, addr: cx.addr, row: c.rowIndex + 1 }, cx.refs);
      cx.root = FK.parseFormula(f, cx.refs);
      if (cx.root) await evalCx(ctx, cx);
    }
    return cx;
  }

  async function evalCx(ctx, cx) {
    for (var round = 0; round < 4; round++) {
      var need = {};
      FK.evalTree(cx.root, {
        getRef: function (i) { return itemVal(cx.items[i]); },
        refLoc: function (i) { var it = cx.items[i]; var b = it && FK.parseA1(it.addr); return b ? { sheet: it.sheet, b: b } : null; },
        cellValue: function (loc) {
          var k = FK.locKey(loc);
          if (k in cx.cache) return cx.cache[k];
          need[k] = loc;
          return undefined;
        }
      });
      var keys = Object.keys(need);
      if (!keys.length) return;
      var rs = keys.map(function (k) {
        var r = ctx.workbook.worksheets.getItem(need[k].sheet).getRange(FK.locAddr(need[k]));
        r.load("cellCount");
        return r;
      });
      await ctx.sync();
      rs.forEach(function (r) { if (r.cellCount <= 5000) r.load("values,valueTypes"); });
      await ctx.sync();
      keys.forEach(function (k, i) {
        var r = rs[i];
        if (r.cellCount > 5000) { cx.cache[k] = undefined; return; }
        if (r.cellCount === 1) {
          var t = r.valueTypes[0][0], v = r.values[0][0];
          cx.cache[k] = t === "Empty" ? 0 : t === "Error" ? { err: String(v) } : v;
        } else cx.cache[k] = { range: r.values };
      });
    }
  }

  // ---- display nodes

  function dnode(parent, props) {
    var d = { parent: parent, depth: parent ? parent.depth + 1 : 0, kids: null, open: false };
    for (var k in props) d[k] = props[k];
    return d;
  }

  function exprNode(parent, cx, en, info) {
    var c = core(en), label;
    if (c.k === "fn") label = c.name;
    else if (c.k === "missing") label = "<missing>";
    else label = cx.formula.slice(en.s, en.e).replace(/^_xl(fn|ws)\./i, "");
    var it = c.k === "ref" ? cx.items[c.ref] : null;
    var open = false;
    if (c.k === "fn") open = c.args.length > 0;
    else if (c.k === "bin" || c.k === "pct") open = true;
    else if (c.k === "unary") open = !(isLeafExpr(c.x) || core(c.x).k === "ref");
    else if (c.k === "ref") open = !!(it && it.cellCount === 1 && it.formula);
    return dnode(parent, { kind: "expr", cx: cx, en: en, c: c, label: label, info: info || "", item: it, canOpen: open });
  }

  function exprKids(d) {
    var c = d.c, cx = d.cx, out = [];
    if (c.k === "fn" && IFS_GROUP[c.name]) {
      var g = IFS_GROUP[c.name](c.args.length);
      g.forEach(function (item) {
        if (item.arg !== undefined) { if (c.args[item.arg] && c.args[item.arg].k !== "missing") out.push(exprNode(d, cx, c.args[item.arg], argName(c.name, item.arg))); }
        else out.push(filterNode(d, cx, c, item.pairs, item.locIdx));
      });
    } else if (c.k === "fn") {
      c.args.forEach(function (a, i) { out.push(exprNode(d, cx, a, argName(c.name, i))); });
    } else if (c.k === "bin" && group(c.op)) {
      var g = group(c.op);
      (function collect(m) {
        var mc = core(m);
        if (m.k === "bin" && group(m.op) === g) { collect(m.l); collect(m.r); }
        else out.push(exprNode(d, cx, m, ""));
      })(c);
    } else if (c.k === "bin") {
      out.push(exprNode(d, cx, c.l, ""));
      out.push(exprNode(d, cx, c.r, ""));
    } else if (c.k === "unary" || c.k === "pct") {
      var x = c.x;
      while (core(x).k === "unary") x = core(x).x;   // --(...) opens straight to its contents
      out.push(exprNode(d, cx, x, ""));
    }
    return out;
  }

  // SUMIF-style functions: the (range, criteria) pairs collapse into one "Filter" row, like Arixcel.
  function pairsFrom(start, n) { var p = []; for (var i = start; i + 1 < n; i += 2) p.push([i, i + 1]); return p; }
  var IFS_GROUP = {
    SUMIF: function (n) { return [{ pairs: [[0, 1]], locIdx: 0 }].concat(n > 2 ? [{ arg: 2 }] : []); },
    AVERAGEIF: function (n) { return [{ pairs: [[0, 1]], locIdx: 0 }].concat(n > 2 ? [{ arg: 2 }] : []); },
    COUNTIF: function () { return [{ pairs: [[0, 1]], locIdx: 0 }]; },
    COUNTIFS: function (n) { return [{ pairs: pairsFrom(0, n), locIdx: 0 }]; },
    SUMIFS: function (n) { return [{ arg: 0 }, { pairs: pairsFrom(1, n), locIdx: 1 }]; },
    AVERAGEIFS: function (n) { return [{ arg: 0 }, { pairs: pairsFrom(1, n), locIdx: 1 }]; },
    MAXIFS: function (n) { return [{ arg: 0 }, { pairs: pairsFrom(1, n), locIdx: 1 }]; },
    MINIFS: function (n) { return [{ arg: 0 }, { pairs: pairsFrom(1, n), locIdx: 1 }]; }
  };

  function filterNode(parent, cx, fnNode, pairs, locIdx) {
    var an = fnNode.args, first = an[pairs[0][0]], last = an[pairs[pairs.length - 1][1]];
    var loc = FK.matchLoc(an, pairs, locIdx);
    return dnode(parent, { kind: "filter", cx: cx, fn: fnNode, pairs: pairs, label: "Filter", info: "", canOpen: true,
                           c: { k: "filter" }, en: { s: first.s, e: last.e, loc: loc } });
  }

  // Children directly under a cell row: top-level terms of its formula.
  function cellKids(d) {
    var cx = d.cx;
    if (!cx.formula) return [];
    if (!cx.root) return cx.refs.map(function (r, i) { return dnode(d, { kind: "expr", cx: cx, en: { k: "ref", ref: i, s: r.start, e: r.end, val: itemVal(cx.items[i]) }, c: { k: "ref", ref: i }, label: r.text, info: "", item: cx.items[i], canOpen: false }); });
    var c = core(cx.root);
    if (c.k === "bin" && group(c.op)) {
      var fake = { parent: d, depth: d.depth, cx: cx, c: c };
      return exprKids(fake).map(function (k) { k.parent = d; k.depth = d.depth + 1; return k; });
    }
    return [exprNode(d, cx, cx.root, "")];
  }

  async function openNode(d) {
    if (!d.kids) {
      if (d.kind === "filter") {
        d.kids = [];
        d.pairs.forEach(function (p) {
          d.kids.push(exprNode(d, d.cx, d.fn.args[p[0]], argName(d.fn.name, p[0])));
          d.kids.push(exprNode(d, d.cx, d.fn.args[p[1]], argName(d.fn.name, p[1])));
        });
      } else if (d.kind === "cell") d.kids = cellKids(d);
      else if (d.kind === "dep") d.kids = await depKids(d);
      else if (d.c.k === "ref") {
        var it = d.item;
        await Excel.run(async function (ctx) {
          var cx = await makeCx(ctx, it.sheet, it.addr);
          var holder = { depth: d.depth, cx: cx };
          d.kids = cellKids(holder).map(function (k) { k.parent = d; k.depth = d.depth + 1; return k; });
        });
      } else d.kids = exprKids(d);
    }
    d.open = true;
  }

  // ---- values / locations for display

  function fmtX(v) {
    if (v === undefined) return "?";
    if (FK.isErr(v)) return v.err;
    if (FK.isRange(v)) return v.filter ? "<filter>" : "<array>";
    if (v && v.missing) return "";
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    if (typeof v === "number") {
      var a = Math.abs(v);
      if (Number.isInteger(v)) return v.toLocaleString("en-US");
      return v.toLocaleString("en-US", { minimumFractionDigits: a >= 100 ? 2 : 0, maximumFractionDigits: a >= 100 ? 2 : a >= 1 ? 3 : 4 });
    }
    return String(v);
  }

  function valueText(d) {
    if (d.kind === "filter") return "<filter>";
    if (d.kind === "cell" || d.kind === "dep") return d.cx ? d.cx.text : d.text;
    var c = d.c, v = d.en.val;
    if (c.k === "missing") return d.parent && d.parent.c && d.parent.c.name === "OFFSET" ? "0" : "";
    if (c.k === "ref" && d.item && d.item.cellCount === 1 && d.item.text && d.item.text.indexOf("##") < 0) return d.item.text;
    if (c.k === "str") return String(v);
    return fmtX(v);
  }

  function locOf(d) {
    if (d.kind === "cell" || d.kind === "dep") return { sheet: d.cx ? d.cx.sheet : d.sheet, addr: d.cx ? d.cx.addr : d.addr };
    var l = d.en.loc;
    if (!l && d.c.k === "ref" && d.item) return { sheet: d.item.sheet, addr: cleanAddr(d.item.addr) };
    return l ? { sheet: l.sheet, addr: FK.locAddr(l) } : null;
  }

  // ---- rendering

  function visible() {
    var out = [];
    (function walk(d) { out.push(d); if (d.open && d.kids) d.kids.forEach(walk); })(X.root);
    return out;
  }

  function guides(d) {
    // tree connector pieces for each level: "v" (line), " " (blank), "t" (tee), "l" (elbow)
    var parts = [];
    var n = d;
    while (n.parent) {
      var sibs = n.parent.kids || [], last = sibs[sibs.length - 1] === n;
      parts.unshift(n === d ? (last ? "l" : "t") : (last ? " " : "v"));
      n = n.parent;
    }
    return parts;
  }

  function renderX() {
    var rows = visible();
    X.rows = rows;
    var html = "";
    rows.forEach(function (d, i) {
      var g = guides(d), ind = "";
      g.forEach(function (p) { ind += '<span class="g g' + (p === " " ? "b" : p) + '"></span>'; });
      var pm = d.canOpen ? '<span class="pm" data-i="' + i + '">' + (d.open ? "−" : "+") + "</span>" : '<span class="pm none"></span>';
      var loc = locOf(d);
      var rootSheet = X.root && X.root.cx ? X.root.cx.sheet : "";
      var locTxt = !loc ? "" : (loc.sheet === rootSheet ? loc.addr : loc.sheet + "!" + loc.addr);
      var cls = "xr" + (d === X.cur ? " sel" : "") + (d.c && d.c.k === "missing" ? " dim" : "");
      html += '<div class="' + cls + '" data-i="' + i + '">' +
        '<span class="xe">' + ind + pm + '<span class="lb">' + esc(d.label) + "</span></span>" +
        '<span class="xi">' + esc(d.info || "") + "</span>" +
        '<span class="xv">' + esc(valueText(d)) + "</span>" +
        '<span class="xl" title="' + esc(locTxt) + '">' + esc(locTxt) + "</span></div>";
    });
    $("xbody").innerHTML = html;
    renderFormulaBox();
    var body = $("xbody");
    var els = body.querySelectorAll(".xr");
    for (var k = 0; k < els.length; k++) {
      els[k].addEventListener("mousedown", function (ev) {
        var d = X.rows[+ev.currentTarget.getAttribute("data-i")];
        if (ev.target.classList.contains("pm")) { ev.preventDefault(); guard(toggleX)(d); return; }
        guard(selectX)(d);
      });
      els[k].addEventListener("dblclick", function (ev) {
        var d = X.rows[+ev.currentTarget.getAttribute("data-i")];
        if (d.canOpen) guard(toggleX)(d);
      });
    }
    var cur = body.querySelector(".xr.sel");
    if (cur && cur.scrollIntoView) cur.scrollIntoView({ block: "nearest" });
  }

  function renderFormulaBox() {
    var d = X.cur || X.root, cx = d.cx || (d.kind === "dep" ? null : X.root.cx);
    var box = $("xformula");
    if (!box || ED.on) return;
    box.title = "Click (or press F2) to edit this cell's formula";
    if (d.kind === "dep") { box.innerHTML = esc(d.cx && d.cx.formula || d.text || ""); return; }
    if (!cx || !cx.formula) { box.innerHTML = esc(cx ? cx.text : ""); return; }
    var f = cx.formula;
    if ((d.kind === "expr" || d.kind === "filter") && d.en && d.en.e > d.en.s) {
      box.innerHTML = esc(f.slice(0, d.en.s)) + '<span class="hl">' + esc(f.slice(d.en.s, d.en.e)) + "</span>" + esc(f.slice(d.en.e));
    } else box.innerHTML = esc(f);
  }

  async function toggleX(d) {
    if (!d.canOpen) return;
    if (d.open) d.open = false;
    else { status("Loading…"); await openNode(d); status(""); }
    renderX();
    focusX();
  }

  function focusX() {
    var g = $("xgrid");
    if (g) { try { g.focus({ preventScroll: true }); } catch (e) { g.focus(); } }
  }

  async function selectX(d) {
    X.cur = d;
    renderX();
    focusX();
    var loc = locOf(d);
    if (loc) {
      await navigate({ sheet: loc.sheet, addr: loc.addr }, !X.pushed);
      X.pushed = true;
    }
  }

  function onXKey(ev) {
    var rows = X.rows, i = rows.indexOf(X.cur), d = X.cur;
    switch (ev.key) {
      case "ArrowDown": if (i < rows.length - 1) guard(selectX)(rows[i + 1]); break;
      case "ArrowUp": if (i > 0) guard(selectX)(rows[i - 1]); break;
      case "ArrowRight": case "+":
        if (d && d.canOpen && !d.open) guard(toggleX)(d);
        else if (d && d.open && d.kids && d.kids.length) guard(selectX)(d.kids[0]);
        break;
      case "ArrowLeft": case "-":
        if (d && d.open) guard(toggleX)(d);
        else if (d && d.parent) guard(selectX)(d.parent);
        break;
      case "F2": startEdit(); break;
      case "Enter": guard(okX)(); break;
      case "Escape": guard(goBack)(); break;
      default: return;
    }
    ev.preventDefault();
  }

  // ---- dependents (same window, rows are cells)

  async function depKids(d) {
    var kids = [];
    await Excel.run(async function (ctx) {
      var sheet = d.cx.sheet, addr = d.cx.addr, b = FK.parseA1(addr), list = null;
      try {
        var deps = ctx.workbook.worksheets.getItem(sheet).getRange(addr).getDirectDependents();
        deps.load("addresses");
        await ctx.sync();
        list = [];
        deps.addresses.forEach(function (s) {
          FK.splitAddressList(s).forEach(function (full) {
            var p = FK.splitSheetAddress(full);
            var bb = FK.parseA1(p.addr);
            if (!bb) return;
            for (var r = bb.r1; r <= bb.r2 && list.length < 300; r++)
              for (var c = bb.c1; c <= bb.c2 && list.length < 300; c++) list.push({ sheet: p.sheet || sheet, addr: FK.numToCol(c) + r });
          });
        });
      } catch (e) {
        list = (e && e.code === "ItemNotFound") ? [] : null;
      }
      if (list === null) list = await scanDependents(ctx, { sheet: sheet, row: b.r1, col: b.c1 });
      for (var i = 0; i < list.length && i < 300; i++) {
        var cx = { sheet: list[i].sheet, addr: list[i].addr };
        var r = ctx.workbook.worksheets.getItem(cx.sheet).getRange(cx.addr);
        r.load("formulas,values,valueTypes,text");
        kids.push({ cx: cx, r: r });
      }
      await ctx.sync();
    });
    return kids.map(function (k) {
      var f = k.r.formulas[0][0];
      k.cx.formula = typeof f === "string" ? f : "";
      k.cx.raw = f;
      k.cx.text = fmtValue(k.r.values[0][0], k.r.valueTypes[0][0], k.r.text[0][0]);
      return dnode(d, { kind: "dep", cx: k.cx, label: k.cx.addr, info: k.cx.formula, canOpen: true });
    });
  }

  // ---- entry points

  async function buildRoot(mode, sheet, addr) {
    await Excel.run(async function (ctx) {
      if (mode === "prec") {
        var cx = await makeCx(ctx, sheet, addr);
        X.root = dnode(null, { kind: "cell", cx: cx, label: cx.addr, info: "", canOpen: !!cx.formula });
      } else {
        var r = ctx.workbook.worksheets.getItem(sheet).getRange(addr);
        r.load("formulas,values,valueTypes,text");
        await ctx.sync();
        var f = r.formulas[0][0];
        var cx2 = { sheet: sheet, addr: cleanAddr(addr), raw: f, formula: typeof f === "string" && f.charAt(0) === "=" ? f : "",
                    text: fmtValue(r.values[0][0], r.valueTypes[0][0], r.text[0][0]) };
        X.root = dnode(null, { kind: "dep", cx: cx2, label: cx2.addr, info: "", canOpen: true });
      }
    });
    if (X.root.canOpen) await openNode(X.root);
  }

  // ---- optional pink highlight on the start cell (always restored)

  var HL_COLOR = "#F7A8F0";
  var HL = { on: false, saved: null };
  function lsGet(k) { try { return window.localStorage.getItem(k); } catch (e) { return null; } }
  function lsSet(k, v) { try { if (v == null) window.localStorage.removeItem(k); else window.localStorage.setItem(k, v); } catch (e) { /* storage off */ } }
  HL.on = false;                 // pink option removed: selection (blue) only, the workbook is never changed
  lsSet("fk.highlight", null);

  function hlMenuLabel() {
    var b = document.querySelector('#xmenu button[data-act="hl"]');
    if (b) b.firstChild.textContent = (HL.on ? "✓ " : "") + "Highlight start cell";
  }

  async function toggleHighlight() {
    HL.on = !HL.on;
    lsSet("fk.highlight", HL.on ? "1" : "0");
    hlMenuLabel();
    if (!HL.on) await restoreHighlight();
    else if (X.root && X.root.cx) await applyHighlight(X.root.cx.sheet, X.root.cx.addr);
    status(HL.on ? "Start cell will show pink (a temporary overlay — your cell formatting is not changed)." : "Start-cell highlight off.");
  }

  // Pink is drawn with a temporary conditional-format rule on top of the cell,
  // so the cell's own fill/format is never changed. Removing the rule restores the look exactly.
  async function applyHighlight(sheet, addr) {
    await restoreHighlight();
    if (!HL.on) return;
    await Excel.run(async function (ctx) {
      var wb = ctx.workbook;
      wb.load("name");
      var r = wb.worksheets.getItem(sheet).getRange(addr);
      var cf = r.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      cf.custom.rule.formula = "=TRUE";
      cf.custom.format.fill.color = HL_COLOR;
      cf.load("id");
      await ctx.sync();
      var saved = { book: wb.name, sheet: sheet, addr: addr, cf: cf.id };
      lsSet("fk.hlSaved", JSON.stringify(saved));   // lets a crash be cleaned up next time
      HL.saved = saved;
    });
  }

  async function restoreHighlight() {
    var saved = HL.saved;
    if (!saved) { try { saved = JSON.parse(lsGet("fk.hlSaved") || "null"); } catch (e) { saved = null; } }
    if (!saved) return;
    await Excel.run(async function (ctx) {
      var wb = ctx.workbook;
      wb.load("name");
      await ctx.sync();
      if (wb.name !== saved.book) return;            // belongs to another workbook
      var ws = wb.worksheets.getItemOrNullObject(saved.sheet);
      await ctx.sync();
      if (!ws.isNullObject && saved.cf) {
        var cfs = ws.getRange(saved.addr).conditionalFormats;
        cfs.load("items/id");
        await ctx.sync();
        cfs.items.forEach(function (c) { if (c.id === saved.cf) c.delete(); });
        await ctx.sync();
      }
      HL.saved = null;
      lsSet("fk.hlSaved", null);
    }).catch(function () { /* keep the record and retry next time */ });
  }

  async function startX(mode) {
    showTab("trace");
    await showPane();
    X.mode = mode;
    ED.on = false;
    var a;
    await Excel.run(async function (ctx) { a = await getActive(ctx); });
    await buildRoot(mode, a.sheet, a.addr);
    if (HL.on) await applyHighlight(a.sheet, a.addr); else await restoreHighlight();
    $("xtitle").textContent = mode === "prec" ? "FormulaKit Explorer" : "FormulaKit Explorer — Dependents";
    $("xheadInfo").textContent = mode === "prec" ? "Info" : "Formula";
    X.cur = X.root;
    X.pushed = false;
    renderX();
    focusX();
  }

  // ---- editing the formula box (like Excel's formula bar)

  var ED = { on: false, target: null };

  function editTarget() {
    var d = X.cur || X.root;
    if (!d) return null;
    var cx = d.cx || X.root.cx;
    var raw = cx.raw != null ? cx.raw : cx.formula;
    return { sheet: cx.sheet, addr: cx.addr, raw: raw == null ? "" : String(raw) };
  }

  function startEdit() {
    if (ED.on || !X.root) return;
    var t = editTarget();
    if (!t) return;
    ED.on = true;
    ED.target = t;
    var box = $("xformula");
    box.classList.add("editing");
    box.innerHTML = '<div class="edlbl">Editing <b>' + esc(t.sheet + "!" + t.addr) + "</b> — Enter saves · Esc cancels · ⇧Enter new line</div>" +
      '<textarea id="xedit" spellcheck="false" autocomplete="off" autocorrect="off" autocapitalize="off"></textarea>';
    var ta = $("xedit");
    ta.value = t.raw;
    function fit() { ta.style.height = "auto"; ta.style.height = Math.min(ta.scrollHeight, 220) + "px"; }
    fit();
    ta.addEventListener("input", fit);
    ta.addEventListener("keydown", function (ev) {
      ev.stopPropagation();
      if (ev.key === "Enter" && !ev.shiftKey) { ev.preventDefault(); guard(saveEdit)(); }
      else if (ev.key === "Escape") { ev.preventDefault(); cancelEdit(); }
    });
    ta.focus();
    ta.setSelectionRange(ta.value.length, ta.value.length);
  }

  function cancelEdit() {
    ED.on = false;
    var box = $("xformula");
    if (box) box.classList.remove("editing");
    renderFormulaBox();
    focusX();
  }

  async function saveEdit() {
    var t = ED.target, ta = $("xedit");
    if (!t || !ta) return;
    var txt = ta.value.replace(/\r?\n/g, "\n");
    try {
      await Excel.run(async function (ctx) {
        ctx.workbook.worksheets.getItem(t.sheet).getRange(t.addr).formulas = [[txt]];
        await ctx.sync();
      });
    } catch (e) {
      status("Excel didn't accept that formula — check brackets, commas and sheet names. (" + friendly(e) + ")", true);
      ta.focus();
      return;
    }
    ED.on = false;
    $("xformula").classList.remove("editing");
    await refreshX();
    status("Saved " + t.sheet + "!" + t.addr + ".");
  }

  function pathOf(d) {
    var p = [];
    while (d) { p.unshift(d.parent ? d.parent.kids.indexOf(d) : 0); d = d.parent; }
    return p.join("/");
  }

  function nodeAt(p) {
    var idx = p.split("/").map(Number), d = X.root;
    for (var k = 1; k < idx.length; k++) {
      if (!d || !d.kids || !d.kids[idx[k]]) return null;
      d = d.kids[idx[k]];
    }
    return d;
  }

  // Rebuild the tree after an edit, keeping the same rows open.
  async function refreshX() {
    var open = visible().filter(function (d) { return d.open; }).map(pathOf);
    var curPath = X.cur ? pathOf(X.cur) : "0";
    await buildRoot(X.mode, X.root.cx.sheet, X.root.cx.addr);
    open.sort(function (a, b) { return a.split("/").length - b.split("/").length; });
    for (var i = 0; i < open.length; i++) {
      var d = nodeAt(open[i]);
      if (d && d.canOpen && !d.open) await openNode(d);
    }
    X.cur = nodeAt(curPath) || X.root;
    renderX();
    focusX();
  }

  async function okX() {
    if (ED.on) { await saveEdit(); if (ED.on) return; }
    X.pushed = false;
    S.back = [];
    await restoreHighlight();
    try { await Office.addin.hide(); } catch (e) { /* older Excel */ }
  }

  async function tracePrecedents() { await startX("prec"); }
  async function traceDependents() { await startX("dep"); }
  async function explore() { await startX("prec"); }

  async function jumpNext() {
    if (!X.root) { await tracePrecedents(); return; }
    var i = X.rows.indexOf(X.cur);
    await selectX(X.rows[(i + 1) % X.rows.length]);
  }

  async function goBack() {
    var item = S.back.pop();
    if (!item) { status("Nothing to go back to."); return; }
    await navigate(item, false);
    S.originPushed = false;
    X.pushed = false;
    if (X.root) { X.cur = X.root; renderX(); }
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

  // ------------------------------------------------------------------ Shortcut editor

  var ACTION_NAMES = [
    ["FK_PRECEDENTS", "Explore precedents"], ["FK_DEPENDENTS", "Explore dependents"], ["FK_NEXT", "Next row in explorer"],
    ["FK_BACK", "Back to start"], ["FK_EXPLORE", "Explore (same as precedents)"], ["FK_MAP", "Model map"],
    ["FK_NUMFMT", "Cycle number format"], ["FK_FONT", "Cycle font colour"], ["FK_AUTOCOLOUR", "Auto-colour"], ["FK_BORDERS", "Cycle borders"]
  ];
  var DEFAULT_KEYS = { FK_PRECEDENTS: "P", FK_DEPENDENTS: "U", FK_NEXT: "J", FK_BACK: "B", FK_EXPLORE: "E", FK_MAP: "G",
                       FK_NUMFMT: "N", FK_FONT: "C", FK_AUTOCOLOUR: "A", FK_BORDERS: "K" };
  var MAC_RESERVED = ["Command+Q", "Command+W", "Command+H", "Command+M", "Command+Tab", "Command+Space",
                      "Command+Option+D", "Command+Option+M", "Command+Option+H", "Command+Option+Escape", "Command+Option+Space"];
  var SC = { cur: {}, ok: false, mac: true, waiting: null, msg: {} };

  function prettyKey(s) {
    if (!s) return "—";
    if (!SC.mac) return String(s).replace(/Command|Cmd/g, "Ctrl").replace(/Option/g, "Alt");
    var sym = { Command: "⌘", Cmd: "⌘", Option: "⌥", Alt: "⌥", Ctrl: "⌃", Control: "⌃", Shift: "⇧", Up: "↑", Down: "↓", Left: "←", Right: "→" };
    return String(s).split("+").map(function (p) { return sym[p] || p.toUpperCase(); }).join("");
  }

  function supportsKeys() {
    try { return !!(Office.context.requirements && Office.context.requirements.isSetSupported("KeyboardShortcuts", "1.1")); }
    catch (e) { return false; }
  }

  async function loadShortcuts() {
    SC.mac = /Mac/i.test(navigator.platform || navigator.userAgent || "");
    SC.ok = supportsKeys();
    if (SC.ok) {
      try { SC.cur = (await Office.actions.getShortcuts()) || {}; } catch (e) { SC.ok = false; }
    }
    if (!SC.ok) {
      SC.cur = {};
      Object.keys(DEFAULT_KEYS).forEach(function (id) { SC.cur[id] = (SC.mac ? "Command+Option+" : "Ctrl+Alt+") + DEFAULT_KEYS[id]; });
      if (!SC.mac) { SC.cur.FK_PRECEDENTS = "Ctrl+Q"; SC.cur.FK_DEPENDENTS = "Ctrl+Shift+Q"; }
    }
    refreshKeyLabels();
    renderKeysPage();
  }

  function refreshKeyLabels() {
    var els = document.querySelectorAll("[data-key]");
    for (var i = 0; i < els.length; i++) els[i].textContent = prettyKey(SC.cur[els[i].getAttribute("data-key")]);
  }

  function renderKeysPage() {
    var t = $("keysTable");
    if (!t) return;
    var html = "";
    ACTION_NAMES.forEach(function (a) {
      var id = a[0], m = SC.msg[id];
      html += "<tr><td>" + esc(a[1]) + (m ? '<div class="kmsg ' + m.cls + '">' + esc(m.text) + "</div>" : "") +
        '</td><td class="r"><button class="kbtn' + (SC.waiting === id ? " wait" : "") + '" data-id="' + id + '"' + (SC.ok ? "" : " disabled") + ">" +
        (SC.waiting === id ? "Press keys…" : esc(prettyKey(SC.cur[id]))) + "</button></td></tr>";
    });
    t.innerHTML = html;
    if (!SC.ok) {
      t.insertAdjacentHTML("afterend", '<p class="note kmsg err" id="kUnsup">This version of Excel doesn\'t let add-ins change shortcuts. Update Excel to use this.</p>');
    }
    var bs = t.querySelectorAll(".kbtn");
    for (var i = 0; i < bs.length; i++) {
      bs[i].addEventListener("click", function (ev) {
        SC.waiting = ev.currentTarget.getAttribute("data-id");
        SC.msg = {};
        renderKeysPage();
        var b = t.querySelector('.kbtn[data-id="' + SC.waiting + '"]');
        if (b) b.focus();
      });
    }
  }

  function comboFromEvent(ev) {
    var c = ev.code || "", key = null;
    if (/^Key[A-Z]$/.test(c)) key = c.slice(3);
    else if (/^Digit\d$/.test(c)) key = c.slice(5);
    else if (c === "Minus") key = "-";
    else if (/^Arrow(Up|Down|Left|Right)$/.test(c)) key = c.slice(5);
    if (!key) return null;
    var mods = [];
    if (SC.mac) {
      if (ev.ctrlKey) mods.push("Ctrl");
      if (ev.metaKey) mods.push("Command");
      if (ev.altKey) mods.push("Option");
      if (ev.shiftKey) mods.push("Shift");
    } else {
      if (ev.ctrlKey) mods.push("Ctrl");
      if (ev.altKey) mods.push("Alt");
      if (ev.shiftKey) mods.push("Shift");
    }
    if (!mods.length) return { error: "Add ⌘, ⌥ or ⌃ — a single key on its own can't be a shortcut." };
    if (mods.length === 1 && mods[0] === "Shift") return { error: "Shift alone isn't allowed — add ⌘, ⌥ or ⌃." };
    return { combo: mods.concat(key).join("+") };
  }

  async function onKeyCapture(ev) {
    if (!SC.waiting) return;
    if (["Meta", "Alt", "Control", "Shift"].indexOf(ev.key) >= 0) return;   // wait for the real key
    ev.preventDefault();
    ev.stopPropagation();
    var id = SC.waiting;
    if (ev.key === "Escape") { SC.waiting = null; renderKeysPage(); return; }
    var r = comboFromEvent(ev);
    if (!r) { SC.msg[id] = { cls: "err", text: "Use a letter, number or arrow key." }; renderKeysPage(); return; }
    if (r.error) { SC.msg[id] = { cls: "err", text: r.error }; renderKeysPage(); return; }
    var combo = r.combo;
    if (MAC_RESERVED.indexOf(combo) >= 0) { SC.msg[id] = { cls: "err", text: prettyKey(combo) + " is used by macOS — pick another." }; renderKeysPage(); return; }
    var clash = Object.keys(SC.cur).filter(function (k) { return k !== id && String(SC.cur[k]).toUpperCase() === combo.toUpperCase(); })[0];
    if (clash) {
      var nm = ACTION_NAMES.filter(function (a) { return a[0] === clash; })[0];
      SC.msg[id] = { cls: "err", text: prettyKey(combo) + " is already used for “" + (nm ? nm[1] : clash) + "”." };
      renderKeysPage();
      return;
    }
    SC.waiting = null;
    var note = "";
    try {
      var used = await Office.actions.areShortcutsInUse([combo]);
      if (used && used[0] && used[0].inUse) note = " Excel also uses it — the first time you press it, Excel will ask which to keep.";
    } catch (e) { /* optional check */ }
    try {
      var o = {};
      o[id] = combo;
      await Office.actions.replaceShortcuts(o);
      SC.cur = (await Office.actions.getShortcuts()) || SC.cur;
      SC.msg[id] = { cls: "ok", text: "Saved: " + prettyKey(SC.cur[id] || combo) + "." + note };
    } catch (e) {
      SC.msg[id] = { cls: "err", text: "Excel didn't accept " + prettyKey(combo) + " (" + friendly(e) + ")." };
    }
    refreshKeyLabels();
    renderKeysPage();
  }

  // Layout-independent preset: digits are the same on Latin and Cyrillic keyboards.
  var DIGIT_KEYS = { FK_PRECEDENTS: "1", FK_DEPENDENTS: "2", FK_NEXT: "3", FK_BACK: "4", FK_EXPLORE: "5", FK_MAP: "6",
                     FK_NUMFMT: "7", FK_FONT: "8", FK_AUTOCOLOUR: "9", FK_BORDERS: "0" };
  async function digitShortcuts() {
    if (!SC.ok) return;
    var o = {}, pre = SC.mac ? "Command+Option+" : "Ctrl+Alt+";
    Object.keys(DIGIT_KEYS).forEach(function (id) { o[id] = pre + DIGIT_KEYS[id]; });
    await Office.actions.replaceShortcuts(o);
    SC.cur = (await Office.actions.getShortcuts()) || o;
    SC.msg = {};
    refreshKeyLabels();
    renderKeysPage();
    status("Shortcuts switched to number keys — they now work on any keyboard layout.");
  }

  async function resetShortcuts() {
    if (!SC.ok) return;
    var o = {};
    ACTION_NAMES.forEach(function (a) { o[a[0]] = null; });
    await Office.actions.replaceShortcuts(o);
    SC.cur = (await Office.actions.getShortcuts()) || {};
    SC.msg = {};
    refreshKeyLabels();
    renderKeysPage();
    status("Shortcuts reset to defaults.");
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

    document.addEventListener("keydown", function (ev) {
      if (!SC.waiting) return;
      if (["Meta", "Alt", "Control", "Shift"].indexOf(ev.key) < 0) { ev.preventDefault(); ev.stopPropagation(); }
      guard(onKeyCapture)(ev);
    }, true);
    if ($("bKeysReset")) $("bKeysReset").addEventListener("click", guard(resetShortcuts));
    if ($("bKeysDigits")) $("bKeysDigits").addEventListener("click", guard(digitShortcuts));
    guard(loadShortcuts)();

    // explorer window
    if ($("xgrid")) $("xgrid").addEventListener("keydown", onXKey);
    if ($("bOK")) $("bOK").addEventListener("click", guard(okX));
    if ($("xformula")) $("xformula").addEventListener("click", function () { if (!ED.on) startEdit(); });
    var menu = $("xmenu");
    if ($("bMore")) $("bMore").addEventListener("click", function (ev) { ev.stopPropagation(); menu.classList.toggle("on"); });
    document.addEventListener("click", function () { if (menu) menu.classList.remove("on"); });
    var acts = {
      prec: tracePrecedents, dep: traceDependents, back: goBack,
      map: function () { showTab("map"); }, format: function () { showTab("format"); }, keys: function () { showTab("keys"); },
      hl: toggleHighlight
    };
    var mb = menu ? menu.querySelectorAll("button") : [];
    for (var m = 0; m < mb.length; m++) {
      mb[m].addEventListener("click", function (ev) {
        menu.classList.remove("on");
        guard(acts[ev.currentTarget.getAttribute("data-act")])();
      });
    }
    hlMenuLabel();
    var backs = document.querySelectorAll("[data-back]");
    for (var q = 0; q < backs.length; q++) backs[q].addEventListener("click", function () { showTab("trace"); });
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
    try {
      Office.addin.onVisibilityModeChanged(function (args) {
        if (args && args.visibilityMode === "Hidden") guard(restoreHighlight)();
      });
    } catch (e) { /* older Excel */ }
    guard(restoreHighlight)();   // undo anything left over from a crash
  });

  // exposed for testing
  if (typeof window !== "undefined") window.FKApp = { S: S, showTab: showTab };
})();
