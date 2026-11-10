/* FormulaKit — formula parsing helpers (pure functions, no Excel calls). */
(function (root) {
  "use strict";

  var MAX_COL = 16384;
  var MAX_ROW = 1048576;

  function isRefChar(ch) {
    if (!ch) return false;
    return /[A-Za-z0-9_.$:!\\?]/.test(ch) || ch.charCodeAt(0) > 127;
  }

  // Position just after the closing quote q (doubled quotes are escapes).
  function skipQuoted(f, i, q) {
    var n = f.length;
    i++;
    while (i < n) {
      if (f[i] === q) {
        if (f[i + 1] === q) i += 2;
        else return i + 1;
      } else i++;
    }
    return n;
  }

  function colToNum(s) {
    var n = 0;
    s = s.toUpperCase();
    for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n;
  }

  function numToCol(n) {
    var s = "";
    while (n > 0) {
      var m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function box(r1, c1, r2, c2) {
    var b = {
      r1: Math.min(r1, r2), r2: Math.max(r1, r2),
      c1: Math.min(c1, c2), c2: Math.max(c1, c2)
    };
    if (b.r1 < 1 || b.c1 < 1 || b.r2 > MAX_ROW || b.c2 > MAX_COL) return null;
    return b;
  }

  // "A1", "$A$1:B5", "A:C", "3:5" -> {r1,c1,r2,c2} (1-based) or null
  function parseA1(addr) {
    if (!addr) return null;
    var a = String(addr).replace(/\$/g, "").toUpperCase();
    var m = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec(a);
    if (m) {
      var c1 = colToNum(m[1]), r1 = +m[2];
      return box(r1, c1, m[4] ? +m[4] : r1, m[3] ? colToNum(m[3]) : c1);
    }
    m = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(a);
    if (m) return box(1, colToNum(m[1]), MAX_ROW, colToNum(m[2]));
    m = /^(\d+):(\d+)$/.exec(a);
    if (m) return box(+m[1], 1, +m[2], MAX_COL);
    return null;
  }

  function overlaps(a, b) {
    return a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
  }

  function unquoteSheet(s) {
    if (s.length >= 2 && s[0] === "'" && s[s.length - 1] === "'") {
      return s.slice(1, -1).replace(/''/g, "'");
    }
    return s;
  }

  function quoteSheet(name) {
    return "'" + String(name).replace(/'/g, "''") + "'";
  }

  // Split "Sheet1!A1:B2, 'My, Sheet'!C3" on commas outside quotes.
  function splitAddressList(s) {
    var out = [], cur = "", inQ = false;
    s = String(s || "");
    for (var i = 0; i < s.length; i++) {
      var ch = s[i];
      if (ch === "'") inQ = !inQ;
      if (ch === "," && !inQ) {
        if (cur.trim()) out.push(cur.trim());
        cur = "";
      } else cur += ch;
    }
    if (cur.trim()) out.push(cur.trim());
    return out;
  }

  // "Sheet1!A1:B2" -> {sheet:"Sheet1", addr:"A1:B2"}
  function splitSheetAddress(full) {
    var p = full.lastIndexOf("!");
    if (p < 0) return { sheet: null, addr: full.replace(/\$/g, "") };
    return { sheet: unquoteSheet(full.slice(0, p)), addr: full.slice(p + 1).replace(/\$/g, "") };
  }

  function classify(text) {
    var t = text, spill = false;
    if (t.charAt(t.length - 1) === "#") { t = t.slice(0, -1); spill = true; }
    if (!t || /^[\d.]+$/.test(t)) return null;

    var sheet = null, body = t;
    var p = t.lastIndexOf("!");
    if (p >= 0) {
      sheet = unquoteSheet(t.slice(0, p));
      body = t.slice(p + 1);
      if (!sheet || sheet.indexOf("[") >= 0 || sheet.indexOf("]") >= 0) return null; // other workbook
    }

    var b = parseA1(body);
    if (b) return { kind: "address", sheet: sheet, addr: body.replace(/\$/g, ""), bounds: b, spill: spill };

    var m = /^([A-Za-z_\\][\w.\\]*)\[(.*)\]$/.exec(body);
    if (m && !sheet) {
      var spec = m[2], cols = [];
      if (spec.indexOf("[") < 0) {
        cols.push(spec.replace(/^@/, ""));
      } else {
        var re = /\[([^\]]*)\]/g, mm;
        while ((mm = re.exec(spec))) {
          if (mm[1].charAt(0) !== "#") cols.push(mm[1].replace(/^@/, ""));
        }
      }
      return {
        kind: "table", table: m[1],
        column: cols.length ? cols[0] : null,
        thisRow: /@|#This Row/i.test(spec)
      };
    }

    if (/^[A-Za-z_\\][\w.\\?]*$/.test(body) && !/^(TRUE|FALSE)$/i.test(body) && !/^_xl/i.test(body)) {
      return { kind: "name", sheet: sheet, name: body };
    }
    return null;
  }

  // All reference-like tokens in an A1 formula, in order of position.
  function findRefs(formula) {
    var f = String(formula == null ? "" : formula);
    var out = [];
    if (f.charAt(0) !== "=") return out;
    var n = f.length, i = 1;
    while (i < n) {
      var ch = f[i];
      if (ch === '"') { i = skipQuoted(f, i, '"'); continue; }
      if (ch === "{") { var j = f.indexOf("}", i); i = j < 0 ? n : j + 1; continue; }
      if (ch === "'" || ch === "[" || isRefChar(ch)) {
        var start = i;
        if (ch === "'") i = skipQuoted(f, i, "'");
        while (i < n) {
          var c = f[i];
          if (isRefChar(c)) i++;
          else if (c === "[") {
            var d = 0;
            while (i < n) {
              if (f[i] === "[") d++;
              else if (f[i] === "]") d--;
              i++;
              if (d === 0) break;
            }
          } else break;
        }
        if (f[i] === "#") i++;
        var text = f.slice(start, i);
        var k = i;
        while (k < n && f[k] === " ") k++;
        if (f[k] !== "(") {
          var tok = classify(text);
          if (tok) { tok.text = text; tok.start = start; tok.end = i; out.push(tok); }
        }
        continue;
      }
      i++;
    }
    return out;
  }

  // Function calls in a formula: [{name, text, start, end, depth}]
  function functionParts(formula) {
    var f = String(formula || ""), n = f.length, out = [], depth = 0, i = 0;
    while (i < n) {
      var ch = f[i];
      if (ch === '"' || ch === "'") { i = skipQuoted(f, i, ch); continue; }
      if (ch === "(") {
        var j = i - 1;
        while (j >= 0 && /[A-Za-z0-9_.]/.test(f[j])) j--;
        var name = f.slice(j + 1, i);
        if (name && !/^\d/.test(name)) {
          var close = matchParen(f, i);
          if (close > 0) {
            out.push({ name: name.replace(/^_xlfn\./i, ""), text: f.slice(j + 1, close + 1), start: j + 1, end: close + 1, depth: depth });
          }
        }
        depth++;
      } else if (ch === ")") depth--;
      i++;
    }
    return out;
  }

  function matchParen(f, open) {
    var depth = 0, i = open, n = f.length;
    while (i < n) {
      var ch = f[i];
      if (ch === '"' || ch === "'") { i = skipQuoted(f, i, ch); continue; }
      if (ch === "(") depth++;
      if (ch === ")") { depth--; if (depth === 0) return i; }
      i++;
    }
    return -1;
  }

  // Prefix unqualified cell addresses with the home sheet so the formula
  // means the same thing on any sheet.
  function qualify(formula, refs, homeSheet) {
    var f = String(formula);
    for (var i = refs.length - 1; i >= 0; i--) {
      var r = refs[i];
      if (r.kind === "address" && !r.sheet) {
        f = f.slice(0, r.start) + quoteSheet(homeSheet) + "!" + f.slice(r.start);
      }
    }
    return f;
  }

  // Short literal for a value, used when substituting values into a formula.
  function literal(v, type) {
    if (type === "Empty") return "0";
    if (type === "Error") return String(v);
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    if (typeof v === "number") {
      if (Number.isInteger(v)) return String(v);
      return String(Math.round(v * 1e6) / 1e6);
    }
    var s = String(v);
    if (s.length > 25) s = s.slice(0, 25) + "…";
    return '"' + s.replace(/"/g, '""') + '"';
  }

  // True if an R1C1 formula contains a typed-in number other than 0 or 1.
  function hasHardcode(f) {
    f = String(f);
    var n = f.length, i = 1, prev = "=";
    while (i < n) {
      var ch = f[i];
      if (ch === '"' || ch === "'") { i = skipQuoted(f, i, ch); prev = ch; continue; }
      if (ch === "[") {
        var d = 0;
        while (i < n) {
          if (f[i] === "[") d++;
          else if (f[i] === "]") d--;
          i++;
          if (d === 0) break;
        }
        prev = "]";
        continue;
      }
      if (ch === "{") { var j0 = f.indexOf("}", i); i = j0 < 0 ? n : j0 + 1; prev = "}"; continue; }
      if (/[0-9.]/.test(ch)) {
        var j = i;
        while (j < n && /[0-9.]/.test(f[j])) j++;
        var num = f.slice(i, j);
        if (!/[A-Za-z0-9_.!$]/.test(prev) && num !== ".") {
          var v = parseFloat(num);
          if (!isNaN(v) && v !== 0 && v !== 1) return true;
        }
        prev = num.charAt(num.length - 1);
        i = j;
        continue;
      }
      prev = ch;
      i++;
    }
    return false;
  }

  // Model map: codes for a block of cells from formulasR1C1 + valueTypes.
  function classifyGrid(F, T) {
    var nR = F.length, nC = nR ? F[0].length : 0;
    var counts = { N: 0, F: 0, "<": 0, "^": 0, L: 0, H: 0, X: 0, E: 0, T: 0 };
    var codes = [];
    function at(r, c) {
      if (r < 0 || c < 0 || r >= nR || c >= nC) return "";
      var v = F[r][c];
      return v == null ? "" : String(v);
    }
    for (var r = 0; r < nR; r++) {
      var row = [];
      for (var c = 0; c < nC; c++) {
        var f = at(r, c), t = T[r][c], code = "";
        if (f.charAt(0) === "=") {
          var L = at(r, c - 1), R = at(r, c + 1), U = at(r - 1, c), D = at(r + 1, c);
          if (t === "Error") code = "E";
          else if (L.charAt(0) === "=" && L === R && L !== f) code = "X";
          else if (U.charAt(0) === "=" && U === D && U !== f) code = "X";
          else if (f === L) code = "<";
          else if (f === U) code = "^";
          else if (hasHardcode(f)) code = "H";
          else if (f.indexOf("!") >= 0) code = "L";
          else code = "F";
        } else if (t === "Empty" || f === "") code = "";
        else if (t === "Error") code = "E";
        else if (t === "String" || t === "Boolean") code = "T";
        else code = "N";
        if (code) counts[code]++;
        row.push(code);
      }
      codes.push(row);
    }
    return { codes: codes, counts: counts };
  }

  // ---------------------------------------------------------------- expression tree

  var PREC = { "=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5 };

  function lexFormula(f, refs) {
    var byStart = {}, toks = [], n = f.length, i = f.charAt(0) === "=" ? 1 : 0, m, j;
    refs.forEach(function (r, k) { byStart[r.start] = k; });
    while (i < n) {
      var ch = f[i], rest = f.slice(i);
      if (/\s/.test(ch)) { i++; continue; }
      if (byStart[i] !== undefined) {
        var r = refs[byStart[i]];
        toks.push({ t: "ref", ref: byStart[i], s: i, e: r.end });
        i = r.end;
        continue;
      }
      if (ch === '"') {
        j = skipQuoted(f, i, '"');
        toks.push({ t: "str", v: f.slice(i + 1, j - 1).replace(/""/g, '"'), s: i, e: j });
        i = j;
        continue;
      }
      if ((m = /^(\d+\.?\d*|\.\d+)(E[+-]?\d+)?/i.exec(rest))) {
        toks.push({ t: "num", v: parseFloat(m[0]), s: i, e: i + m[0].length });
        i += m[0].length;
        continue;
      }
      if ((m = /^#(N\/A|DIV\/0!|REF!|VALUE!|NAME\?|NUM!|NULL!|SPILL!|CALC!)/i.exec(rest))) {
        toks.push({ t: "err", v: m[0].toUpperCase(), s: i, e: i + m[0].length });
        i += m[0].length;
        continue;
      }
      if ((m = /^([A-Za-z_][\w.]*)\s*\(/.exec(rest))) {
        toks.push({ t: "fn", name: m[1].replace(/^_xl(fn|ws)\./i, "").toUpperCase(), s: i, e: i + m[0].length });
        i += m[0].length;
        continue;
      }
      if ((m = /^(TRUE|FALSE)(?![\w.(])/i.exec(rest))) {
        toks.push({ t: "bool", v: m[1].toUpperCase() === "TRUE", s: i, e: i + m[0].length });
        i += m[0].length;
        continue;
      }
      if ((m = /^(<=|>=|<>)/.exec(rest))) {
        toks.push({ t: "op", v: m[0], s: i, e: i + 2 });
        i += 2;
        continue;
      }
      if ("+-*/^&=<>%,()".indexOf(ch) >= 0) {
        toks.push({ t: "op", v: ch, s: i, e: i + 1 });
        i++;
        continue;
      }
      if (ch === "{") {
        j = f.indexOf("}", i);
        j = j < 0 ? n : j + 1;
        toks.push({ t: "unk", s: i, e: j });
        i = j;
        continue;
      }
      // unresolved name / external link / anything else
      j = i;
      if (ch === "'") j = skipQuoted(f, i, "'");
      while (j < n && (isRefChar(f[j]) || f[j] === "[" || f[j] === "]" || f[j] === "#")) j++;
      if (j === i) j = i + 1;
      toks.push({ t: "unk", s: i, e: j });
      i = j;
    }
    return toks;
  }

  // Parse an A1 formula into a tree. Nodes: {k, s, e, ...}. Returns null if it can't parse.
  function parseFormula(formula, refs) {
    var f = String(formula), toks, p = 0;
    try { toks = lexFormula(f, refs); } catch (e) { return null; }
    function peek() { return toks[p]; }
    function isOp(t, v) { return !!t && t.t === "op" && t.v === v; }

    function primary() {
      var t = toks[p++], node;
      if (!t) throw new Error("end");
      if (t.t === "num" || t.t === "str" || t.t === "bool" || t.t === "err") node = { k: t.t, v: t.v, s: t.s, e: t.e };
      else if (t.t === "ref") node = { k: "ref", ref: t.ref, s: t.s, e: t.e };
      else if (t.t === "unk") node = { k: "unk", s: t.s, e: t.e };
      else if (t.t === "fn") {
        var args = [];
        if (isOp(peek(), ")")) p++;
        else {
          for (;;) {
            var nx = peek();
            if (!nx) throw new Error("end");
            if (isOp(nx, ",") || isOp(nx, ")")) args.push({ k: "missing", s: nx.s, e: nx.s });
            else args.push(expr(0));
            if (isOp(peek(), ",")) { p++; continue; }
            if (isOp(peek(), ")")) { p++; break; }
            throw new Error("bad args");
          }
        }
        node = { k: "fn", name: t.name, args: args, s: t.s, e: toks[p - 1].e };
      } else if (isOp(t, "(")) {
        var inner = expr(0);
        if (!isOp(peek(), ")")) throw new Error("no )");
        p++;
        node = { k: "paren", x: inner, s: t.s, e: toks[p - 1].e };
      } else if (isOp(t, "-") || isOp(t, "+")) {
        var x = expr(6);
        node = { k: "unary", op: t.v, x: x, s: t.s, e: x.e };
      } else throw new Error("unexpected");
      while (isOp(peek(), "%")) {
        var q = toks[p++];
        node = { k: "pct", x: node, s: node.s, e: q.e };
      }
      return node;
    }

    function expr(minPrec) {
      var left = primary();
      for (;;) {
        var t = peek();
        if (!t || t.t !== "op" || !(t.v in PREC) || PREC[t.v] < minPrec) break;
        p++;
        var right = expr(PREC[t.v] + 1);
        left = { k: "bin", op: t.v, l: left, r: right, s: left.s, e: right.e };
      }
      return left;
    }

    try {
      var root = expr(0);
      if (p < toks.length) return null;
      return root;
    } catch (e) {
      return null;
    }
  }

  // ---- evaluation (values: number | string | boolean | {err} | {range: 2D} | undefined = unknown)

  function isErr(v) { return !!v && typeof v === "object" && "err" in v; }
  function isRange(v) { return !!v && typeof v === "object" && "range" in v; }
  function E(code) { return { err: code }; }

  function scalar(v) {
    if (isRange(v)) {
      var g = v.range;
      return g.length === 1 && g[0].length === 1 ? (g[0][0] === "" ? 0 : g[0][0]) : undefined;
    }
    return v;
  }

  function toNum(v) {
    v = scalar(v);
    if (v === undefined || isErr(v)) return v;
    if (v && v.missing) return 0;
    if (typeof v === "number") return v;
    if (typeof v === "boolean") return v ? 1 : 0;
    var s = String(v).trim();
    if (s !== "" && !isNaN(+s)) return +s;
    return E("#VALUE!");
  }

  function toStr(v) {
    v = scalar(v);
    if (v === undefined || isErr(v)) return v;
    if (v && v.missing) return "";
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    if (typeof v === "number") return String(Math.round(v * 1e10) / 1e10);
    return String(v);
  }

  function toBool(v) {
    v = scalar(v);
    if (v === undefined || isErr(v)) return v;
    if (typeof v === "boolean") return v;
    if (typeof v === "number") return v !== 0;
    var s = String(v).toUpperCase();
    if (s === "TRUE") return true;
    if (s === "FALSE") return false;
    return E("#VALUE!");
  }

  function rank(v) { return typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2; }

  function compare(op, a, b) {
    a = scalar(a); b = scalar(b);
    if (a === undefined || b === undefined) return undefined;
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    var c;
    if (rank(a) !== rank(b)) c = rank(a) - rank(b);
    else if (typeof a === "string") { var x = a.toLowerCase(), y = b.toLowerCase(); c = x < y ? -1 : x > y ? 1 : 0; }
    else c = (a === b) ? 0 : (a < b ? -1 : 1);
    switch (op) {
      case "=": return c === 0;
      case "<>": return c !== 0;
      case "<": return c < 0;
      case ">": return c > 0;
      case "<=": return c <= 0;
      default: return c >= 0;
    }
  }

  function binop(op, l, r) {
    if (op === "&") {
      var s1 = toStr(l), s2 = toStr(r);
      if (s1 === undefined || s2 === undefined) return undefined;
      if (isErr(s1)) return s1;
      if (isErr(s2)) return s2;
      return s1 + s2;
    }
    if (PREC[op] === 1) return compare(op, l, r);
    var a = toNum(l), b = toNum(r);
    if (a === undefined || b === undefined) return undefined;
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    switch (op) {
      case "+": return a + b;
      case "-": return a - b;
      case "*": return a * b;
      case "/": return b === 0 ? E("#DIV/0!") : a / b;
      case "^": var pw = Math.pow(a, b); return isFinite(pw) ? pw : E("#NUM!");
    }
    return undefined;
  }

  // Numbers from arguments: ranges contribute only their numbers; direct values are coerced.
  function nums(args) {
    var out = [];
    for (var i = 0; i < args.length; i++) {
      var v = args[i];
      if (v === undefined) return undefined;
      if (isErr(v)) return v;
      if (v && v.missing) continue;
      if (isRange(v)) {
        for (var r = 0; r < v.range.length; r++) {
          for (var c = 0; c < v.range[r].length; c++) {
            var x = v.range[r][c];
            if (typeof x === "number") out.push(x);
            else if (typeof x === "string" && /^#/.test(x)) return E(x);
          }
        }
      } else {
        var n = toNum(v);
        if (isErr(n)) return n;
        out.push(n);
      }
    }
    return out;
  }

  function roundX(x, d, mode) {
    var f = Math.pow(10, d), y = Math.abs(x) * f, s = x < 0 ? -1 : 1;
    y = +y.toPrecision(15);
    var r = mode === "up" ? Math.ceil(y) : mode === "down" ? Math.floor(y) : Math.round(y);
    return s * r / f;
  }

  function flatRange(v) {
    if (isRange(v)) return [].concat.apply([], v.range);
    return [scalar(v)];
  }

  function callFn(name, a) {
    var ns, x, i;
    switch (name) {
      case "SUM": ns = nums(a); if (!Array.isArray(ns)) return ns; return ns.reduce(function (s, v) { return s + v; }, 0);
      case "PRODUCT": ns = nums(a); if (!Array.isArray(ns)) return ns; return ns.reduce(function (s, v) { return s * v; }, 1);
      case "AVERAGE": ns = nums(a); if (!Array.isArray(ns)) return ns; return ns.length ? ns.reduce(function (s, v) { return s + v; }, 0) / ns.length : E("#DIV/0!");
      case "MIN": ns = nums(a); if (!Array.isArray(ns)) return ns; return ns.length ? Math.min.apply(null, ns) : 0;
      case "MAX": ns = nums(a); if (!Array.isArray(ns)) return ns; return ns.length ? Math.max.apply(null, ns) : 0;
      case "COUNT": ns = nums(a); if (!Array.isArray(ns)) return ns; return ns.length;
      case "ABS": x = toNum(a[0]); return typeof x === "number" ? Math.abs(x) : x;
      case "INT": x = toNum(a[0]); return typeof x === "number" ? Math.floor(x) : x;
      case "SQRT": x = toNum(a[0]); return typeof x === "number" ? (x < 0 ? E("#NUM!") : Math.sqrt(x)) : x;
      case "ROUND": case "ROUNDUP": case "ROUNDDOWN":
        x = toNum(a[0]); var d = a.length > 1 ? toNum(a[1]) : 0;
        if (typeof x !== "number") return x;
        if (typeof d !== "number") return d;
        return roundX(x, Math.trunc(d), name === "ROUNDUP" ? "up" : name === "ROUNDDOWN" ? "down" : "");
      case "POWER": return binop("^", a[0], a[1]);
      case "MOD":
        x = toNum(a[0]); var m = toNum(a[1]);
        if (typeof x !== "number") return x;
        if (typeof m !== "number") return m;
        return m === 0 ? E("#DIV/0!") : x - m * Math.floor(x / m);
      case "IF":
        var c = toBool(a[0]);
        if (c === undefined || isErr(c)) return c;
        if (c) return a.length > 1 && !(a[1] && a[1].missing) ? scalar(a[1]) : (a.length > 1 ? 0 : true);
        return a.length > 2 && !(a[2] && a[2].missing) ? scalar(a[2]) : (a.length > 2 ? 0 : false);
      case "IFERROR":
        if (a[0] === undefined) return undefined;
        return isErr(scalar(a[0])) ? scalar(a[1]) : scalar(a[0]);
      case "AND": case "OR":
        var vals = [];
        for (i = 0; i < a.length; i++) {
          var parts = flatRange(a[i]);
          for (var k = 0; k < parts.length; k++) {
            var b = parts[k];
            if (b === undefined) return undefined;
            if (isErr(b)) return b;
            if (b === "" ) continue;
            b = toBool(b);
            if (isErr(b)) return b;
            vals.push(b);
          }
        }
        return name === "AND" ? vals.every(Boolean) : vals.some(Boolean);
      case "NOT": x = toBool(a[0]); return typeof x === "boolean" ? !x : x;
      case "SUMPRODUCT":
        if (!a.length) return undefined;
        var arrs = [];
        for (i = 0; i < a.length; i++) {
          if (a[i] === undefined) return undefined;
          arrs.push(flatRange(a[i]));
        }
        var len = arrs[0].length, tot = 0;
        for (i = 0; i < arrs.length; i++) if (arrs[i].length !== len) return E("#VALUE!");
        for (var j = 0; j < len; j++) {
          var prod = 1;
          for (i = 0; i < arrs.length; i++) prod *= typeof arrs[i][j] === "number" ? arrs[i][j] : 0;
          tot += prod;
        }
        return tot;
    }
    return undefined;
  }

  // Evaluates every node (sets node.val). getRef(i) returns the value for refs[i].
  function evalTree(node, getRef) {
    var v, i;
    switch (node.k) {
      case "num": case "str": case "bool": v = node.v; break;
      case "err": v = E(node.v); break;
      case "missing": v = { missing: true }; break;
      case "unk": v = undefined; break;
      case "ref": v = getRef(node.ref); break;
      case "paren": v = evalTree(node.x, getRef); break;
      case "unary":
        var x = toNum(evalTree(node.x, getRef));
        v = typeof x === "number" ? (node.op === "-" ? -x : x) : x;
        break;
      case "pct":
        var y = toNum(evalTree(node.x, getRef));
        v = typeof y === "number" ? y / 100 : y;
        break;
      case "bin":
        v = binop(node.op, evalTree(node.l, getRef), evalTree(node.r, getRef));
        break;
      case "fn":
        var args = [];
        for (i = 0; i < node.args.length; i++) args.push(evalTree(node.args[i], getRef));
        v = callFn(node.name, args);
        break;
    }
    node.val = v;
    return v;
  }

  // Prefix unqualified addresses inside f.slice(s, e) with the home sheet.
  function qualifySlice(f, refs, s, e, homeSheet) {
    var out = "", pos = s;
    refs.forEach(function (r) {
      if (r.start >= s && r.end <= e) {
        out += f.slice(pos, r.start);
        if (r.kind === "address" && !r.sheet) out += quoteSheet(homeSheet) + "!";
        out += f.slice(r.start, r.end);
        pos = r.end;
      }
    });
    return out + f.slice(pos, e);
  }

  var api = {
    parseFormula: parseFormula, evalTree: evalTree, qualifySlice: qualifySlice,
    isErr: isErr, isRange: isRange, PREC: PREC,
    findRefs: findRefs, functionParts: functionParts, qualify: qualify,
    parseA1: parseA1, overlaps: overlaps, colToNum: colToNum, numToCol: numToCol,
    quoteSheet: quoteSheet, unquoteSheet: unquoteSheet,
    splitAddressList: splitAddressList, splitSheetAddress: splitSheetAddress,
    literal: literal, hasHardcode: hasHardcode, classifyGrid: classifyGrid
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.FK = api;
})(this);
