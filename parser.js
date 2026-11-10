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

  var api = {
    findRefs: findRefs, functionParts: functionParts, qualify: qualify,
    parseA1: parseA1, overlaps: overlaps, colToNum: colToNum, numToCol: numToCol,
    quoteSheet: quoteSheet, unquoteSheet: unquoteSheet,
    splitAddressList: splitAddressList, splitSheetAddress: splitSheetAddress,
    literal: literal, hasHardcode: hasHardcode, classifyGrid: classifyGrid
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.FK = api;
})(this);
