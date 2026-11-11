/* OM 305 Excel Practice — formula engine.
   A small Excel formula calculator that runs in the browser. It supports the functions taught in OM 305,
   A1 references ($ locking, other sheets, whole columns), Excel Table structured references, and fill-down/right. */
(function (root) {
"use strict";

// ================================================================== values
// number | string | boolean | null (blank) | {err} | Range | {arr: [[...]]}
const ERR = (code, why) => ({ err: code, why: why || "" });
const isErr = (v) => v !== null && typeof v === "object" && v.err !== undefined;
const isRange = (v) => v !== null && typeof v === "object" && v.range === true;
const isArr = (v) => v !== null && typeof v === "object" && Array.isArray(v.arr);

// ================================================================== functions students may use
const TAUGHT = [
  "SUM", "AVERAGE", "COUNT", "COUNTA", "MAX", "MIN",
  "IF", "AND", "OR", "IFERROR",
  "COUNTIFS", "SUMIFS", "AVERAGEIFS", "SUMPRODUCT",
  "VLOOKUP", "INDEX", "MATCH",
  "TODAY", "NOW", "MONTH", "YEAR", "DAY", "ROUND",
  "ISNUMBER", "ISTEXT", "ISBLANK", "ISERROR",
  "LEN", "LEFT", "RIGHT", "TRIM", "SEARCH", "SUBSTITUTE", "CONCATENATE",
  "UPPER", "LOWER", "PROPER", "VALUE", "TEXT",
];
// accepted even though they aren't the focus (common cousins of taught functions)
const EXTRAS = ["COUNTIF", "SUMIF", "AVERAGEIF", "NOT", "DATE", "ABS", "INT", "ROUNDUP", "ROUNDDOWN", "MID", "FIND"];
// real Excel functions we don't accept, with a friendly nudge
const NOT_IN_CLASS = {
  XLOOKUP: "XLOOKUP isn't part of OM 305. Use VLOOKUP or INDEX/MATCH.",
  XMATCH: "XMATCH isn't part of OM 305. Use MATCH.",
  HLOOKUP: "HLOOKUP isn't part of OM 305. Use INDEX/MATCH.",
  LOOKUP: "LOOKUP isn't part of OM 305. Use VLOOKUP or INDEX/MATCH.",
  IFS: "IFS isn't part of OM 305. Nest IF functions instead.",
  SWITCH: "SWITCH isn't part of OM 305. Use nested IFs or a VLOOKUP table.",
  CHOOSE: "CHOOSE isn't part of OM 305. Use nested IFs or a VLOOKUP table.",
  TEXTBEFORE: "TEXTBEFORE isn't part of OM 305. Use LEFT with SEARCH.",
  TEXTAFTER: "TEXTAFTER isn't part of OM 305. Use RIGHT with LEN and SEARCH.",
  TEXTSPLIT: "TEXTSPLIT isn't part of OM 305. Use LEFT/RIGHT with SEARCH (or Text to Columns).",
  TEXTJOIN: "TEXTJOIN isn't part of OM 305. Use CONCATENATE or &.",
  CONCAT: "CONCAT isn't part of OM 305. Use CONCATENATE or &.",
  MAXIFS: "MAXIFS isn't part of OM 305.",
  MINIFS: "MINIFS isn't part of OM 305.",
  FILTER: "FILTER isn't part of OM 305.", UNIQUE: "UNIQUE isn't part of OM 305.", SORT: "SORT isn't part of OM 305.",
  SORTBY: "SORTBY isn't part of OM 305.", LET: "LET isn't part of OM 305.", LAMBDA: "LAMBDA isn't part of OM 305.",
  OFFSET: "OFFSET isn't part of OM 305.", INDIRECT: "INDIRECT isn't part of OM 305.",
  IFNA: "IFNA isn't part of OM 305. Use IFERROR.",
  DATEDIF: "DATEDIF isn't part of OM 305. Subtract the dates and divide by 365.",
  YEARFRAC: "YEARFRAC isn't part of OM 305. Subtract the dates and divide by 365.",
  CLEAN: "CLEAN isn't part of OM 305. TRIM is the one to use.",
};

// ================================================================== cell addresses
function colToNum(s) { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; }
function numToCol(n) { let s = ""; n += 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function addr(r, c, rabs, cabs) { return (cabs ? "$" : "") + numToCol(c) + (rabs ? "$" : "") + (r + 1); }
function parseAddr(a) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a.trim());
  if (!m) return null;
  return { r: +m[2] - 1, c: colToNum(m[1]) };
}
function sheetPrefix(name) { return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name + "!" : "'" + name.replace(/'/g, "''") + "'!"; }

// ================================================================== workbook
// spec: { sheets: [{ name, rows: [[...]], fmt: {colIndex: fmt} }], tables: [{ name, sheet, top, left, ncols, nrows }] }
class Workbook {
  constructor(spec) {
    this.sheets = {};
    this.order = [];
    for (const s of spec.sheets) {
      this.sheets[s.name.toLowerCase()] = { name: s.name, rows: s.rows.map((r) => r.slice()), fmt: Object.assign({}, s.fmt || {}) };
      this.order.push(s.name);
    }
    this.tables = {};
    for (const t of spec.tables || []) {
      const sh = this.sheet(t.sheet);
      const cols = [];
      for (let i = 0; i < t.ncols; i++) cols.push(String(sh.rows[t.top][t.left + i]));
      this.tables[t.name.toLowerCase()] = { name: t.name, sheet: sh.name, top: t.top, left: t.left, ncols: t.ncols, nrows: t.nrows, cols };
    }
  }
  sheet(name) { return this.sheets[String(name).toLowerCase()] || null; }
  get(sheet, r, c) {
    const sh = typeof sheet === "string" ? this.sheet(sheet) : sheet;
    if (!sh || r < 0 || c < 0) return null;
    const row = sh.rows[r];
    if (!row) return null;
    const v = row[c];
    return v === undefined ? null : v;
  }
  set(sheet, r, c, v) {
    const sh = this.sheet(sheet);
    while (sh.rows.length <= r) sh.rows.push([]);
    const row = sh.rows[r];
    while (row.length <= c) row.push(null);
    row[c] = v;
  }
  extent(sheet) {
    const sh = this.sheet(sheet);
    let w = 0;
    for (const r of sh.rows) if (r.length > w) w = r.length;
    return { rows: sh.rows.length, cols: w };
  }
  tableAt(sheet, r, c) {
    for (const t of Object.values(this.tables)) {
      if (t.sheet.toLowerCase() !== String(sheet).toLowerCase()) continue;
      if (r > t.top && r <= t.top + t.nrows && c >= t.left && c < t.left + t.ncols) return t;
    }
    return null;
  }
}

function mkRange(sheet, r1, c1, r2, c2) {
  return { range: true, sheet, r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}

// ================================================================== tokenizer
class FormulaError extends Error {
  constructor(msg) { super(msg); this.friendly = msg; }
}

function tokenize(src) {
  const toks = [];
  let i = 0;
  const n = src.length;
  const push = (t, v, extra) => toks.push(Object.assign({ t, v, pos: i }, extra || {}));
  while (i < n) {
    const ch = src[i];
    if (/\s/.test(ch)) { i++; continue; }
    const rest = src.slice(i);
    let m;
    if (ch === '"') {
      let j = i + 1, s = "";
      for (;;) {
        if (j >= n) throw new FormulaError("A text value is missing its closing quote mark (\").");
        if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; }
        s += src[j++];
      }
      push("str", s); i = j + 1; continue;
    }
    if (ch === "“" || ch === "”") throw new FormulaError("That formula has curly quotes (“ ”). Excel needs straight quotes (\"). This usually happens when a formula is copied from Word.");
    if ((m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?(?![A-Za-z_:$])/.exec(rest))) {
      push("num", parseFloat(m[0])); i += m[0].length; continue;
    }
    // sheet prefix: 'Sheet Name'! or Sheet1!
    let sheet = null;
    if ((m = /^'((?:[^']|'')+)'!/.exec(rest)) || (m = /^([A-Za-z_][A-Za-z0-9_.]*)!/.exec(rest))) {
      sheet = m[1].replace(/''/g, "'");
      i += m[0].length;
      const r2 = src.slice(i);
      if ((m = /^\$?([A-Za-z]{1,3})\s*:\s*\$?([A-Za-z]{1,3})(?![A-Za-z0-9_(])/.exec(r2))) {
        const a = /^(\$?)([A-Za-z]{1,3})\s*:\s*(\$?)([A-Za-z]{1,3})/.exec(r2);
        push("cols", null, { sheet, c1: colToNum(a[2]), c2: colToNum(a[4]), a1: !!a[1], a2: !!a[3] });
        i += m[0].length; continue;
      }
      if ((m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/.exec(r2))) {
        push("ref", null, { sheet, c: colToNum(m[2]), r: +m[4] - 1, cabs: !!m[1], rabs: !!m[3] });
        i += m[0].length; continue;
      }
      throw new FormulaError("After the sheet name and ! Excel expects a cell or range, like " + sheetPrefix(sheet) + "A2 or " + sheetPrefix(sheet) + "A:B.");
    }
    if ((m = /^(\$?)([A-Za-z]{1,3})\s*:\s*(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(\[])/.exec(rest))) {
      push("cols", null, { sheet: null, c1: colToNum(m[2]), c2: colToNum(m[4]), a1: !!m[1], a2: !!m[3] });
      i += m[0].length; continue;
    }
    if ((m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(\[.])/.exec(rest)) && +m[4] > 0) {
      push("ref", null, { sheet: null, c: colToNum(m[2]), r: +m[4] - 1, cabs: !!m[1], rabs: !!m[3] });
      i += m[0].length; continue;
    }
    if ((m = /^([A-Za-z_][A-Za-z0-9_.]*)\s*\(/.exec(rest))) {
      push("func", m[1].toUpperCase()); i += m[0].length; continue;
    }
    if ((m = /^(TRUE|FALSE)(?![A-Za-z0-9_(\[])/i.exec(rest))) {
      push("bool", m[1].toUpperCase() === "TRUE"); i += m[0].length; continue;
    }
    // structured references: Table[...] or [...]
    if ((m = /^([A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(rest))) {
      const table = m[1] || null;
      let j = i + (table ? table.length : 0);
      let depth = 0, k = j;
      for (; k < n; k++) {
        if (src[k] === "[") depth++;
        else if (src[k] === "]") { depth--; if (depth === 0) break; }
      }
      if (k >= n) throw new FormulaError("A structured reference is missing its closing ].");
      push("sref", null, { table, spec: src.slice(j + 1, k) });
      i = k + 1; continue;
    }
    if ((m = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(rest))) {
      push("name", m[0]); i += m[0].length; continue;
    }
    if ((m = /^(<=|>=|<>)/.exec(rest))) { push("op", m[0]); i += 2; continue; }
    if ("+-*/^&=<>%:".includes(ch)) { push("op", ch); i++; continue; }
    if (ch === "(" || ch === ")" || ch === ",") { push(ch, ch); i++; continue; }
    if (ch === ";") throw new FormulaError("Use commas (,) between function arguments, not semicolons.");
    if (ch === "{" || ch === "}") throw new FormulaError("Curly-brace arrays aren't used in OM 305.");
    throw new FormulaError("Excel doesn't understand the character \"" + ch + "\" here.");
  }
  return toks;
}

// ================================================================== parser (Pratt)
const BIN = {
  "=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1,
  "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5,
};

function parse(formula) {
  let src = String(formula).trim();
  if (!src.startsWith("=")) throw new FormulaError("Formulas start with an equals sign (=).");
  src = src.slice(1);
  if (!src.trim()) throw new FormulaError("There's nothing after the equals sign.");
  const toks = tokenize(src);
  let p = 0;
  const peek = () => toks[p];
  const next = () => toks[p++];
  const funcs = new Set();

  function expect(t, msg) {
    const tok = next();
    if (!tok || tok.t !== t) throw new FormulaError(msg);
    return tok;
  }

  function primary() {
    const tok = next();
    if (!tok) throw new FormulaError("The formula ends too early. Is something missing at the end?");
    if (tok.t === "num") return { k: "num", v: tok.v };
    if (tok.t === "str") return { k: "str", v: tok.v };
    if (tok.t === "bool") return { k: "bool", v: tok.v };
    if (tok.t === "op" && (tok.v === "-" || tok.v === "+")) {
      const e = unary();
      return tok.v === "-" ? { k: "neg", e } : e;
    }
    if (tok.t === "ref") {
      let node = { k: "ref", sheet: tok.sheet, r: tok.r, c: tok.c, rabs: tok.rabs, cabs: tok.cabs };
      if (peek() && peek().t === "op" && peek().v === ":" ) {
        next();
        const b = next();
        if (!b || b.t !== "ref") throw new FormulaError("A range needs a cell on both sides of the colon, like A2:B10.");
        if (b.sheet && tok.sheet && b.sheet.toLowerCase() !== tok.sheet.toLowerCase()) throw new FormulaError("A range can't stretch across two sheets.");
        node = { k: "range", a: node, b: { k: "ref", sheet: tok.sheet, r: b.r, c: b.c, rabs: b.rabs, cabs: b.cabs } };
      }
      return node;
    }
    if (tok.t === "cols") return { k: "cols", sheet: tok.sheet, c1: tok.c1, c2: tok.c2, a1: tok.a1, a2: tok.a2 };
    if (tok.t === "sref") return { k: "sref", table: tok.table, spec: tok.spec };
    if (tok.t === "name") return { k: "name", v: tok.v };
    if (tok.t === "func") {
      funcs.add(tok.v);
      const args = [];
      if (peek() && peek().t === ")") { next(); return { k: "func", name: tok.v, args }; }
      for (;;) {
        if (peek() && (peek().t === "," || peek().t === ")")) args.push({ k: "missing" });
        else args.push(expr(0));
        const t2 = next();
        if (!t2) throw new FormulaError("A parenthesis is missing. " + tok.v + "( was never closed with ).");
        if (t2.t === ")") break;
        if (t2.t !== ",") throw new FormulaError("Something is off inside " + tok.v + "(. Check the commas between its arguments.");
      }
      return { k: "func", name: tok.v, args };
    }
    if (tok.t === "(") {
      const e = expr(0);
      expect(")", "A parenthesis is missing. Every ( needs a matching ).");
      return { k: "paren", e };
    }
    if (tok.t === ")") throw new FormulaError("There's an extra ) or something missing before it.");
    if (tok.t === ",") throw new FormulaError("There's a comma that isn't inside a function.");
    throw new FormulaError("The formula has \"" + (tok.v || "") + "\" where Excel expected a value.");
  }
  function unary() {
    let e = primary();
    while (peek() && peek().t === "op" && peek().v === "%") { next(); e = { k: "pct", e }; }
    return e;
  }
  function expr(minPrec) {
    let left = unary();
    for (;;) {
      const tok = peek();
      if (!tok || tok.t !== "op" || !(tok.v in BIN)) break;
      const prec = BIN[tok.v];
      if (prec <= minPrec) break;
      next();
      const right = expr(prec);
      left = { k: "bin", op: tok.v, a: left, b: right };
    }
    return left;
  }
  const ast = expr(0);
  if (p < toks.length) {
    const t = toks[p];
    if (t.t === ")") throw new FormulaError("There's an extra ) in the formula.");
    if (t.t === "name" || t.t === "ref") throw new FormulaError("Two values are next to each other with nothing between them. Is an operator (+ - * / &) or comma missing?");
    throw new FormulaError("Excel couldn't read the end of the formula. Check for a missing operator or parenthesis.");
  }
  return { ast, funcs: [...funcs] };
}

// ================================================================== coercion helpers
function fmtGeneral(v) {
  if (typeof v !== "number") return String(v);
  if (!isFinite(v)) return "#NUM!";
  if (Number.isInteger(v)) return String(v);
  let s = String(parseFloat(v.toPrecision(15)));
  if (s.includes("e")) s = String(parseFloat(v.toPrecision(10)));
  return s;
}
function parseNumText(s) {
  let t = String(s).trim();
  if (t === "") return null;
  let pct = false, neg = false;
  if (t.endsWith("%")) { pct = true; t = t.slice(0, -1).trim(); }
  if (/^\(.*\)$/.test(t)) { neg = true; t = t.slice(1, -1); }
  if (t.startsWith("-")) { neg = !neg; t = t.slice(1); }
  if (t.startsWith("$")) t = t.slice(1);
  if (/^\d{1,3}(,\d{3})+(\.\d*)?$/.test(t)) t = t.replace(/,/g, "");
  if (!/^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t)) {
    const d = parseDateText(String(s).trim());
    return d;
  }
  let v = parseFloat(t);
  if (neg) v = -v;
  if (pct) v /= 100;
  return v;
}
function parseDateText(s) {
  let m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(s);
  if (m) {
    let y = +m[3]; if (y < 100) y += y < 30 ? 2000 : 1900;
    return dateSerial(y, +m[1], +m[2]);
  }
  m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
  if (m) return dateSerial(+m[1], +m[2], +m[3]);
  return null;
}
function toNum(v) {
  if (typeof v === "number") return v;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (v === null || v === undefined) return 0;
  if (isErr(v)) return v;
  if (typeof v === "string") {
    const x = parseNumText(v);
    if (x === null) return ERR("#VALUE!", "Excel tried to do math with the text \"" + v + "\".");
    return x;
  }
  return ERR("#VALUE!");
}
function toStr(v) {
  if (typeof v === "string") return v;
  if (typeof v === "number") return fmtGeneral(v);
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (v === null || v === undefined) return "";
  if (isErr(v)) return v;
  return "";
}
function toBool(v) {
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  if (v === null || v === undefined) return false;
  if (isErr(v)) return v;
  if (typeof v === "string") {
    const u = v.toUpperCase();
    if (u === "TRUE") return true;
    if (u === "FALSE") return false;
    return ERR("#VALUE!", "Excel expected TRUE or FALSE but got the text \"" + v + "\".");
  }
  return ERR("#VALUE!");
}
const typeRank = (v) => (typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2);
function cmp(a, b) {
  if (a === null) a = typeof b === "string" ? "" : typeof b === "boolean" ? false : 0;
  if (b === null) b = typeof a === "string" ? "" : typeof a === "boolean" ? false : 0;
  const ra = typeRank(a), rb = typeRank(b);
  if (ra !== rb) return ra < rb ? -1 : 1;
  if (ra === 0) return Math.abs(a - b) < 1e-12 * Math.max(1, Math.abs(a), Math.abs(b)) ? 0 : a < b ? -1 : 1;
  if (ra === 1) { const x = a.toLowerCase(), y = b.toLowerCase(); return x === y ? 0 : x < y ? -1 : 1; }
  return a === b ? 0 : a ? 1 : -1;
}
function wildRegex(pat) {
  let re = "";
  for (let i = 0; i < pat.length; i++) {
    const ch = pat[i];
    if (ch === "~" && i + 1 < pat.length && "*?~".includes(pat[i + 1])) { re += "\\" + pat[++i]; continue; }
    if (ch === "*") re += "[\\s\\S]*";
    else if (ch === "?") re += "[\\s\\S]";
    else re += ch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }
  return new RegExp("^" + re + "$", "i");
}
const hasWild = (s) => /[*?]/.test(s);

// criteria for COUNTIFS / SUMIFS / AVERAGEIFS
function makeCriterion(c) {
  if (isErr(c)) return () => false;
  if (typeof c === "number") return (v) => (typeof v === "number" && Math.abs(v - c) < 1e-9) || (typeof v === "string" && parseNumText(v) === c);
  if (typeof c === "boolean") return (v) => v === c;
  if (c === null) return (v) => v === null || v === "";
  const s = String(c);
  const m = /^(<=|>=|<>|=|<|>)?([\s\S]*)$/.exec(s);
  const op = m[1] || "=";
  const rest = m[2];
  const num = rest.trim() === "" ? null : parseNumText(rest);
  const boolV = rest.toUpperCase() === "TRUE" ? true : rest.toUpperCase() === "FALSE" ? false : null;
  if (num !== null) {
    return (v) => {
      let x = v;
      if (typeof x !== "number") {
        if (op === "<>") return true;
        if ((op === "=") && typeof x === "string" && parseNumText(x) === num) return true;
        return false;
      }
      const d = cmp(x, num);
      return op === "=" ? d === 0 : op === "<>" ? d !== 0 : op === "<" ? d < 0 : op === ">" ? d > 0 : op === "<=" ? d <= 0 : d >= 0;
    };
  }
  if (boolV !== null) return (v) => (op === "<>" ? v !== boolV : op === "=" ? v === boolV : false);
  if (rest === "") {
    if (op === "=") return (v) => v === null || v === "";
    if (op === "<>") return (v) => !(v === null || v === "");
  }
  if (op === "=" || op === "<>") {
    const re = wildRegex(rest);
    const hit = (v) => typeof v === "string" && re.test(v);
    return op === "=" ? hit : (v) => !hit(v);
  }
  return (v) => {
    if (typeof v !== "string") return false;
    const d = cmp(v, rest);
    return op === "<" ? d < 0 : op === ">" ? d > 0 : op === "<=" ? d <= 0 : d >= 0;
  };
}

// ================================================================== dates (Excel serial numbers)
const EPOCH = Date.UTC(1899, 11, 30);
function dateSerial(y, m, d) { return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / 86400000); }
function serialParts(s) {
  s = Math.floor(s);
  // Excel's calendar starts at 1/0/1900 and (on purpose) includes a fake 2/29/1900
  if (s < 1) return { y: 1900, m: 1, d: 0, dow: 6 };
  if (s === 60) return { y: 1900, m: 2, d: 29, dow: 3 };
  if (s < 60) s += 1;
  const d = new Date(EPOCH + s * 86400000);
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), dow: d.getUTCDay() };
}
let FIXED_NOW = null; // for testing
function nowSerial() {
  const t = FIXED_NOW ? new Date(FIXED_NOW) : new Date();
  const day = dateSerial(t.getFullYear(), t.getMonth() + 1, t.getDate());
  return day + (t.getHours() * 3600 + t.getMinutes() * 60 + t.getSeconds()) / 86400;
}
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

// ================================================================== number formats (TEXT function + display)
function groupThousands(intStr) { return intStr.replace(/\B(?=(\d{3})+(?!\d))/g, ","); }
function roundHalfAway(x, n) {
  const p = Math.pow(10, n);
  const y = Math.abs(x) * p;
  const r = Math.round(parseFloat(y.toPrecision(15)));
  return (Math.sign(x) || 1) * r / p;
}
function formatNumberCode(v, code) {
  // returns string or null if code not understood
  let c = code;
  const sections = splitSections(c);
  if (sections.length > 1) {
    if (v < 0 && sections[1] !== undefined) { c = sections[1]; v = -v; }
    else if (v === 0 && sections[2] !== undefined) c = sections[2];
    else c = sections[0];
  }
  if (/[ymdhs]/i.test(c.replace(/"[^"]*"/g, "")) && !/[0#]/.test(c.replace(/"[^"]*"/g, ""))) return formatDateCode(v, c);
  let pre = "", post = "", body = "";
  const m = /^((?:"[^"]*"|\\.|[^0#?,.])*)([0#?,.]+)((?:"[^"]*"|\\.|[^0#?])*)$/.exec(c);
  if (!m) return null;
  pre = unquote(m[1]); body = m[2]; post = unquote(m[3]);
  let x = v;
  const pctCount = (m[1] + m[3]).replace(/"[^"]*"/g, "").split("%").length - 1;
  for (let i = 0; i < pctCount; i++) x *= 100;
  const dot = body.indexOf(".");
  const intPart = dot >= 0 ? body.slice(0, dot) : body;
  const decPart = dot >= 0 ? body.slice(dot + 1).replace(/,/g, "") : "";
  const decs = decPart.length;
  const useComma = intPart.includes(",");
  const neg = x < 0;
  const rx = roundHalfAway(Math.abs(x), decs);
  let [ip, dp] = rx.toFixed(decs).split(".");
  const minInt = (intPart.replace(/,/g, "").match(/0/g) || []).length;
  if (ip === "0" && minInt === 0) ip = "";
  while (ip.length < minInt) ip = "0" + ip;
  if (useComma) ip = groupThousands(ip);
  if (dp !== undefined) {
    // optional # decimals trim
    const optional = decPart.split("").filter((ch) => ch === "#").length;
    let k = 0;
    while (k < optional && dp.endsWith("0")) { dp = dp.slice(0, -1); k++; }
  }
  let s = ip + (decs ? "." + dp : "");
  if (s.endsWith(".")) s = s.slice(0, -1);
  return (neg && sections.length === 1 ? "-" : "") + pre + s + post;
}
function splitSections(c) {
  const out = []; let cur = "", q = false;
  for (const ch of c) {
    if (ch === '"') q = !q;
    if (ch === ";" && !q) { out.push(cur); cur = ""; } else cur += ch;
  }
  out.push(cur);
  return out;
}
function unquote(s) { return s.replace(/"([^"]*)"/g, "$1").replace(/\\(.)/g, "$1"); }
function formatDateCode(v, code) {
  const p = serialParts(v);
  let out = "", i = 0;
  const c = code;
  while (i < c.length) {
    const ch = c[i];
    if (ch === '"') { const j = c.indexOf('"', i + 1); out += c.slice(i + 1, j < 0 ? c.length : j); i = j < 0 ? c.length : j + 1; continue; }
    if (ch === "\\") { out += c[i + 1] || ""; i += 2; continue; }
    const run = (/^(y+|m+|d+)/i.exec(c.slice(i)) || [null])[0];
    if (run) {
      const L = run.length, t = run[0].toLowerCase();
      if (t === "y") out += L <= 2 ? String(p.y).slice(-2) : String(p.y);
      else if (t === "m") out += L === 1 ? p.m : L === 2 ? String(p.m).padStart(2, "0") : L === 3 ? MONTHS[p.m - 1].slice(0, 3) : L === 5 ? MONTHS[p.m - 1][0] : MONTHS[p.m - 1];
      else out += L === 1 ? p.d : L === 2 ? String(p.d).padStart(2, "0") : L === 3 ? DAYS[p.dow].slice(0, 3) : DAYS[p.dow];
      i += L; continue;
    }
    out += ch; i++;
  }
  return out;
}

// display formats used by the grid: money, money0, pct, pct1, int, dec1, dec2, date, text, general
function display(v, fmt) {
  if (v === null || v === undefined) return "";
  if (isErr(v)) return v.err;
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "string") return v;
  switch (fmt) {
    case "money": return formatNumberCode(v, "$#,##0.00;($#,##0.00)").replace(/^\((.*)\)$/, "-$1");
    case "money0": return formatNumberCode(v, "$#,##0;-$#,##0");
    case "pct": return formatNumberCode(v, "0%");
    case "pct1": return formatNumberCode(v, "0.0%");
    case "int": return formatNumberCode(v, "#,##0");
    case "dec1": return formatNumberCode(v, "#,##0.0");
    case "dec2": return formatNumberCode(v, "#,##0.00");
    case "date": return v >= 1 && v < 2958466 ? formatDateCode(v, "m/d/yyyy") : fmtGeneral(v);
    case "datetime": return formatDateCode(v, "m/d/yyyy") + " " + timeStr(v);
    case "phone": { const s = String(Math.round(v)); return s.length === 10 ? "(" + s.slice(0, 3) + ") " + s.slice(3, 6) + "-" + s.slice(6) : s; }
    default: return fmtGeneral(v);
  }
}
function timeStr(v) {
  const secs = Math.round((v - Math.floor(v)) * 86400);
  let h = Math.floor(secs / 3600); const mi = Math.floor((secs % 3600) / 60);
  const ap = h >= 12 ? "PM" : "AM"; h = h % 12 || 12;
  return h + ":" + String(mi).padStart(2, "0") + " " + ap;
}

// ================================================================== evaluation
// ctx: { wb, sheet, r, c (cell being calculated), dr, dc (fill offset from where the formula was written) }
function shiftRef(node, ctx) {
  const r = node.rabs ? node.r : node.r + ctx.dr;
  const c = node.cabs ? node.c : node.c + ctx.dc;
  return { r, c };
}
function sheetOf(name, ctx) {
  if (!name) return ctx.wb.sheet(ctx.sheet).name;
  const sh = ctx.wb.sheet(name);
  if (!sh) throw new EvalStop(ERR("#REF!", "There's no sheet named \"" + name + "\" in this workbook. Check the spelling of the sheet name."));
  return sh.name;
}
class EvalStop extends Error { constructor(v) { super(v.err); this.val = v; } }

function circularCheck(sheet, r1, c1, r2, c2, ctx) {
  if (ctx.noCirc) return;
  if (sheet.toLowerCase() === String(ctx.sheet).toLowerCase() && ctx.r >= r1 && ctx.r <= r2 && ctx.c >= c1 && ctx.c <= c2) {
    throw new EvalStop(ERR("#CIRC", "Your formula refers to its own cell (" + addr(ctx.r, ctx.c) + "). That's a circular reference. Refer to other cells instead."));
  }
}

function evalNode(node, ctx) {
  switch (node.k) {
    case "num": return node.v;
    case "str": return node.v;
    case "bool": return node.v;
    case "missing": return null;
    case "paren": return evalNode(node.e, ctx);
    case "ref": {
      const sheet = sheetOf(node.sheet, ctx);
      const { r, c } = shiftRef(node, ctx);
      if (r < 0 || c < 0) return ERR("#REF!", "When the formula was copied, a reference slid off the edge of the sheet. Use $ to lock it.");
      circularCheck(sheet, r, c, r, c, ctx);
      return mkRange(sheet, r, c, r, c);
    }
    case "range": {
      const sheet = sheetOf(node.a.sheet, ctx);
      const a = shiftRef(node.a, ctx), b = shiftRef(node.b, ctx);
      if (a.r < 0 || a.c < 0 || b.r < 0 || b.c < 0) return ERR("#REF!", "When the formula was copied, a range slid off the edge of the sheet. Use $ to lock it.");
      const rg = mkRange(sheet, a.r, a.c, b.r, b.c);
      circularCheck(sheet, rg.r1, rg.c1, rg.r2, rg.c2, ctx);
      return rg;
    }
    case "cols": {
      const sheet = sheetOf(node.sheet, ctx);
      const c1 = node.a1 ? node.c1 : node.c1 + ctx.dc;
      const c2 = node.a2 ? node.c2 : node.c2 + ctx.dc;
      const ext = ctx.wb.extent(sheet);
      const rg = mkRange(sheet, 0, c1, Math.max(ext.rows - 1, 0), c2);
      rg.wholeCols = true;
      circularCheck(sheet, rg.r1, rg.c1, rg.r2, rg.c2, ctx);
      return rg;
    }
    case "sref": return evalStructured(node, ctx);
    case "name": {
      const t = ctx.wb.tables[node.v.toLowerCase()];
      if (t) return mkRange(t.sheet, t.top + 1, t.left, t.top + t.nrows, t.left + t.ncols - 1);
      const u = node.v.toUpperCase();
      if (NOT_IN_CLASS[u] !== undefined || FUNCS[u]) return ERR("#NAME?", u + " is a function, so it needs parentheses after it: " + u + "( ... ).");
      if (/^[A-Za-z]{1,3}\d+$/.test(node.v)) return ERR("#NAME?", "\"" + node.v + "\" isn't a valid cell.");
      return ERR("#NAME?", "Excel doesn't recognize \"" + node.v + "\". If it's meant to be text, put it in quotes: \"" + node.v + "\".");
    }
    case "neg": return mapUnary(evalNode(node.e, ctx), (x) => { const n = toNum(x); return isErr(n) ? n : -n; }, ctx);
    case "pct": return mapUnary(evalNode(node.e, ctx), (x) => { const n = toNum(x); return isErr(n) ? n : n / 100; }, ctx);
    case "bin": return evalBinary(node.op, evalNode(node.a, ctx), evalNode(node.b, ctx), ctx);
    case "func": return callFunc(node, ctx);
  }
  return ERR("#VALUE!");
}

function evalStructured(node, ctx) {
  const wb = ctx.wb;
  let t;
  if (node.table) {
    t = wb.tables[node.table.toLowerCase()];
    if (!t) return ERR("#NAME?", "There's no table named \"" + node.table + "\".");
  } else {
    t = wb.tableAt(ctx.sheet, ctx.r, ctx.c);
    if (!t) return ERR("#NAME?", "A reference like [@Column] only works inside an Excel Table.");
  }
  let spec = node.spec.trim();
  let thisRow = false;
  let area = "data";
  if (spec.startsWith("@")) { thisRow = true; spec = spec.slice(1).trim(); }
  // split top-level comma parts: [#This Row],[Col]
  const parts = [];
  if (spec.startsWith("[")) {
    let depth = 0, cur = "";
    for (const ch of spec) {
      if (ch === "[") { depth++; if (depth === 1) { cur = ""; continue; } }
      if (ch === "]") { depth--; if (depth === 0) { parts.push(cur); continue; } }
      if (depth >= 1) cur += ch;
      else if (ch === ":") parts.push(":");
    }
  } else if (spec !== "") parts.push(spec);
  const cols = [];
  let range = false;
  for (const p of parts) {
    const pl = p.trim().toLowerCase();
    if (pl === "#this row") thisRow = true;
    else if (pl === "#data") area = "data";
    else if (pl === "#all") area = "all";
    else if (pl === "#headers") area = "headers";
    else if (pl === ":") range = true;
    else cols.push(p.replace(/'(.)/g, "$1"));
  }
  const colIdx = (name) => {
    const i = t.cols.findIndex((c) => c.toLowerCase() === name.trim().toLowerCase());
    if (i < 0) throw new EvalStop(ERR("#REF!", "The table \"" + t.name + "\" doesn't have a column named \"" + name + "\". Check the spelling."));
    return t.left + i;
  };
  let c1 = t.left, c2 = t.left + t.ncols - 1;
  if (cols.length === 1) c1 = c2 = colIdx(cols[0]);
  else if (cols.length === 2 && range) { c1 = colIdx(cols[0]); c2 = colIdx(cols[1]); }
  else if (cols.length > 1) return ERR("#REF!", "Excel couldn't read that structured reference.");
  let r1 = t.top + 1, r2 = t.top + t.nrows;
  if (area === "all") r1 = t.top;
  if (area === "headers") r1 = r2 = t.top;
  if (thisRow) {
    if (ctx.sheet.toLowerCase() !== t.sheet.toLowerCase() || ctx.r <= t.top || ctx.r > t.top + t.nrows) {
      return ERR("#VALUE!", "[@" + (cols[0] || "") + "] means \"this row of the table\", so it only works in a cell on the same row of that table.");
    }
    r1 = r2 = ctx.r;
  }
  const rg = mkRange(t.sheet, r1, c1, r2, c2);
  circularCheck(t.sheet, rg.r1, rg.c1, rg.r2, rg.c2, ctx);
  return rg;
}

// value of a range as a 2D array
function rangeVals(rg, wb) {
  const out = [];
  for (let r = rg.r1; r <= rg.r2; r++) {
    const row = [];
    for (let c = rg.c1; c <= rg.c2; c++) row.push(wb.get(rg.sheet, r, c));
    out.push(row);
  }
  return out;
}
function toMatrix(v, ctx) {
  if (isRange(v)) return rangeVals(v, ctx.wb);
  if (isArr(v)) return v.arr;
  return [[v]];
}
// a single value from something that might be a range (1 cell) or array
function scalar(v, ctx) {
  if (isRange(v)) {
    if (v.r1 === v.r2 && v.c1 === v.c2) return ctx.wb.get(v.sheet, v.r1, v.c1);
    // implicit intersection on the current row/column (how classic Excel treats A:A in a single cell)
    if (v.c1 === v.c2 && ctx.r >= v.r1 && ctx.r <= v.r2 && v.sheet.toLowerCase() === ctx.sheet.toLowerCase() && ctx.allowIntersect) return ctx.wb.get(v.sheet, ctx.r, v.c1);
    return ERR("#SPILL!", "Part of your formula gives a whole range of values where Excel needs one value.");
  }
  if (isArr(v)) {
    if (v.arr.length === 1 && v.arr[0].length === 1) return v.arr[0][0];
    return ERR("#SPILL!", "Part of your formula gives a whole list of values where Excel needs one value.");
  }
  return v;
}
function mapUnary(v, f, ctx) {
  const multi = (isRange(v) && !(v.r1 === v.r2 && v.c1 === v.c2)) || (isArr(v) && !(v.arr.length === 1 && v.arr[0].length === 1));
  if (multi) return { arr: toMatrix(v, ctx).map((row) => row.map((x) => (isErr(x) ? x : f(x)))) };
  const x = scalar(v, ctx);
  return isErr(x) ? x : f(x);
}

function evalBinary(op, a, b, ctx) {
  const multiA = (isRange(a) && !(a.r1 === a.r2 && a.c1 === a.c2)) || (isArr(a) && !(a.arr.length === 1 && a.arr[0].length === 1));
  const multiB = (isRange(b) && !(b.r1 === b.r2 && b.c1 === b.c2)) || (isArr(b) && !(b.arr.length === 1 && b.arr[0].length === 1));
  if (multiA || multiB) {
    const A = multiA ? toMatrix(a, ctx) : [[scalar(a, ctx)]];
    const B = multiB ? toMatrix(b, ctx) : [[scalar(b, ctx)]];
    const rows = Math.max(A.length, B.length), cols = Math.max(A[0].length, B[0].length);
    const out = [];
    for (let i = 0; i < rows; i++) {
      const row = [];
      for (let j = 0; j < cols; j++) {
        const x = A.length === 1 ? (A[0].length === 1 ? A[0][0] : A[0][j]) : A[i] ? (A[i].length === 1 ? A[i][0] : A[i][j]) : undefined;
        const y = B.length === 1 ? (B[0].length === 1 ? B[0][0] : B[0][j]) : B[i] ? (B[i].length === 1 ? B[i][0] : B[i][j]) : undefined;
        row.push(x === undefined || y === undefined ? ERR("#N/A") : binScalar(op, x, y));
      }
      out.push(row);
    }
    return { arr: out };
  }
  return binScalar(op, scalar(a, ctx), scalar(b, ctx));
}
function binScalar(op, x, y) {
  if (isErr(x)) return x;
  if (isErr(y)) return y;
  if (op === "&") return toStr(x) + toStr(y);
  if (op in { "=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1 }) {
    const d = cmp(x, y);
    return op === "=" ? d === 0 : op === "<>" ? d !== 0 : op === "<" ? d < 0 : op === ">" ? d > 0 : op === "<=" ? d <= 0 : d >= 0;
  }
  const a = toNum(x), b = toNum(y);
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  switch (op) {
    case "+": return a + b;
    case "-": return a - b;
    case "*": return a * b;
    case "/": return b === 0 ? ERR("#DIV/0!", "The formula divides by zero (or by an empty cell).") : a / b;
    case "^": { const r = Math.pow(a, b); return isFinite(r) ? r : ERR("#NUM!"); }
  }
  return ERR("#VALUE!");
}

// ---------------------------------------------------------------- function library
// helpers
function flatValues(v, ctx) {
  // for ranges/arrays: all values; for scalars: [v]
  if (isRange(v) || isArr(v)) return { fromRange: true, vals: [].concat(...toMatrix(v, ctx)) };
  return { fromRange: false, vals: [v] };
}
function numbersFrom(args, ctx, opts) {
  // Excel SUM/AVERAGE/MAX/MIN semantics
  const out = [];
  for (const a of args) {
    const { fromRange, vals } = flatValues(a, ctx);
    for (const v of vals) {
      if (isErr(v)) throw new EvalStop(v);
      if (fromRange) { if (typeof v === "number") out.push(v); }
      else {
        if (v === null) { if (opts && opts.countBlankArg) out.push(0); continue; }
        const n = toNum(v);
        if (isErr(n)) throw new EvalStop(n);
        out.push(n);
      }
    }
  }
  return out;
}
const S = (v, ctx) => { const s = scalar(v, ctx); if (isErr(s)) throw new EvalStop(s); return s; };
const N = (v, ctx) => { const n = toNum(S(v, ctx)); if (isErr(n)) throw new EvalStop(n); return n; };
const T = (v, ctx) => toStr(S(v, ctx));
const B = (v, ctx) => { const b = toBool(S(v, ctx)); if (isErr(b)) throw new EvalStop(b); return b; };
function needArgs(name, args, min, max) {
  if (args.length < min || args.length > max) {
    const want = min === max ? min : min + (max === Infinity ? " or more" : " to " + max);
    throw new EvalStop(ERR("#ARGS", name + " needs " + want + " argument" + (want === 1 ? "" : "s") + ", but your formula gives it " + args.length + "."));
  }
}
function rangeArg(name, v, which) {
  if (!isRange(v) && !isArr(v)) throw new EvalStop(ERR("#VALUE!", "The " + which + " in " + name + " needs to be a range of cells, like A2:A50."));
  return v;
}
function dims(v, ctx) {
  if (isRange(v)) return [v.r2 - v.r1 + 1, v.c2 - v.c1 + 1];
  const m = toMatrix(v, ctx);
  return [m.length, m[0].length];
}
function ifsCore(name, ctx, rangesAndCrit) {
  // returns list of booleans (flattened) of matches
  let size = null, mask = null;
  for (let i = 0; i < rangesAndCrit.length; i += 2) {
    const rg = rangeArg(name, rangesAndCrit[i], "criteria range");
    const d = dims(rg, ctx);
    if (size && (size[0] !== d[0] || size[1] !== d[1])) throw new EvalStop(ERR("#VALUE!", "In " + name + ", every range has to be the same size (same number of rows). One of yours is a different size."));
    size = d;
    const crit = rangesAndCrit[i + 1] === undefined ? null : S(rangesAndCrit[i + 1], ctx);
    const test = makeCriterion(crit);
    const vals = [].concat(...toMatrix(rg, ctx));
    const m = vals.map((v) => test(v));
    mask = mask ? mask.map((x, j) => x && m[j]) : m;
  }
  return { mask, size };
}
function lookupEq(target, v) {
  if (typeof target === "string" && typeof v === "string") return hasWild(target) ? wildRegex(target).test(v) : target.toLowerCase() === v.toLowerCase();
  if (target === null) return false;
  return cmp(target, v) === 0 && typeRank(target) === typeRank(v === null ? target : v) && v !== null;
}
function approxIndex(list, target) {
  // like Excel's binary search on sorted data: last item <= target of the same type
  let lo = 0, hi = list.length - 1, ans = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    let v = list[mid];
    // skip blanks / other types toward the low side
    let k = mid;
    while (k >= lo && (list[k] === null || typeRank(list[k]) !== typeRank(target))) k--;
    if (k < lo) { lo = mid + 1; continue; }
    v = list[k];
    if (cmp(v, target) <= 0) { ans = k; lo = mid + 1; } else hi = k - 1;
  }
  return ans;
}

const FUNCS = {
  SUM(a, ctx) { needArgs("SUM", a, 1, 255); return numbersFrom(a, ctx).reduce((s, x) => s + x, 0); },
  AVERAGE(a, ctx) { needArgs("AVERAGE", a, 1, 255); const n = numbersFrom(a, ctx); return n.length ? n.reduce((s, x) => s + x, 0) / n.length : ERR("#DIV/0!", "AVERAGE didn't find any numbers to average."); },
  MAX(a, ctx) { needArgs("MAX", a, 1, 255); const n = numbersFrom(a, ctx); return n.length ? Math.max(...n) : 0; },
  MIN(a, ctx) { needArgs("MIN", a, 1, 255); const n = numbersFrom(a, ctx); return n.length ? Math.min(...n) : 0; },
  COUNT(a, ctx) {
    needArgs("COUNT", a, 1, 255);
    let k = 0;
    for (const x of a) {
      const { fromRange, vals } = flatValues(x, ctx);
      for (const v of vals) {
        if (typeof v === "number") k++;
        else if (!fromRange && (typeof v === "boolean" || (typeof v === "string" && parseNumText(v) !== null))) k++;
      }
    }
    return k;
  },
  COUNTA(a, ctx) {
    needArgs("COUNTA", a, 1, 255);
    let k = 0;
    for (const x of a) {
      if (x && x.k === "missing") continue;
      const { vals } = flatValues(x, ctx);
      for (const v of vals) if (v !== null) k++;
    }
    return k;
  },
  IF() { return ERR("#VALUE!"); },   // handled in callFunc
  AND(a, ctx) {
    needArgs("AND", a, 1, 255);
    let r = true;
    for (const x of a) {
      const { fromRange, vals } = flatValues(x, ctx);
      for (const v of vals) {
        if (isErr(v)) throw new EvalStop(v);
        if (fromRange && typeof v !== "boolean" && typeof v !== "number") continue;
        if (!B(v, ctx)) r = false;
      }
    }
    return r;
  },
  OR(a, ctx) {
    needArgs("OR", a, 1, 255);
    let r = false;
    for (const x of a) {
      const { fromRange, vals } = flatValues(x, ctx);
      for (const v of vals) {
        if (isErr(v)) throw new EvalStop(v);
        if (fromRange && typeof v !== "boolean" && typeof v !== "number") continue;
        if (B(v, ctx)) r = true;
      }
    }
    return r;
  },
  NOT(a, ctx) { needArgs("NOT", a, 1, 1); return !B(a[0], ctx); },
  IFERROR() { return ERR("#VALUE!"); },   // handled in callFunc
  COUNTIFS(a, ctx) {
    if (a.length < 2 || a.length % 2) throw new EvalStop(ERR("#ARGS", "COUNTIFS takes pairs: a range, then its criteria (range1, criteria1, range2, criteria2, ...)."));
    return ifsCore("COUNTIFS", ctx, a).mask.filter(Boolean).length;
  },
  COUNTIF(a, ctx) { needArgs("COUNTIF", a, 2, 2); return ifsCore("COUNTIF", ctx, a).mask.filter(Boolean).length; },
  SUMIFS(a, ctx) {
    if (a.length < 3 || a.length % 2 === 0) throw new EvalStop(ERR("#ARGS", "SUMIFS starts with the range to add up, then pairs of range and criteria: SUMIFS(sum_range, range1, criteria1, ...)."));
    const sr = rangeArg("SUMIFS", a[0], "sum range");
    const { mask, size } = ifsCore("SUMIFS", ctx, a.slice(1));
    const d = dims(sr, ctx);
    if (d[0] !== size[0] || d[1] !== size[1]) throw new EvalStop(ERR("#VALUE!", "In SUMIFS, the sum range and the criteria ranges have to be the same size."));
    const vals = [].concat(...toMatrix(sr, ctx));
    let s = 0;
    vals.forEach((v, i) => { if (mask[i]) { if (isErr(v)) throw new EvalStop(v); if (typeof v === "number") s += v; } });
    return s;
  },
  SUMIF(a, ctx) {
    needArgs("SUMIF", a, 2, 3);
    const rg = rangeArg("SUMIF", a[0], "range");
    const { mask } = ifsCore("SUMIF", ctx, [a[0], a[1]]);
    const sr = a[2] ? rangeArg("SUMIF", a[2], "sum range") : rg;
    const vals = [].concat(...toMatrix(isRange(sr) ? mkRange(sr.sheet, sr.r1, sr.c1, sr.r1 + (rg.r2 - rg.r1), sr.c1 + (rg.c2 - rg.c1)) : sr, ctx));
    let s = 0;
    vals.forEach((v, i) => { if (mask[i] && typeof v === "number") s += v; });
    return s;
  },
  AVERAGEIFS(a, ctx) {
    if (a.length < 3 || a.length % 2 === 0) throw new EvalStop(ERR("#ARGS", "AVERAGEIFS starts with the range to average, then pairs of range and criteria."));
    const sr = rangeArg("AVERAGEIFS", a[0], "average range");
    const { mask, size } = ifsCore("AVERAGEIFS", ctx, a.slice(1));
    const d = dims(sr, ctx);
    if (d[0] !== size[0] || d[1] !== size[1]) throw new EvalStop(ERR("#VALUE!", "In AVERAGEIFS, the average range and the criteria ranges have to be the same size."));
    const vals = [].concat(...toMatrix(sr, ctx));
    let s = 0, k = 0;
    vals.forEach((v, i) => { if (mask[i] && typeof v === "number") { s += v; k++; } });
    return k ? s / k : ERR("#DIV/0!", "No rows matched the criteria, so there was nothing to average.");
  },
  AVERAGEIF(a, ctx) {
    needArgs("AVERAGEIF", a, 2, 3);
    const rg = rangeArg("AVERAGEIF", a[0], "range");
    const { mask } = ifsCore("AVERAGEIF", ctx, [a[0], a[1]]);
    const sr = a[2] ? rangeArg("AVERAGEIF", a[2], "average range") : rg;
    const vals = [].concat(...toMatrix(isRange(sr) ? mkRange(sr.sheet, sr.r1, sr.c1, sr.r1 + (rg.r2 - rg.r1), sr.c1 + (rg.c2 - rg.c1)) : sr, ctx));
    let s = 0, k = 0;
    vals.forEach((v, i) => { if (mask[i] && typeof v === "number") { s += v; k++; } });
    return k ? s / k : ERR("#DIV/0!", "No rows matched the criteria, so there was nothing to average.");
  },
  SUMPRODUCT(a, ctx) {
    needArgs("SUMPRODUCT", a, 1, 255);
    let size = null;
    const mats = a.map((x) => {
      const m = toMatrix(x, ctx);
      if (size && (size[0] !== m.length || size[1] !== m[0].length)) throw new EvalStop(ERR("#VALUE!", "SUMPRODUCT's ranges have to be the same size. One of yours has a different number of rows."));
      size = [m.length, m[0].length];
      return m;
    });
    let s = 0;
    for (let i = 0; i < size[0]; i++) for (let j = 0; j < size[1]; j++) {
      let p = 1;
      for (const m of mats) {
        const v = m[i][j];
        if (isErr(v)) throw new EvalStop(v);
        p *= typeof v === "number" ? v : 0;
      }
      s += p;
    }
    return s;
  },
  VLOOKUP(a, ctx) {
    needArgs("VLOOKUP", a, 3, 4);
    const target = S(a[0], ctx);
    const rg = rangeArg("VLOOKUP", a[1], "table array");
    const col = Math.trunc(N(a[2], ctx));
    const approx = a.length < 4 || a[3] === null ? true : B(a[3], ctx);
    const m = toMatrix(rg, ctx);
    if (col < 1) return ERR("#VALUE!", "VLOOKUP's column number has to be 1 or more.");
    if (col > m[0].length) return ERR("#REF!", "VLOOKUP's column number (" + col + ") is bigger than the number of columns in the table array (" + m[0].length + ").");
    const first = m.map((r) => r[0]);
    let idx = -1;
    if (target === null) return ERR("#N/A", "VLOOKUP was asked to look up an empty cell.");
    if (!approx) idx = first.findIndex((v) => lookupEq(target, v));
    else idx = approxIndex(first, target);
    if (idx < 0) return ERR("#N/A", approx ? "VLOOKUP (approximate match) found nothing smaller than or equal to " + toStr(target) + "." : "VLOOKUP couldn't find " + (typeof target === "string" ? "\"" + target + "\"" : toStr(target)) + " in the first column of the table array.");
    const v = m[idx][col - 1];
    return v === null ? 0 : v;
  },
  MATCH(a, ctx) {
    needArgs("MATCH", a, 2, 3);
    const target = S(a[0], ctx);
    const rg = rangeArg("MATCH", a[1], "lookup array");
    const type = a.length < 3 || a[2] === null ? 1 : Math.trunc(N(a[2], ctx));
    const m = toMatrix(rg, ctx);
    if (m.length > 1 && m[0].length > 1) return ERR("#N/A", "MATCH needs a single row or a single column, not a block of cells.");
    const list = m.length > 1 ? m.map((r) => r[0]) : m[0];
    let idx = -1;
    if (type === 0) idx = list.findIndex((v) => lookupEq(target, v));
    else if (type > 0) idx = approxIndex(list, target);
    else {
      for (let i = 0; i < list.length; i++) {
        if (list[i] === null || typeRank(list[i]) !== typeRank(target)) continue;
        if (cmp(list[i], target) >= 0) idx = i; else break;
      }
    }
    if (idx < 0) return ERR("#N/A", "MATCH couldn't find " + (typeof target === "string" ? "\"" + target + "\"" : toStr(target)) + (type === 0 ? "." : ". (Did you mean to put 0 at the end for an exact match?)"));
    return idx + 1;
  },
  INDEX(a, ctx) {
    needArgs("INDEX", a, 2, 3);
    const rg = rangeArg("INDEX", a[0], "array");
    const m = toMatrix(rg, ctx);
    let r = Math.trunc(N(a[1], ctx));
    let c = a.length >= 3 && a[2] !== null ? Math.trunc(N(a[2], ctx)) : null;
    if (c === null) {
      if (m.length === 1) { c = r; r = 1; } else c = 1;
      if (m.length > 1 && m[0].length > 1 && a.length < 3) return ERR("#REF!", "When INDEX gets a block with several columns, it also needs a column number: INDEX(range, row, column).");
    }
    if (r < 0 || c < 0) return ERR("#VALUE!");
    if (r > m.length || c > m[0].length) return ERR("#REF!", "INDEX was asked for row " + r + (m[0].length > 1 ? ", column " + c : "") + ", which is outside the range.");
    if (r === 0 || c === 0) {
      if (!isRange(rg)) return ERR("#VALUE!");
      if (r === 0 && c === 0) return rg;
      if (r === 0) return mkRange(rg.sheet, rg.r1, rg.c1 + c - 1, rg.r2, rg.c1 + c - 1);
      return mkRange(rg.sheet, rg.r1 + r - 1, rg.c1, rg.r1 + r - 1, rg.c2);
    }
    if (isRange(rg)) return mkRange(rg.sheet, rg.r1 + r - 1, rg.c1 + c - 1, rg.r1 + r - 1, rg.c1 + c - 1);
    return m[r - 1][c - 1];
  },
  TODAY(a) { needArgs("TODAY", a, 0, 0); return Math.floor(nowSerial()); },
  NOW(a) { needArgs("NOW", a, 0, 0); return nowSerial(); },
  DATE(a, ctx) { needArgs("DATE", a, 3, 3); let y = Math.trunc(N(a[0], ctx)); if (y < 1900) y += 1900; return dateSerial(y, Math.trunc(N(a[1], ctx)), Math.trunc(N(a[2], ctx))); },
  YEAR(a, ctx) { needArgs("YEAR", a, 1, 1); return serialParts(dateArg(a[0], ctx, "YEAR")).y; },
  MONTH(a, ctx) { needArgs("MONTH", a, 1, 1); return serialParts(dateArg(a[0], ctx, "MONTH")).m; },
  DAY(a, ctx) { needArgs("DAY", a, 1, 1); return serialParts(dateArg(a[0], ctx, "DAY")).d; },
  ROUND(a, ctx) { needArgs("ROUND", a, 2, 2); return roundHalfAway(N(a[0], ctx), Math.trunc(N(a[1], ctx))); },
  ROUNDUP(a, ctx) { needArgs("ROUNDUP", a, 2, 2); const x = N(a[0], ctx), p = Math.pow(10, Math.trunc(N(a[1], ctx))); return Math.sign(x) * Math.ceil(parseFloat((Math.abs(x) * p).toPrecision(15))) / p; },
  ROUNDDOWN(a, ctx) { needArgs("ROUNDDOWN", a, 2, 2); const x = N(a[0], ctx), p = Math.pow(10, Math.trunc(N(a[1], ctx))); return Math.sign(x) * Math.floor(parseFloat((Math.abs(x) * p).toPrecision(15))) / p; },
  ABS(a, ctx) { needArgs("ABS", a, 1, 1); return Math.abs(N(a[0], ctx)); },
  INT(a, ctx) { needArgs("INT", a, 1, 1); return Math.floor(N(a[0], ctx)); },
  ISNUMBER(a, ctx) { needArgs("ISNUMBER", a, 1, 1); const v = safeScalar(a[0], ctx); return typeof v === "number"; },
  ISTEXT(a, ctx) { needArgs("ISTEXT", a, 1, 1); const v = safeScalar(a[0], ctx); return typeof v === "string"; },
  ISBLANK(a, ctx) { needArgs("ISBLANK", a, 1, 1); const v = safeScalar(a[0], ctx); return v === null && isRange(a[0]); },
  ISERROR(a, ctx) { needArgs("ISERROR", a, 1, 1); const v = safeScalar(a[0], ctx); return isErr(v) && v.err !== "#CIRC" && v.err !== "#ARGS"; },
  LEN(a, ctx) { needArgs("LEN", a, 1, 1); return T(a[0], ctx).length; },
  LEFT(a, ctx) { needArgs("LEFT", a, 1, 2); const s = T(a[0], ctx); const k = a.length > 1 && a[1] !== null ? Math.trunc(N(a[1], ctx)) : 1; if (k < 0) return ERR("#VALUE!", "LEFT can't take a negative number of characters."); return s.slice(0, k); },
  RIGHT(a, ctx) { needArgs("RIGHT", a, 1, 2); const s = T(a[0], ctx); const k = a.length > 1 && a[1] !== null ? Math.trunc(N(a[1], ctx)) : 1; if (k < 0) return ERR("#VALUE!", "RIGHT can't take a negative number of characters."); return k === 0 ? "" : s.slice(-k); },
  MID(a, ctx) { needArgs("MID", a, 3, 3); const s = T(a[0], ctx); const st = Math.trunc(N(a[1], ctx)), k = Math.trunc(N(a[2], ctx)); if (st < 1 || k < 0) return ERR("#VALUE!"); return s.substr(st - 1, k); },
  TRIM(a, ctx) { needArgs("TRIM", a, 1, 1); return T(a[0], ctx).replace(/ +/g, " ").replace(/^ | $/g, ""); },
  UPPER(a, ctx) { needArgs("UPPER", a, 1, 1); return T(a[0], ctx).toUpperCase(); },
  LOWER(a, ctx) { needArgs("LOWER", a, 1, 1); return T(a[0], ctx).toLowerCase(); },
  PROPER(a, ctx) {
    needArgs("PROPER", a, 1, 1);
    const s = T(a[0], ctx);
    let out = "", prevLetter = false;
    for (const ch of s) {
      const isLetter = ch.toLowerCase() !== ch.toUpperCase();
      out += isLetter ? (prevLetter ? ch.toLowerCase() : ch.toUpperCase()) : ch;
      prevLetter = isLetter;
    }
    return out;
  },
  SUBSTITUTE(a, ctx) {
    needArgs("SUBSTITUTE", a, 3, 4);
    const s = T(a[0], ctx), old = T(a[1], ctx), neu = T(a[2], ctx);
    if (old === "") return s;
    if (a.length === 4 && a[3] !== null) {
      const k = Math.trunc(N(a[3], ctx));
      if (k < 1) return ERR("#VALUE!");
      let idx = -1, from = 0;
      for (let i = 0; i < k; i++) { idx = s.indexOf(old, from); if (idx < 0) return s; from = idx + old.length; }
      return s.slice(0, idx) + neu + s.slice(idx + old.length);
    }
    return s.split(old).join(neu);
  },
  SEARCH(a, ctx) {
    needArgs("SEARCH", a, 2, 3);
    const f = T(a[0], ctx), s = T(a[1], ctx);
    const st = a.length > 2 && a[2] !== null ? Math.trunc(N(a[2], ctx)) : 1;
    if (st < 1 || st > s.length + 1) return ERR("#VALUE!", "SEARCH's start position is outside the text.");
    if (f === "") return st;
    if (hasWild(f)) {
      const re = wildRegex(f);
      const inner = new RegExp(re.source.slice(1, -1), "i");
      for (let i = st - 1; i < s.length; i++) {
        for (let j = s.length; j > i; j--) if (re.test(s.slice(i, j))) return i + 1;
      }
      void inner;
      return ERR("#VALUE!", "SEARCH couldn't find \"" + f + "\" in \"" + s + "\".");
    }
    const i = s.toLowerCase().indexOf(f.toLowerCase(), st - 1);
    return i < 0 ? ERR("#VALUE!", "SEARCH couldn't find \"" + f + "\" in \"" + s + "\".") : i + 1;
  },
  FIND(a, ctx) {
    needArgs("FIND", a, 2, 3);
    const f = T(a[0], ctx), s = T(a[1], ctx);
    const st = a.length > 2 && a[2] !== null ? Math.trunc(N(a[2], ctx)) : 1;
    const i = s.indexOf(f, st - 1);
    return i < 0 ? ERR("#VALUE!", "FIND couldn't find \"" + f + "\" in \"" + s + "\". (FIND is case-sensitive.)") : i + 1;
  },
  CONCATENATE(a, ctx) {
    needArgs("CONCATENATE", a, 1, 255);
    return a.map((x) => {
      if ((isRange(x) && !(x.r1 === x.r2 && x.c1 === x.c2))) throw new EvalStop(ERR("#VALUE!", "CONCATENATE joins single cells. List them one at a time: CONCATENATE(A2, \" \", B2)."));
      return T(x, ctx);
    }).join("");
  },
  VALUE(a, ctx) {
    needArgs("VALUE", a, 1, 1);
    const v = S(a[0], ctx);
    if (typeof v === "number") return v;
    if (typeof v === "boolean") return ERR("#VALUE!", "VALUE can't turn TRUE/FALSE into a number.");
    const x = parseNumText(toStr(v));
    return x === null ? ERR("#VALUE!", "VALUE couldn't turn \"" + toStr(v) + "\" into a number.") : x;
  },
  TEXT(a, ctx) {
    needArgs("TEXT", a, 2, 2);
    const v = S(a[0], ctx), code = T(a[1], ctx);
    let x = typeof v === "number" ? v : typeof v === "string" ? parseNumText(v) : toNum(v);
    if (x === null || isErr(x)) return toStr(v);
    const out = formatNumberCode(x, code);
    return out === null ? ERR("#VALUE!", "This practice tool doesn't understand the format code \"" + code + "\". Try codes like \"0.00\", \"$#,##0\", \"0%\", \"mm/dd/yyyy\", or \"mmmm\".") : out;
  },
};
function dateArg(v, ctx, name) {
  const s = S(v, ctx);
  if (s === null) return 0;
  const n = toNum(s);
  if (isErr(n)) throw new EvalStop(ERR("#VALUE!", name + " needs a date, but got the text \"" + toStr(s) + "\"."));
  if (n < 0) throw new EvalStop(ERR("#NUM!"));
  return n;
}
function safeScalar(v, ctx) {
  try { return scalar(v, ctx); } catch (e) { if (e instanceof EvalStop) return e.val; throw e; }
}

// functions whose arguments should stay as ranges (everything else gets scalars)
const KEEP_RANGE = new Set(["SUM", "AVERAGE", "MAX", "MIN", "COUNT", "COUNTA", "AND", "OR", "COUNTIFS", "COUNTIF", "SUMIFS", "SUMIF", "AVERAGEIFS", "AVERAGEIF", "SUMPRODUCT", "VLOOKUP", "MATCH", "INDEX", "ISBLANK", "CONCATENATE", "IFERROR", "ISERROR", "ISNUMBER", "ISTEXT"]);

function callFunc(node, ctx) {
  const name = node.name;
  if (NOT_IN_CLASS[name] !== undefined) throw new EvalStop(ERR("#CLASS", NOT_IN_CLASS[name] || name + " isn't part of OM 305."));
  const f = FUNCS[name];
  if (!f) throw new EvalStop(ERR("#NAME?", "Excel doesn't have a function called " + name + ". Check the spelling."));
  const allowed = ctx.allowed;
  if (allowed && !allowed.has(name)) throw new EvalStop(ERR("#CLASS", name + " isn't one of the functions from OM 305."));
  // IF / IFERROR evaluate lazily so the unused branch can't cause trouble
  if (name === "IF") {
    needArgs("IF", node.args, 1, 3);
    const test = evalArg(node.args[0], ctx, false);
    const t = B(test, ctx);
    const pick = t ? node.args[1] : node.args[2];
    if (pick === undefined) return t ? true : false;
    if (pick.k === "missing") return 0;
    return S(evalArg(pick, ctx, false), ctx);
  }
  if (name === "IFERROR") {
    needArgs("IFERROR", node.args, 2, 2);
    let v;
    try { v = safeScalar(evalArg(node.args[0], ctx, false), ctx); } catch (e) { if (e instanceof EvalStop) v = e.val; else throw e; }
    if (isErr(v) && !["#CIRC", "#ARGS", "#CLASS", "#NAME?"].includes(v.err)) return node.args[1].k === "missing" ? 0 : S(evalArg(node.args[1], ctx, false), ctx);
    if (isErr(v) && v.err !== "#NAME?") throw new EvalStop(v);
    return v === null ? 0 : v;
  }
  const keep = KEEP_RANGE.has(name);
  const args = node.args.map((a) => (a.k === "missing" ? null : evalArg(a, ctx, keep)));
  return f(args, ctx);
}
function evalArg(a, ctx, keep) {
  try {
    return evalNode(a, ctx);
  } catch (e) {
    if (e instanceof EvalStop) {
      if (["#CIRC", "#ARGS", "#CLASS", "#NAME?"].includes(e.val.err)) throw e;
      return e.val;
    }
    throw e;
  }
}

// ================================================================== public API
// Evaluate a parsed formula in a cell. Returns a final scalar value (number/string/bool/null/error).
function evaluate(parsed, wb, sheet, r, c, dr, dc, opts) {
  const ctx = Object.assign({ wb, sheet, r, c, dr: dr || 0, dc: dc || 0 }, opts || {});
  let v;
  try {
    v = evalNode(parsed.ast, ctx);
    if (isRange(v) && v.c1 === v.c2 && !(v.r1 === v.r2)) {
      v = ERR("#SPILL!", "Your formula returns a whole column of values instead of one value. Did you mean to look up one item, or add them up with SUM?");
    } else v = scalar(v, ctx);
  } catch (e) {
    if (e instanceof EvalStop) v = e.val;
    else if (e instanceof RangeError) v = ERR("#VALUE!", "That formula is too complicated for the practice tool.");
    else throw e;
  }
  if (v === null) v = 0; // an empty result shows as 0 in Excel
  if (typeof v === "number" && !isFinite(v)) v = ERR("#NUM!");
  return v;
}

// Fill a formula into a block (like typing it in the top-left cell and dragging down/right).
// target: { sheet, r, c, rows, cols }. Writes results into the workbook. Returns 2D array of values.
function fill(formula, wb, target, opts) {
  const parsed = typeof formula === "string" ? parse(formula) : formula;
  const out = [];
  for (let i = 0; i < target.rows; i++) {
    const row = [];
    for (let j = 0; j < (target.cols || 1); j++) {
      const v = evaluate(parsed, wb, target.sheet, target.r + i, target.c + j, i, j, opts);
      wb.set(target.sheet, target.r + i, target.c + j, v);
      row.push(v);
    }
    out.push(row);
  }
  return out;
}

// Toggle $ signs on the reference at/left of the cursor (Excel's F4 key). Returns { text, cursor }.
function cycleRef(text, cursor) {
  const re = /(('(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?:(:)(\$?)([A-Za-z]{1,3})(\$?)(\d+))?/g;
  let m;
  while ((m = re.exec(text))) {
    const start = m.index, end = m.index + m[0].length;
    if (cursor >= start && cursor <= end) {
      if (start > 0 && /[A-Za-z0-9_"]/.test(text[start - 1]) && !m[1]) continue;
      const next = (cabs, rabs) => (!cabs && !rabs ? ["$", "$"] : cabs && rabs ? ["", "$"] : !cabs && rabs ? ["$", ""] : ["", ""]);
      const [c1, r1] = next(!!m[3], !!m[5]);
      let s = (m[1] || "") + c1 + m[4] + r1 + m[6];
      if (m[7]) { const [c2, r2] = next(!!m[8], !!m[10]); s += ":" + c2 + m[9] + r2 + m[11]; }
      return { text: text.slice(0, start) + s + text.slice(end), cursor: start + s.length };
    }
  }
  return { text, cursor };
}

const API = {
  Workbook, parse, evaluate, fill, display, fmtGeneral, cmp, toStr, isErr, cycleRef,
  addr, numToCol, colToNum, parseAddr, sheetPrefix, dateSerial, serialParts, FormulaError,
  TAUGHT, EXTRAS, NOT_IN_CLASS, FUNCS, formatNumberCode,
  setFixedNow(d) { FIXED_NOW = d; },
};
if (typeof module !== "undefined" && module.exports) module.exports = API;
else root.XL = API;
})(typeof self !== "undefined" ? self : this);
