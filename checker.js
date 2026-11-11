/* OM 305 Excel Practice — answer checking.
   Runs the student's formula and the model formula on the shown data AND on two hidden versions of the data,
   so typing in numbers instead of cell references doesn't pass. Then explains what went wrong. */
(function (root) {
"use strict";
const XL = root.XL || require("./engine.js");
const Bank = root.Bank || require("./problems.js");
const ALLOWED = new Set(XL.TAUGHT.concat(XL.EXTRAS));
const ALT_SEEDS = [7777, 31337];

function instantiate(prob, pSeed, dSeed) {
  const b = prob.build(Bank.makeR(pSeed, dSeed));
  const wb = new XL.Workbook({ sheets: b.sheets, tables: b.tables });
  for (const p of b.prefill || []) {
    const a = XL.parseAddr(p.cell);
    const t = { sheet: p.sheet, r: a.r, c: a.c, rows: p.rows || 1, cols: p.cols || 1 };
    XL.fill(p.f, wb, t);
    if (p.fmt) markFmt(wb, t, p.fmt);
  }
  const a = XL.parseAddr(b.target.cell);
  const target = { sheet: b.target.sheet, r: a.r, c: a.c, rows: b.target.rows || 1, cols: b.target.cols || 1, fmt: b.target.fmt };
  if (target.fmt) markFmt(wb, target, target.fmt);
  return { b, wb, target };
}
function markFmt(wb, t, fmt) {
  const sh = wb.sheet(t.sheet);
  for (let i = 0; i < t.rows; i++) for (let j = 0; j < (t.cols || 1); j++) sh.fmt[(t.r + i) + ":" + (t.c + j)] = fmt;
}

function walk(node, fn) {
  if (!node || typeof node !== "object") return;
  fn(node);
  for (const k of ["e", "a", "b"]) if (node[k]) walk(node[k], fn);
  if (node.args) node.args.forEach((x) => walk(x, fn));
}
function hasRefs(ast) { let f = false; walk(ast, (n) => { if (["ref", "range", "cols", "sref", "name"].includes(n.k)) f = true; }); return f; }
function hasStructured(ast, wb) { let f = false; walk(ast, (n) => { if (n.k === "sref" || (n.k === "name" && wb.tables[n.v.toLowerCase()])) f = true; }); return f; }
function relativeRanges(ast) {
  let f = false;
  walk(ast, (n) => { if (n.k === "range" && (!n.a.rabs || !n.b.rabs)) f = true; });
  return f;
}

const GENERIC_WHY = {
  "#N/A": "#N/A means a lookup couldn't find what it was looking for. Check the lookup value, the range, and that you used FALSE (or 0) for an exact match.",
  "#VALUE!": "#VALUE! usually means the formula tried to do math with text, or SEARCH couldn't find what it was looking for.",
  "#REF!": "#REF! means a reference points outside the range or sheet, for example a VLOOKUP column number bigger than the table.",
  "#DIV/0!": "#DIV/0! means something was divided by zero or by an empty cell. If that comes from copying the formula down, a $ might be missing.",
  "#NAME?": "#NAME? means Excel saw a word it doesn't know. Is a function misspelled, or is text missing its quotes?",
  "#SPILL!": "Your formula returns a whole list of values instead of one value.",
  "#NUM!": "#NUM! means the math produced an impossible number.",
  "#CIRC": "Your formula refers to its own cell (a circular reference).",
};

function numLike(s) { return typeof s === "string" && /^\s*-?\d+(\.\d+)?\s*$/.test(s); }
function same(e, g, strict) {
  if (XL.isErr(e) || XL.isErr(g)) return XL.isErr(e) && XL.isErr(g) && e.err === g.err;
  if (typeof e === "number" && typeof g === "number") return Math.abs(e - g) <= 1e-7 * Math.max(1, Math.abs(e));
  if (!strict) {
    if (typeof e === "string" && typeof g === "number" && numLike(e)) return Math.abs(parseFloat(e) - g) < 1e-9;
    if (typeof e === "number" && typeof g === "string" && numLike(g)) return false;
  }
  return e === g;
}
function show(v) {
  if (v === "") return "a blank cell (\"\")";
  if (typeof v === "string") return "\"" + v + "\"";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (XL.isErr(v)) return v.err;
  return XL.fmtGeneral(Math.round(v * 1e6) / 1e6);
}
function cellName(t, i, j) { return XL.addr(t.r + i, t.c + j); }

// pre-checks that don't need to run anything
function lint(prob, b, parsed, wb) {
  for (const f of parsed.funcs) {
    if (XL.NOT_IN_CLASS[f] !== undefined) return XL.NOT_IN_CLASS[f] || f + " isn't part of OM 305.";
    if (!XL.FUNCS[f]) return "Excel doesn't have a function called " + f + ". Check the spelling.";
    if (!ALLOWED.has(f)) return f + " isn't one of the functions we use in OM 305.";
    if (b.forbid && b.forbid.includes(f)) return "This problem is practice for " + (b.require || []).filter((x) => x !== "IFERROR").join("/") + ", so " + f + " isn't accepted here.";
  }
  const missing = (b.require || []).filter((f) => !parsed.funcs.includes(f));
  if (missing.length) return "This one is practice for " + (b.require || []).join(" + ") + ". Your formula needs to use " + missing.join(" and ") + ".";
  if (b.requireStructured && !hasStructured(parsed.ast, wb)) return "This one is practice for Excel Table formulas. Use structured references like [@Facility] instead of cell addresses.";
  if (!hasRefs(parsed.ast) && !b.allowConst) return "Your formula doesn't refer to any cells. Formulas should point at the data (like B2) so they update when the data changes.";
  return null;
}

function runFormula(parsed, inst) {
  const wb = inst.wb;
  return XL.fill(parsed, wb, inst.target, { allowed: ALLOWED });
}

// Just calculate (the "Enter" key). Returns { wb, target, values, error }
function preview(prob, pSeed, dSeed, formula) {
  const inst = instantiate(prob, pSeed, dSeed);
  let parsed;
  try { parsed = XL.parse(formula); } catch (e) { if (e instanceof XL.FormulaError) return { inst, error: e.friendly, syntax: true }; throw e; }
  for (const f of parsed.funcs) {
    if (XL.NOT_IN_CLASS[f] !== undefined) return { inst, error: XL.NOT_IN_CLASS[f] || f + " isn't part of OM 305.", syntax: true };
    if (!XL.FUNCS[f]) return { inst, error: "Excel doesn't have a function called " + f + ". Check the spelling.", syntax: true };
  }
  const vals = runFormula(parsed, inst);
  const bad = firstSpecial(vals);
  if (bad) return { inst, error: bad.why, syntax: true };
  const err = firstError(vals, inst.target);
  return { inst, values: vals, error: err ? err.msg : null };
}
function firstSpecial(vals) {
  for (const row of vals) for (const v of row) if (XL.isErr(v) && (v.err === "#ARGS" || v.err === "#CLASS")) return v;
  return null;
}
function firstError(vals, t) {
  for (let i = 0; i < vals.length; i++) for (let j = 0; j < vals[i].length; j++) {
    const v = vals[i][j];
    if (XL.isErr(v)) return { i, j, v, msg: cellName(t, i, j) + " shows " + (v.err === "#CIRC" ? "a circular reference" : v.err) + ". " + (v.why || GENERIC_WHY[v.err] || "") };
  }
  return null;
}

// Full check. Returns { ok, msg, kind, inst (student's view on shown data), wrong: [[i,j]...] }
function check(prob, pSeed, dSeed, formula) {
  const shown = instantiate(prob, pSeed, dSeed);
  const b = shown.b;
  let parsed;
  try { parsed = XL.parse(formula); } catch (e) {
    if (e instanceof XL.FormulaError) return { ok: false, kind: "syntax", msg: e.friendly, inst: shown };
    throw e;
  }
  const lintMsg = lint(prob, b, parsed, shown.wb);
  if (lintMsg) {
    try { runFormula(parsed, shown); } catch (e) { /* show what we can */ }
    return { ok: false, kind: "rule", msg: lintMsg, inst: shown };
  }
  const strict = !!(b.strictType || (Bank.LEVEL_INFO[prob.part + "|" + prob.level] || {}).strictType);

  // shown data
  const expInst = instantiate(prob, pSeed, dSeed);
  const expected = XL.fill(b.solution, expInst.wb, expInst.target);
  const got = runFormula(parsed, shown);
  const special = firstSpecial(got);
  if (special) return { ok: false, kind: "syntax", msg: special.why, inst: shown };

  const t = shown.target;
  const wrong = [];
  for (let i = 0; i < t.rows; i++) for (let j = 0; j < t.cols; j++) if (!same(expected[i][j], got[i][j], strict)) wrong.push([i, j]);

  if (!wrong.length) {
    // hidden data versions
    for (const k of ALT_SEEDS) {
      const ie = instantiate(prob, pSeed, dSeed + k), iu = instantiate(prob, pSeed, dSeed + k);
      const e2 = XL.fill(b.solution, ie.wb, ie.target);
      const g2 = runFormula(parsed, iu);
      for (let i = 0; i < t.rows; i++) for (let j = 0; j < t.cols; j++) {
        if (!same(e2[i][j], g2[i][j], strict)) {
          return { ok: false, kind: "hardcode", inst: shown, wrong: [],
            msg: "Your formula gives the right answers for this data, but when the tool changed the numbers in the data, it no longer did. Did you type a number (or a cell's value) into the formula instead of referring to the cell? Point at the cells so the formula updates when the data changes." };
        }
      }
    }
    return { ok: true, inst: shown, expected, got };
  }

  const res = { ok: false, inst: shown, wrong, expected, got };
  // known mistakes
  for (const m of b.mistakes || []) {
    try {
      const im = instantiate(prob, pSeed, dSeed);
      const mv = XL.fill(m.f, im.wb, im.target);
      let match = true;
      for (let i = 0; i < t.rows && match; i++) for (let j = 0; j < t.cols; j++) if (!same(mv[i][j], got[i][j], true)) { match = false; break; }
      if (match) return Object.assign(res, { kind: "mistake", msg: m.msg });
    } catch (e) { /* ignore broken mistake formula */ }
  }
  // errors in the student's results
  const err = firstError(got, t);
  if (err && !XL.isErr(expected[err.i][err.j])) {
    let extra = "";
    if (t.rows > 1 && same(expected[0][0], got[0][0], strict) && relativeRanges(parsed.ast)) extra = " Your first cell works, so a range probably slid when the formula was copied down. Lock it with $ signs.";
    return Object.assign(res, { kind: "error", msg: err.msg + extra });
  }
  // type / formatting style problems on the first wrong cell
  const [wi, wj] = wrong[0];
  const e = expected[wi][wj], g = got[wi][wj];
  const where = cellName(t, wi, wj);
  let msg = null;
  if (typeof e === "number" && typeof g === "string" && numLike(g) && Math.abs(parseFloat(g) - e) < 1e-9) msg = where + " shows " + show(g) + ", which is text that looks like a number. This one asks for a real number. Wrap it in VALUE().";
  else if (typeof e === "string" && typeof g === "string" && e.toLowerCase() === g.toLowerCase()) msg = where + " shows " + show(g) + " but it should be " + show(e) + ". Check the capitalization.";
  else if (typeof e === "string" && typeof g === "string" && e.replace(/\s+/g, "") === g.replace(/\s+/g, "")) msg = where + " shows " + show(g) + " but it should be " + show(e) + ". Look closely at the spaces.";
  else if (e === "" && g === 0) msg = where + " shows 0 but should look blank. Use \"\" (two quote marks) for the blank part of your IF.";
  else if (e === "" && g === " ") msg = where + " has a space in it. For a blank cell, use \"\" with nothing between the quotes.";
  else if (typeof e === "number" && typeof g === "number" && e !== 0 && Math.abs(g / e - 100) < 1e-6) msg = where + " is 100 times too big. Percentages in Excel are decimals: 4% is 0.04.";
  else if (typeof e === "number" && typeof g === "number" && e !== 0 && Math.abs(g - e) / Math.abs(e) < 0.02) msg = where + " gives " + show(g) + ", very close to the right answer, " + show(e) + ". Check your rounding or order of operations.";
  else if (e === "" && g === false) msg = where + " shows FALSE. When IF's test is false and there's no third part, Excel shows FALSE. Add \"\" as the value_if_false to leave it blank.";
  else if (typeof e === "boolean" && typeof g === "string") msg = where + " shows " + show(g) + ", but this one wants TRUE/FALSE. Use the function on its own, without wrapping it in IF.";
  // first cell right, others wrong -> copying problem
  const firstOk = same(expected[0][0], got[0][0], strict);
  if (!msg && firstOk && wrong.length && (t.rows > 1 || t.cols > 1)) {
    const w = wrong[0];
    msg = "Your formula is right in " + cellName(t, 0, 0) + ", but not after it's copied: " + cellName(t, w[0], w[1]) + " gives " + show(got[w[0]][w[1]]) + " instead of " + show(expected[w[0]][w[1]]) + ". ";
    msg += t.cols > 1 ? "In a grid, decide which part of each reference should stay put: $A2 locks the column, B$1 locks the row." : "Check your $ signs. A reference that should stay put (like a total, a rate, or a lookup range) needs to be locked, and the row's own cells shouldn't be.";
    return Object.assign(res, { kind: "copy", msg });
  }
  if (!msg) {
    msg = where + ": your formula gives " + show(g) + ", but it should be " + show(e) + ".";
    if (wrong.length > 1) msg += " (" + wrong.length + " of " + t.rows * t.cols + " cells don't match.)";
  }
  return Object.assign(res, { kind: "wrong", msg });
}

const Checker = { instantiate, check, preview, same, ALLOWED };
if (typeof module !== "undefined" && module.exports) module.exports = Checker;
else root.Checker = Checker;
})(typeof self !== "undefined" ? self : this);
