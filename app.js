/* OM 305 Excel Practice — runs entirely in the browser. */
"use strict";

// ------------------------------------------------------------------ settings you can change
const CONFIG = {
  LOCK_PARTS: false,       // true = students must finish (or test out of) a part before the next one opens
  JUMP_AHEAD: true,        // locked levels offer a one-shot "Jump ahead" challenge problem
  TEST_OUT_STREAK: 2,      // first-try correct answers in a row needed to test out of a level
  MIN_LEVEL_SIZE: 3,       // levels smaller than this can't be tested out of
  STRUGGLE_FAILS: 3,       // wrong checks on one problem before the level turns off test-out
  STRUGGLE_HINTS: 2,       // hints on one problem before the level turns off test-out
  NO_TEST_OUT_PARTS: [7],  // exam-style practice: do every question
  DRILL_LENGTH: 8,         // formulas per speed drill
  DRILL_SKIP_PENALTY: 15,  // seconds added for skipping a drill formula
};
const STORE_KEY = "om305-excel-v1";
const APP_VERSION = "2026-11-11a";

const SIGS = {
  SUM: ["number1", "[number2]", "..."], AVERAGE: ["number1", "[number2]", "..."], COUNT: ["value1", "[value2]", "..."], COUNTA: ["value1", "[value2]", "..."],
  MAX: ["number1", "[number2]", "..."], MIN: ["number1", "[number2]", "..."],
  IF: ["logical_test", "[value_if_true]", "[value_if_false]"], AND: ["logical1", "[logical2]", "..."], OR: ["logical1", "[logical2]", "..."], NOT: ["logical"],
  IFERROR: ["value", "value_if_error"],
  COUNTIFS: ["criteria_range1", "criteria1", "[criteria_range2, criteria2]", "..."], COUNTIF: ["range", "criteria"],
  SUMIFS: ["sum_range", "criteria_range1", "criteria1", "[criteria_range2, criteria2]", "..."], SUMIF: ["range", "criteria", "[sum_range]"],
  AVERAGEIFS: ["average_range", "criteria_range1", "criteria1", "[criteria_range2, criteria2]", "..."], AVERAGEIF: ["range", "criteria", "[average_range]"],
  SUMPRODUCT: ["array1", "[array2]", "..."],
  VLOOKUP: ["lookup_value", "table_array", "col_index_num", "[range_lookup]"],
  INDEX: ["array", "row_num", "[column_num]"], MATCH: ["lookup_value", "lookup_array", "[match_type]"],
  TODAY: [], NOW: [], DATE: ["year", "month", "day"], YEAR: ["serial_number"], MONTH: ["serial_number"], DAY: ["serial_number"],
  ROUND: ["number", "num_digits"], ROUNDUP: ["number", "num_digits"], ROUNDDOWN: ["number", "num_digits"], ABS: ["number"], INT: ["number"],
  ISNUMBER: ["value"], ISTEXT: ["value"], ISBLANK: ["value"], ISERROR: ["value"],
  LEN: ["text"], LEFT: ["text", "[num_chars]"], RIGHT: ["text", "[num_chars]"], MID: ["text", "start_num", "num_chars"], TRIM: ["text"],
  SEARCH: ["find_text", "within_text", "[start_num]"], FIND: ["find_text", "within_text", "[start_num]"],
  SUBSTITUTE: ["text", "old_text", "new_text", "[instance_num]"], CONCATENATE: ["text1", "[text2]", "..."],
  UPPER: ["text"], LOWER: ["text"], PROPER: ["text"], VALUE: ["text"], TEXT: ["value", "format_text"],
};
const FUNC_NAMES = Object.keys(SIGS).sort();

// ------------------------------------------------------------------ state
let state = loadState();
let view = { screen: "welcome", pid: null, hintsShown: 0, feedback: null, solved: false };
let PARTS = [];
const BYID = {};

function loadState() {
  try {
    const s = JSON.parse(localStorage.getItem(STORE_KEY));
    if (s && s.probs) { s.drills = s.drills || {}; s.seeds = s.seeds || {}; s.drafts = s.drafts || {}; return s; }
  } catch (e) { /* storage blocked: start fresh */ }
  return { name: "", probs: {}, levels: {}, drills: {}, seeds: {}, drafts: {}, current: null };
}
function saveState() { try { localStorage.setItem(STORE_KEY, JSON.stringify(state)); } catch (e) { /* ignore */ } }
function ps(id) { return state.probs[id] || (state.probs[id] = { status: null, fails: 0, hints: 0, firstTry: false }); }
function ls(key) { return state.levels[key] || (state.levels[key] = { mastered: false, struggle: false, streak: 0 }); }
const levelKey = (p) => p.part + "|" + p.level;
function seedFor(id) {
  if (!state.seeds[id]) { state.seeds[id] = { p: 1 + Math.floor(Math.random() * 1e6), d: 1 + Math.floor(Math.random() * 1e6) }; saveState(); }
  return state.seeds[id];
}

// ------------------------------------------------------------------ structure
function buildParts() {
  for (const p of Bank.BANK) {
    BYID[p.id] = p;
    let part = PARTS.find((x) => x.num === p.part);
    if (!part) { part = { num: p.part, title: p.partTitle, levels: [] }; PARTS.push(part); }
    let lv = part.levels.find((l) => l.key === levelKey(p));
    if (!lv) { lv = { key: levelKey(p), name: p.level, part: p.part, ids: [] }; part.levels.push(lv); }
    lv.ids.push(p.id);
  }
}
const levelOf = (id) => { const p = BYID[id]; return PARTS.find((x) => x.num === p.part).levels.find((l) => l.key === levelKey(p)); };
const partOf = (num) => PARTS.find((x) => x.num === num);
const isDone = (id) => { const s = state.probs[id]; return !!(s && (s.status === "solved" || s.status === "skipped")) || ls(levelOf(id).key).mastered; };
const levelComplete = (lv) => ls(lv.key).mastered || lv.ids.every((id) => isDone(id));
const partComplete = (part) => part.levels.every(levelComplete);
function partUnlocked(part) {
  if (!CONFIG.LOCK_PARTS) return true;
  const i = PARTS.indexOf(part);
  return i === 0 || partComplete(PARTS[i - 1]) || part.levels.some((lv) => lv.ids.some((id) => state.probs[id] && state.probs[id].status));
}
function levelUnlocked(lv) {
  const part = partOf(lv.part);
  if (!partUnlocked(part)) return false;
  const i = part.levels.indexOf(lv);
  return i === 0 || ls(lv.key).unlocked || levelComplete(part.levels[i - 1]) || lv.ids.some((id) => state.probs[id] && (state.probs[id].status || state.probs[id].fails));
}
function canTestOut(lv) {
  const info = Bank.LEVEL_INFO[lv.key] || {};
  return !info.noTestOut && !CONFIG.NO_TEST_OUT_PARTS.includes(lv.part) && lv.ids.length >= CONFIG.MIN_LEVEL_SIZE;
}
const solvedCount = () => Object.values(state.probs).filter((x) => x.status === "solved").length;
function jumpCandidate(lv) {
  const lvs = ls(lv.key);
  if (lvs.jumpWaitUntil != null && solvedCount() <= lvs.jumpWaitUntil) return null;
  for (let k = lv.ids.length - 1; k >= 0; k--) {
    const s = state.probs[lv.ids[k]];
    if (!s || (!s.status && !s.jumpFailed && !s.fails && !s.hints)) return lv.ids[k];
  }
  return null;
}
function nextUndone(fromId) {
  const all = Bank.BANK.map((p) => p.id);
  const start = fromId ? all.indexOf(fromId) + 1 : 0;
  for (let k = 0; k < all.length; k++) {
    const id = all[(start + k) % all.length];
    if (!isDone(id) && levelUnlocked(levelOf(id))) return id;
  }
  return null;
}
function counts(ids) {
  let solved = 0, tested = 0, skipped = 0;
  for (const id of ids) {
    const s = state.probs[id];
    if (s && s.status === "solved") solved++;
    else if (s && s.status === "skipped") skipped++;
    else if (ls(levelOf(id).key).mastered) tested++;
  }
  return { solved, tested, skipped, total: ids.length };
}

// ------------------------------------------------------------------ dom helpers
const $ = (sel, el) => (el || document).querySelector(sel);
function h(tag, attrs, ...kids) {
  const el = document.createElement(tag);
  if (attrs) for (const k of Object.keys(attrs)) {
    const v = attrs[k];
    if (v === null || v === undefined || v === false) continue;
    if (k === "class") el.className = v;
    else if (k === "html") el.innerHTML = v;
    else if (k.startsWith("on")) el.addEventListener(k.slice(2), v);
    else el.setAttribute(k, v === true ? "" : v);
  }
  for (const kid of kids.flat()) if (kid !== null && kid !== undefined && kid !== false) el.append(kid.nodeType ? kid : document.createTextNode(String(kid)));
  return el;
}
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

// ================================================================== workbook widget
// Shows the sheets, the formula bar, and handles click-to-reference, F4, autocomplete, and argument hints.
class Sheetbook {
  constructor(opts) {
    this.inst = opts.inst;                 // { wb, target, b }
    this.onEnter = opts.onEnter;           // (formula, ctrl) => void
    this.initial = opts.initial || "";
    this.onInput = opts.onInput || (() => {});
    this.viewSheet = opts.inst.b.active || opts.inst.target.sheet;
    this.selected = null;
    this.drag = null;
    this.acIndex = 0;
    this.wrong = [];
    this.el = this.build();
    this.renderGrid();
  }
  get t() { return this.inst.target; }
  setInst(inst, wrong) { this.inst = inst; this.wrong = wrong || []; this.renderGrid(); }
  build() {
    const t = this.t;
    const anchor = XL.addr(t.r, t.c);
    const endCell = XL.addr(t.r + t.rows - 1, t.c + t.cols - 1);
    this.input = h("input", { class: "fx-input", type: "text", spellcheck: "false", autocomplete: "off", "aria-label": "Formula for " + anchor, placeholder: "=" });
    this.input.value = this.initial;
    this.sig = h("div", { class: "sig", hidden: true });
    this.ac = h("div", { class: "ac", hidden: true, role: "listbox" });
    this.info = h("div", { class: "cell-info" });
    this.gridWrap = h("div", { class: "grid-wrap" });
    this.tabs = h("div", { class: "tabs", role: "tablist" });
    const fillNote = t.rows * t.cols > 1 ? (t.cols > 1 ? `fills ${anchor}:${endCell}` : `copies down to ${endCell}`) : "";
    const box = h("div", { class: "book" },
      h("div", { class: "fx-bar" },
        h("div", { class: "name-box", title: "The cell your formula goes in" }, anchor + (t.sheet !== this.viewSheet ? "" : "")),
        h("div", { class: "fx-label" }, "fx"),
        h("div", { class: "fx-field" }, this.input, this.sig, this.ac),
      ),
      h("div", { class: "fx-sub" },
        h("span", null, "Formula for ", h("b", null, t.sheet + "!" + anchor), fillNote ? " · " + fillNote : ""),
        h("span", { class: "keys" }, h("kbd", null, "Enter"), " calculate  ", h("kbd", null, "Ctrl"), "+", h("kbd", null, "Enter"), " check  ", h("kbd", null, "F4"), " $ signs  ", h("kbd", null, "Tab"), " finish a function name"),
      ),
      this.gridWrap, this.tabs, this.info,
    );
    this.input.addEventListener("keydown", (e) => this.onKey(e));
    this.input.addEventListener("input", () => { this.refreshAssist(); this.renderHighlights(); this.onInput(this.input.value); });
    this.input.addEventListener("click", () => this.refreshAssist());
    this.input.addEventListener("keyup", (e) => { if (e.key.startsWith("Arrow") || e.key === "Home" || e.key === "End") this.refreshAssist(); });
    this.input.addEventListener("blur", () => setTimeout(() => { this.ac.hidden = true; this.sig.hidden = true; }, 150));
    document.addEventListener("mouseup", () => { this.drag = null; });
    return box;
  }
  focus() { this.input.focus(); const n = this.input.value.length; this.input.setSelectionRange(n, n); this.refreshAssist(); }
  value() { return this.input.value; }

  // ---------------------------------------------------------------- grid
  renderGrid() {
    const wb = this.inst.wb, t = this.t;
    const sh = wb.sheet(this.viewSheet);
    const ext = wb.extent(this.viewSheet);
    let nr = ext.rows + 1, nc = ext.cols + 1;
    if (t.sheet === this.viewSheet) { nr = Math.max(nr, t.r + t.rows + 1); nc = Math.max(nc, t.c + t.cols + 1); }
    nr = Math.max(nr, 8); nc = Math.max(nc, 5);
    const table = h("table", { class: "grid" });
    const thead = h("tr", null, h("th", { class: "corner" }));
    for (let c = 0; c < nc; c++) thead.append(h("th", { class: "colh" }, XL.numToCol(c)));
    table.append(h("thead", null, thead));
    const tbody = h("tbody");
    const tables = Object.values(wb.tables).filter((x) => x.sheet === sh.name);
    for (let r = 0; r < nr; r++) {
      const tr = h("tr", null, h("th", { class: "rowh" }, String(r + 1)));
      for (let c = 0; c < nc; c++) {
        const v = wb.get(this.viewSheet, r, c);
        const fmt = sh.fmt[r + ":" + c] || (r > 0 ? sh.fmt[c] : null);
        const td = h("td", { "data-r": r, "data-c": c });
        const txt = XL.display(v, fmt);
        td.textContent = txt;
        if (typeof v === "number" || typeof v === "boolean" || (v && v.err)) td.classList.add("num");
        if (v && v.err) td.classList.add("err");
        if (typeof v === "boolean") td.classList.add("ctr");
        const inTable = tables.find((x) => r >= x.top && r <= x.top + x.nrows && c >= x.left && c < x.left + x.ncols);
        if (inTable) { td.classList.add(r === inTable.top ? "tbl-head" : (r - inTable.top) % 2 ? "tbl-row" : "tbl-row2"); }
        else if (r === 0 && typeof v === "string") td.classList.add("hdr");
        if (t.sheet === this.viewSheet && r >= t.r && r < t.r + t.rows && c >= t.c && c < t.c + t.cols) {
          td.classList.add("target");
          if (r === t.r && c === t.c) td.classList.add("anchor");
          if (this.wrong.some(([i, j]) => t.r + i === r && t.c + j === c)) td.classList.add("wrong");
        }
        td.addEventListener("mousedown", (e) => this.cellDown(e, r, c));
        td.addEventListener("mouseenter", () => this.cellOver(r, c));
        tr.append(td);
      }
      tbody.append(tr);
    }
    table.append(tbody);
    this.gridWrap.replaceChildren(table);
    this.tableEl = table;
    // sheet tabs
    this.tabs.replaceChildren(...wb.order.map((name) => {
      const b = h("button", { class: "tab" + (name === this.viewSheet ? " active" : ""), role: "tab", type: "button" }, name);
      b.addEventListener("mousedown", (e) => { e.preventDefault(); this.viewSheet = name; this.renderGrid(); });
      return b;
    }));
    const tbl = Object.values(wb.tables).filter((x) => x.sheet === sh.name).map((x) => x.name);
    if (tbl.length) this.tabs.append(h("span", { class: "tbl-note" }, "Excel Table: " + tbl.join(", ")));
    this.renderHighlights();
  }
  cell(r, c) { return this.tableEl && this.tableEl.querySelector(`td[data-r="${r}"][data-c="${c}"]`); }

  // where can a reference be inserted? (right after = ( , operator, or replacing a reference just inserted)
  insertPoint() {
    const v = this.input.value, pos = this.input.selectionStart, end = this.input.selectionEnd;
    if (document.activeElement !== this.input) return null;
    if (!v.startsWith("=")) return null;
    if (end > pos) return { start: pos, end };
    const before = v.slice(0, pos);
    if (/[=(,+\-*/^&<>:]\s*$/.test(before) || before === "=") return { start: pos, end: pos };
    return null;
  }
  cellDown(e, r, c) {
    // Shift+click extends the reference just inserted into a range
    if (e.shiftKey && this.lastInsert && this.lastAnchor && document.activeElement === this.input && this.input.selectionStart === this.lastInsert.end && this.lastAnchor.sheet === this.viewSheet) {
      e.preventDefault();
      this.putRef(this.lastAnchor.start, this.lastInsert.end, this.lastAnchor.r, this.lastAnchor.c, r, c);
      return;
    }
    const ip = this.insertPoint();
    if (ip || (this.lastInsert && document.activeElement === this.input && this.input.selectionStart === this.lastInsert.end)) {
      e.preventDefault();
      const at = ip || { start: this.lastInsert.start, end: this.lastInsert.end };
      this.drag = { r0: r, c0: c, start: at.start };
      this.lastAnchor = { r, c, start: at.start, sheet: this.viewSheet };
      this.putRef(at.start, at.end, r, c, r, c);
      return;
    }
    this.selected = [r, c];
    this.showInfo(r, c);
    this.renderHighlights();
  }
  cellOver(r, c) {
    if (!this.drag) return;
    this.putRef(this.drag.start, this.lastInsert.end, this.drag.r0, this.drag.c0, r, c);
  }
  putRef(start, end, r1, c1, r2, c2) {
    let ref = XL.addr(Math.min(r1, r2), Math.min(c1, c2));
    if (r1 !== r2 || c1 !== c2) ref += ":" + XL.addr(Math.max(r1, r2), Math.max(c1, c2));
    if (this.viewSheet !== this.t.sheet) ref = XL.sheetPrefix(this.viewSheet) + ref;
    const v = this.input.value;
    this.input.value = v.slice(0, start) + ref + v.slice(end);
    const pos = start + ref.length;
    this.input.setSelectionRange(pos, pos);
    this.lastInsert = { start, end: pos };
    this.renderHighlights();
    this.onInput(this.input.value);
  }
  showInfo(r, c) {
    const v = this.inst.wb.get(this.viewSheet, r, c);
    const sh = this.inst.wb.sheet(this.viewSheet);
    const fmt = sh.fmt[r + ":" + c] || (r > 0 ? sh.fmt[c] : null);
    let txt = this.viewSheet + "!" + XL.addr(r, c) + ": ";
    if (v === null) txt += "(empty)";
    else if (typeof v === "string") txt += "\"" + v + "\" (text)";
    else if (typeof v === "number") txt += XL.display(v, fmt) + (fmt === "date" ? "  (stored as the number " + v + ")" : fmt && fmt !== "int" ? "  (" + XL.fmtGeneral(v) + ")" : "");
    else txt += XL.display(v, fmt);
    this.info.textContent = txt;
  }

  // colored outlines for the references in the formula (like Excel)
  renderHighlights() {
    if (!this.tableEl) return;
    this.tableEl.querySelectorAll("td.ref").forEach((td) => { td.classList.remove("ref", "rt", "rb", "rl", "rr"); td.style.removeProperty("--rc"); });
    this.tableEl.querySelectorAll("td.sel").forEach((td) => td.classList.remove("sel"));
    if (this.selected) { const td = this.cell(...this.selected); if (td) td.classList.add("sel"); }
    let parsed;
    try { parsed = XL.parse(this.input.value || "="); } catch (e) { return; }
    const colors = ["#2f6fdb", "#c0392b", "#7d3cbd", "#1f8a5b", "#c76b00", "#0f8b8d"];
    let k = 0;
    const wb = this.inst.wb, t = this.t;
    const visit = (n) => {
      if (!n || typeof n !== "object") return;
      let rg = null;
      if (n.k === "ref") rg = { sheet: n.sheet || t.sheet, r1: n.r, c1: n.c, r2: n.r, c2: n.c };
      else if (n.k === "range") rg = { sheet: n.a.sheet || t.sheet, r1: n.a.r, c1: n.a.c, r2: n.b.r, c2: n.b.c };
      else if (n.k === "cols") { const ext = wb.extent(n.sheet || t.sheet); rg = { sheet: n.sheet || t.sheet, r1: 0, c1: n.c1, r2: ext.rows - 1, c2: n.c2 }; }
      else if (n.k === "sref" || n.k === "name") {
        const tb = n.k === "name" ? wb.tables[n.v.toLowerCase()] : n.table ? wb.tables[n.table.toLowerCase()] : wb.tableAt(t.sheet, t.r, t.c);
        if (tb) {
          let c1 = tb.left, c2 = tb.left + tb.ncols - 1, r1 = tb.top + 1, r2 = tb.top + tb.nrows;
          if (n.k === "sref") {
            const m = /^@?\[?([^\]]*)\]?$/.exec(n.spec.trim());
            const name = m ? m[1].replace(/^@/, "").replace(/^\[|\]$/g, "") : "";
            const i = tb.cols.findIndex((x) => x.toLowerCase() === name.toLowerCase());
            if (i >= 0) { c1 = c2 = tb.left + i; }
            if (n.spec.trim().startsWith("@")) r1 = r2 = t.r;
          }
          rg = { sheet: tb.sheet, r1, c1, r2, c2 };
        }
      }
      if (rg && rg.sheet.toLowerCase() === this.viewSheet.toLowerCase()) {
        const color = colors[k++ % colors.length];
        const r1 = Math.min(rg.r1, rg.r2), r2 = Math.max(rg.r1, rg.r2), c1 = Math.min(rg.c1, rg.c2), c2 = Math.max(rg.c1, rg.c2);
        for (let r = r1; r <= Math.min(r2, 200); r++) for (let c = c1; c <= c2; c++) {
          const td = this.cell(r, c);
          if (!td) continue;
          td.classList.add("ref");
          td.style.setProperty("--rc", color);
          if (r === r1) td.classList.add("rt");
          if (r === r2) td.classList.add("rb");
          if (c === c1) td.classList.add("rl");
          if (c === c2) td.classList.add("rr");
        }
      } else if (rg) k++;
      for (const key of ["e", "a", "b"]) if (n[key] && n.k !== "range") visit(n[key]);
      if (n.args) n.args.forEach(visit);
    };
    visit(parsed.ast);
  }

  // ---------------------------------------------------------------- typing help
  onKey(e) {
    if (!this.ac.hidden) {
      const items = [...this.ac.children];
      if (e.key === "ArrowDown") { e.preventDefault(); this.acIndex = (this.acIndex + 1) % items.length; this.paintAc(); return; }
      if (e.key === "ArrowUp") { e.preventDefault(); this.acIndex = (this.acIndex - 1 + items.length) % items.length; this.paintAc(); return; }
      if (e.key === "Tab") { e.preventDefault(); this.acceptAc(items[this.acIndex].dataset.name); return; }
      if (e.key === "Escape") { this.ac.hidden = true; return; }
    }
    if (e.key === "F4") {
      e.preventDefault();
      const r = XL.cycleRef(this.input.value, this.input.selectionStart);
      this.input.value = r.text;
      this.input.setSelectionRange(r.cursor, r.cursor);
      this.lastInsert = null;
      this.renderHighlights();
      this.onInput(this.input.value);
      return;
    }
    if (e.key === "Enter") {
      e.preventDefault();
      this.ac.hidden = true;
      this.onEnter(this.input.value, e.ctrlKey || e.metaKey || e.shiftKey);
      return;
    }
    if (!["Shift", "Control", "Alt", "Meta"].includes(e.key)) this.lastInsert = null;
  }
  currentWord() {
    const v = this.input.value, pos = this.input.selectionStart;
    const m = /([A-Za-z][A-Za-z0-9.]*)$/.exec(v.slice(0, pos));
    if (!m) return null;
    const start = pos - m[1].length;
    const before = v.slice(0, start);
    if (!(/[=(,+\-*/^&<>]\s*$/.test(before))) return null;
    if (/"/.test(before) && (before.split('"').length - 1) % 2 === 1) return null;
    return { word: m[1], start, end: pos };
  }
  refreshAssist() {
    // autocomplete
    const cw = this.currentWord();
    const matches = cw ? FUNC_NAMES.filter((f) => f.startsWith(cw.word.toUpperCase()) && Checker.ALLOWED.has(f)).sort((a, b) => a.length - b.length || a.localeCompare(b)) : [];
    if (cw && matches.length && !(matches.length === 1 && matches[0] === cw.word.toUpperCase() && this.input.value[cw.end] === "(")) {
      this.acWord = cw;
      this.acIndex = Math.min(this.acIndex, matches.length - 1);
      this.ac.replaceChildren(...matches.slice(0, 7).map((f) => {
        const it = h("div", { class: "ac-item", "data-name": f, role: "option" }, h("b", null, f), h("span", null, "(" + SIGS[f].join(", ") + ")"));
        it.addEventListener("mousedown", (e) => { e.preventDefault(); this.acceptAc(f); });
        return it;
      }));
      this.ac.hidden = false;
      this.paintAc();
    } else { this.ac.hidden = true; this.acIndex = 0; }
    // signature of the function the cursor is inside
    const v = this.input.value.slice(0, this.input.selectionStart);
    let depth = 0, inStr = false;
    const stack = [];
    for (let i = 0; i < v.length; i++) {
      const ch = v[i];
      if (ch === '"') { inStr = !inStr; continue; }
      if (inStr) continue;
      if (ch === "(") {
        const m = /([A-Za-z][A-Za-z0-9.]*)\s*$/.exec(v.slice(0, i));
        stack.push({ name: m ? m[1].toUpperCase() : null, arg: 0 });
        depth++;
      } else if (ch === ")") { stack.pop(); depth--; }
      else if (ch === "," && stack.length) stack[stack.length - 1].arg++;
    }
    const top = stack.filter((s) => s.name && SIGS[s.name]).pop();
    if (top && this.ac.hidden) {
      const args = SIGS[top.name];
      let idx = top.arg;
      if (idx >= args.length) { const rep = args.indexOf("..."); idx = rep > 0 ? Math.max(rep - 1 - ((idx - rep) % 2 === 0 ? 0 : 0), rep - 1) : args.length - 1; }
      this.sig.innerHTML = "<b>" + top.name + "</b>(" + args.map((a, i) => (i === idx ? "<u>" + esc(a) + "</u>" : esc(a))).join(", ") + ")";
      this.sig.hidden = false;
    } else this.sig.hidden = true;
  }
  paintAc() { [...this.ac.children].forEach((c, i) => c.classList.toggle("on", i === this.acIndex)); }
  acceptAc(name) {
    const cw = this.acWord || this.currentWord();
    if (!cw) return;
    const v = this.input.value;
    const after = v.slice(cw.end);
    const ins = name + (after.startsWith("(") ? "" : "(");
    this.input.value = v.slice(0, cw.start) + ins + after;
    const pos = cw.start + ins.length + (after.startsWith("(") ? 1 : 0);
    this.input.setSelectionRange(pos, pos);
    this.ac.hidden = true;
    this.refreshAssist();
    this.onInput(this.input.value);
  }
}

// ================================================================== sidebar
function renderSidebar() {
  const sb = $("#sidebar");
  const kids = [];
  kids.push(h("button", { class: "drill-link" + (view.screen === "drills" || view.screen === "drill" ? " current" : ""), onclick: () => { closeMenu(); showDrills(); } }, h("span", { class: "bolt" }, "⚡"), "Speed Drills"));
  for (const part of PARTS) {
    const done = partComplete(part), unlocked = partUnlocked(part);
    const c = counts(part.levels.flatMap((l) => l.ids));
    const cs = state.collapsed && state.collapsed[part.num];
    const collapsed = cs !== undefined ? cs : view.pid && BYID[view.pid] ? BYID[view.pid].part !== part.num : part.num !== 1;
    const levels = part.levels.map((lv) => {
      const lvs = ls(lv.key), un = levelUnlocked(lv), comp = levelComplete(lv);
      const lc = counts(lv.ids);
      let badge = lc.solved + lc.tested + lc.skipped + "/" + lc.total, bcls = "";
      if (lvs.mastered) { badge = "tested out"; bcls = "mastered"; }
      else if (comp) { badge = "done ✓"; bcls = "done"; }
      else if (!un && CONFIG.JUMP_AHEAD && jumpCandidate(lv)) { badge = "🔒 jump"; bcls = "jump"; }
      else if (!un) badge = "🔒";
      const dots = h("div", { class: "dots" }, lv.ids.map((id) => {
        const s = state.probs[id];
        const cls = s && s.status === "solved" ? "ok" : s && s.status === "skipped" ? "skip" : lvs.mastered ? "tested" : id === view.pid ? "cur" : "";
        return h("i", { class: cls, title: BYID[id].title });
      }));
      return h("button", {
        class: "level" + (view.pid && levelOf(view.pid) === lv ? " current" : ""), disabled: !un && !(CONFIG.JUMP_AHEAD && jumpCandidate(lv)) ? true : null,
        title: un ? null : "Locked. Finish the level above, or click to try a one-shot jump-ahead challenge.",
        onclick: () => { closeMenu(); if (un) gotoLevel(lv); else startJump(lv); },
      }, h("div", { class: "lname" }, h("span", null, lv.name), h("span", { class: "badge " + bcls }, badge)), dots);
    });
    kids.push(h("div", { class: "part" + (done ? " done" : "") + (unlocked ? "" : " locked") + (collapsed ? " collapsed" : "") },
      h("button", { class: "part-head", onclick: () => { state.collapsed = state.collapsed || {}; state.collapsed[part.num] = !collapsed; saveState(); renderSidebar(); } },
        h("span", { class: "num" }, done ? "✓" : part.num), h("span", null, part.title), h("span", { class: "meta" }, (c.solved + c.tested + c.skipped) + "/" + c.total)),
      h("div", { class: "levels" }, levels)));
  }
  sb.replaceChildren(...kids);
  // top progress
  const all = counts(Bank.BANK.map((p) => p.id));
  const pct = Math.round(100 * (all.solved + all.tested + all.skipped) / all.total);
  $("#topProgress").replaceChildren(h("div", { class: "bar" }, h("span", { style: "width:" + pct + "%" })), h("span", null, pct + "% · " + all.solved + " solved"));
}
function closeMenu() { document.body.classList.remove("menu-open"); }
function gotoLevel(lv) {
  const id = lv.ids.find((x) => !isDone(x)) || lv.ids[0];
  goto(id);
}

// ================================================================== problem screen
let book = null;
function goto(id, opts) {
  view = { screen: "problem", pid: id, hintsShown: 0, feedback: null, solved: false, jump: opts && opts.jump ? opts.jump : null };
  state.current = id;
  saveState();
  render();
  window.scrollTo(0, 0);
}
function startJump(lv) {
  const pid = jumpCandidate(lv);
  if (!pid) { alert("You've used your jump-ahead try for this level. Solve another problem anywhere, then you can try a different one."); return; }
  goto(pid, { jump: { key: lv.key } });
}

function renderProblem() {
  const p = BYID[view.pid];
  const lv = levelOf(p.id), lvs = ls(lv.key), s = ps(p.id);
  const seed = seedFor(p.id);
  const inst = Checker.instantiate(p, seed.p, seed.d);
  const b = inst.b;
  const info = Bank.LEVEL_INFO[lv.key] || {};
  const idx = lv.ids.indexOf(p.id);
  const main = $("#main");
  const out = [];

  out.push(h("div", { class: "crumbs" }, "Part " + p.part + " · " + p.partTitle + "  ›  " + lv.name + "  ›  " + (idx + 1) + " of " + lv.ids.length));
  if (view.jump) out.push(h("div", { class: "banner jump" }, h("b", null, "Jump-ahead challenge. "), "Solve this one with no hints and one check. Get it right and this level opens, and the easier levels before it count as mastered."));
  if (lvs.mastered && s.status !== "solved") out.push(h("div", { class: "banner info" }, "You tested out of this level, so this problem is optional. It's still good practice!"));
  if (lvs.struggle && lvs.mastered === false && canTestOut(lv) && !view.jump && idx === 0) { /* quiet */ }

  const intro = h("details", { class: "intro", open: !lv.ids.some((id) => state.probs[id] && state.probs[id].status) ? true : null },
    h("summary", null, "Level notes: " + lv.name), h("div", { class: "intro-body", html: info.intro || "" }));
  out.push(intro);

  out.push(h("h1", { class: "ptitle" }, p.title, s.status === "solved" ? h("span", { class: "chip ok" }, "solved") : null));
  out.push(h("div", { class: "prompt", html: b.prompt }));

  const draft = state.drafts[p.id] || "";
  book = new Sheetbook({ inst, initial: draft, onEnter: (f, ctrl) => (ctrl ? doCheck() : doRun()), onInput: (v) => { state.drafts[p.id] = v; saveStateSoon(); } });
  out.push(book.el);

  const fb = h("div", { class: "feedback", id: "feedback", "aria-live": "polite" });
  out.push(fb);

  const actions = h("div", { class: "actions" },
    h("button", { class: "btn ghost", onclick: doRun, title: "Enter" }, "Calculate"),
    h("button", { class: "btn", onclick: doCheck, title: "Ctrl+Enter" }, view.jump ? "Check (one try)" : "Check answer"),
    h("span", { class: "spacer" }),
    !view.jump && b.hints && view.hintsShown < b.hints.length ? h("button", { class: "btn ghost", onclick: showHint }, "Hint (" + (b.hints.length - view.hintsShown) + " left)") : null,
    !view.jump && s.status !== "solved" ? h("button", { class: "btn ghost subtle", onclick: skip }, "Skip") : null,
    h("button", { class: "btn ghost subtle", onclick: newVersion, title: "Same problem, new names and numbers" }, "New data"),
  );
  out.push(actions);
  const hints = h("div", { class: "hints", id: "hints" });
  (b.hints || []).slice(0, view.hintsShown).forEach((t, i) => hints.append(h("div", { class: "hint" }, h("b", null, "Hint " + (i + 1) + ": "), h("span", { html: esc(t) }))));
  out.push(hints);
  out.push(h("div", { id: "after" }));
  main.replaceChildren(h("div", { class: "problem" }, ...out));
  if (view.feedback) paintFeedback(view.feedback);
  if (s.status === "solved") showAfter(true);
  setTimeout(() => book.focus(), 30);
}
let saveTimer = null;
function saveStateSoon() { clearTimeout(saveTimer); saveTimer = setTimeout(saveState, 400); }

function paintFeedback(fbk) {
  const el = $("#feedback");
  if (!el) return;
  el.className = "feedback " + fbk.cls;
  el.replaceChildren(...[h("div", { class: "fb-title" }, fbk.title), fbk.msg ? h("div", { class: "fb-msg" }, fbk.msg) : null].filter(Boolean));
}
function doRun() {
  const p = BYID[view.pid], seed = seedFor(p.id);
  const f = book.value().trim();
  if (!f) return;
  const r = Checker.preview(p, seed.p, seed.d, f);
  book.setInst(r.inst);
  if (r.error) paintFeedback({ cls: r.syntax ? "bad" : "warn", title: r.syntax ? "Excel can't use that formula yet" : "Calculated, with an error", msg: r.error });
  else paintFeedback({ cls: "neutral", title: "Calculated", msg: "Look over the highlighted cells. When you think they're right, press Check answer (Ctrl+Enter)." });
}
function doCheck() {
  const p = BYID[view.pid], seed = seedFor(p.id), s = ps(p.id), lv = levelOf(p.id), lvs = ls(lv.key);
  const f = book.value().trim();
  if (!f) { paintFeedback({ cls: "warn", title: "Type a formula first", msg: "Formulas start with =" }); return; }
  const r = Checker.check(p, seed.p, seed.d, f);
  book.setInst(r.inst, r.ok ? [] : r.wrong);
  const wasSolved = s.status === "solved";
  if (r.ok) {
    const firstTry = !s.fails && !s.hints && !wasSolved;
    if (!wasSolved) { s.status = "solved"; s.firstTry = firstTry; s.solvedAt = Date.now(); }
    let extra = "";
    if (view.jump) {
      lvs.unlocked = true;
      const part = partOf(lv.part);
      part.levels.slice(0, part.levels.indexOf(lv)).forEach((l) => { if (!levelComplete(l)) ls(l.key).mastered = true; });
      extra = "Jump-ahead passed! This level is open, and the levels before it count as mastered.";
      view.jump = null;
    } else if (!wasSolved) {
      if (firstTry) lvs.streak = (lvs.streak || 0) + 1; else lvs.streak = 0;
      if (canTestOut(lv) && !lvs.struggle && !lvs.mastered && lvs.streak >= CONFIG.TEST_OUT_STREAK && lv.ids.indexOf(p.id) < lv.ids.length - 1) {
        lvs.mastered = true;
        extra = "You got " + CONFIG.TEST_OUT_STREAK + " in a row on the first try, so you tested out of this level! The rest of it is optional.";
      }
    }
    saveState();
    view.feedback = { cls: "good", title: pickCheer(), msg: extra };
    paintFeedback(view.feedback);
    renderSidebar();
    showAfter(false);
  } else {
    if (view.jump) {
      s.jumpFailed = true;
      lvs.jumpWaitUntil = solvedCount();
      saveState();
      view.feedback = { cls: "bad", title: "Not quite. This level stays locked for now.", msg: r.msg + " Work through the earlier levels, or solve another problem anywhere to get a new jump-ahead try." };
      paintFeedback(view.feedback);
      view.jump = null;
      renderSidebar();
      return;
    }
    if (!wasSolved) {
      s.fails = (s.fails || 0) + 1;
      lvs.streak = 0;
      if (s.fails >= CONFIG.STRUGGLE_FAILS && !lvs.struggle) { lvs.struggle = true; }
    }
    saveState();
    const titles = { syntax: "Excel can't use that formula yet", rule: "Almost. One rule for this problem", hardcode: "Works here, but it's hard-coded", copy: "Check the copy-down", mistake: "Close!", error: "There's an error in the results", wrong: "Not yet" };
    view.feedback = { cls: r.kind === "syntax" || r.kind === "rule" ? "warn" : "bad", title: titles[r.kind] || "Not yet", msg: r.msg };
    paintFeedback(view.feedback);
    renderSidebar();
  }
}
const CHEERS = ["Correct!", "Nailed it!", "That's right!", "Excel-lent!", "Yes! Correct.", "Looks great!"];
function pickCheer() { return CHEERS[Math.floor(Math.random() * CHEERS.length)]; }
function showHint() {
  const p = BYID[view.pid], s = ps(p.id), lv = levelOf(p.id), lvs = ls(lv.key);
  view.hintsShown++;
  if (s.status !== "solved") {
    s.hints = Math.max(s.hints || 0, view.hintsShown);
    if (s.hints >= CONFIG.STRUGGLE_HINTS) lvs.struggle = true;
    lvs.streak = 0;
  }
  saveState();
  const fbSave = view.feedback;
  const cur = book.value();
  renderProblem();
  book.input.value = cur;
  view.feedback = fbSave;
  if (fbSave) paintFeedback(fbSave);
}
function skip() {
  const p = BYID[view.pid];
  ps(p.id).status = "skipped";
  ls(levelOf(p.id).key).streak = 0;
  saveState();
  const n = nextUndone(p.id);
  if (n) goto(n); else render();
}
function newVersion() {
  const p = BYID[view.pid];
  state.seeds[p.id] = { p: 1 + Math.floor(Math.random() * 1e6), d: 1 + Math.floor(Math.random() * 1e6) };
  delete state.drafts[p.id];
  saveState();
  view.feedback = null;
  view.hintsShown = 0;
  renderProblem();
}
function showAfter(quiet) {
  const p = BYID[view.pid], seed = seedFor(p.id);
  const inst = Checker.instantiate(p, seed.p, seed.d);
  const nxt = nextUndone(p.id);
  const el = $("#after");
  if (!el) return;
  el.replaceChildren(h("div", { class: "solution" },
    h("div", { class: "sol-label" }, "One way to write it:"),
    h("code", { class: "sol-code" }, inst.b.solution),
    h("div", { class: "sol-note" }, "Other formulas that give the same results count, too."),
  ), h("div", { class: "actions" },
    h("button", { class: "btn ghost", onclick: newVersion }, "Practice again with new data"),
    h("span", { class: "spacer" }),
    nxt ? h("button", { class: "btn", id: "nextBtn", onclick: () => goto(nxt) }, "Next problem →") : h("button", { class: "btn", onclick: () => { view = { screen: "welcome" }; render(); } }, "All done! 🎉"),
  ));
  if (!quiet) { const nb = $("#nextBtn"); if (nb) setTimeout(() => nb.focus(), 50); }
}

// ================================================================== speed drills
let drill = null;
function showDrills() { view = { screen: "drills" }; render(); }
function renderDrills() {
  const cards = Bank.DRILLS.map((d) => {
    const st = state.drills[d.key] || {};
    return h("button", { class: "drill-card", onclick: () => startDrill(d.key) },
      h("div", { class: "dc-name" }, d.name),
      h("div", { class: "dc-desc" }, d.desc),
      h("div", { class: "dc-best" }, st.best ? "Best: " + fmtTime(st.best) + " · " + st.runs + " run" + (st.runs === 1 ? "" : "s") : "No runs yet"));
  });
  $("#main").replaceChildren(h("div", { class: "drills" },
    h("h1", null, "⚡ Speed Drills"),
    h("p", { class: "lead" }, `${CONFIG.DRILL_LENGTH} formulas against the clock. Each one has fresh data. Press Enter to check, and a correct answer jumps straight to the next one. There are no hints, but you can skip (+${CONFIG.DRILL_SKIP_PENALTY} seconds). Beat your best time!`),
    h("div", { class: "drill-grid" }, cards),
    h("p", { class: "muted small" }, "Speed tips: click cells instead of typing addresses, press F4 for $ signs, and press Tab to finish function names."),
  ));
}
function startDrill(key) {
  let pool = Bank.BANK.filter((p) => p.drill && (key === "mixed" || p.drill === key));
  const picks = [];
  const shuffled = pool.slice().sort(() => Math.random() - 0.5);
  for (let i = 0; i < CONFIG.DRILL_LENGTH; i++) picks.push(shuffled[i % shuffled.length]);
  drill = { key, items: picks.map((p) => ({ pid: p.id, p: 1 + Math.floor(Math.random() * 1e6), d: 1 + Math.floor(Math.random() * 1e6), time: 0, skipped: false, misses: 0 })), i: 0, start: performance.now(), penalty: 0, itemStart: performance.now(), timer: null };
  view = { screen: "drill" };
  render();
}
function fmtTime(sec) { const m = Math.floor(sec / 60), s = sec - m * 60; return (m ? m + ":" + String(Math.floor(s)).padStart(2, "0") : Math.floor(s)) + "." + Math.floor((s % 1) * 10) + (m ? "" : "s"); }
function renderDrillItem() {
  const it = drill.items[drill.i], p = BYID[it.pid];
  const inst = Checker.instantiate(p, it.p, it.d);
  const timerEl = h("div", { class: "timer", id: "timer" }, "0.0s");
  const head = h("div", { class: "drill-head" },
    h("div", { class: "drill-prog" }, Bank.DRILLS.find((d) => d.key === drill.key).name + " · " + (drill.i + 1) + " / " + drill.items.length),
    h("div", { class: "drill-dots" }, drill.items.map((x, k) => h("i", { class: k < drill.i ? (x.skipped ? "skip" : "ok") : k === drill.i ? "cur" : "" }))),
    timerEl,
    h("button", { class: "btn ghost small", onclick: () => { clearInterval(drill.timer); drill = null; showDrills(); } }, "Quit"));
  book = new Sheetbook({ inst, onEnter: () => drillCheck() });
  const fb = h("div", { class: "feedback", id: "feedback" });
  $("#main").replaceChildren(h("div", { class: "problem drill-run" }, head,
    h("div", { class: "drill-topic" }, p.level),
    h("div", { class: "prompt", html: inst.b.prompt }), book.el, fb,
    h("div", { class: "actions" }, h("button", { class: "btn", onclick: drillCheck }, "Check (Enter)"), h("span", { class: "spacer" }), h("button", { class: "btn ghost subtle", onclick: drillSkip }, "Skip (+" + CONFIG.DRILL_SKIP_PENALTY + "s)"))));
  clearInterval(drill.timer);
  drill.timer = setInterval(() => { const t = $("#timer"); if (t && drill) t.textContent = fmtTime((performance.now() - drill.start) / 1000 + drill.penalty); }, 100);
  drill.itemStart = performance.now();
  setTimeout(() => book.focus(), 30);
}
function drillCheck() {
  const it = drill.items[drill.i], p = BYID[it.pid];
  const r = Checker.check(p, it.p, it.d, book.value());
  book.setInst(r.inst, r.ok ? [] : r.wrong);
  if (r.ok) {
    it.time = (performance.now() - drill.itemStart) / 1000;
    paintFeedback({ cls: "good", title: "✓" });
    setTimeout(drillNext, 350);
  } else {
    it.misses++;
    paintFeedback({ cls: r.kind === "syntax" || r.kind === "rule" ? "warn" : "bad", title: "Not yet", msg: r.msg });
  }
}
function drillSkip() {
  const it = drill.items[drill.i];
  it.skipped = true; it.time = (performance.now() - drill.itemStart) / 1000;
  drill.penalty += CONFIG.DRILL_SKIP_PENALTY;
  drillNext();
}
function drillNext() {
  if (!drill) return;
  drill.i++;
  if (drill.i < drill.items.length) { renderDrillItem(); return; }
  clearInterval(drill.timer);
  const total = (performance.now() - drill.start) / 1000 + drill.penalty;
  const st = state.drills[drill.key] || (state.drills[drill.key] = { best: null, runs: 0, history: [] });
  const pb = !st.best || total < st.best;
  const prev = st.best;
  st.runs++; st.history = (st.history || []).concat([{ t: Math.round(total * 10) / 10, at: Date.now(), skips: drill.items.filter((x) => x.skipped).length }]).slice(-10);
  if (pb) st.best = Math.round(total * 10) / 10;
  saveState();
  const d = Bank.DRILLS.find((x) => x.key === drill.key);
  const rows = drill.items.map((x, k) => h("tr", null, h("td", null, k + 1), h("td", null, BYID[x.pid].level), h("td", { class: "num" }, x.skipped ? "skipped" : fmtTime(x.time)), h("td", { class: "num" }, x.misses || "")));
  $("#main").replaceChildren(h("div", { class: "drills" },
    h("h1", null, pb ? "🏆 New personal best!" : "Drill complete"),
    h("div", { class: "big-time" }, fmtTime(total)),
    h("p", { class: "lead" }, d.name + (drill.penalty ? " · includes +" + drill.penalty + "s of skip penalties" : "") + (prev && !pb ? " · your best is " + fmtTime(prev) : "") + (prev && pb ? " · beat your old best of " + fmtTime(prev) : "")),
    h("table", { class: "drill-table" }, h("thead", null, h("tr", null, h("th", null, "#"), h("th", null, "Topic"), h("th", null, "Time"), h("th", null, "Misses"))), h("tbody", null, rows)),
    h("div", { class: "actions" }, h("button", { class: "btn", onclick: () => startDrill(drill.key) }, "Run it again"), h("button", { class: "btn ghost", onclick: showDrills }, "Other drills")),
  ));
  renderSidebar();
}

// ================================================================== welcome + report
function renderWelcome() {
  const resume = state.current && BYID[state.current] ? state.current : null;
  $("#main").replaceChildren(h("div", { class: "welcome" },
    h("div", { class: "hero" },
      h("div", { class: "hero-mark" }, "fx"),
      h("div", null, h("h1", null, "OM 305 Excel Practice"), h("p", { class: "lead" }, "Practice the functions from class in a real spreadsheet grid, and get faster at them. Nothing to install, and nothing is sent anywhere."))),
    h("div", { class: "actions" },
      h("button", { class: "btn big", onclick: () => goto(resume || nextUndone(null) || Bank.BANK[0].id) }, resume ? "Pick up where you left off →" : "Start practicing →"),
      h("button", { class: "btn ghost big", onclick: showDrills }, "⚡ Speed Drills")),
    h("div", { class: "how" },
      h("div", null, h("h3", null, "How it works"), h("ul", null,
        h("li", null, "Type a formula for the highlighted cell. It's copied down for you, like dragging the fill handle."),
        h("li", null, h("b", null, "Enter"), " calculates. ", h("b", null, "Check answer"), " (Ctrl+Enter) grades it. The checker also tries your formula on different data, so hard-coded numbers won't pass."),
        h("li", null, "Click cells (or drag across a range) while typing to insert references. ", h("b", null, "F4"), " adds $ signs, and ", h("b", null, "Tab"), " finishes a function name."),
        h("li", null, "Get the first two problems in a level right on the first try with no hints, and you ", h("b", null, "test out"), " of the rest."))),
      h("div", null, h("h3", null, "What's covered"), h("ul", null, PARTS.map((p) => h("li", null, h("b", null, "Part " + p.num + ": "), p.title + " (" + p.levels.reduce((a, l) => a + l.ids.length, 0) + " problems)")))),
    ),
    h("p", { class: "muted small" }, "Just for practice: not graded and won't count toward or against your grade. Your progress is saved in this browser only."),
  ));
}
function reportText() {
  const lines = [];
  lines.push("OM 305 Excel Practice: progress report");
  lines.push("Name: " + (state.name || "(not entered)"));
  lines.push("Date: " + new Date().toLocaleString());
  const all = counts(Bank.BANK.map((p) => p.id));
  lines.push("Overall: " + all.solved + " solved, " + all.tested + " tested out, " + all.skipped + " skipped, of " + all.total);
  for (const part of PARTS) {
    const c = counts(part.levels.flatMap((l) => l.ids));
    lines.push("");
    lines.push("Part " + part.num + ": " + part.title + " (" + (c.solved + c.tested) + "/" + c.total + ")");
    for (const lv of part.levels) {
      const lc = counts(lv.ids), lvs = ls(lv.key);
      lines.push("  " + lv.name + ": " + (lvs.mastered ? "tested out" : lc.solved + "/" + lc.total + " solved") + (lc.skipped ? ", " + lc.skipped + " skipped" : ""));
    }
  }
  const dr = Bank.DRILLS.filter((d) => state.drills[d.key] && state.drills[d.key].best);
  if (dr.length) {
    lines.push("");
    lines.push("Speed drill personal bests:");
    for (const d of dr) lines.push("  " + d.name + ": " + fmtTime(state.drills[d.key].best) + " (" + state.drills[d.key].runs + " runs)");
  }
  return lines.join("\n");
}
function openReport() {
  const dlg = $("#reportDialog");
  $("#studentName").value = state.name || "";
  $("#reportBody").replaceChildren(h("pre", { class: "report" }, reportText()));
  dlg.showModal();
}

// ================================================================== main render
function render() {
  renderSidebar();
  if (view.screen === "problem") renderProblem();
  else if (view.screen === "drills") renderDrills();
  else if (view.screen === "drill") renderDrillItem();
  else renderWelcome();
}

function init() {
  buildParts();
  $("#reportBtn").addEventListener("click", openReport);
  $("#studentName").addEventListener("input", (e) => { state.name = e.target.value; saveState(); $("#reportBody").replaceChildren(h("pre", { class: "report" }, reportText())); });
  $("#copyReportBtn").addEventListener("click", async () => {
    try { await navigator.clipboard.writeText(reportText()); $("#copyReportBtn").textContent = "Copied ✓"; }
    catch (e) { const r = document.createRange(); r.selectNodeContents($("#reportBody")); getSelection().removeAllRanges(); getSelection().addRange(r); $("#copyReportBtn").textContent = "Press Ctrl+C"; }
    setTimeout(() => ($("#copyReportBtn").textContent = "Copy report"), 2000);
  });
  $("#resetBtn").addEventListener("click", () => {
    if (!confirm("Erase all of your progress and drill times on this computer? This can't be undone.")) return;
    state = { name: state.name, probs: {}, levels: {}, drills: {}, seeds: {}, drafts: {}, current: null };
    saveState(); $("#reportDialog").close(); view = { screen: "welcome" }; render();
  });
  $("#menuBtn").addEventListener("click", () => document.body.classList.toggle("menu-open"));
  $("#homeBtn").addEventListener("click", () => { if (drill) { clearInterval(drill.timer); drill = null; } view = { screen: "welcome" }; render(); });
  render();
}
document.addEventListener("DOMContentLoaded", init);
