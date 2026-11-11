/* OM 305 Excel Practice — problem bank.
   Every problem is a template: build(R) makes the workbook, prompt, and model formula.
   R.p = random numbers for the problem's wording (stays the same for a student's attempt)
   R.d = random numbers for the data (the checker swaps in different data to catch hard-coded answers). */
(function (root) {
"use strict";
const XL = root.XL || (typeof require !== "undefined" ? require("./engine.js") : null);

// ================================================================== random helpers
function rand(seed) {
  let a = (seed >>> 0) || 1;
  const next = () => {
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
  const r = {
    next,
    int: (lo, hi) => lo + Math.floor(next() * (hi - lo + 1)),
    pick: (arr) => arr[Math.floor(next() * arr.length)],
    shuffle: (arr) => { const b = arr.slice(); for (let i = b.length - 1; i > 0; i--) { const j = Math.floor(next() * (i + 1)); [b[i], b[j]] = [b[j], b[i]]; } return b; },
    sample: (arr, n) => r.shuffle(arr).slice(0, n),
    money: (lo, hi) => Math.round((lo + next() * (hi - lo)) * 100) / 100,
    chance: (p) => next() < p,
    step: (lo, hi, st) => lo + st * Math.floor(next() * (Math.floor((hi - lo) / st) + 1)),
  };
  return r;
}
function makeR(pSeed, dSeed) { return { p: rand(pSeed * 7919 + 17), d: rand(dSeed * 104729 + 31) }; }

function todaySerial() { const t = new Date(); return XL.dateSerial(t.getFullYear(), t.getMonth() + 1, t.getDate()); }
function thisMonth() { return new Date().getMonth() + 1; }
const D = (y, m, d) => XL.dateSerial(y, m, d);
const col = (i) => XL.numToCol(i);

// ================================================================== name lists
const FIRST = ["James", "Mary", "Michael", "Patricia", "Robert", "Jennifer", "David", "Linda", "William", "Elizabeth", "Joseph", "Susan", "Thomas", "Jessica", "Charles", "Sarah", "Daniel", "Karen", "Matthew", "Lisa", "Anthony", "Nancy", "Mark", "Betty", "Steven", "Ashley", "Andrew", "Emily", "Joshua", "Megan", "Kevin", "Hannah", "Brian", "Olivia", "Tyler", "Grace", "Caleb", "Ava", "Logan", "Chloe", "Jalen", "Maria", "Devin", "Aaliyah", "Carlos", "Priya"];
const LAST = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis", "Rodriguez", "Martinez", "Lopez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Martin", "Lee", "Thompson", "White", "Harris", "Clark", "Lewis", "Robinson", "Walker", "Young", "Allen", "King", "Wright", "Scott", "Hill", "Green", "Adams", "Baker", "Nelson", "Carter", "Mitchell", "Turner", "Patel", "Nguyen", "Brooks"];
const CITIES = [["Tuscaloosa", "AL"], ["Birmingham", "AL"], ["Huntsville", "AL"], ["Mobile", "AL"], ["Montgomery", "AL"], ["Auburn", "AL"], ["Hoover", "AL"], ["Dothan", "AL"], ["Atlanta", "GA"], ["Nashville", "TN"], ["Jackson", "MS"], ["Pensacola", "FL"], ["Memphis", "TN"], ["Chattanooga", "TN"]];
const REGIONS = ["North", "South", "East", "West"];
const FACILITIES = ["Bessemer", "Chelsea", "Huntsville", "Pell City", "Spanish Fort", "Northport", "Opelika"];
const PRODUCTS = [
  ["Trail Tent", "Camping", 189.99], ["Dome Tent", "Camping", 129.5], ["Sleeping Bag", "Camping", 64.99], ["Camp Stove", "Camping", 54.25],
  ["Kayak Paddle", "Water", 79.0], ["Life Vest", "Water", 49.99], ["Dry Bag", "Water", 24.5], ["Fishing Rod", "Water", 89.95],
  ["Hiking Boots", "Footwear", 139.0], ["Trail Runners", "Footwear", 109.99], ["Wool Socks", "Footwear", 18.0], ["Sandals", "Footwear", 42.75],
  ["Rain Jacket", "Apparel", 99.0], ["Fleece Pullover", "Apparel", 59.5], ["Sun Hat", "Apparel", 26.99], ["Base Layer", "Apparel", 44.0],
  ["Headlamp", "Gear", 34.99], ["Water Filter", "Gear", 39.95], ["First Aid Kit", "Gear", 29.0], ["Trekking Poles", "Gear", 74.5],
];
const CATEGORIES = ["Camping", "Water", "Footwear", "Apparel", "Gear"];

function people(d, n) {
  const used = new Set(), out = [];
  while (out.length < n) {
    const f = d.pick(FIRST), l = d.pick(LAST);
    if (used.has(f + l)) continue;
    used.add(f + l);
    out.push({ first: f, last: l });
  }
  return out;
}
function ids(d, n, prefix, lo, hi) {
  const s = new Set();
  while (s.size < n) s.add(d.int(lo, hi));
  return [...s].map((x) => (prefix || "") + x);
}

// ================================================================== sheet helpers
// sheet(name, headers, rows, fmtByHeader, extras) — extras: [[cell, value, fmt]]
function sheet(name, headers, rows, fmt, extras) {
  const grid = [headers.slice()].concat(rows.map((r) => r.slice()));
  const f = {};
  if (fmt) for (const k of Object.keys(fmt)) { const i = headers.indexOf(k); if (i >= 0) f[i] = fmt[k]; }
  for (const e of extras || []) {
    const a = XL.parseAddr(e[0]);
    while (grid.length <= a.r) grid.push([]);
    while (grid[a.r].length <= a.c) grid[a.r].push(null);
    grid[a.r][a.c] = e[1];
    if (e[2]) f[a.r + ":" + a.c] = e[2];
  }
  return { name, rows: grid, fmt: f };
}
// letters for each header
function L(headers, start) { const o = {}; headers.forEach((h, i) => (o[h] = col((start || 0) + i))); return o; }
function tgt(sheetName, cell, rows, header, fmt, cols) { return { sheet: sheetName, cell, rows, cols: cols || 1, header, fmt }; }
const last = (n) => n + 1; // last data row number when data starts on row 2

// ================================================================== common datasets
function orders(d, n, opt) {
  opt = opt || {};
  const reps = people(d, 5).map((p) => p.first + " " + p.last);
  const start = todaySerial() - 120;
  const rows = [];
  const oid = ids(d, n, "", 10100, 10999).sort();
  for (let i = 0; i < n; i++) {
    const pr = d.pick(PRODUCTS);
    rows.push([+oid[i], start + d.int(0, 110), d.pick(reps), d.pick(REGIONS), pr[0], pr[1], d.int(1, 40), pr[2]]);
  }
  rows.sort((a, b) => a[1] - b[1]);
  return { headers: ["Order ID", "Order Date", "Rep", "Region", "Product", "Category", "Units", "Unit Price"], rows, reps, fmt: { "Order Date": "date", "Unit Price": "money" } };
}
// guarantee the param value appears in a column (so counts aren't 0)
function ensure(rows, ci, value, d, times) {
  const have = rows.filter((r) => r[ci] === value).length;
  for (let k = have; k < (times || 2); k++) rows[d.int(0, rows.length - 1)][ci] = value;
}

// ================================================================== the bank
const BANK = [];
const LEVEL_INFO = {};
let CUR = { part: 0, partTitle: "", level: "" };
function part(num, title) { CUR.part = num; CUR.partTitle = title; }
function level(name, intro, opts) { CUR.level = name; LEVEL_INFO[CUR.part + "|" + name] = Object.assign({ intro }, opts || {}); }
function add(id, title, build, opts) {
  BANK.push(Object.assign({ id, title, build, part: CUR.part, partTitle: CUR.partTitle, level: CUR.level }, opts || {}));
}

// ################################################################## PART 1
part(1, "Counting & Conditional Math");

level("Warm-up: SUM, AVERAGE, COUNT, MAX", `
<p>Quick ones to get your fingers moving. Each answer goes in a single cell.</p>
<ul><li><code>=SUM(G2:G15)</code> adds a range. <code>=AVERAGE(...)</code>, <code>=MAX(...)</code>, and <code>=MIN(...)</code> work the same way.</li>
<li><code>=COUNT(range)</code> counts cells that hold <b>numbers</b>. <code>=COUNTA(range)</code> counts cells that aren't empty.</li></ul>
<p><b>Speed tip:</b> type <code>=SUM(</code>, then click the first cell and Shift+click the last cell to select the range. In real Excel, Ctrl+Shift+↓ selects to the bottom of the data.</p>`);

function warmup(R, kind) {
  const n = 14, o = orders(R.d, n);
  const h = o.headers.concat(["", ""]);
  const C = L(o.headers);
  const labels = { sum: "Total Units", avg: "Average Unit Price", count: "Orders With Units", counta: "Orders With a Rep", spread: "Price Spread" };
  if (kind === "count") o.rows.forEach((r) => { if (R.d.chance(0.3)) r[6] = "TBD"; });
  if (kind === "counta") o.rows.forEach((r) => { if (R.d.chance(0.25)) r[2] = null; });
  const s = sheet("Orders", o.headers, o.rows, o.fmt, [["J2", labels[kind]]]);
  void h;
  const rng = (c) => c + "2:" + c + last(n);
  const sol = {
    sum: "=SUM(" + rng(C.Units) + ")",
    avg: "=AVERAGE(" + rng(C["Unit Price"]) + ")",
    count: "=COUNT(" + rng(C.Units) + ")",
    counta: "=COUNTA(" + rng(C.Rep) + ")",
    spread: "=MAX(" + rng(C["Unit Price"]) + ")-MIN(" + rng(C["Unit Price"]) + ")",
  }[kind];
  return { sheets: [s], target: tgt("Orders", "K2", 1, null, kind === "avg" || kind === "spread" ? "money" : "int"), solution: sol };
}
add("p1-sum", "Total units", (R) => Object.assign(warmup(R, "sum"), {
  prompt: `In <b>K2</b>, add up all the <b>Units</b> sold.`,
  hints: ["Use SUM with the Units column (G).", "=SUM(G2:G15)"],
}), { drill: "math" });
add("p1-avg", "Average price", (R) => Object.assign(warmup(R, "avg"), {
  prompt: `In <b>K2</b>, find the average <b>Unit Price</b>.`,
  hints: ["AVERAGE works like SUM.", "Unit Price is column H."],
}), { drill: "math" });
add("p1-count", "Count the numbers", (R) => Object.assign(warmup(R, "count"), {
  prompt: `Some orders say <b>TBD</b> instead of a number of units. In <b>K2</b>, count how many orders have a <b>number</b> in the Units column.`,
  hints: ["One counting function counts only cells with numbers; the other counts anything that isn't empty.", "COUNT ignores text like TBD."],
  mistakes: [{ f: "=COUNTA(G2:G15)", msg: "COUNTA counts the TBD cells too. You want the function that only counts numbers." }],
}), { drill: "math" });
add("p1-counta", "Count non-blanks", (R) => Object.assign(warmup(R, "counta"), {
  prompt: `Some orders don't have a rep assigned yet (the Rep cell is empty). In <b>K2</b>, count how many orders <b>do</b> have a rep.`,
  hints: ["The reps are names (text), so COUNT won't see them.", "COUNTA counts every cell that isn't empty."],
  mistakes: [{ f: "=COUNT(C2:C15)", msg: "COUNT only counts numbers, and the Rep column holds names. Use the counting function that counts any non-empty cell." }],
}), { drill: "math" });
add("p1-spread", "MAX minus MIN", (R) => Object.assign(warmup(R, "spread"), {
  prompt: `In <b>K2</b>, find the difference between the highest and lowest <b>Unit Price</b> (highest minus lowest).`,
  hints: ["You can do math with functions: =MAX(...)-MIN(...)", "Both use the Unit Price column, H2:H15."],
}), { drill: "math" });

level("Absolute references ($)", `
<p>When a formula is copied down, its cell references move with it. Add <code>$</code> to <b>lock</b> a reference that should stay put.</p>
<p><code>=B2/$B$12</code> copied down becomes <code>=B3/$B$12</code>, then <code>=B4/$B$12</code>, and so on. B2 moves; $B$12 stays.</p>
<p><b>Speed tip:</b> right after typing or clicking a cell reference, press <b>F4</b> to add the $ signs. Press it again to cycle through B$12 → $B12 → B12. F4 works in this practice grid, too.</p>
<p>Here you type the formula for the <b>first</b> highlighted cell, and the tool copies it down for you, just like dragging the fill handle.</p>`);

add("p1-pct-total", "Percent of total", (R) => {
  const n = R.p.int(7, 10);
  const reps = people(R.d, n);
  const rows = reps.map((p) => [p.first + " " + p.last, R.d.int(12, 95) * 1000]);
  const total = n + 2;
  const s = sheet("Sales", ["Rep", "Sales", "% of Total"], rows, { Sales: "money0" }, [["A" + total, "Total"]]);
  return {
    prompt: `<b>B${total}</b> has the total sales. In <b>C2</b>, find each rep's share of the total (their Sales ÷ the total). Lock the total so the formula still works when it's copied down to C${n + 1}.`,
    sheets: [s], prefill: [{ sheet: "Sales", cell: "B" + total, rows: 1, f: `=SUM(B2:B${n + 1})`, fmt: "money0" }],
    target: tgt("Sales", "C2", n, null, "pct1"),
    solution: `=B2/$B$${total}`,
    hints: ["Divide the rep's sales by the total.", `The total cell must not move when copied, so lock it: $B$${total}. (Click it and press F4.)`],
  };
}, { drill: "math" });
add("p1-tax", "Price with tax", (R) => {
  const n = R.p.int(8, 12);
  const items = R.d.sample(PRODUCTS, n);
  const rate = R.d.pick([0.06, 0.065, 0.07, 0.075, 0.08, 0.09, 0.1]);
  const s = sheet("Menu", ["Item", "Price", "Price with Tax"], items.map((x) => [x[0], x[2]]), { Price: "money" }, [["E1", "Sales Tax Rate"], ["F1", rate, "pct1"]]);
  return {
    prompt: `The sales tax rate is in <b>F1</b>. In <b>C2</b>, find each item's price including tax: Price × (1 + tax rate). Don't type the rate into your formula. Refer to F1 so the formula updates if the rate changes.`,
    sheets: [s], target: tgt("Menu", "C2", n, null, "money"),
    solution: "=B2*(1+$F$1)",
    hints: ["Price × (1 + rate)", "Lock F1 with $ so it doesn't slide down when copied: $F$1"],
  };
}, { drill: "math" });
add("p1-diff-avg", "Difference from average", (R) => {
  const n = R.p.int(8, 11);
  const stores = R.d.sample(FACILITIES.concat(["Tuscaloosa", "Hoover", "Vestavia", "Cullman", "Prattville"]), n);
  const rows = stores.map((s) => [s, R.d.int(150, 900) * 100]);
  const ar = n + 2;
  const s = sheet("Stores", ["Store", "Revenue", "Vs. Average"], rows, { Revenue: "money0" }, [["A" + ar, "Average"]]);
  return {
    prompt: `<b>B${ar}</b> holds the average revenue. In <b>C2</b>, find how far each store is above (positive) or below (negative) the average: Revenue − Average.`,
    sheets: [s], prefill: [{ sheet: "Stores", cell: "B" + ar, rows: 1, f: `=AVERAGE(B2:B${n + 1})`, fmt: "money0" }],
    target: tgt("Stores", "C2", n, null, "money0"),
    solution: `=B2-$B$${ar}`,
    hints: ["Revenue minus the average.", `Lock the average: $B$${ar}`],
  };
}, { drill: "math" });
add("p1-fx", "Currency conversion", (R) => {
  const n = R.p.int(8, 12);
  const items = R.d.sample(PRODUCTS, n);
  const fx = R.d.pick([0.88, 0.91, 0.92, 0.94, 0.79, 0.81]);
  const cur = fx > 0.85 ? "Euros" : "Pounds";
  const s = sheet("Export", ["Product", "Cost (USD)", "Cost (" + cur + ")"], items.map((x) => [x[0], Math.round(x[2] * 0.55 * 100) / 100]), { "Cost (USD)": "money" }, [["E1", "1 USD ="], ["F1", fx], ["G1", cur]]);
  return {
    prompt: `<b>F1</b> holds the exchange rate (how many ${cur} one U.S. dollar buys). In <b>C2</b>, convert each cost to ${cur}: Cost (USD) × exchange rate.`,
    sheets: [s], target: tgt("Export", "C2", n, null, "dec2"),
    solution: "=B2*$F$1",
    hints: ["Multiply the USD cost by the rate.", "Lock F1: $F$1"],
  };
}, { drill: "math" });

level("Mixed references", `
<p>A <b>mixed reference</b> locks only the column or only the row.</p>
<ul><li><code>$A2</code> locks the <b>column</b>, so it always reads column A but the row still moves.</li>
<li><code>B$1</code> locks the <b>row</b>, so it always reads row 1 but the column still moves.</li></ul>
<p>These problems fill a whole <b>grid</b>, down <i>and</i> across. Write the formula for the top-left cell so it works everywhere.</p>`);

add("p1-grid-qty", "Price × quantity grid", (R) => {
  const prices = R.d.sample([4.5, 6, 7.25, 8.99, 12, 15.5, 19.99, 24, 3.75, 10], 6).sort((a, b) => a - b);
  const qty = R.p.pick([[1, 5, 10, 25, 50], [2, 4, 6, 12, 24], [10, 20, 50, 100, 250]]);
  const rows = prices.map((p) => [p].concat(qty.map(() => null)));
  const s = sheet("Quote", ["Unit Price →  Qty →"].concat(qty), rows, null, []);
  for (let i = 1; i <= 6; i++) s.fmt[i + ":0"] = "money";
  return {
    prompt: `Build a price quote grid. Unit prices are down column A, quantities are across row 1. In <b>B2</b>, write <b>one</b> formula (Unit Price × Quantity) that works for the whole grid B2:F7.`,
    sheets: [s], target: tgt("Quote", "B2", 6, null, "money", 5),
    solution: "=$A2*B$1",
    hints: ["The price is always in column A, so lock the column: $A2", "The quantity is always in row 1, so lock the row: B$1"],
    mistakes: [{ f: "=A2*B1", msg: "Without $ signs, both references slide when copied. Lock column A ($A2) and row 1 (B$1)." }, { f: "=$A$2*$B$1", msg: "Fully locking both references makes every cell the same. Lock only the column of the price ($A2) and only the row of the quantity (B$1)." }],
  };
}, { drill: "math" });
add("p1-grid-markup", "Markup grid", (R) => {
  const costs = R.d.sample([12, 18, 25, 32, 40, 55, 64, 75, 90, 110], 6).sort((a, b) => a - b);
  const mk = R.p.pick([[0.1, 0.2, 0.3, 0.4], [0.15, 0.25, 0.35, 0.5], [0.05, 0.1, 0.15, 0.2]]);
  const s = sheet("Markup", ["Cost ↓  Markup →"].concat(mk), costs.map((c) => [c, null, null, null, null]), null, []);
  for (let j = 1; j <= 4; j++) s.fmt["0:" + j] = "pct";
  for (let i = 1; i <= 6; i++) s.fmt[i + ":0"] = "money";
  return {
    prompt: `Costs are down column A and markup percentages are across row 1. In <b>B2</b>, write one formula for the selling price, Cost × (1 + Markup), that fills the whole grid B2:E7.`,
    sheets: [s], target: tgt("Markup", "B2", 6, null, "money", 4),
    solution: "=$A2*(1+B$1)",
    hints: ["Cost × (1 + markup)", "Cost is always in column A ($A2); markup is always in row 1 (B$1)."],
  };
}, { drill: "math" });
add("p1-grid-fuel", "Fuel grid", (R) => {
  const miles = R.d.sample([120, 250, 310, 480, 520, 640, 775, 900], 6).sort((a, b) => a - b);
  const mpg = R.p.pick([[6, 7, 8, 9], [18, 22, 26, 30], [5.5, 6.5, 7.5, 8.5]]);
  const s = sheet("Fuel", ["Miles ↓  MPG →"].concat(mpg), miles.map((m) => [m, null, null, null, null]), null, []);
  return {
    prompt: `A dispatcher wants the gallons of fuel needed for each trip. Trip miles are down column A and truck MPG is across row 1. In <b>B2</b>, write one formula, Miles ÷ MPG, that fills B2:E7.`,
    sheets: [s], target: tgt("Fuel", "B2", 6, null, "dec1", 4),
    solution: "=$A2/B$1",
    hints: ["Miles divided by MPG.", "Lock the column of the miles ($A2) and the row of the MPG (B$1)."],
  };
}, { drill: "math" });

level("COUNTIFS", `
<p><code>=COUNTIFS(range1, criteria1, range2, criteria2, ...)</code> counts rows that meet <b>every</b> condition.</p>
<ul><li>Text criteria go in quotes: <code>=COUNTIFS(D2:D15,"East")</code>. Or point to a cell holding the text: <code>=COUNTIFS(D2:D15,J2)</code>.</li>
<li>Comparisons go in quotes too: <code>"&gt;=10"</code>. To compare to a cell, glue the operator on with &amp;: <code>"&gt;="&amp;K1</code>.</li>
<li>Referring to another sheet looks like <code>Orders!D2:D15</code> or <code>Orders!D:D</code>.</li></ul>`);

function ordersSheet(R, n, extras) {
  const o = orders(R.d, n);
  return { o, C: L(o.headers), s: sheet("Orders", o.headers, o.rows, o.fmt, extras) };
}
add("p1-cif-region", "Count one region", (R) => {
  const reg = R.p.pick(REGIONS), n = 15;
  const { o, s } = ordersSheet(R, n, [["J1", "Region"], ["K1", "Orders"], ["J2", reg]]);
  ensure(o.rows, 3, reg, R.d);
  return {
    prompt: `In <b>K2</b>, count how many orders came from the <b>${reg}</b> region (J2 already says ${reg}).`,
    sheets: [s], target: tgt("Orders", "K2", 1, null, "int"),
    solution: `=COUNTIFS(D2:D${n + 1},J2)`,
    hints: ["Region is column D.", `=COUNTIFS(D2:D${n + 1}, ...) and the criteria is the region: J2 or "${reg}".`],
  };
}, { drill: "math" });
add("p1-cif-summary", "Count by region (summary)", (R) => {
  const n = 16;
  const o = orders(R.d, n);
  const s1 = sheet("Orders", o.headers, o.rows, o.fmt);
  const s2 = sheet("Summary", ["Region", "# of Orders"], REGIONS.map((r) => [r]));
  return {
    prompt: `On the <b>Summary</b> sheet, in <b>B2</b>, count the orders on the <b>Orders</b> sheet for the region in A2. Your formula will be copied down to B5, so lock anything that shouldn't move.`,
    sheets: [s2, s1], target: tgt("Summary", "B2", 4, null, "int"),
    solution: `=COUNTIFS(Orders!$D$2:$D$${n + 1},A2)`,
    hints: ["The range is the Region column on the Orders sheet: Orders!D2:D17 (click the Orders tab while typing to grab it).", "Lock the range ($D$2:$D$17) or use the whole column (Orders!D:D). A2 should NOT be locked, so it moves to A3, A4..."],
  };
}, { drill: "math" });
add("p1-cif-two", "Two conditions", (R) => {
  const reg = R.p.pick(REGIONS), cat = R.p.pick(CATEGORIES), n = 18;
  const { o, s } = ordersSheet(R, n, [["J1", "Region"], ["J2", reg], ["K1", "Category"], ["K2", cat], ["L1", "Orders"]]);
  for (let k = 0; k < 2; k++) { const r = o.rows[R.d.int(0, n - 1)]; r[3] = reg; const pr = PRODUCTS.filter((x) => x[1] === cat)[R.d.int(0, 3)]; r[4] = pr[0]; r[5] = cat; r[7] = pr[2]; }
  return {
    prompt: `In <b>L2</b>, count the orders that are from the <b>${reg}</b> region <b>and</b> in the <b>${cat}</b> category.`,
    sheets: [s], target: tgt("Orders", "L2", 1, null, "int"),
    solution: `=COUNTIFS(D2:D${n + 1},J2,F2:F${n + 1},K2)`,
    hints: ["COUNTIFS takes pairs: range1, criteria1, range2, criteria2.", "Region is column D, Category is column F."],
  };
}, { drill: "math" });
add("p1-cif-cell", "At least a cell value", (R) => {
  const n = 16;
  const { o, s } = ordersSheet(R, n, [["J1", "Minimum Units"], ["K1", R.d.int(12, 30)], ["J2", "Big Orders"]]);
  void o;
  return {
    prompt: `<b>K1</b> holds a minimum number of units. In <b>K2</b>, count the orders with Units <b>greater than or equal to</b> K1. Refer to K1 (don't type its number) so the count updates when K1 changes.`,
    sheets: [s], target: tgt("Orders", "K2", 1, null, "int"),
    solution: `=COUNTIFS(G2:G${n + 1},">="&K1)`,
    hints: ["The operator goes in quotes and gets glued to the cell with &.", `=COUNTIFS(G2:G${n + 1}, ">="&K1)`],
    mistakes: [{ f: `=COUNTIFS(G2:G${n + 1},">=K1")`, msg: "Inside quotes, K1 is just the letters K and 1. Put only the operator in quotes and glue the cell on with &: \">=\"&K1" }],
  };
}, { drill: "math" });
add("p1-cif-between", "Between two values", (R) => {
  const n = 16;
  const lo = R.d.int(5, 12), hi = lo + R.d.int(8, 15);
  const { s } = ordersSheet(R, n, [["J1", "Low"], ["K1", lo], ["J2", "High"], ["K2", hi], ["J3", "Orders in Range"]]);
  return {
    prompt: `In <b>K3</b>, count orders whose Units are <b>between</b> the Low value (K1) and the High value (K2), including both ends.`,
    sheets: [s], target: tgt("Orders", "K3", 1, null, "int"),
    solution: `=COUNTIFS(G2:G${n + 1},">="&K1,G2:G${n + 1},"<="&K2)`,
    hints: ["You can use the same range twice in COUNTIFS with two different conditions.", "\">=\"&K1 for the low end, \"<=\"&K2 for the high end."],
  };
}, { drill: "math" });

level("SUMIFS & AVERAGEIFS", `
<p><code>=SUMIFS(sum_range, range1, criteria1, ...)</code> adds up the rows that meet every condition. Notice that the range to <b>add</b> comes <b>first</b>.</p>
<p><code>=AVERAGEIFS(average_range, range1, criteria1, ...)</code> works the same way but averages.</p>
<p>All the ranges must be the same size, for example all rows 2 through 16.</p>`);

function ordersWithTotal(R, n) {
  const o = orders(R.d, n);
  const h = o.headers.concat(["Total"]);
  const rows = o.rows.map((r) => r.concat([Math.round(r[6] * r[7] * 100) / 100]));
  return { o, h, rows, fmt: Object.assign({ Total: "money" }, o.fmt), C: L(h) };
}
add("p1-sif-rep", "Units for one rep", (R) => {
  const n = 16, o = orders(R.d, n);
  const rep = o.reps[R.p.int(0, 4)];
  ensure(o.rows, 2, rep, R.d, 3);
  const s = sheet("Orders", o.headers, o.rows, o.fmt, [["J1", "Rep"], ["J2", rep], ["K1", "Units Sold"]]);
  return {
    prompt: `In <b>K2</b>, add up the <b>Units</b> sold by <b>${rep}</b> (the name in J2).`,
    sheets: [s], target: tgt("Orders", "K2", 1, null, "int"),
    solution: `=SUMIFS(G2:G${n + 1},C2:C${n + 1},J2)`,
    hints: ["SUMIFS: what to add first (Units, column G), then the condition range (Rep, column C) and the criteria.", `=SUMIFS(G2:G${n + 1}, C2:C${n + 1}, J2)`],
    mistakes: [{ f: `=SUMIFS(C2:C${n + 1},G2:G${n + 1},J2)`, msg: "The range to add up goes first in SUMIFS. Here that's Units (G)." }],
  };
}, { drill: "math" });
add("p1-sif-summary", "Revenue by region (summary)", (R) => {
  const n = 16, t = ordersWithTotal(R, n);
  const s1 = sheet("Orders", t.h, t.rows, t.fmt);
  const s2 = sheet("Summary", ["Region", "Total Revenue"], REGIONS.map((r) => [r]));
  return {
    prompt: `On the <b>Summary</b> sheet, in <b>B2</b>, add up the <b>Total</b> column on the Orders sheet for the region in A2. It'll be copied down to B5.`,
    sheets: [s2, s1], target: tgt("Summary", "B2", 4, null, "money"),
    solution: `=SUMIFS(Orders!$I$2:$I$${n + 1},Orders!$D$2:$D$${n + 1},A2)`,
    hints: ["Sum range: Orders!I2:I17 (Total). Criteria range: Orders!D2:D17 (Region).", "Lock both ranges with $ (or use whole columns like Orders!I:I) so they don't slide when copied down."],
  };
}, { drill: "math" });
add("p1-sif-two", "Sum with two conditions", (R) => {
  const n = 18, t = ordersWithTotal(R, n);
  const reg = R.p.pick(REGIONS), cat = R.p.pick(CATEGORIES);
  for (let k = 0; k < 2; k++) { const r = t.rows[R.d.int(0, n - 1)]; r[3] = reg; r[5] = cat; }
  const s = sheet("Orders", t.h, t.rows, t.fmt, [["K1", "Region"], ["K2", reg], ["L1", "Category"], ["L2", cat], ["M1", "Revenue"]]);
  return {
    prompt: `In <b>M2</b>, add up the <b>Total</b> for orders in the <b>${reg}</b> region <b>and</b> the <b>${cat}</b> category.`,
    sheets: [s], target: tgt("Orders", "M2", 1, null, "money"),
    solution: `=SUMIFS(I2:I${n + 1},D2:D${n + 1},K2,F2:F${n + 1},L2)`,
    hints: ["Sum range first (Total, column I), then two range/criteria pairs.", "Region is D, Category is F."],
  };
}, { drill: "math" });
add("p1-aif-cat", "Average for a category", (R) => {
  const n = 18, o = orders(R.d, n), cat = R.p.pick(CATEGORIES);
  for (let k = 0; k < 3; k++) { const r = o.rows[R.d.int(0, n - 1)]; const pr = PRODUCTS.filter((x) => x[1] === cat)[R.d.int(0, 3)]; r[4] = pr[0]; r[5] = cat; r[7] = pr[2]; }
  const s = sheet("Orders", o.headers, o.rows, o.fmt, [["J1", "Category"], ["J2", cat], ["K1", "Avg Units"]]);
  return {
    prompt: `In <b>K2</b>, find the average number of <b>Units</b> per order for the <b>${cat}</b> category.`,
    sheets: [s], target: tgt("Orders", "K2", 1, null, "dec2"),
    solution: `=AVERAGEIFS(G2:G${n + 1},F2:F${n + 1},J2)`,
    hints: ["AVERAGEIFS: the range to average first.", "Average Units (G) where Category (F) matches J2."],
  };
}, { drill: "math" });
add("p1-aif-cell", "Average above a cutoff", (R) => {
  const n = 16, t = ordersWithTotal(R, n);
  const s = sheet("Orders", t.h, t.rows, t.fmt, [["K1", "Min Units"], ["L1", R.d.int(10, 25)], ["K2", "Avg Total"]]);
  return {
    prompt: `In <b>L2</b>, find the average <b>Total</b> for orders with Units of at least the number in <b>L1</b>.`,
    sheets: [s], target: tgt("Orders", "L2", 1, null, "money"),
    solution: `=AVERAGEIFS(I2:I${n + 1},G2:G${n + 1},">="&L1)`,
    hints: ["Average the Total column (I) where Units (G) meets the condition.", "The criteria is \">=\"&L1"],
  };
}, { drill: "math" });

level("Wildcards", `
<p>In COUNTIFS, SUMIFS, AVERAGEIFS (and exact-match lookups), text criteria can use wildcards:</p>
<ul><li><code>*</code> means "any number of characters". <code>"*Tent*"</code> means "contains Tent", <code>"BK*"</code> means "starts with BK", and <code>"*Kit"</code> means "ends with Kit".</li>
<li><code>?</code> means "exactly one character".</li></ul>
<p>To use a cell: <code>"*"&amp;J2&amp;"*"</code>.</p>`);

add("p1-wild-contains", "Contains a word", (R) => {
  const n = 16, o = orders(R.d, n);
  const word = R.p.pick(["Tent", "Trail", "Bag", "Kit"]);
  const matches = PRODUCTS.filter((x) => x[0].includes(word));
  for (let k = 0; k < 3; k++) { const r = o.rows[R.d.int(0, n - 1)]; const pr = R.d.pick(matches); r[4] = pr[0]; r[5] = pr[1]; r[7] = pr[2]; }
  const s = sheet("Orders", o.headers, o.rows, o.fmt, [["J1", "Word"], ["J2", word], ["K1", "Orders"]]);
  return {
    prompt: `In <b>K2</b>, count the orders whose <b>Product</b> name contains the word <b>${word}</b> anywhere in it.`,
    sheets: [s], target: tgt("Orders", "K2", 1, null, "int"),
    solution: `=COUNTIFS(E2:E${n + 1},"*"&J2&"*")`,
    hints: ["Use * on both sides of the word.", `"*${word}*" (or "*"&J2&"*")`],
  };
}, { drill: "math" });
add("p1-wild-starts", "Starts with", (R) => {
  const n = 15;
  const pre = ["BK", "TR", "CP", "WT"];
  const pick = R.p.pick(pre);
  const rows = [];
  const skus = ids(R.d, n, "", 1000, 9999);
  for (let i = 0; i < n; i++) rows.push([R.d.pick(pre) + "-" + skus[i], R.d.int(0, 80), R.d.money(5, 120)]);
  ensure(rows, 0, pick + "-" + skus[0], R.d, 2);
  const s = sheet("Inventory", ["SKU", "On Hand", "Cost"], rows, { Cost: "money" }, [["E1", "Prefix"], ["E2", pick], ["F1", "# of SKUs"]]);
  return {
    prompt: `SKU codes start with a 2-letter prefix. In <b>F2</b>, count how many SKUs <b>start with</b> <b>${pick}</b>.`,
    sheets: [s], target: tgt("Inventory", "F2", 1, null, "int"),
    solution: `=COUNTIFS(A2:A${n + 1},E2&"*")`,
    hints: ["Starts with: the prefix followed by *.", `"${pick}*" (or E2&"*")`],
  };
}, { drill: "math" });
add("p1-wild-ends", "Ends with (SUMIFS)", (R) => {
  const n = 16;
  const suffix = R.p.pick(["Kit", "Pack", "Set"]);
  const names = ["Repair", "Starter", "Cook", "Travel", "Gift", "Fishing", "Hydration", "Cleaning"];
  const tails = ["Kit", "Pack", "Set", "Bundle"];
  const rows = [];
  for (let i = 0; i < n; i++) rows.push([R.d.pick(names) + " " + R.d.pick(tails), R.d.int(2, 60)]);
  for (let k = 0; k < 2; k++) rows[R.d.int(0, n - 1)][0] = R.d.pick(names) + " " + suffix;
  const s = sheet("Sales", ["Item", "Units"], rows, null, [["D1", "Ends With"], ["D2", suffix], ["E1", "Units"]]);
  return {
    prompt: `In <b>E2</b>, add up the <b>Units</b> for items whose name <b>ends with</b> <b>${suffix}</b>.`,
    sheets: [s], target: tgt("Sales", "E2", 1, null, "int"),
    solution: `=SUMIFS(B2:B${n + 1},A2:A${n + 1},"*"&D2)`,
    hints: ["Ends with: * then the word.", `=SUMIFS(B2:B${n + 1}, A2:A${n + 1}, "*${suffix}")`],
  };
}, { drill: "math" });

level("SUMPRODUCT", `
<p><code>=SUMPRODUCT(range1, range2)</code> multiplies the ranges row by row, then adds up all the products. You get the total revenue in one cell without a helper column.</p>
<p><code>=SUMPRODUCT(B2:B10, C2:C10)</code> is the same as <code>B2*C2 + B3*C3 + ... + B10*C10</code>. The ranges have to be the same size.</p>`);

add("p1-sp-revenue", "Total revenue", (R) => {
  const n = 15, o = orders(R.d, n);
  const s = sheet("Orders", o.headers, o.rows, o.fmt, [["J1", "Total Revenue"]]);
  return {
    prompt: `In <b>J2</b>, use SUMPRODUCT to find total revenue: the sum of Units × Unit Price over every order.`,
    sheets: [s], target: tgt("Orders", "J2", 1, null, "money"),
    solution: `=SUMPRODUCT(G2:G${n + 1},H2:H${n + 1})`, require: ["SUMPRODUCT"],
    hints: ["Units is G, Unit Price is H.", `=SUMPRODUCT(G2:G${n + 1}, H2:H${n + 1})`],
  };
}, { drill: "math" });
add("p1-sp-grade", "Weighted grade", (R) => {
  const cats = [["Homework", 0.2], ["Quizzes", 0.15], ["In-Class", 0.15], ["Exam 1", 0.25], ["Exam 2", 0.25]];
  const rows = cats.map((c) => [c[0], c[1], R.d.int(62, 100) / 100]);
  const s = sheet("Grade", ["Category", "Weight", "Score"], rows, { Weight: "pct", Score: "pct" }, [["E1", "Final Grade"]]);
  return {
    prompt: `In <b>E2</b>, calculate the weighted final grade: each Score × its Weight, all added up. Use SUMPRODUCT.`,
    sheets: [s], target: tgt("Grade", "E2", 1, null, "pct1"),
    solution: "=SUMPRODUCT(B2:B6,C2:C6)", require: ["SUMPRODUCT"],
    hints: ["Weights are B2:B6 and scores are C2:C6.", "=SUMPRODUCT(B2:B6, C2:C6)"],
  };
}, { drill: "math" });
add("p1-sp-inventory", "Inventory value", (R) => {
  const n = R.p.int(10, 14);
  const items = R.d.sample(PRODUCTS, n);
  const rows = items.map((x) => [x[0], R.d.int(0, 60), Math.round(x[2] * 0.5 * 100) / 100]);
  const s = sheet("Inventory", ["Item", "On Hand", "Unit Cost"], rows, { "Unit Cost": "money" }, [["E1", "Inventory Value"]]);
  return {
    prompt: `In <b>E2</b>, find the total value of the inventory (On Hand × Unit Cost for every item, added up) with SUMPRODUCT.`,
    sheets: [s], target: tgt("Inventory", "E2", 1, null, "money"),
    solution: `=SUMPRODUCT(B2:B${n + 1},C2:C${n + 1})`, require: ["SUMPRODUCT"],
    hints: ["On Hand is column B, Unit Cost is column C.", "Both ranges must cover the same rows."],
  };
}, { drill: "math" });

// ################################################################## PART 2
part(2, "Logical Functions");

level("IF basics", `
<p><code>=IF(logical_test, value_if_true, value_if_false)</code></p>
<ul><li>Text answers go in quotes: <code>=IF(B2&gt;=70,"Pass","Fail")</code>.</li>
<li>To leave a cell looking blank, use two quote marks: <code>""</code>.</li>
<li>The true/false parts can be math: <code>=IF(B2&gt;C2, B2*0.05, 0)</code>.</li>
<li>Comparison operators: <code>=  &lt;&gt;  &gt;  &lt;  &gt;=  &lt;=</code></li></ul>`);

add("p2-if-reorder", "Reorder flag", (R) => {
  const n = R.p.int(10, 14), items = R.d.sample(PRODUCTS, n);
  const rows = items.map((x) => { const rp = R.d.int(5, 25); return [x[0], Math.max(0, rp + R.d.int(-10, 15)), rp]; });
  const s = sheet("Inventory", ["Item", "On Hand", "Reorder Point", "Reorder?"], rows);
  return {
    prompt: `In <b>D2</b>, show <b>"Yes"</b> if On Hand is <b>at or below</b> the Reorder Point. Otherwise leave the cell blank ("").`,
    sheets: [s], target: tgt("Inventory", "D2", n),
    solution: '=IF(B2<=C2,"Yes","")',
    hints: ["At or below means <=.", '=IF(B2<=C2, "Yes", "")'],
    mistakes: [{ f: '=IF(B2<C2,"Yes","")', msg: "Close! \"At or below\" includes equal, so use <= instead of <." }],
  };
}, { drill: "logic" });
add("p2-if-pass", "Pass / Fail", (R) => {
  const n = R.p.int(10, 14), ppl = people(R.d, n);
  const cut = R.d.pick([60, 65, 70, 75]);
  const rows = ppl.map((p) => [p.first + " " + p.last, R.d.int(45, 100)]);
  const s = sheet("Scores", ["Student", "Score", "Result"], rows, null, [["E1", "Passing Score"], ["F1", cut]]);
  return {
    prompt: `<b>F1</b> holds the passing score. In <b>C2</b>, show <b>"Pass"</b> if the Score is greater than or equal to the passing score, otherwise <b>"Fail"</b>.`,
    sheets: [s], target: tgt("Scores", "C2", n),
    solution: '=IF(B2>=$F$1,"Pass","Fail")',
    hints: ["Compare the score to F1, and lock F1 since the formula gets copied down.", '=IF(B2>=$F$1, "Pass", "Fail")'],
  };
}, { drill: "logic" });
add("p2-if-bonus", "Bonus if over quota", (R) => {
  const n = R.p.int(8, 12), ppl = people(R.d, n);
  const rate = R.d.pick([0.03, 0.04, 0.05, 0.06]);
  const rows = ppl.map((p) => { const q = R.d.int(40, 80) * 1000; return [p.first + " " + p.last, q + R.d.int(-15, 20) * 1000, q]; });
  const s = sheet("Reps", ["Rep", "Sales", "Quota", "Bonus"], rows, { Sales: "money0", Quota: "money0" }, [["F1", "Bonus Rate"], ["G1", rate, "pct"]]);
  return {
    prompt: `Reps who sell <b>more than</b> their quota earn a bonus of Sales × the bonus rate in <b>G1</b>. In <b>D2</b>, calculate the bonus (0 for reps who didn't beat their quota).`,
    sheets: [s], target: tgt("Reps", "D2", n, null, "money"),
    solution: "=IF(B2>C2,B2*$G$1,0)",
    hints: ["The true part can be math: B2*$G$1", "=IF(B2>C2, B2*$G$1, 0)"],
  };
}, { drill: "logic" });
add("p2-if-text", "IF on text", (R) => {
  const n = R.p.int(10, 14);
  const fee = R.p.pick([[25, 10], [30, 12], [19.99, 7.99]]);
  const rows = ids(R.d, n, "S-", 2000, 2999).map((id) => [id, R.d.pick(["Standard", "Express", "Standard", "Ground"])]);
  ensure(rows, 1, "Express", R.d);
  const s = sheet("Shipments", ["Shipment", "Method", "Fee"], rows);
  return {
    prompt: `In <b>C2</b>, charge <b>${fee[0]}</b> if the Method is <b>"Express"</b>. Anything else costs <b>${fee[1]}</b>.`,
    sheets: [s], target: tgt("Shipments", "C2", n, null, "money"),
    solution: `=IF(B2="Express",${fee[0]},${fee[1]})`,
    hints: ["Text being compared needs quotes: B2=\"Express\"", `=IF(B2="Express", ${fee[0]}, ${fee[1]})`],
  };
}, { drill: "logic" });

level("IF with AND / OR", `
<ul><li><code>AND(test1, test2, ...)</code> is TRUE only if <b>every</b> test is true.</li>
<li><code>OR(test1, test2, ...)</code> is TRUE if <b>at least one</b> test is true.</li></ul>
<p>Put them inside IF: <code>=IF(AND(C2="Huntsville", D2="FT"), 0.04, 0.03)</code>.</p>`);

add("p2-and-raise", "Raise % (AND)", (R) => {
  const n = R.p.int(10, 14), fac = R.p.pick(FACILITIES.slice(0, 5));
  const ppl = people(R.d, n), idl = ids(R.d, n, "", 1000, 9999);
  const rows = ppl.map((p, i) => [+idl[i], p.first + " " + p.last, R.d.pick(FACILITIES.slice(0, 5)), R.d.pick(["FT", "FT", "PT"])]);
  rows[0].splice(2, 2, fac, "FT");
  rows[1].splice(2, 2, fac, "PT");
  rows[2].splice(3, 1, "FT");
  const s = sheet("Staff", ["Emp ID", "Name", "Facility", "Job Status", "Raise %"], rows);
  return {
    prompt: `<b>Full-time ("FT")</b> employees at the <b>${fac}</b> facility get a <b>4%</b> raise. Everyone else gets <b>3%</b>. In <b>E2</b>, show the raise as a decimal (0.04 or 0.03).`,
    sheets: [s], target: tgt("Staff", "E2", n, null, "pct"),
    solution: `=IF(AND(C2="${fac}",D2="FT"),0.04,0.03)`,
    hints: ["Both things must be true, so use AND.", `=IF(AND(C2="${fac}", D2="FT"), 0.04, 0.03)`],
    mistakes: [{ f: `=IF(OR(C2="${fac}",D2="FT"),0.04,0.03)`, msg: "OR gives the raise if either is true. Here both must be true, so use AND." }],
  };
}, { drill: "logic" });
add("p2-or-priority", "Priority flag (OR)", (R) => {
  const n = R.p.int(10, 14), lim = R.p.pick([500, 750, 1000]);
  const rows = ids(R.d, n, "PO-", 300, 999).map((id) => [id, R.d.pick(["On Time", "On Time", "Late", "On Time"]), R.d.int(80, 1600)]);
  ensure(rows, 1, "Late", R.d);
  const s = sheet("POs", ["PO", "Status", "Amount", "Priority"], rows, { Amount: "money0" });
  return {
    prompt: `A purchase order is a priority if its Status is <b>"Late"</b> <b>or</b> its Amount is <b>over ${lim}</b>. In <b>D2</b>, show <b>"High"</b> for priority orders and <b>"Normal"</b> for the rest.`,
    sheets: [s], target: tgt("POs", "D2", n),
    solution: `=IF(OR(B2="Late",C2>${lim}),"High","Normal")`,
    hints: ["Either condition is enough, so use OR.", `=IF(OR(B2="Late", C2>${lim}), "High", "Normal")`],
  };
}, { drill: "logic" });
add("p2-or-ship", "Free shipping", (R) => {
  const n = R.p.int(10, 14), ship = R.p.pick([6.99, 8.99, 9.5]);
  const rows = ids(R.d, n, "", 5000, 5999).map((id) => [+id, R.d.money(15, 160), R.d.pick(["Y", "N", "N"])]);
  const s = sheet("Carts", ["Order", "Subtotal", "Member", "Shipping"], rows, { Subtotal: "money" }, [["F1", "Free Shipping At"], ["G1", R.d.pick([50, 75, 100]), "money0"]]);
  return {
    prompt: `Shipping is <b>free (0)</b> if the Subtotal is at least the amount in <b>G1</b> <b>or</b> the customer is a Member ("Y"). Otherwise shipping costs <b>${ship}</b>. Fill in <b>D2</b>.`,
    sheets: [s], target: tgt("Carts", "D2", n, null, "money"),
    solution: `=IF(OR(B2>=$G$1,C2="Y"),0,${ship})`,
    hints: ["OR with two tests: B2>=$G$1 and C2=\"Y\".", "Lock G1 since the formula is copied down."],
  };
}, { drill: "logic" });
add("p2-and-dean", "Dean's list (AND)", (R) => {
  const n = R.p.int(10, 14), ppl = people(R.d, n);
  const rows = ppl.map((p) => [p.first + " " + p.last, Math.round(R.d.int(250, 400)) / 100, R.d.pick([9, 12, 12, 15, 15, 18])]);
  const s = sheet("Students", ["Student", "GPA", "Hours", "Dean's List"], rows, { GPA: "dec2" });
  return {
    prompt: `A student makes the Dean's List with a <b>GPA of 3.5 or higher</b> <b>and</b> at least <b>12 hours</b>. In <b>D2</b>, show <b>"Yes"</b> or <b>"No"</b>.`,
    sheets: [s], target: tgt("Students", "D2", n),
    solution: '=IF(AND(B2>=3.5,C2>=12),"Yes","No")',
    hints: ["Both must be true: AND(B2>=3.5, C2>=12)", "Wrap it in IF to show Yes or No."],
  };
}, { drill: "logic" });

level("Nested IF", `
<p>When there are more than two outcomes, put an IF <b>inside</b> another IF's false part:</p>
<p><code>=IF(B2&gt;=90,"A",IF(B2&gt;=80,"B",IF(B2&gt;=70,"C","F")))</code></p>
<p>Excel checks the tests in order and stops at the first true one, so start from the top tier. Count your closing parentheses: one per IF.</p>
<p>You can also mix AND/OR in: <code>=IF(OR(D2="FT",AND(D2="PT",E2&gt;=2)),"Yes","")</code>.</p>`);

add("p2-nest-grade", "Letter grades", (R) => {
  const n = R.p.int(10, 14), ppl = people(R.d, n);
  const rows = ppl.map((p) => [p.first + " " + p.last, R.d.int(52, 100)]);
  const s = sheet("Grades", ["Student", "Average", "Letter"], rows);
  return {
    prompt: `In <b>C2</b>, assign a letter grade: <b>A</b> for 90 and up, <b>B</b> for 80–89, <b>C</b> for 70–79, <b>D</b> for 60–69, and <b>F</b> below 60.`,
    sheets: [s], target: tgt("Grades", "C2", n),
    solution: '=IF(B2>=90,"A",IF(B2>=80,"B",IF(B2>=70,"C",IF(B2>=60,"D","F"))))',
    hints: ["Start with the highest grade: IF(B2>=90,\"A\", ...)", "Each false part holds the next IF. The last one ends with \"D\",\"F\" and then four closing parentheses."],
  };
}, { drill: "logic" });
add("p2-nest-ship", "Shipping tiers", (R) => {
  const n = R.p.int(10, 14), t = R.p.pick([[5, 20, 7.5, 15, 30], [10, 50, 9, 18, 40], [2, 10, 4.5, 11, 25]]);
  const rows = ids(R.d, n, "PK", 100, 999).map((id) => [id, Math.round(R.d.int(5, t[1] * 25)) / 10]);
  const s = sheet("Packages", ["Package", "Weight (lb)", "Shipping"], rows);
  return {
    prompt: `In <b>C2</b>, find the shipping charge: packages <b>under ${t[0]} lb</b> cost <b>${t[2]}</b>, packages <b>under ${t[1]} lb</b> cost <b>${t[3]}</b>, and anything heavier costs <b>${t[4]}</b>.`,
    sheets: [s], target: tgt("Packages", "C2", n, null, "money"),
    solution: `=IF(B2<${t[0]},${t[2]},IF(B2<${t[1]},${t[3]},${t[4]}))`,
    hints: [`Start with the lightest tier: IF(B2<${t[0]}, ${t[2]}, ...)`, `The false part checks the next tier: IF(B2<${t[1]}, ${t[3]}, ${t[4]})`],
  };
}, { drill: "logic" });
add("p2-nest-vacation", "Vacation eligibility", (R) => {
  const n = R.p.int(10, 14), ppl = people(R.d, n);
  const rows = ppl.map((p) => [p.first + " " + p.last, R.d.pick(["FT", "PT", "PT"]), Math.round(R.d.int(1, 120)) / 10]);
  rows[0][1] = "PT"; rows[0][2] = 2; rows[1][1] = "PT"; rows[1][2] = 1.9;
  const s = sheet("Staff", ["Name", "Job Status", "Years of Service", "Eligible for Vacation?"], rows, { "Years of Service": "dec1" });
  return {
    prompt: `Employees are eligible for vacation if they're full-time (<b>"FT"</b>), <b>or</b> if they're part-time (<b>"PT"</b>) <b>and</b> have <b>2 or more</b> years of service. In <b>D2</b>, show <b>"Yes"</b> if eligible; otherwise leave it blank ("").`,
    sheets: [s], target: tgt("Staff", "D2", n),
    solution: '=IF(OR(B2="FT",AND(B2="PT",C2>=2)),"Yes","")',
    hints: ["OR( full-time , AND( part-time , years>=2 ) )", '=IF(OR(B2="FT", AND(B2="PT", C2>=2)), "Yes", "")'],
  };
}, { drill: "logic" });
add("p2-nest-days", "Vacation days by facility", (R) => {
  const n = R.p.int(10, 14), facs = R.p.sample(FACILITIES, 4), days = R.p.sample([3, 5, 6, 7, 8, 9, 10], 4);
  const ppl = people(R.d, n);
  const rows = ppl.map((p) => [p.first + " " + p.last, R.d.pick(facs)]);
  const s = sheet("Staff", ["Name", "Facility", "Vacation Days"], rows);
  const list = facs.map((f, i) => `${f}: <b>${days[i]}</b>`).join(", ");
  return {
    prompt: `Each facility gives a different number of vacation days. ${list}. In <b>C2</b>, use nested IFs to show each employee's days. (You can test the first three; anyone left must be at the fourth.)`,
    sheets: [s], target: tgt("Staff", "C2", n, null, "int"),
    solution: `=IF(B2="${facs[0]}",${days[0]},IF(B2="${facs[1]}",${days[1]},IF(B2="${facs[2]}",${days[2]},${days[3]})))`,
    hints: [`=IF(B2="${facs[0]}", ${days[0]}, IF(B2="${facs[1]}", ...`, "Three IFs are enough: the last false part is the fourth facility's days."],
  };
}, { drill: "logic" });

level("IS functions & IFERROR", `
<ul><li><code>ISBLANK(A2)</code> is TRUE if A2 is empty.</li>
<li><code>ISNUMBER(A2)</code> and <code>ISTEXT(A2)</code> check what kind of value is in a cell.</li>
<li><code>ISERROR(x)</code> is TRUE if x is any error (#N/A, #DIV/0!, #VALUE!...).</li>
<li><code>=IFERROR(formula, value_if_error)</code> shows the formula's answer, or a backup value if it errors.</li></ul>`);

add("p2-is-blank", "Missing email (ISBLANK)", (R) => {
  const n = R.p.int(10, 14), ppl = people(R.d, n);
  const rows = ppl.map((p) => [p.first, p.last, R.d.chance(0.3) ? null : (p.first[0] + p.last).toLowerCase() + "@crimson.ua.edu"]);
  rows[0][2] = null;
  const s = sheet("Contacts", ["First", "Last", "Email", "Check"], rows);
  return {
    prompt: `In <b>D2</b>, use ISBLANK to show <b>"Missing"</b> if the Email cell is empty, otherwise <b>"OK"</b>.`,
    sheets: [s], target: tgt("Contacts", "D2", n),
    solution: '=IF(ISBLANK(C2),"Missing","OK")', require: ["ISBLANK"],
    hints: ["ISBLANK(C2) gives TRUE or FALSE, so put it in IF's test.", '=IF(ISBLANK(C2), "Missing", "OK")'],
  };
}, { drill: "logic" });
add("p2-iferror-div", "Cost per unit (IFERROR)", (R) => {
  const n = R.p.int(10, 14), items = R.d.sample(PRODUCTS, n);
  const rows = items.map((x) => [x[0], R.d.int(200, 3000), R.d.chance(0.25) ? 0 : R.d.int(5, 120)]);
  rows[0][2] = 0;
  const s = sheet("Production", ["Item", "Total Cost", "Units Made", "Cost per Unit"], rows, { "Total Cost": "money0" });
  return {
    prompt: `In <b>D2</b>, divide Total Cost by Units Made. Some items made 0 units, which causes #DIV/0!. Wrap the division in IFERROR so those show <b>0</b> instead.`,
    sheets: [s], target: tgt("Production", "D2", n, null, "money"),
    solution: "=IFERROR(B2/C2,0)", require: ["IFERROR"],
    hints: ["=IFERROR( your formula , what to show if it errors )", "=IFERROR(B2/C2, 0)"],
  };
}, { drill: "logic" });
add("p2-is-number", "Valid IDs (ISNUMBER)", (R) => {
  const n = R.p.int(10, 14), ppl = people(R.d, n);
  const rows = ppl.map((p) => [p.first + " " + p.last, R.d.chance(0.3) ? R.d.pick(["pending", "TBD", "N/A", "unknown"]) : R.d.int(10000, 99999)]);
  rows[1][1] = "pending";
  const s = sheet("Badges", ["Employee", "Badge ID", "Status"], rows);
  return {
    prompt: `Badge IDs should be numbers. In <b>C2</b>, show <b>"Valid"</b> if the Badge ID is a number, otherwise <b>"Check"</b>.`,
    sheets: [s], target: tgt("Badges", "C2", n),
    solution: '=IF(ISNUMBER(B2),"Valid","Check")', require: ["ISNUMBER"],
    hints: ["ISNUMBER(B2) is TRUE for numbers.", '=IF(ISNUMBER(B2), "Valid", "Check")'],
  };
}, { drill: "logic" });
add("p2-is-text", "Text in a number column (ISTEXT)", (R) => {
  const n = R.p.int(10, 14);
  const rows = ids(R.d, n, "R", 100, 999).map((id) => [id, R.d.chance(0.3) ? R.d.pick(["five", "n/a", "4 stars", "great"]) : R.d.int(1, 5)]);
  rows[2][1] = "n/a";
  const s = sheet("Survey", ["Response", "Rating", "Is Text?"], rows);
  return {
    prompt: `Some survey ratings were typed as words. In <b>C2</b>, use ISTEXT to show TRUE if the Rating is text and FALSE if it's a number. (Just ISTEXT, no IF needed.)`,
    sheets: [s], target: tgt("Survey", "C2", n),
    solution: "=ISTEXT(B2)", require: ["ISTEXT"],
    hints: ["ISTEXT returns TRUE or FALSE all by itself.", "=ISTEXT(B2)"],
  };
}, { drill: "logic" });
add("p2-is-error", "Flag errors (ISERROR)", (R) => {
  const n = R.p.int(10, 14);
  const rows = ids(R.d, n, "L", 10, 99).map((id) => [id, R.d.int(100, 900), R.d.chance(0.25) ? 0 : R.d.int(2, 30)]);
  rows[0][2] = 0;
  const s = sheet("Lines", ["Line", "Output", "Hours", "Flag"], rows);
  return {
    prompt: `Output ÷ Hours gives units per hour, but some lines logged 0 hours. In <b>D2</b>, use ISERROR to show <b>"Check"</b> if Output/Hours would be an error, otherwise leave it blank ("").`,
    sheets: [s], target: tgt("Lines", "D2", n),
    solution: '=IF(ISERROR(B2/C2),"Check","")', require: ["ISERROR"],
    hints: ["ISERROR(B2/C2) is TRUE when the division fails.", '=IF(ISERROR(B2/C2), "Check", "")'],
  };
}, { drill: "logic" });

level("Logic in Excel Tables", `
<p>When data is formatted as an <b>Excel Table</b>, formulas can use column names instead of cell addresses:</p>
<ul><li><code>[@Facility]</code> means "the Facility value in <b>this row</b>".</li>
<li>Column names with spaces need an extra set of brackets: <code>[@[Job Status]]</code>.</li></ul>
<p><code>=IF(AND([@Facility]="Huntsville",[@[Job Status]]="FT"),0.04,0.03)</code></p>
<p>These problems ask you to use structured references. In real Excel, clicking a cell inside a table types them for you.</p>`);

function staffTable(R, n, extra) {
  const ppl = people(R.d, n), idl = ids(R.d, n, "", 1000, 9999);
  const rows = ppl.map((p, i) => [+idl[i], p.first, p.last, R.d.pick(FACILITIES.slice(0, 5)), R.d.pick(["FT", "FT", "PT"])]);
  const h = ["Employee ID", "First Name", "Last Name", "Facility", "Job Status"].concat(extra);
  return { rows: rows.map((r) => r.concat(extra.map(() => null))), h };
}
add("p2-tbl-status", "Full-time label", (R) => {
  const n = R.p.int(10, 13), t = staffTable(R, n, ["Status Label"]);
  return {
    prompt: `The data is a Table named <b>Staff</b>. In <b>F2</b>, show <b>"Full-time"</b> if Job Status is "FT", otherwise <b>"Part-time"</b>. Use a structured reference like <code>[@[Job Status]]</code>.`,
    sheets: [sheet("Staff", t.h, t.rows)], tables: [{ name: "Staff", sheet: "Staff", top: 0, left: 0, ncols: 6, nrows: n }],
    target: tgt("Staff", "F2", n), requireStructured: true,
    solution: '=IF([@[Job Status]]="FT","Full-time","Part-time")',
    hints: ["Job Status has a space, so it's written [@[Job Status]].", '=IF([@[Job Status]]="FT", "Full-time", "Part-time")'],
  };
}, { drill: "logic" });
add("p2-tbl-raise", "Raise % in a Table", (R) => {
  const n = R.p.int(10, 13), t = staffTable(R, n, ["Raise %"]), fac = R.p.pick(FACILITIES.slice(0, 5));
  t.rows[0][3] = fac; t.rows[0][4] = "FT";
  return {
    prompt: `Table <b>Staff</b>: full-time ("FT") employees at <b>${fac}</b> get <b>0.04</b>; everyone else gets <b>0.03</b>. Fill in <b>F2</b> using structured references.`,
    sheets: [sheet("Staff", t.h, t.rows)], tables: [{ name: "Staff", sheet: "Staff", top: 0, left: 0, ncols: 6, nrows: n }],
    target: tgt("Staff", "F2", n, null, "pct"), requireStructured: true,
    solution: `=IF(AND([@Facility]="${fac}",[@[Job Status]]="FT"),0.04,0.03)`,
    hints: ["[@Facility] and [@[Job Status]]", `=IF(AND([@Facility]="${fac}", [@[Job Status]]="FT"), 0.04, 0.03)`],
  };
}, { drill: "logic" });
add("p2-tbl-reorder", "Reorder flag in a Table", (R) => {
  const n = R.p.int(10, 13), items = R.d.sample(PRODUCTS, n);
  const rows = items.map((x) => { const rp = R.d.int(5, 25); return [x[0], x[1], Math.max(0, rp + R.d.int(-10, 15)), rp, null]; });
  return {
    prompt: `Table <b>Stock</b>: in <b>E2</b>, show <b>"Order"</b> if On Hand is less than the Reorder Point, otherwise <b>"OK"</b>. Use structured references.`,
    sheets: [sheet("Stock", ["Item", "Category", "On Hand", "Reorder Point", "Action"], rows)], tables: [{ name: "Stock", sheet: "Stock", top: 0, left: 0, ncols: 5, nrows: n }],
    target: tgt("Stock", "E2", n), requireStructured: true,
    solution: '=IF([@[On Hand]]<[@[Reorder Point]],"Order","OK")',
    hints: ["Both column names have spaces: [@[On Hand]] and [@[Reorder Point]]", "Less than is <."],
  };
}, { drill: "logic" });

// ################################################################## PART 3
part(3, "Lookup Functions");

level("VLOOKUP: exact match", `
<p><code>=VLOOKUP(lookup_value, table_array, col_index_num, FALSE)</code></p>
<ul><li>Looks for the value in the <b>first column</b> of the table array, then returns the value from column number <i>col_index_num</i> (1 = the first column of the table array).</li>
<li><b>FALSE</b> means exact match. Use it for IDs, names, and products.</li>
<li>Lock the table array with $ (or use whole columns like <code>Products!A:C</code>) so it doesn't slide when copied down.</li></ul>`);

function productSheet(R, k) {
  const list = R.d.sample(PRODUCTS, k);
  return { list, s: sheet("Products", ["Product", "Category", "Price"], list.map((x) => [x[0], x[1], x[2]]), { Price: "money" }) };
}
function lineItems(R, n, list, extraCols) {
  const rows = ids(R.d, n, "", 20100, 20999).sort().map((id) => [+id, R.d.pick(list)[0], R.d.int(1, 12)].concat(extraCols.map(() => null)));
  return sheet("Orders", ["Order ID", "Product", "Qty"].concat(extraCols), rows);
}
add("p3-vl-price", "Look up the price", (R) => {
  const k = 10, { list, s } = productSheet(R, k), n = R.p.int(10, 14);
  const o = lineItems(R, n, list, ["Unit Price"]);
  return {
    prompt: `In <b>D2</b>, use VLOOKUP to bring in each product's <b>Price</b> from the <b>Products</b> sheet (Price is the 3rd column there).`,
    sheets: [o, s], target: tgt("Orders", "D2", n, null, "money"),
    solution: `=VLOOKUP(B2,Products!$A$2:$C$${k + 1},3,FALSE)`, require: ["VLOOKUP"],
    hints: ["Look up B2 in the Products table: Products!$A$2:$C$11 (or Products!A:C).", "Price is column 3, and use FALSE for an exact match."],
    mistakes: [{ f: `=VLOOKUP(B2,Products!$A$2:$C$${k + 1},3)`, msg: "Leaving off the last argument means approximate match, which gives wrong answers on unsorted lists. Add FALSE at the end." }, { f: `=VLOOKUP(B2,Products!A2:C${k + 1},3,FALSE)`, msg: "The first rows work, but the table array slides down as the formula is copied. Lock it with $ signs (F4)." }],
  };
}, { drill: "lookup" });
add("p3-vl-cat", "Look up the category", (R) => {
  const k = 10, { list, s } = productSheet(R, k), n = R.p.int(10, 14);
  const o = lineItems(R, n, list, ["Category"]);
  return {
    prompt: `In <b>D2</b>, use VLOOKUP to show each product's <b>Category</b> from the Products sheet.`,
    sheets: [o, s], target: tgt("Orders", "D2", n),
    solution: `=VLOOKUP(B2,Products!$A$2:$C$${k + 1},2,FALSE)`, require: ["VLOOKUP"],
    hints: ["Category is the 2nd column of the Products table.", "Remember FALSE and the $ signs."],
  };
}, { drill: "lookup" });
add("p3-vl-name", "Employee name by ID", (R) => {
  const k = 12, ppl = people(R.d, k), idl = ids(R.d, k, "", 100, 999).map(Number);
  const emp = sheet("Employees", ["Emp ID", "Name", "Hourly Rate"], ppl.map((p, i) => [idl[i], p.first + " " + p.last, R.d.pick([14, 15.5, 16, 17.25, 18, 19.5, 21])]), { "Hourly Rate": "money" });
  const n = R.p.int(10, 14);
  const ts = sheet("Timesheet", ["Week", "Emp ID", "Hours", "Name"], Array.from({ length: n }, (_, i) => [1 + Math.floor(i / 5), R.d.pick(idl), R.d.int(12, 40), null]));
  return {
    prompt: `In <b>D2</b> of the Timesheet, use VLOOKUP to show each employee's <b>Name</b> from the <b>Employees</b> sheet based on their Emp ID.`,
    sheets: [ts, emp], target: tgt("Timesheet", "D2", n),
    solution: `=VLOOKUP(B2,Employees!$A$2:$C$${k + 1},2,FALSE)`, require: ["VLOOKUP"],
    hints: ["The lookup value is the Emp ID in B2.", "Name is the 2nd column of the Employees table."],
  };
}, { drill: "lookup" });
add("p3-vl-pay", "VLOOKUP inside math", (R) => {
  const k = 12, ppl = people(R.d, k), idl = ids(R.d, k, "", 100, 999).map(Number);
  const emp = sheet("Employees", ["Emp ID", "Name", "Hourly Rate"], ppl.map((p, i) => [idl[i], p.first + " " + p.last, R.d.pick([14, 15.5, 16, 17.25, 18, 19.5, 21])]), { "Hourly Rate": "money" });
  const n = R.p.int(10, 14);
  const ts = sheet("Timesheet", ["Week", "Emp ID", "Hours", "Pay"], Array.from({ length: n }, (_, i) => [1 + Math.floor(i / 5), R.d.pick(idl), R.d.int(12, 40), null]));
  return {
    prompt: `In <b>D2</b>, calculate Pay: Hours × the employee's <b>Hourly Rate</b>, which you'll look up from the Employees sheet.`,
    sheets: [ts, emp], target: tgt("Timesheet", "D2", n, null, "money"),
    solution: `=C2*VLOOKUP(B2,Employees!$A$2:$C$${k + 1},3,FALSE)`, require: ["VLOOKUP"],
    hints: ["VLOOKUP can sit inside a bigger formula: =C2*VLOOKUP(...)", "Hourly Rate is the 3rd column."],
  };
}, { drill: "lookup" });

level("VLOOKUP: approximate match", `
<p>Use <b>TRUE</b> (approximate match) to find which <b>tier</b> or <b>bracket</b> a number falls into.</p>
<p><code>=VLOOKUP(C2, Tiers!$A$2:$B$4, 2, TRUE)</code></p>
<ul><li>The first column of the table holds each tier's <b>starting</b> value, sorted smallest to largest.</li>
<li>VLOOKUP finds the largest starting value that's ≤ the number you're looking up.</li></ul>`);

function tierProblem(R, cfg) {
  const n = R.p.int(10, 14);
  const tiers = sheet(cfg.tsheet, cfg.theads, cfg.tiers, cfg.tfmt);
  const rows = Array.from({ length: n }, () => cfg.row(R));
  const main = sheet(cfg.sheet, cfg.heads, rows, cfg.fmt);
  return {
    prompt: cfg.prompt,
    sheets: [main, tiers], target: tgt(cfg.sheet, cfg.cell, n, null, cfg.outfmt),
    solution: cfg.solution, require: ["VLOOKUP"],
    hints: cfg.hints,
    mistakes: [{ f: cfg.solution.replace(",TRUE)", ",FALSE)"), msg: "FALSE only finds exact matches, so most values give #N/A. For tiers, use TRUE (approximate match)." }],
  };
}
add("p3-va-bonus", "Holiday bonus tiers", (R) => {
  const b = R.p.pick([[100, 250, 500], [150, 300, 600], [75, 200, 400]]);
  return tierProblem(R, {
    tsheet: "Bonus", theads: ["Years", "Bonus"], tiers: [[0, b[0]], [5, b[1]], [10, b[2]]], tfmt: { Bonus: "money0" },
    sheet: "Staff", heads: ["Name", "Facility", "Years of Service", "Holiday Bonus"], fmt: { "Years of Service": "dec1" },
    row: (R2) => { const p = people(R2.d, 1)[0]; return [p.first + " " + p.last, R2.d.pick(FACILITIES.slice(0, 5)), R2.d.int(2, 160) / 10, null]; },
    cell: "D2", outfmt: "money0",
    prompt: `Holiday bonuses depend on years of service: under 5 years earns $${b[0]}, 5 to under 10 earns $${b[1]}, and 10 or more earns $${b[2]}. The tiers are on the <b>Bonus</b> sheet. In <b>D2</b>, look up each bonus with an approximate-match VLOOKUP.`,
    solution: "=VLOOKUP(C2,Bonus!$A$2:$B$4,2,TRUE)",
    hints: ["Look up the Years of Service (C2) in Bonus!$A$2:$B$4.", "Use TRUE at the end for approximate match."],
  });
}, { drill: "lookup" });
add("p3-va-ship", "Shipping rate by weight", (R) => {
  const t = R.p.pick([[[0, 6.5], [5, 11], [20, 19.5], [50, 42]], [[0, 4.25], [2, 7.5], [10, 14], [30, 27.5]]]);
  return tierProblem(R, {
    tsheet: "Rates", theads: ["Min Weight", "Rate"], tiers: t, tfmt: { Rate: "money" },
    sheet: "Packages", heads: ["Package", "Weight (lb)", "Rate"], fmt: {},
    row: (R2) => ["PK" + R2.d.int(100, 999), R2.d.int(5, t[3][0] * 15) / 10, null],
    cell: "C2", outfmt: "money",
    prompt: `The <b>Rates</b> sheet lists the starting weight of each shipping tier. In <b>C2</b>, find each package's rate with VLOOKUP (approximate match).`,
    solution: "=VLOOKUP(B2,Rates!$A$2:$B$5,2,TRUE)",
    hints: ["Look up the weight (B2) in Rates!$A$2:$B$5.", "Rate is column 2; TRUE for approximate."],
  });
}, { drill: "lookup" });
add("p3-va-discount", "Quantity discount", (R) => {
  const t = R.p.pick([[[1, 0], [10, 0.05], [25, 0.1], [100, 0.15]], [[1, 0], [12, 0.04], [50, 0.08], [144, 0.12]]]);
  return tierProblem(R, {
    tsheet: "Discounts", theads: ["Min Qty", "Discount"], tiers: t, tfmt: { Discount: "pct" },
    sheet: "Orders", heads: ["Order", "Qty", "Discount"], fmt: {},
    row: (R2) => [R2.d.int(30000, 39999), R2.d.int(1, 180), null],
    cell: "C2", outfmt: "pct",
    prompt: `Bigger orders get bigger discounts (see the <b>Discounts</b> sheet). In <b>C2</b>, look up each order's discount from its Qty.`,
    solution: "=VLOOKUP(B2,Discounts!$A$2:$B$5,2,TRUE)",
    hints: ["The table's first column is the minimum quantity of each tier.", "=VLOOKUP(B2, Discounts!$A$2:$B$5, 2, TRUE)"],
  });
}, { drill: "lookup" });
add("p3-va-grade", "Grade scale lookup", (R) => tierProblem(R, {
  tsheet: "Scale", theads: ["Min Score", "Grade"], tiers: [[0, "F"], [60, "D"], [70, "C"], [80, "B"], [90, "A"]],
  sheet: "Grades", heads: ["Student", "Average", "Letter"], fmt: {},
  row: (R2) => { const p = people(R2.d, 1)[0]; return [p.first + " " + p.last, R2.d.int(50, 100), null]; },
  cell: "C2",
  prompt: `Instead of nested IFs, use the grade scale on the <b>Scale</b> sheet. In <b>C2</b>, look up each student's letter grade with an approximate-match VLOOKUP.`,
  solution: "=VLOOKUP(B2,Scale!$A$2:$B$6,2,TRUE)",
  hints: ["Look up B2 in Scale!$A$2:$B$6.", "Column 2, TRUE."],
}), { drill: "lookup" });

level("INDEX / MATCH", `
<p>VLOOKUP can only look to the <b>right</b>. INDEX/MATCH can return a column on <b>either side</b>.</p>
<ul><li><code>MATCH(lookup_value, lookup_column, 0)</code> gives the <b>position</b> (row number in the range) where the value is found. 0 = exact match.</li>
<li><code>INDEX(return_column, position)</code> gives the value at that position.</li></ul>
<p><code>=INDEX(Names!$A$2:$A$16, MATCH(A2, Names!$C$2:$C$16, 0))</code></p>
<p>These problems are practice for INDEX/MATCH, so VLOOKUP isn't accepted here.</p>`);

function namesSheet(R, k) {
  const ppl = people(R.d, k), idl = ids(R.d, k, "", 1000, 9999).map(Number);
  return { idl, s: sheet("Names", ["First Name", "Last Name", "Employee ID"], ppl.map((p, i) => [p.first, p.last, idl[i]])) };
}
add("p3-im-first", "First name (INDEX/MATCH)", (R) => {
  const k = 15, { idl, s } = namesSheet(R, k), n = R.p.int(10, 13);
  const emps = R.d.sample(idl, n);
  const main = sheet("Employees", ["Employee ID", "First Name", "Facility"], emps.map((id) => [id, null, R.d.pick(FACILITIES.slice(0, 5))]));
  return {
    prompt: `On the <b>Names</b> sheet the Employee ID is in the <b>last</b> column, so VLOOKUP can't reach the names. In <b>B2</b>, use INDEX/MATCH to bring in each employee's <b>First Name</b>.`,
    sheets: [main, s], target: tgt("Employees", "B2", n), require: ["INDEX", "MATCH"], forbid: ["VLOOKUP"],
    solution: `=INDEX(Names!$A$2:$A$${k + 1},MATCH(A2,Names!$C$2:$C$${k + 1},0))`,
    hints: ["MATCH(A2, Names!$C$2:$C$16, 0) finds the row of the ID.", "INDEX(Names!$A$2:$A$16, that MATCH) returns the first name."],
    mistakes: [{ f: `=INDEX(Names!$A$2:$A$${k + 1},MATCH(A2,Names!$C$2:$C$${k + 1}))`, msg: "Without the 0 at the end, MATCH does an approximate match. Add 0 for an exact match." }],
  };
}, { drill: "lookup" });
add("p3-im-last", "Last name (INDEX/MATCH)", (R) => {
  const k = 15, { idl, s } = namesSheet(R, k), n = R.p.int(10, 13);
  const emps = R.d.sample(idl, n);
  const main = sheet("Employees", ["Employee ID", "Facility", "Last Name"], emps.map((id) => [id, R.d.pick(FACILITIES.slice(0, 5)), null]));
  return {
    prompt: `In <b>C2</b>, use INDEX/MATCH to bring in each employee's <b>Last Name</b> from the Names sheet, matching on Employee ID.`,
    sheets: [main, s], target: tgt("Employees", "C2", n), require: ["INDEX", "MATCH"], forbid: ["VLOOKUP"],
    solution: `=INDEX(Names!$B$2:$B$${k + 1},MATCH(A2,Names!$C$2:$C$${k + 1},0))`,
    hints: ["The return column is Last Name: Names!$B$2:$B$16.", "Match A2 against the Employee ID column with 0 for exact."],
  };
}, { drill: "lookup" });
add("p3-im-sku", "Price by SKU (INDEX/MATCH)", (R) => {
  const k = 12, list = R.d.sample(PRODUCTS, k), skus = ids(R.d, k, "SKU-", 100, 999);
  const cat = sheet("Catalog", ["Product Name", "Price", "SKU"], list.map((x, i) => [x[0], x[2], skus[i]]), { Price: "money" });
  const n = R.p.int(10, 13);
  const main = sheet("Cart", ["SKU", "Qty", "Price"], Array.from({ length: n }, () => [R.d.pick(skus), R.d.int(1, 6), null]));
  return {
    prompt: `In <b>C2</b>, use INDEX/MATCH to find each item's <b>Price</b> on the <b>Catalog</b> sheet by its SKU (the SKU is in the last column there).`,
    sheets: [main, cat], target: tgt("Cart", "C2", n, null, "money"), require: ["INDEX", "MATCH"], forbid: ["VLOOKUP"],
    solution: `=INDEX(Catalog!$B$2:$B$${k + 1},MATCH(A2,Catalog!$C$2:$C$${k + 1},0))`,
    hints: ["INDEX the Price column; MATCH the SKU column.", "Lock both ranges with $."],
  };
}, { drill: "lookup" });
add("p3-match-pos", "Leaderboard position (MATCH)", (R) => {
  const k = 10, ppl = people(R.d, k).map((p) => p.first + " " + p.last);
  const board = sheet("Leaderboard", ["Rep (ranked)", "Sales"], ppl.map((p, i) => [p, (k - i) * 9000 + R.d.int(0, 8000)]), { Sales: "money0" });
  const n = R.p.int(6, 9);
  const main = sheet("Lookup", ["Rep", "Rank"], R.d.sample(ppl, n).map((p) => [p, null]));
  return {
    prompt: `The <b>Leaderboard</b> lists reps from 1st place down. In <b>B2</b>, use MATCH to find each rep's rank (their position in the list).`,
    sheets: [main, board], target: tgt("Lookup", "B2", n, null, "int"), require: ["MATCH"],
    solution: `=MATCH(A2,Leaderboard!$A$2:$A$${k + 1},0)`,
    hints: ["MATCH returns the position, which is exactly the rank here.", "=MATCH(A2, Leaderboard!$A$2:$A$11, 0)"],
  };
}, { drill: "lookup" });

level("INDEX / MATCH / MATCH", `
<p>For a two-way table (row labels down the side, column labels across the top), use two MATCHes:</p>
<p><code>=INDEX(table_values, MATCH(row_item, row_labels, 0), MATCH(col_item, col_labels, 0))</code></p>
<p>The table values don't include the labels. The row labels and column labels must line up with them.</p>`);

add("p3-imm-rate", "Rate by zone and weight", (R) => {
  const classes = ["A", "B", "C", "D", "E"], zones = ["Zone 1", "Zone 2", "Zone 3", "Zone 4", "Zone 5"];
  const rates = classes.map((c, i) => [c].concat(zones.map((z, j) => Math.round((5 + i * 3.5 + j * 2.25 + R.d.int(0, 100) / 100) * 100) / 100)));
  const tbl = sheet("Rates", ["Class"].concat(zones), rates);
  for (let i = 1; i <= 5; i++) for (let j = 1; j <= 5; j++) tbl.fmt[i + ":" + j] = "money";
  const n = R.p.int(10, 13);
  const main = sheet("Shipments", ["Shipment", "Weight Class", "Zone", "Rate"], Array.from({ length: n }, () => ["SH" + R.d.int(1000, 9999), R.d.pick(classes), R.d.pick(zones), null]));
  return {
    prompt: `The <b>Rates</b> sheet has weight classes down column A and zones across row 1. In <b>D2</b>, use INDEX/MATCH/MATCH to find each shipment's rate.`,
    sheets: [main, tbl], target: tgt("Shipments", "D2", n, null, "money"), require: ["INDEX", "MATCH"],
    solution: "=INDEX(Rates!$B$2:$F$6,MATCH(B2,Rates!$A$2:$A$6,0),MATCH(C2,Rates!$B$1:$F$1,0))",
    hints: ["The values are Rates!$B$2:$F$6. Row labels: Rates!$A$2:$A$6. Column labels: Rates!$B$1:$F$1.", "=INDEX(values, MATCH(B2, row labels, 0), MATCH(C2, column labels, 0))"],
  };
}, { drill: "lookup" });
add("p3-imm-size", "Price by style and size", (R) => {
  const styles = ["Crew Tee", "Polo", "Hoodie", "Quarter Zip", "Rain Shell"], sizes = ["S", "M", "L", "XL"];
  const vals = styles.map((st, i) => [st].concat(sizes.map((z, j) => 14 + i * 9 + j * 2 + R.d.pick([0, 0.5, 0.99]))));
  const tbl = sheet("Prices", ["Style"].concat(sizes), vals);
  for (let i = 1; i <= 5; i++) for (let j = 1; j <= 4; j++) tbl.fmt[i + ":" + j] = "money";
  const n = R.p.int(10, 13);
  const main = sheet("Orders", ["Order", "Style", "Size", "Price"], Array.from({ length: n }, () => [R.d.int(800, 999), R.d.pick(styles), R.d.pick(sizes), null]));
  return {
    prompt: `In <b>D2</b>, use INDEX/MATCH/MATCH to look up each item's price on the <b>Prices</b> sheet by Style (down the side) and Size (across the top).`,
    sheets: [main, tbl], target: tgt("Orders", "D2", n, null, "money"), require: ["INDEX", "MATCH"],
    solution: "=INDEX(Prices!$B$2:$E$6,MATCH(B2,Prices!$A$2:$A$6,0),MATCH(C2,Prices!$B$1:$E$1,0))",
    hints: ["Values: Prices!$B$2:$E$6", "Match the style in column A and the size in row 1."],
  };
}, { drill: "lookup" });
add("p3-imm-box", "Lookup box (one cell)", (R) => {
  const stores = ["Tuscaloosa", "Hoover", "Huntsville", "Mobile"];
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const rows = months.map((m) => [m].concat(stores.map(() => R.d.int(20, 90) * 1000)));
  const s = sheet("Sales", ["Month"].concat(stores), rows, null, [["G1", "Month"], ["H1", R.d.pick(months)], ["G2", "Store"], ["H2", R.d.pick(stores)], ["G3", "Sales"]]);
  for (let i = 1; i <= 12; i++) for (let j = 1; j <= 4; j++) s.fmt[i + ":" + j] = "money0";
  return {
    prompt: `Build a lookup box: in <b>H3</b>, return the sales for the month in <b>H1</b> and the store in <b>H2</b> using INDEX/MATCH/MATCH.`,
    sheets: [s], target: tgt("Sales", "H3", 1, null, "money0"), require: ["INDEX", "MATCH"],
    solution: "=INDEX(B2:E13,MATCH(H1,A2:A13,0),MATCH(H2,B1:E1,0))",
    hints: ["Values B2:E13, month labels A2:A13, store labels B1:E1.", "=INDEX(B2:E13, MATCH(H1, A2:A13, 0), MATCH(H2, B1:E1, 0))"],
  };
}, { drill: "lookup" });

level("Lookups with IF & IFERROR", `
<ul><li><code>=IFERROR(VLOOKUP(...), "Not Found")</code> replaces #N/A with a friendly message.</li>
<li><code>=IF(J2="Yes", VLOOKUP(...), "")</code> only looks something up when a condition is met.</li>
<li>Inside Excel Tables you can use the <b>table name</b> as the table array: <code>=VLOOKUP([@Product], Products, 3, FALSE)</code>.</li></ul>`);

add("p3-ie-notfound", "Not Found (IFERROR + VLOOKUP)", (R) => {
  const k = 10, { list, s } = productSheet(R, k), n = R.p.int(10, 13);
  const others = PRODUCTS.filter((x) => !list.includes(x));
  const rows = ids(R.d, n, "", 20100, 20999).sort().map((id) => [+id, R.d.chance(0.25) ? R.d.pick(others)[0] : R.d.pick(list)[0], null]);
  rows[1][1] = others[0][0];
  const o = sheet("Orders", ["Order ID", "Product", "Price"], rows);
  return {
    prompt: `Some products aren't on the Products sheet yet. In <b>C2</b>, look up each Price (3rd column of Products) and show <b>"Not Found"</b> instead of #N/A.`,
    sheets: [o, s], target: tgt("Orders", "C2", n, null, "money"), require: ["IFERROR", "VLOOKUP"],
    solution: `=IFERROR(VLOOKUP(B2,Products!$A$2:$C$${k + 1},3,FALSE),"Not Found")`,
    hints: ["Wrap the whole VLOOKUP: =IFERROR(VLOOKUP(...), \"Not Found\")", "VLOOKUP(B2, Products!$A$2:$C$11, 3, FALSE)"],
  };
}, { drill: "lookup" });
add("p3-if-vl", "Vacation days if eligible", (R) => {
  const facs = FACILITIES.slice(0, 5), days = R.p.sample([3, 5, 6, 7, 9, 10], 5);
  const vac = sheet("Data Tables", ["Facility", "Vacation Days"], facs.map((f, i) => [f, days[i]]));
  const n = R.p.int(10, 13), ppl = people(R.d, n);
  const main = sheet("Staff", ["Name", "Facility", "Eligible?", "# of Vacation Days"], ppl.map((p) => [p.first + " " + p.last, R.d.pick(facs), R.d.pick(["Yes", "Yes", ""]), null]));
  main.rows.forEach((r) => { if (r[2] === "") r[2] = null; });
  return {
    prompt: `If the employee is eligible (C = "Yes"), look up their facility's vacation days on the <b>Data Tables</b> sheet. Otherwise leave the cell blank (""). Fill in <b>D2</b>.`,
    sheets: [main, vac], target: tgt("Staff", "D2", n, null, "int"), require: ["IF", "VLOOKUP"],
    solution: `=IF(C2="Yes",VLOOKUP(B2,'Data Tables'!$A$2:$B$6,2,FALSE),"")`,
    hints: ["=IF(C2=\"Yes\", VLOOKUP(...), \"\")", "The sheet name has a space, so it's written 'Data Tables'!$A$2:$B$6 with single quotes. (Clicking the tab while typing adds them for you.)"],
  };
}, { drill: "lookup" });
add("p3-tbl-vl", "VLOOKUP with Tables", (R) => {
  const k = 10, list = R.d.sample(PRODUCTS, k);
  const prod = sheet("Products", ["Product", "Category", "Price"], list.map((x) => [x[0], x[1], x[2]]), { Price: "money" });
  const n = R.p.int(10, 13);
  const sales = sheet("Sales", ["Order ID", "Product", "Qty", "Price"], ids(R.d, n, "", 40100, 40999).sort().map((id) => [+id, R.d.pick(list)[0], R.d.int(1, 9), null]));
  return {
    prompt: `Both sheets are Excel Tables: <b>Sales</b> and <b>Products</b>. In <b>D2</b>, use VLOOKUP with structured references to bring in the Price: look up <code>[@Product]</code> in the table <code>Products</code>.`,
    sheets: [sales, prod], target: tgt("Sales", "D2", n, null, "money"), requireStructured: true, require: ["VLOOKUP"],
    tables: [{ name: "Sales", sheet: "Sales", top: 0, left: 0, ncols: 4, nrows: n }, { name: "Products", sheet: "Products", top: 0, left: 0, ncols: 3, nrows: k }],
    solution: "=VLOOKUP([@Product],Products,3,FALSE)",
    hints: ["The whole table can be the table array: just type Products.", "=VLOOKUP([@Product], Products, 3, FALSE)"],
  };
}, { drill: "lookup" });
add("p3-ie-im", "Check ID (IFERROR + INDEX/MATCH)", (R) => {
  const k = 15, { idl, s } = namesSheet(R, k), n = R.p.int(10, 13);
  const emps = Array.from({ length: n }, () => (R.d.chance(0.2) ? R.d.int(1000, 9999) : R.d.pick(idl)));
  emps[2] = 1000 + ((idl[0] + 17) % 9000);
  while (idl.includes(emps[2])) emps[2]++;
  const main = sheet("Badges", ["Employee ID", "First Name"], emps.map((id) => [id, null]));
  return {
    prompt: `Some badge IDs were typed wrong. In <b>B2</b>, use INDEX/MATCH to get the First Name from the Names sheet, and show <b>"Check ID"</b> when the ID isn't found.`,
    sheets: [main, s], target: tgt("Badges", "B2", n), require: ["IFERROR", "INDEX", "MATCH"], forbid: ["VLOOKUP"],
    solution: `=IFERROR(INDEX(Names!$A$2:$A$${k + 1},MATCH(A2,Names!$C$2:$C$${k + 1},0)),"Check ID")`,
    hints: ["Start with the INDEX/MATCH, then wrap it in IFERROR.", "=IFERROR(INDEX(...,MATCH(...,0)), \"Check ID\")"],
  };
}, { drill: "lookup" });

// ################################################################## PART 4
part(4, "Dates & Rounding");

level("YEAR, MONTH, DAY", `
<p>Excel stores dates as numbers (days since 1/1/1900), so you can pull pieces out of them:</p>
<ul><li><code>=YEAR(B2)</code> gives 2026, <code>=MONTH(B2)</code> gives 1–12, and <code>=DAY(B2)</code> gives 1–31.</li></ul>
<p>They work well inside IF: <code>=IF(MONTH(B2)&lt;=3,"Q1","")</code></p>`);

function datedOrders(R, n, extra) {
  const base = D(R.d.pick([2023, 2024, 2025]), R.d.int(1, 12), 1);
  const rows = ids(R.d, n, "", 60100, 60999).sort().map((id) => [+id, base + R.d.int(0, 700), R.d.money(20, 900)].concat(extra.map(() => null)));
  rows.sort((a, b) => a[1] - b[1]);
  return sheet("Orders", ["Order ID", "Order Date", "Amount"].concat(extra), rows, { "Order Date": "date", Amount: "money" });
}
add("p4-year", "Order year", (R) => {
  const n = R.p.int(10, 13);
  return { prompt: "In <b>D2</b>, pull the <b>year</b> out of each Order Date.", sheets: [datedOrders(R, n, ["Year"])], target: tgt("Orders", "D2", n), solution: "=YEAR(B2)", hints: ["=YEAR(date)", "=YEAR(B2)"] };
}, { drill: "dates" });
add("p4-month", "Order month", (R) => {
  const n = R.p.int(10, 13);
  return { prompt: "In <b>D2</b>, show the <b>month number</b> (1–12) of each Order Date.", sheets: [datedOrders(R, n, ["Month"])], target: tgt("Orders", "D2", n), solution: "=MONTH(B2)", hints: ["=MONTH(date)", "=MONTH(B2)"] };
}, { drill: "dates" });
add("p4-day", "Day of the month", (R) => {
  const n = R.p.int(10, 13);
  return { prompt: "In <b>D2</b>, show the <b>day of the month</b> (1–31) for each Order Date.", sheets: [datedOrders(R, n, ["Day"])], target: tgt("Orders", "D2", n), solution: "=DAY(B2)", hints: ["=DAY(date)", "=DAY(B2)"] };
}, { drill: "dates" });
add("p4-q", "Which half of the year", (R) => {
  const n = R.p.int(10, 13);
  return {
    prompt: `In <b>D2</b>, show <b>"H1"</b> if the order was placed in January–June (month 6 or earlier), otherwise <b>"H2"</b>.`,
    sheets: [datedOrders(R, n, ["Half"])], target: tgt("Orders", "D2", n), solution: '=IF(MONTH(B2)<=6,"H1","H2")', require: ["MONTH"],
    hints: ["MONTH(B2) gives the month number.", '=IF(MONTH(B2)<=6, "H1", "H2")'],
  };
}, { drill: "dates" });

level("TODAY", `
<p><code>=TODAY()</code> returns today's date and updates every day the file is opened. (NOW() also includes the time.)</p>
<ul><li>Subtracting dates gives the number of days between them: <code>=C2-TODAY()</code>.</li>
<li><code>=IF(MONTH(C2)=MONTH(TODAY()),"Yes","")</code> checks whether a date is in the current month, whatever month it is when the file is opened.</li></ul>`);

add("p4-days-until", "Days until due", (R) => {
  const n = R.p.int(10, 13), t = todaySerial();
  const rows = ids(R.d, n, "INV-", 1000, 1999).map((id) => [id, R.d.money(100, 5000), t + R.d.int(1, 90)]);
  const s = sheet("Invoices", ["Invoice", "Amount", "Due Date", "Days Until Due"], rows, { Amount: "money", "Due Date": "date" });
  return {
    prompt: `In <b>D2</b>, find how many days remain until each invoice is due (Due Date minus today). Use TODAY() so it stays current.`,
    sheets: [s], target: tgt("Invoices", "D2", n, null, "int"), solution: "=C2-TODAY()", require: ["TODAY"],
    hints: ["Subtract today's date from the due date.", "=C2-TODAY()"],
  };
}, { drill: "dates" });
add("p4-bday", "Birthday this month", (R) => {
  const n = R.p.int(10, 13), ppl = people(R.d, n), m = thisMonth();
  const rows = ppl.map((p) => [p.first + " " + p.last, D(R.d.int(1965, 2004), R.d.chance(0.35) ? m : R.d.int(1, 12), R.d.int(1, 28))]);
  rows[0][1] = D(1990 + R.d.int(0, 10), m, R.d.int(1, 28));
  const s = sheet("Staff", ["Name", "Birth Date", "Birthday?"], rows, { "Birth Date": "date" });
  return {
    prompt: `In <b>C2</b>, show <b>"Yes"</b> if the employee's birthday is in the <b>current month</b>, otherwise leave it blank (""). It should still work next month without editing, so use TODAY().`,
    sheets: [s], target: tgt("Staff", "C2", n), solution: '=IF(MONTH(B2)=MONTH(TODAY()),"Yes","")', require: ["TODAY", "MONTH"],
    hints: ["Compare the birth month to this month: MONTH(B2)=MONTH(TODAY())", '=IF(MONTH(B2)=MONTH(TODAY()), "Yes", "")'],
  };
}, { drill: "dates" });
add("p4-overdue", "Overdue?", (R) => {
  const n = R.p.int(10, 13), t = todaySerial();
  const rows = ids(R.d, n, "WO-", 500, 999).map((id) => [id, t + R.d.int(-40, 40)]);
  const s = sheet("Work Orders", ["Work Order", "Due Date", "Status"], rows, { "Due Date": "date" });
  return {
    prompt: `In <b>C2</b>, show <b>"Overdue"</b> if the Due Date is <b>before today</b>, otherwise leave it blank ("").`,
    sheets: [s], target: tgt("Work Orders", "C2", n), solution: '=IF(B2<TODAY(),"Overdue","")', require: ["TODAY"],
    hints: ["Earlier dates are smaller numbers, so before today means B2<TODAY().", '=IF(B2<TODAY(), "Overdue", "")'],
  };
}, { drill: "dates" });

level("ROUND", `
<p><code>=ROUND(number, digits)</code></p>
<ul><li><code>=ROUND(B2,2)</code> rounds to 2 decimal places (cents). <code>=ROUND(B2,0)</code> rounds to a whole number.</li>
<li>Negative digits round to the left of the decimal point: <code>=ROUND(B2,-2)</code> rounds to the nearest hundred.</li>
<li>You can round a whole calculation: <code>=ROUND((TODAY()-C2)/365,1)</code>.</li></ul>
<p>Formatting a cell only changes how it <i>looks</i>. ROUND changes the actual value.</p>`);

add("p4-round-unit", "Unit cost to the cent", (R) => {
  const n = R.p.int(10, 13), items = R.d.sample(PRODUCTS, n);
  const rows = items.map((x) => [x[0], R.d.money(100, 2000), R.d.int(3, 97)]);
  const s = sheet("Costs", ["Item", "Batch Cost", "Units", "Unit Cost"], rows, { "Batch Cost": "money" });
  return {
    prompt: `In <b>D2</b>, calculate the unit cost (Batch Cost ÷ Units) <b>rounded to 2 decimal places</b>.`,
    sheets: [s], target: tgt("Costs", "D2", n, null, "dec2"), solution: "=ROUND(B2/C2,2)", require: ["ROUND"],
    hints: ["Put the division inside ROUND.", "=ROUND(B2/C2, 2)"],
  };
}, { drill: "dates" });
add("p4-round-yos", "Years of service", (R) => {
  const n = R.p.int(10, 13), ppl = people(R.d, n), t = todaySerial();
  const rows = ppl.map((p) => [p.first + " " + p.last, t - R.d.int(60, 6000)]);
  const s = sheet("Staff", ["Name", "Hire Date", "Years of Service"], rows, { "Hire Date": "date" });
  return {
    prompt: `In <b>C2</b>, calculate each employee's years of service: (today − Hire Date) ÷ 365, <b>rounded to 1 decimal place</b>.`,
    sheets: [s], target: tgt("Staff", "C2", n, null, "dec1"), solution: "=ROUND((TODAY()-B2)/365,1)", require: ["ROUND", "TODAY"],
    hints: ["Days worked: TODAY()-B2. Divide by 365, with parentheses around the subtraction.", "=ROUND((TODAY()-B2)/365, 1)"],
    mistakes: [{ f: "=ROUND(TODAY()-B2/365,1)", msg: "Order of operations! Division happens before subtraction, so put parentheses around TODAY()-B2." }],
  };
}, { drill: "dates" });
add("p4-round-hundred", "Round to the nearest hundred", (R) => {
  const n = R.p.int(8, 11);
  const depts = R.d.sample(["Receiving", "Shipping", "Quality", "Maintenance", "Assembly", "Packaging", "Purchasing", "Planning", "Safety", "Warehouse", "IT"], n);
  const s = sheet("Budget", ["Department", "Request", "Rounded"], depts.map((d) => [d, R.d.int(1200, 98000) + R.d.int(0, 99) / 100]), { Request: "money" });
  return {
    prompt: `In <b>C2</b>, round each budget request to the <b>nearest hundred dollars</b>.`,
    sheets: [s], target: tgt("Budget", "C2", n, null, "money0"), solution: "=ROUND(B2,-2)", require: ["ROUND"],
    hints: ["Negative digits round to the left of the decimal point.", "=ROUND(B2, -2)"],
  };
}, { drill: "dates" });
add("p4-round-avg", "Rounded average", (R) => {
  const n = 14, ppl = people(R.d, n);
  const s = sheet("Ratings", ["Customer", "Rating"], ppl.map((p) => [p.first + " " + p.last, R.d.int(1, 10)]), null, [["D1", "Avg Rating"]]);
  return {
    prompt: `In <b>D2</b>, find the average Rating <b>rounded to 1 decimal place</b>.`,
    sheets: [s], target: tgt("Ratings", "D2", 1, null, "dec1"), solution: `=ROUND(AVERAGE(B2:B${n + 1}),1)`, require: ["ROUND", "AVERAGE"],
    hints: ["Put AVERAGE inside ROUND.", `=ROUND(AVERAGE(B2:B${n + 1}), 1)`],
  };
}, { drill: "dates" });

// ################################################################## PART 5
part(5, "Text Functions");

level("LEFT, RIGHT, LEN", `
<ul><li><code>=LEFT(text, n)</code> gives the first n characters, and <code>=RIGHT(text, n)</code> gives the last n.</li>
<li><code>=LEN(text)</code> counts the characters, including spaces.</li></ul>
<p><code>=LEFT("35401-2231",5)</code> gives "35401".</p>`);

function customers(R, n) {
  const ppl = people(R.d, n);
  return ppl.map((p) => {
    const c = R.d.pick(CITIES);
    const zip = String(R.d.int(30000, 39999)) + "-" + String(R.d.int(1000, 9999));
    const area = R.d.pick(["205", "256", "251", "334", "659", "938"]);
    return { first: p.first, last: p.last, city: c[0], st: c[1], zip, phone: "1-" + area + R.d.int(2000000, 9999999), email: (p.first + "." + p.last).toLowerCase() + R.d.pick(["@gmail.com", "@crimson.ua.edu", "@yahoo.com", "@outlook.com"]) };
  });
}
add("p5-zip5", "5-digit ZIP", (R) => {
  const n = R.p.int(10, 13), c = customers(R, n);
  const s = sheet("Customers", ["Name", "ZIP+4", "ZIP"], c.map((x) => [x.first + " " + x.last, x.zip, null]));
  return { prompt: `In <b>C2</b>, keep only the first 5 digits of each ZIP+4 code.`, sheets: [s], target: tgt("Customers", "C2", n), solution: "=LEFT(B2,5)", require: ["LEFT"], hints: ["You want characters from the left side.", "=LEFT(B2, 5)"] };
}, { drill: "text" });
add("p5-last4", "Last 4 digits", (R) => {
  const n = R.p.int(10, 13), c = customers(R, n);
  const s = sheet("Customers", ["Name", "Phone", "Last 4"], c.map((x) => [x.first + " " + x.last, x.phone, null]));
  return { prompt: `In <b>C2</b>, show the <b>last 4 digits</b> of each phone number.`, sheets: [s], target: tgt("Customers", "C2", n), solution: "=RIGHT(B2,4)", require: ["RIGHT"], hints: ["Characters from the right side.", "=RIGHT(B2, 4)"] };
}, { drill: "text" });
add("p5-len", "Code length check", (R) => {
  const n = R.p.int(10, 13);
  const rows = Array.from({ length: n }, () => { const k = R.d.pick([8, 8, 8, 7, 9]); let s = ""; for (let i = 0; i < k; i++) s += "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"[R.d.int(0, 31)]; return [s, null]; });
  const s = sheet("Codes", ["Promo Code", "Length"], rows);
  return { prompt: `Promo codes should be 8 characters. In <b>B2</b>, count the characters in each code.`, sheets: [s], target: tgt("Codes", "B2", n), solution: "=LEN(A2)", require: ["LEN"], hints: ["LEN counts characters.", "=LEN(A2)"] };
}, { drill: "text" });
add("p5-dept", "Department prefix", (R) => {
  const n = R.p.int(10, 13);
  const pre = ["CMP", "WTR", "FTW", "APP", "GER"];
  const rows = Array.from({ length: n }, () => [R.d.pick(pre) + R.d.int(10000, 99999), null]);
  const s = sheet("Items", ["Item Code", "Dept"], rows);
  return { prompt: `The first 3 letters of each item code are the department. In <b>B2</b>, pull them out.`, sheets: [s], target: tgt("Items", "B2", n), solution: "=LEFT(A2,3)", require: ["LEFT"], hints: ["The first 3 characters.", "=LEFT(A2, 3)"] };
}, { drill: "text" });

level("UPPER, LOWER, PROPER, TRIM", `
<ul><li><code>=UPPER(A2)</code> makes text ALL CAPS, and <code>=LOWER(A2)</code> makes it all lowercase.</li>
<li><code>=PROPER(A2)</code> Capitalizes The First Letter Of Each Word.</li>
<li><code>=TRIM(A2)</code> removes extra spaces at the start and end, and squeezes double spaces down to one.</li></ul>
<p>You can stack them: <code>=UPPER(TRIM(A2))</code>.</p>`);

add("p5-proper", "Proper case names", (R) => {
  const n = R.p.int(10, 13), c = customers(R, n);
  const s = sheet("Customers", ["FIRST", "LAST", "First Name"], c.map((x) => [x.first.toUpperCase(), x.last.toUpperCase(), null]));
  return { prompt: `The names came in ALL CAPS. In <b>C2</b>, convert the first name to proper case (like "Mary").`, sheets: [s], target: tgt("Customers", "C2", n), solution: "=PROPER(A2)", require: ["PROPER"], hints: ["PROPER capitalizes the first letter of each word.", "=PROPER(A2)"] };
}, { drill: "text" });
add("p5-upper-trim", "Clean state codes", (R) => {
  const n = R.p.int(10, 13), c = customers(R, n);
  const s = sheet("Customers", ["City", "State (raw)", "State"], c.map((x) => [x.city, R.d.pick([" ", "", "  "]) + x.st.toLowerCase() + R.d.pick(["", " "]), null]));
  s.rows[1][1] = " " + c[0].st.toLowerCase();
  return { prompt: `After Text to Columns, the state codes look like <code>" al"</code>: lowercase with extra spaces. In <b>C2</b>, clean them up so they look like <b>AL</b>.`, sheets: [s], target: tgt("Customers", "C2", n), solution: "=UPPER(TRIM(B2))", require: ["UPPER", "TRIM"], hints: ["TRIM removes the spaces; UPPER makes it capital.", "=UPPER(TRIM(B2))"] };
}, { drill: "text" });
add("p5-lower", "Lowercase emails", (R) => {
  const n = R.p.int(10, 13), c = customers(R, n);
  const mess = (e) => e.split("").map((ch) => (R.d.chance(0.3) ? ch.toUpperCase() : ch)).join("");
  const s = sheet("Customers", ["Name", "Email (raw)", "Email"], c.map((x) => [x.first + " " + x.last, mess(x.email), null]));
  return { prompt: `Emails were typed with random capital letters. In <b>C2</b>, make each one all lowercase.`, sheets: [s], target: tgt("Customers", "C2", n), solution: "=LOWER(B2)", require: ["LOWER"], hints: ["The opposite of UPPER.", "=LOWER(B2)"] };
}, { drill: "text" });
add("p5-trim", "Remove extra spaces", (R) => {
  const n = R.p.int(10, 13), items = R.d.sample(PRODUCTS, n);
  const sp = () => R.d.pick(["", " ", "  ", "   "]);
  const s = sheet("Items", ["Product (raw)", "Product"], items.map((x) => [sp() + x[0].replace(" ", R.d.pick([" ", "  ", "   "])) + sp(), null]));
  s.rows[1][0] = "  " + items[0][0] + " ";
  return { prompt: `Product names have extra spaces in front, behind, and between words. In <b>B2</b>, clean them up.`, sheets: [s], target: tgt("Items", "B2", n), solution: "=TRIM(A2)", require: ["TRIM"], hints: ["One function handles all of the extra spaces.", "=TRIM(A2)"] };
}, { drill: "text" });

level("Joining text: CONCATENATE and &", `
<p>Two ways to glue text together:</p>
<ul><li><code>=CONCATENATE(A2," ",B2)</code></li>
<li><code>=A2&amp;" "&amp;B2</code> (the &amp; sign does the same thing)</li></ul>
<p>Remember to add the spaces or punctuation yourself, in quotes. You can wrap the result too: <code>=PROPER(CONCATENATE(A2," ",B2))</code>.</p>`);

add("p5-full", "Full name", (R) => {
  const n = R.p.int(10, 13), c = customers(R, n);
  const s = sheet("Customers", ["First", "Last", "Full Name"], c.map((x) => [x.first, x.last, null]));
  return { prompt: `In <b>C2</b>, join the first and last name with a space between them, like "Mary Smith".`, sheets: [s], target: tgt("Customers", "C2", n), solution: '=A2&" "&B2', hints: ["Don't forget the space in quotes.", '=A2&" "&B2  or  =CONCATENATE(A2," ",B2)'] };
}, { drill: "text" });
add("p5-full-proper", "Full name, proper case", (R) => {
  const n = R.p.int(10, 13), c = customers(R, n);
  const s = sheet("Customers", ["FIRST", "LAST", "Full Name"], c.map((x) => [x.first.toUpperCase(), x.last.toUpperCase(), null]));
  return { prompt: `The names are in ALL CAPS. In <b>C2</b>, make a full name in proper case, like "Mary Smith".`, sheets: [s], target: tgt("Customers", "C2", n), solution: '=PROPER(CONCATENATE(A2," ",B2))', require: ["PROPER"], hints: ["Join them first, then PROPER the result (or PROPER each piece).", '=PROPER(CONCATENATE(A2," ",B2))'] };
}, { drill: "text" });
add("p5-lastfirst", "Last, First", (R) => {
  const n = R.p.int(10, 13), c = customers(R, n);
  const s = sheet("Roster", ["First", "Last", "Sort Name"], c.map((x) => [x.first, x.last, null]));
  return { prompt: `In <b>C2</b>, write each name as <b>Last, First</b>, like "Smith, Mary" (comma and space).`, sheets: [s], target: tgt("Roster", "C2", n), solution: '=B2&", "&A2', hints: ["Last name first, then \", \", then first name.", '=B2&", "&A2'] };
}, { drill: "text" });
add("p5-code", "Build a code", (R) => {
  const n = R.p.int(10, 13);
  const rows = Array.from({ length: n }, () => [R.d.int(100, 999), R.d.pick(REGIONS), null]);
  const s = sheet("Accounts", ["Account #", "Region", "Code"], rows);
  return { prompt: `Account codes are the first 3 letters of the region in capitals, a dash, then the account number, like <b>EAS-417</b>. Build them in <b>C2</b>.`, sheets: [s], target: tgt("Accounts", "C2", n), solution: '=UPPER(LEFT(B2,3))&"-"&A2', require: ["UPPER", "LEFT"], hints: ["UPPER(LEFT(B2,3)) gives the letters.", '=UPPER(LEFT(B2,3))&"-"&A2'] };
}, { drill: "text" });

level("SUBSTITUTE & VALUE", `
<ul><li><code>=SUBSTITUTE(text, old, new)</code> replaces every <i>old</i> with <i>new</i>. Use <code>""</code> as <i>new</i> to delete it.</li>
<li><code>=VALUE(text)</code> turns text that looks like a number into a real number (so it lines up right and you can do math with it).</li></ul>
<p><code>=VALUE(SUBSTITUTE("1-2055550142","1-",""))</code> gives 2055550142 as a number.</p>
<p>In this level, answers that are text-that-looks-like-a-number aren't accepted where a real number is asked for.</p>`, { strictType: true });

add("p5-phone", "Phone as a number", (R) => {
  const n = R.p.int(10, 13), c = customers(R, n);
  const s = sheet("Customers", ["Name", "Phone (raw)", "Phone"], c.map((x) => [x.first + " " + x.last, x.phone, null]));
  return { prompt: `Phones are stored as text like <code>1-2055550142</code>. In <b>C2</b>, remove the leading <b>"1-"</b> and turn the result into a <b>number</b>.`, sheets: [s], target: tgt("Customers", "C2", n, null, "phone"), solution: '=VALUE(SUBSTITUTE(B2,"1-",""))', require: ["VALUE", "SUBSTITUTE"], strictType: true, hints: ['SUBSTITUTE(B2,"1-","") removes the 1-.', "Wrap it in VALUE to make it a number."] };
}, { drill: "text" });
add("p5-zipnum", "ZIP as a number", (R) => {
  const n = R.p.int(10, 13), c = customers(R, n);
  const s = sheet("Customers", ["Name", "ZIP+4", "ZIP"], c.map((x) => [x.first + " " + x.last, x.zip, null]));
  return { prompt: `In <b>C2</b>, take the first 5 characters of the ZIP+4 and make them a real <b>number</b>.`, sheets: [s], target: tgt("Customers", "C2", n), solution: "=VALUE(LEFT(B2,5))", require: ["VALUE", "LEFT"], strictType: true, hints: ["LEFT gets the 5 digits as text.", "=VALUE(LEFT(B2,5))"] };
}, { drill: "text" });
add("p5-street", "Spell out Street", (R) => {
  const n = R.p.int(10, 13);
  const names = ["Main", "Oak", "Hackberry", "University", "Paul Bryant", "Greensboro", "Jack Warner", "Bryant", "Queen City", "Lurleen Wallace", "McFarland"];
  const rows = Array.from({ length: n }, () => [R.d.int(100, 9999) + " " + R.d.pick(names) + R.d.pick([" St.", " St.", " Ave.", " Blvd."]), null]);
  rows[0][0] = "220 " + names[R.d.int(0, 10)] + " St.";
  const s = sheet("Addresses", ["Address (raw)", "Address"], rows);
  return { prompt: `In <b>B2</b>, replace <b>"St."</b> with <b>"Street"</b> in each address (leave other addresses unchanged).`, sheets: [s], target: tgt("Addresses", "B2", n), solution: '=SUBSTITUTE(A2,"St.","Street")', require: ["SUBSTITUTE"], hints: ["SUBSTITUTE(text, old, new)", '=SUBSTITUTE(A2, "St.", "Street")'] };
}, { drill: "text" });
add("p5-price", "Price text to number", (R) => {
  const n = R.p.int(10, 13), items = R.d.sample(PRODUCTS, n);
  const s = sheet("Import", ["Product", "Price (text)", "Price"], items.map((x) => [x[0], "$" + x[2].toFixed(2) + " USD", null]));
  return { prompt: `Prices were imported as text like <code>$64.99 USD</code>. In <b>C2</b>, remove <b>" USD"</b> (with the space) and convert the rest to a <b>number</b>.`, sheets: [s], target: tgt("Import", "C2", n, null, "money"), solution: '=VALUE(SUBSTITUTE(B2," USD",""))', require: ["VALUE", "SUBSTITUTE"], strictType: true, hints: ['SUBSTITUTE(B2," USD","") leaves "$64.99".', "VALUE understands the $ sign."] };
}, { drill: "text" });

level("SEARCH", `
<p><code>=SEARCH(find_text, within_text)</code> gives the <b>position</b> of find_text (1 = the first character). It isn't case-sensitive. If the text isn't found, you get #VALUE!.</p>
<ul><li>Everything before the @: <code>=LEFT(C2, SEARCH("@",C2)-1)</code></li>
<li>Everything after the space: <code>=RIGHT(A2, LEN(A2)-SEARCH(" ",A2))</code></li>
<li>Does it contain a word? <code>=IF(ISNUMBER(SEARCH("organic",B2)),"Yes","No")</code></li></ul>`);

add("p5-at", "Position of @", (R) => {
  const n = R.p.int(10, 13), c = customers(R, n);
  const s = sheet("Customers", ["Name", "Email", "@ Position"], c.map((x) => [x.first + " " + x.last, x.email, null]));
  return { prompt: `In <b>C2</b>, find the position of the <b>@</b> sign in each email.`, sheets: [s], target: tgt("Customers", "C2", n), solution: '=SEARCH("@",B2)', require: ["SEARCH"], hints: ["SEARCH(what, where)", '=SEARCH("@", B2)'] };
}, { drill: "text" });
add("p5-user", "Username from email", (R) => {
  const n = R.p.int(10, 13), c = customers(R, n);
  const s = sheet("Customers", ["Name", "Email", "Username"], c.map((x) => [x.first + " " + x.last, x.email, null]));
  return { prompt: `In <b>C2</b>, pull out the username: everything <b>before</b> the @.`, sheets: [s], target: tgt("Customers", "C2", n), solution: '=LEFT(B2,SEARCH("@",B2)-1)', require: ["LEFT", "SEARCH"], hints: ["The username is the LEFT part. How many characters? One less than the position of @.", '=LEFT(B2, SEARCH("@",B2)-1)'], mistakes: [{ f: '=LEFT(B2,SEARCH("@",B2))', msg: "That includes the @ sign. Subtract 1 from the SEARCH result." }] };
}, { drill: "text" });
add("p5-firstname", "First name from full name", (R) => {
  const n = R.p.int(10, 13), c = customers(R, n);
  const s = sheet("Roster", ["Full Name", "First"], c.map((x) => [x.first + " " + x.last, null]));
  return { prompt: `In <b>B2</b>, pull out the first name: everything before the space.`, sheets: [s], target: tgt("Roster", "B2", n), solution: '=LEFT(A2,SEARCH(" ",A2)-1)', require: ["LEFT", "SEARCH"], hints: ['SEARCH(" ",A2) finds the space.', '=LEFT(A2, SEARCH(" ",A2)-1)'] };
}, { drill: "text" });
add("p5-lastname", "Last name from full name", (R) => {
  const n = R.p.int(10, 13), c = customers(R, n);
  const s = sheet("Roster", ["Full Name", "Last"], c.map((x) => [x.first + " " + x.last, null]));
  return { prompt: `In <b>B2</b>, pull out the last name: everything after the space.`, sheets: [s], target: tgt("Roster", "B2", n), solution: '=RIGHT(A2,LEN(A2)-SEARCH(" ",A2))', require: ["RIGHT", "LEN", "SEARCH"], hints: ["How many characters are after the space? The total length minus the position of the space.", '=RIGHT(A2, LEN(A2)-SEARCH(" ",A2))'] };
}, { drill: "text" });
add("p5-contains", "Contains a word", (R) => {
  const n = R.p.int(10, 13), word = R.p.pick(["organic", "local", "gluten-free"]);
  const base = ["Apple Butter", "Pecan Pie", "Honey", "Peach Salsa", "Grits", "Pepper Jelly", "Sweet Tea Mix", "Granola", "Coffee", "Hot Sauce"];
  const rows = Array.from({ length: n }, () => [(R.d.chance(0.4) ? R.d.pick(["Organic ", "Local ", "Gluten-Free "]) : "") + R.d.pick(base) + (R.d.chance(0.2) ? " (" + R.d.pick(["organic", "local"]) + ")" : ""), null]);
  rows[0][0] = word.charAt(0).toUpperCase() + word.slice(1) + " " + base[0];
  const s = sheet("Products", ["Product", "Tagged?"], rows);
  return { prompt: `In <b>B2</b>, show <b>"Yes"</b> if the product name contains <b>"${word}"</b> anywhere (any capitalization), otherwise <b>"No"</b>.`, sheets: [s], target: tgt("Products", "B2", n), solution: `=IF(ISNUMBER(SEARCH("${word}",A2)),"Yes","No")`, require: ["SEARCH"], hints: ["SEARCH gives a number if it finds the word and #VALUE! if not.", `=IF(ISNUMBER(SEARCH("${word}",A2)), "Yes", "No")`] };
}, { drill: "text" });
add("p5-city", "City from \"City, ST\"", (R) => {
  const n = R.p.int(10, 13), c = customers(R, n);
  const s = sheet("Customers", ["Name", "Hometown", "City"], c.map((x) => [x.first + " " + x.last, x.city + ", " + x.st, null]));
  return { prompt: `In <b>C2</b>, pull the city out of the Hometown: everything before the comma.`, sheets: [s], target: tgt("Customers", "C2", n), solution: '=LEFT(B2,SEARCH(",",B2)-1)', require: ["LEFT", "SEARCH"], hints: ["Like the username problem, but search for the comma.", '=LEFT(B2, SEARCH(",",B2)-1)'] };
}, { drill: "text" });

level("TEXT", `
<p><code>=TEXT(value, "format")</code> turns a number or date into text in a format you choose:</p>
<ul><li><code>"mmmm"</code> gives the month name (January), <code>"mmm"</code> gives Jan, <code>"dddd"</code> gives the weekday (Monday), and <code>"mm/dd/yyyy"</code> gives 01/05/2026.</li>
<li><code>"$#,##0.00"</code> gives $1,234.50, <code>"0%"</code> gives 15%, and <code>"0.0"</code> gives 3.1.</li></ul>
<p>Handy for labels: <code>="Total: "&amp;TEXT(C2,"$#,##0.00")</code>.</p>`);

add("p5-text-month", "Month name", (R) => {
  const n = R.p.int(10, 13);
  return { prompt: `In <b>D2</b>, show the <b>full month name</b> (like "March") for each Order Date.`, sheets: [datedOrders(R, n, ["Month"])], target: tgt("Orders", "D2", n), solution: '=TEXT(B2,"mmmm")', require: ["TEXT"], hints: ['Four m\'s give the full month name.', '=TEXT(B2, "mmmm")'] };
}, { drill: "text" });
add("p5-text-day", "Weekday name", (R) => {
  const n = R.p.int(10, 13);
  return { prompt: `In <b>D2</b>, show the <b>day of the week</b> (like "Monday") each order was placed.`, sheets: [datedOrders(R, n, ["Weekday"])], target: tgt("Orders", "D2", n), solution: '=TEXT(B2,"dddd")', require: ["TEXT"], hints: ["Four d's give the full weekday name.", '=TEXT(B2, "dddd")'] };
}, { drill: "text" });
add("p5-text-label", "Dollar label", (R) => {
  const n = R.p.int(10, 13);
  return { prompt: `In <b>D2</b>, build a label like <b>Total: $1,234.50</b> from the Amount (comma, 2 decimals).`, sheets: [datedOrders(R, n, ["Label"])], target: tgt("Orders", "D2", n), solution: '="Total: "&TEXT(C2,"$#,##0.00")', require: ["TEXT"], hints: ['TEXT(C2,"$#,##0.00") formats the amount.', '="Total: "&TEXT(C2, "$#,##0.00")'], mistakes: [{ f: '="Total: "&C2', msg: "Gluing a number onto text drops its formatting (no $, commas, or fixed decimals). Wrap the amount in TEXT with a format code." }] };
}, { drill: "text" });

// ################################################################## PART 6 — combined formulas
part(6, "Mix It Up: Combined Formulas");

level("Two functions together", `
<p>Real formulas nest functions from <b>different topics</b> inside each other. Build them from the inside out:</p>
<ol><li>Figure out the inner piece first. For example, <code>LEFT(A2,3)</code> gives a department code.</li>
<li>Use that piece as an argument of the outer function: <code>=VLOOKUP(LEFT(A2,3), Depts!$A$2:$B$6, 2, FALSE)</code>.</li></ol>
<p><b>Tip:</b> type the inner part alone first and press <b>Enter</b> to see what it gives. Then wrap it.</p>`);

add("c-round-avgifs", "Rounded category average", (R) => {
  const n = 16, o = orders(R.d, n), cat = R.p.pick(CATEGORIES);
  for (let k = 0; k < 3; k++) { const r = o.rows[R.d.int(0, n - 1)]; const pr = PRODUCTS.filter((x) => x[1] === cat)[R.d.int(0, 3)]; r[4] = pr[0]; r[5] = cat; r[7] = pr[2]; }
  const s = sheet("Orders", o.headers, o.rows, o.fmt, [["J1", "Category"], ["J2", cat], ["K1", "Avg Price"]]);
  return {
    prompt: `In <b>K2</b>, find the average <b>Unit Price</b> for the <b>${cat}</b> category (J2), <b>rounded to 2 decimal places</b>.`,
    sheets: [s], target: tgt("Orders", "K2", 1, null, "money"),
    solution: `=ROUND(AVERAGEIFS(H2:H${n + 1},F2:F${n + 1},J2),2)`, require: ["ROUND", "AVERAGEIFS"],
    hints: ["Inside: AVERAGEIFS(Unit Price, Category, J2). Outside: ROUND(..., 2).", `=ROUND(AVERAGEIFS(H2:H${n + 1}, F2:F${n + 1}, J2), 2)`],
  };
}, { drill: "combo" });
add("c-if-countifs", "Was it ever ordered?", (R) => {
  const k = 10, list = R.d.sample(PRODUCTS, k), n = 14;
  const used = list.slice(0, 6);
  const orders_ = sheet("Orders", ["Order ID", "Product", "Qty"], ids(R.d, n, "", 20100, 20999).sort().map((id) => [+id, R.d.pick(used)[0], R.d.int(1, 9)]));
  const prod = sheet("Products", ["Product", "Category", "Price", "Ordered?"], R.d.shuffle(list).map((x) => [x[0], x[1], x[2], null]), { Price: "money" });
  return {
    prompt: `On the <b>Products</b> sheet, in <b>D2</b>, show <b>"Yes"</b> if the product appears <b>at least once</b> on the Orders sheet, otherwise <b>"No"</b>.`,
    sheets: [prod, orders_], target: tgt("Products", "D2", k), require: ["IF", "COUNTIFS"],
    solution: `=IF(COUNTIFS(Orders!$B$2:$B$${n + 1},A2)>0,"Yes","No")`,
    hints: ["COUNTIFS(Orders!$B$2:$B$15, A2) counts how many times the product was ordered.", "Wrap it: =IF(COUNTIFS(...)>0, \"Yes\", \"No\")"],
  };
}, { drill: "combo" });
add("c-vl-left", "Look up a piece of a code", (R) => {
  const codes = [["CMP", "Camping"], ["WTR", "Water Sports"], ["FTW", "Footwear"], ["APP", "Apparel"], ["GER", "Gear"]];
  const n = R.p.int(10, 13);
  const depts = sheet("Depts", ["Code", "Department"], codes);
  const items = sheet("Items", ["Item Code", "On Hand", "Department"], Array.from({ length: n }, () => [R.d.pick(codes)[0] + R.d.int(10000, 99999), R.d.int(0, 80), null]));
  return {
    prompt: `The first 3 letters of each item code are a department code. In <b>C2</b>, look up the <b>Department</b> name on the <b>Depts</b> sheet using those 3 letters.`,
    sheets: [items, depts], target: tgt("Items", "C2", n), require: ["VLOOKUP", "LEFT"],
    solution: "=VLOOKUP(LEFT(A2,3),Depts!$A$2:$B$6,2,FALSE)",
    hints: ["LEFT(A2,3) gives the code. Use it as VLOOKUP's lookup value.", "=VLOOKUP(LEFT(A2,3), Depts!$A$2:$B$6, 2, FALSE)"],
  };
}, { drill: "combo" });
add("c-proper-email", "First name from an email", (R) => {
  const n = R.p.int(10, 13), c = customers(R, n);
  const s = sheet("Customers", ["Email", "First Name"], c.map((x) => [x.email, null]));
  return {
    prompt: `Emails look like <code>mary.smith@gmail.com</code>. In <b>B2</b>, pull out the first name (everything before the <b>period</b>) and capitalize it properly, like <b>Mary</b>.`,
    sheets: [s], target: tgt("Customers", "B2", n), require: ["PROPER", "LEFT", "SEARCH"],
    solution: '=PROPER(LEFT(A2,SEARCH(".",A2)-1))',
    hints: ['LEFT(A2, SEARCH(".",A2)-1) gets "mary".', 'Wrap it in PROPER: =PROPER(LEFT(A2, SEARCH(".",A2)-1))'],
  };
}, { drill: "combo" });
add("c-text-sumifs", "Formatted summary label", (R) => {
  const n = 16, t = ordersWithTotal(R, n), reg = R.p.pick(REGIONS);
  ensure(t.rows, 3, reg, R.d, 3);
  const s = sheet("Orders", t.h, t.rows, t.fmt, [["K1", "Region"], ["K2", reg], ["L1", "Label"]]);
  return {
    prompt: `In <b>L2</b>, build a label like <b>${reg}: $12,345</b>: the region name from K2, a colon and a space, then that region's total of the <b>Total</b> column formatted with a dollar sign, commas, and no decimals.`,
    sheets: [s], target: tgt("Orders", "L2", 1), require: ["TEXT", "SUMIFS"],
    solution: `=K2&": "&TEXT(SUMIFS(I2:I${n + 1},D2:D${n + 1},K2),"$#,##0")`,
    hints: [`SUMIFS(I2:I${n + 1}, D2:D${n + 1}, K2) gives the number. TEXT(..., "$#,##0") formats it.`, 'Glue it together: =K2&": "&TEXT(SUMIFS(...), "$#,##0")'],
  };
}, { drill: "combo" });
add("c-sp-cond", "SUMPRODUCT with a condition", (R) => {
  const n = 16, reg = R.p.pick(REGIONS);
  const { o, s } = ordersSheet(R, n, [["J1", "Region"], ["J2", reg], ["K1", "Revenue"]]);
  ensure(o.rows, 3, reg, R.d, 3);
  return {
    prompt: `In <b>K2</b>, use one SUMPRODUCT to find the revenue (Units × Unit Price) for only the <b>${reg}</b> region. Multiply in a test, <code>(D2:D17=J2)</code>, which counts as 1 for matching rows and 0 for the rest.`,
    sheets: [s], target: tgt("Orders", "K2", 1, null, "money"), require: ["SUMPRODUCT"],
    solution: `=SUMPRODUCT((D2:D${n + 1}=J2)*G2:G${n + 1}*H2:H${n + 1})`,
    hints: ["(D2:D17=J2) gives TRUE/FALSE for each row, and multiplying turns those into 1s and 0s.", `=SUMPRODUCT((D2:D${n + 1}=J2)*G2:G${n + 1}*H2:H${n + 1})`],
  };
}, { drill: "combo" });

level("Three or more", `
<p>Now there are three or more layers. Talk it through in plain English first:</p>
<p><i>"If the birthday is this month, look up the gift for their facility, otherwise 0."</i> becomes
<code>=IF(MONTH(C2)=MONTH(TODAY()), VLOOKUP(B2, Gifts!$A$2:$B$6, 2, FALSE), 0)</code></p>
<p>Count your parentheses. Each function you open needs its own <code>)</code>.</p>`);

add("c-iferror-vl-trim", "Messy lookup values", (R) => {
  const k = 10, { list, s } = productSheet(R, k), n = R.p.int(10, 13);
  const others = PRODUCTS.filter((x) => !list.includes(x));
  const sp = () => R.d.pick(["", " ", "  "]);
  const rows = ids(R.d, n, "", 20100, 20999).sort().map((id) => [+id, sp() + (R.d.chance(0.2) ? R.d.pick(others)[0] : R.d.pick(list)[0]) + sp(), null]);
  rows[0][1] = " " + list[0][0] + "  "; rows[1][1] = others[0][0];
  const o = sheet("Orders", ["Order ID", "Product", "Price"], rows);
  return {
    prompt: `The product names on Orders have stray spaces, and a few products aren't on the Products sheet at all. In <b>C2</b>, clean up the name, look up its <b>Price</b> (3rd column of Products), and show <b>"Not Found"</b> if it isn't there.`,
    sheets: [o, s], target: tgt("Orders", "C2", n, null, "money"), require: ["IFERROR", "VLOOKUP", "TRIM"],
    solution: `=IFERROR(VLOOKUP(TRIM(B2),Products!$A$2:$C$${k + 1},3,FALSE),"Not Found")`,
    hints: ["Use TRIM(B2) as the lookup value.", "=IFERROR(VLOOKUP(TRIM(B2), Products!$A$2:$C$11, 3, FALSE), \"Not Found\")"],
  };
}, { drill: "combo" });
add("c-if-month-vl", "Birthday gift card", (R) => {
  const facs = FACILITIES.slice(0, 5), amts = R.p.sample([15, 20, 25, 30, 40, 50], 5);
  const gifts = sheet("Gifts", ["Facility", "Gift Card"], facs.map((f, i) => [f, amts[i]]), { "Gift Card": "money0" });
  const n = R.p.int(10, 13), ppl = people(R.d, n), m = thisMonth();
  const rows = ppl.map((p) => [p.first + " " + p.last, R.d.pick(facs), D(R.d.int(1965, 2004), R.d.chance(0.35) ? m : R.d.int(1, 12), R.d.int(1, 28)), null]);
  rows[0][2] = D(1991, m, 12);
  const s = sheet("Staff", ["Name", "Facility", "Birth Date", "Gift"], rows, { "Birth Date": "date" });
  return {
    prompt: `Employees with a birthday in the <b>current month</b> get a gift card, and the amount depends on their facility (see the <b>Gifts</b> sheet). In <b>D2</b>, show the gift amount, or <b>0</b> if their birthday isn't this month.`,
    sheets: [s, gifts], target: tgt("Staff", "D2", n, null, "money0"), require: ["IF", "MONTH", "TODAY", "VLOOKUP"],
    solution: "=IF(MONTH(C2)=MONTH(TODAY()),VLOOKUP(B2,Gifts!$A$2:$B$6,2,FALSE),0)",
    hints: ["Test: MONTH(C2)=MONTH(TODAY()). True part: a VLOOKUP of the facility. False part: 0.", "=IF(MONTH(C2)=MONTH(TODAY()), VLOOKUP(B2, Gifts!$A$2:$B$6, 2, FALSE), 0)"],
  };
}, { drill: "combo" });
add("c-if-and-search", "Rush orders", (R) => {
  const n = R.p.int(10, 13);
  const svc = ["Ground", "Express Air", "2-Day Express", "Standard", "Freight", "Next-Day express"];
  const rows = ids(R.d, n, "SO-", 100, 999).map((id) => [id, R.d.pick(svc), R.d.money(40, 900)]);
  rows[0][1] = "Express Air";
  const s = sheet("Orders", ["Order", "Service", "Amount", "Handling"], rows, { Amount: "money" }, [["F1", "Rush Minimum"], ["G1", R.d.pick([150, 200, 250, 300]), "money0"]]);
  return {
    prompt: `An order is <b>"Rush"</b> if its Service contains the word <b>express</b> (any capitalization, anywhere) <b>and</b> its Amount is at least the minimum in <b>G1</b>. Otherwise it's <b>"Normal"</b>. Fill in <b>D2</b>.`,
    sheets: [s], target: tgt("Orders", "D2", n), require: ["IF", "AND", "ISNUMBER", "SEARCH"],
    solution: '=IF(AND(ISNUMBER(SEARCH("express",B2)),C2>=$G$1),"Rush","Normal")',
    hints: ['"Contains express" is ISNUMBER(SEARCH("express",B2)).', 'Put both tests in AND: =IF(AND(ISNUMBER(SEARCH("express",B2)), C2>=$G$1), "Rush", "Normal")'],
  };
}, { drill: "combo" });
add("c-vl-round-yos", "Bonus from years of service", (R) => {
  const b = R.p.pick([[100, 250, 500], [150, 300, 600]]);
  const tiers = sheet("Bonus", ["Years", "Bonus"], [[0, b[0]], [5, b[1]], [10, b[2]]], { Bonus: "money0" });
  const n = R.p.int(10, 13), ppl = people(R.d, n), t = todaySerial();
  const s = sheet("Staff", ["Name", "Hire Date", "Holiday Bonus"], ppl.map((p) => [p.first + " " + p.last, t - R.d.int(100, 5500), null]), { "Hire Date": "date" });
  return {
    prompt: `In <b>C2</b>, calculate years of service from the Hire Date ((today − Hire Date) ÷ 365, rounded to 1 decimal), and use it to look up the holiday bonus tier on the <b>Bonus</b> sheet, all in one formula.`,
    sheets: [s, tiers], target: tgt("Staff", "C2", n, null, "money0"), require: ["VLOOKUP", "ROUND", "TODAY"],
    solution: "=VLOOKUP(ROUND((TODAY()-B2)/365,1),Bonus!$A$2:$B$4,2,TRUE)",
    hints: ["The years: ROUND((TODAY()-B2)/365, 1). That's the lookup value.", "=VLOOKUP(ROUND((TODAY()-B2)/365,1), Bonus!$A$2:$B$4, 2, TRUE)"],
  };
}, { drill: "combo" });
add("c-imm-clean", "Two-way lookup with messy input", (R) => {
  const styles = ["Crew Tee", "Polo", "Hoodie", "Quarter Zip", "Rain Shell"], sizes = ["S", "M", "L", "XL"];
  const tbl = sheet("Prices", ["Style"].concat(sizes), styles.map((st, i) => [st].concat(sizes.map((z, j) => 14 + i * 9 + j * 2))));
  for (let i = 1; i <= 5; i++) for (let j = 1; j <= 4; j++) tbl.fmt[i + ":" + j] = "money";
  const n = R.p.int(10, 13);
  const rows = Array.from({ length: n }, () => [R.d.int(800, 999), R.d.pick(styles), R.d.pick(["", " "]) + R.d.pick(sizes).toLowerCase() + R.d.pick(["", " ", "  "]), null]);
  rows[0][2] = " xl ";
  const main = sheet("Orders", ["Order", "Style", "Size (typed)", "Price"], rows);
  return {
    prompt: `Sizes were typed sloppily (like <code>" xl "</code>). In <b>D2</b>, use INDEX/MATCH/MATCH to look up the price on the <b>Prices</b> sheet, cleaning the size with <b>UPPER</b> and <b>TRIM</b> before matching it.`,
    sheets: [main, tbl], target: tgt("Orders", "D2", n, null, "money"), require: ["INDEX", "MATCH", "TRIM", "UPPER"],
    solution: "=INDEX(Prices!$B$2:$E$6,MATCH(B2,Prices!$A$2:$A$6,0),MATCH(UPPER(TRIM(C2)),Prices!$B$1:$E$1,0))",
    hints: ["The second MATCH looks up UPPER(TRIM(C2)) in Prices!$B$1:$E$1.", "=INDEX(Prices!$B$2:$E$6, MATCH(B2, Prices!$A$2:$A$6, 0), MATCH(UPPER(TRIM(C2)), Prices!$B$1:$E$1, 0))"],
  };
}, { drill: "combo" });
add("c-countifs-avg", "Above-average orders in a region", (R) => {
  const n = 16, reg = R.p.pick(REGIONS);
  const { o, s } = ordersSheet(R, n, [["J1", "Region"], ["J2", reg], ["K1", "Big Orders"]]);
  ensure(o.rows, 3, reg, R.d, 4);
  return {
    prompt: `In <b>K2</b>, count the orders from the <b>${reg}</b> region (J2) whose Units are <b>at least the average</b> Units of all orders. Calculate the average inside the formula. Don't type it in.`,
    sheets: [s], target: tgt("Orders", "K2", 1, null, "int"), require: ["COUNTIFS", "AVERAGE"],
    solution: `=COUNTIFS(D2:D${n + 1},J2,G2:G${n + 1},">="&AVERAGE(G2:G${n + 1}))`,
    hints: [`The second criteria is ">="&AVERAGE(G2:G${n + 1}).`, `=COUNTIFS(D2:D${n + 1}, J2, G2:G${n + 1}, ">="&AVERAGE(G2:G${n + 1}))`],
  };
}, { drill: "combo" });

level("Boss formulas", `
<p>The big ones. Each problem uses functions from three or more topics. Build them one layer at a time, and press Enter to test each layer before you wrap it.</p>
<p>Testing out is off here: every boss counts.</p>`, { noTestOut: true });

add("c-boss-code", "Account code builder", (R) => {
  const n = R.p.int(10, 13), base = todaySerial();
  const rows = Array.from({ length: n }, () => [R.d.int(1, 9999), R.d.pick(REGIONS), base - R.d.int(30, 2500), null]);
  rows[0][0] = R.d.int(1, 99);
  const s = sheet("Accounts", ["Account #", "Region", "Opened", "Code"], rows, { Opened: "date" });
  return {
    prompt: `Account codes look like <b>EAS-0042-2024</b>: the first 3 letters of the Region in capitals, a dash, the account number padded to <b>4 digits</b> with TEXT (format <code>"0000"</code>), a dash, then the <b>year</b> the account was opened. Build it in <b>D2</b>.`,
    sheets: [s], target: tgt("Accounts", "D2", n), require: ["UPPER", "LEFT", "TEXT", "YEAR"],
    solution: '=UPPER(LEFT(B2,3))&"-"&TEXT(A2,"0000")&"-"&YEAR(C2)',
    hints: ['Three pieces: UPPER(LEFT(B2,3)), TEXT(A2,"0000"), and YEAR(C2).', 'Glue them with &"-"&: =UPPER(LEFT(B2,3))&"-"&TEXT(A2,"0000")&"-"&YEAR(C2)'],
  };
}, { drill: "combo" });
add("c-boss-commission", "Commission if over quota", (R) => {
  const lv = ["Associate", "Senior", "Lead", "Manager"], rates = R.p.pick([[0.02, 0.03, 0.04, 0.05], [0.015, 0.025, 0.035, 0.045]]);
  const rt = sheet("Rates", ["Level", "Rate"], lv.map((x, i) => [x, rates[i]]), { Rate: "pct1" });
  const n = R.p.int(10, 13), ppl = people(R.d, n);
  const rows = ppl.map((p) => { const q = R.d.int(30, 70) * 1000; return [p.first + " " + p.last, q + R.d.int(-15, 25) * 1000 + R.d.int(0, 999), q, R.d.pick(lv), null]; });
  const s = sheet("Reps", ["Rep", "Sales", "Quota", "Level", "Commission"], rows, { Sales: "money0", Quota: "money0" });
  return {
    prompt: `Reps who reach their quota (Sales <b>at or above</b> Quota) earn Sales × the commission rate for their <b>Level</b> (see the <b>Rates</b> sheet). Reps below quota earn 0. Round the commission to <b>2 decimals</b>. Fill in <b>E2</b>.`,
    sheets: [s, rt], target: tgt("Reps", "E2", n, null, "money"), require: ["ROUND", "IF", "VLOOKUP"],
    solution: "=ROUND(IF(B2>=C2,B2*VLOOKUP(D2,Rates!$A$2:$B$5,2,FALSE),0),2)",
    hints: ["Middle layer: IF(B2>=C2, B2*VLOOKUP(D2, Rates!$A$2:$B$5, 2, FALSE), 0)", "Wrap the whole IF in ROUND(..., 2)."],
  };
}, { drill: "combo" });
add("c-boss-lastname", "Phone by last name", (R) => {
  const k = 12, lasts = R.d.sample(LAST, k);
  const dir = sheet("Directory", ["Last Name", "Phone Ext"], lasts.map((l) => [l, R.d.int(1000, 4999)]));
  const n = R.p.int(9, 12);
  const rows = Array.from({ length: n }, () => [R.d.pick(FIRST) + " " + (R.d.chance(0.2) ? R.d.pick(LAST.filter((x) => !lasts.includes(x))) : R.d.pick(lasts)), null]);
  rows[1][0] = "Pat " + LAST.filter((x) => !lasts.includes(x))[0];
  const s = sheet("Callers", ["Full Name", "Ext"], rows);
  return {
    prompt: `The <b>Directory</b> is organized by last name only. In <b>B2</b>, pull the last name out of the Full Name, look up the phone extension, and show <b>"Unlisted"</b> if that last name isn't in the directory.`,
    sheets: [s, dir], target: tgt("Callers", "B2", n), require: ["IFERROR", "VLOOKUP", "RIGHT", "LEN", "SEARCH"],
    solution: '=IFERROR(VLOOKUP(RIGHT(A2,LEN(A2)-SEARCH(" ",A2)),Directory!$A$2:$B$13,2,FALSE),"Unlisted")',
    hints: ['Last name: RIGHT(A2, LEN(A2)-SEARCH(" ",A2)). That\'s the lookup value.', 'Then =IFERROR(VLOOKUP(that, Directory!$A$2:$B$13, 2, FALSE), "Unlisted")'],
  };
}, { drill: "combo" });
add("c-boss-status", "Invoice status", (R) => {
  const n = R.p.int(10, 13), t = todaySerial();
  const rows = ids(R.d, n, "INV-", 2000, 2999).map((id) => { const due = t + R.d.int(-45, 30); return [id, R.d.money(100, 4000), due, R.d.chance(0.4) ? due - R.d.int(-5, 10) : null, null]; });
  rows[0][2] = t - 10; rows[0][3] = null; rows[1][2] = t + 10; rows[1][3] = null;
  const s = sheet("Invoices", ["Invoice", "Amount", "Due Date", "Paid On", "Status"], rows, { Amount: "money", "Due Date": "date", "Paid On": "date" });
  return {
    prompt: `In <b>E2</b>: if <b>Paid On</b> has a date, show <b>"Paid"</b>. If it's empty, show <b>"Overdue"</b> when the Due Date is before today, or <b>"Open"</b> otherwise.`,
    sheets: [s], target: tgt("Invoices", "E2", n), require: ["IF", "ISBLANK", "TODAY"],
    solution: '=IF(ISBLANK(D2),IF(C2<TODAY(),"Overdue","Open"),"Paid")',
    hints: ["Outer IF: is D2 blank? (ISBLANK)", 'The true part is another IF: IF(C2<TODAY(), "Overdue", "Open"). The false part is "Paid".'],
  };
}, { drill: "combo" });

// ################################################################## PART 7
part(7, "Exam-Style Practice");

// Scenario 1: HR workbook built in steps
function hrScenario(R) {
  const n = 14, ppl = people(R.d, n), idl = ids(R.d, n, "", 1000, 9999).map(Number), t = todaySerial(), m = thisMonth();
  const facs = FACILITIES.slice(0, 5), days = [5, 6, 3, 7, 9];
  const rows = idl.map((id, i) => [id, R.d.pick(facs), R.d.pick(["FT", "FT", "PT"]), t - R.d.int(200, 5500), D(R.d.int(1966, 2004), R.d.chance(0.25) ? m : R.d.int(1, 12), R.d.int(1, 28))]);
  rows[0][1] = "Huntsville"; rows[0][2] = "FT";
  rows[1][2] = "PT"; rows[1][3] = t - 1100;
  const H = ["Employee ID", "Facility", "Job Status", "Hire Date", "Birth Date", "First Name", "Birthday?", "Raise %", "Years of Service", "Eligible for Vacation?", "# of Vacation Days", "Holiday Bonus"];
  const main = sheet("Employees", H, rows.map((r) => r.concat([null, null, null, null, null, null, null])), { "Hire Date": "date", "Birth Date": "date", "Raise %": "pct", "Years of Service": "dec1", "Holiday Bonus": "money0" });
  const names = sheet("Employee Names", ["First Name", "Last Name", "Employee ID"], R.d.shuffle(ppl.map((p, i) => [p.first, p.last, idl[i]])));
  const dt = sheet("Data Tables", ["Facility", "Vacation Days", "", "Years", "Bonus"], facs.map((f, i) => [f, days[i], null, i < 3 ? [0, 5, 10][i] : null, i < 3 ? [100, 250, 500][i] : null]));
  dt.fmt[4] = "money0";
  return { sheets: [main, names, dt], n };
}
const HR_STEPS = [
  { col: "F", title: "First Name", f: "=INDEX('Employee Names'!$A$2:$A$15,MATCH(A2,'Employee Names'!$C$2:$C$15,0))", fmt: null, req: ["INDEX", "MATCH"], forbid: ["VLOOKUP"],
    prompt: "Bring in each employee's <b>First Name</b> from the <b>Employee Names</b> sheet, matching on Employee ID. (The ID is in the last column there, so use INDEX/MATCH.)", hints: ["MATCH the ID in 'Employee Names'!$C$2:$C$15 with 0.", "INDEX 'Employee Names'!$A$2:$A$15 at that position."] },
  { col: "G", title: "Birthday?", f: '=IF(MONTH(E2)=MONTH(TODAY()),"Yes","")', req: ["MONTH", "TODAY"],
    prompt: "Show <b>\"Yes\"</b> if the employee's birthday is in the current month, otherwise blank (\"\"). It must keep working as the months change.", hints: ["Compare MONTH(E2) with MONTH(TODAY()).", '=IF(MONTH(E2)=MONTH(TODAY()),"Yes","")'] },
  { col: "H", title: "Raise %", f: '=IF(AND(B2="Huntsville",C2="FT"),0.04,0.03)', fmt: "pct",
    prompt: "Full-time (\"FT\") employees at the <b>Huntsville</b> facility get a <b>4%</b> raise; everyone else gets <b>3%</b>.", hints: ["Two conditions that must both be true: AND.", '=IF(AND(B2="Huntsville",C2="FT"),0.04,0.03)'] },
  { col: "I", title: "Years of Service", f: "=ROUND((TODAY()-D2)/365,1)", fmt: "dec1", req: ["ROUND", "TODAY"],
    prompt: "Calculate years of service from the Hire Date: (today − Hire Date) ÷ 365, rounded to <b>1 decimal</b>.", hints: ["(TODAY()-D2)/365", "Wrap it in ROUND(..., 1)."] },
  { col: "J", title: "Eligible for Vacation?", f: '=IF(OR(C2="FT",AND(C2="PT",I2>=2)),"Yes","")',
    prompt: "Employees are eligible for vacation if they're full-time, <b>or</b> part-time with <b>2 or more</b> years of service. Show <b>\"Yes\"</b> or blank (\"\").", hints: ["OR(full-time, AND(part-time, years>=2))", "Years of Service is in column I."] },
  { col: "K", title: "# of Vacation Days", f: "=IF(J2=\"Yes\",VLOOKUP(B2,'Data Tables'!$A$2:$B$6,2,FALSE),\"\")", fmt: "int", req: ["VLOOKUP"],
    prompt: "If the employee is eligible, look up their facility's vacation days on the <b>Data Tables</b> sheet (A2:B6). Otherwise blank (\"\").", hints: ["=IF(J2=\"Yes\", VLOOKUP(...), \"\")", "'Data Tables'!$A$2:$B$6, column 2, FALSE"] },
  { col: "L", title: "Holiday Bonus", f: "=VLOOKUP(I2,'Data Tables'!$D$2:$E$4,2,TRUE)", fmt: "money0", req: ["VLOOKUP"],
    prompt: "Everyone gets a holiday bonus based on years of service: under 5 years $100, 5 to under 10 $250, 10+ $500. The tiers are on <b>Data Tables</b> (D2:E4). Use an approximate-match VLOOKUP.", hints: ["Look up the years (I2) in 'Data Tables'!$D$2:$E$4.", "Use TRUE for approximate match."] },
];

// Scenario 2: shipping / 3PL workbook built in steps
function shipScenario(R) {
  const n = 15, carriers = ["FedEx", "UPS", "XPO", "Old Dominion"];
  const custs = ["ACME SUPPLY", "TIDE TOOLING", "GULF COAST FOODS", "MAGIC CITY METALS", "RIVER REGION PAPER", "SHOALS TEXTILES"];
  const t = todaySerial();
  const rows = ids(R.d, n, "SH-", 4000, 4999).sort().map((id, i) => {
    const c = R.d.pick(CITIES.slice(0, 8)), ship = t - R.d.int(10, 80), prom = R.d.pick([2, 3, 3, 5]);
    return [id, R.d.pick(custs), i < 8 ? carriers[i % 4] : R.d.pick(carriers), c[0].toLowerCase() + ", " + c[1].toLowerCase(), String(R.d.int(35000, 36999)) + "-" + R.d.int(1000, 9999), R.d.int(5, 900), ship, prom, ship + Math.max(1, prom + R.d.int(-1, 3))];
  });
  const H = ["Shipment", "Customer (raw)", "Carrier", "Origin", "ZIP+4", "Weight (lb)", "Ship Date", "Promised Days", "Delivered", "Customer", "ZIP", "Days in Transit", "On Time?", "Rate"];
  const main = sheet("Shipments", H, rows.map((r) => r.concat([null, null, null, null, null])), { "Ship Date": "date", Delivered: "date", Rate: "money" });
  const rates = sheet("Rates", ["Min Weight", "Rate"], [[0, 45], [100, 85], [250, 160], [500, 290]], { Rate: "money" });
  const summary = sheet("Summary", ["Carrier", "On-Time Shipments", "Avg Days in Transit"], carriers.map((c) => [c, null, null]));
  return { sheets: [main, rates, summary], n, carriers };
}
const SHIP_STEPS = [
  { col: "J", title: "Customer", f: "=PROPER(B2)", req: ["PROPER"], prompt: "Customer names came in ALL CAPS. Convert them to proper case.", hints: ["One text function does it.", "=PROPER(B2)"] },
  { col: "K", title: "ZIP", f: "=VALUE(LEFT(E2,5))", req: ["VALUE", "LEFT"], strictType: true, prompt: "Take the first 5 digits of the ZIP+4 and make them a real <b>number</b>.", hints: ["LEFT(E2,5) gets the digits as text.", "Wrap it in VALUE."] },
  { col: "L", title: "Days in Transit", f: "=I2-G2", fmt: "int", prompt: "How many days was each shipment in transit? (Delivered − Ship Date)", hints: ["Subtracting dates gives days.", "=I2-G2"] },
  { col: "M", title: "On Time?", f: '=IF(L2<=H2,"Yes","No")', prompt: "A shipment is on time if its Days in Transit is <b>at or below</b> its Promised Days. Show <b>\"Yes\"</b> or <b>\"No\"</b>.", hints: ["Compare L2 with H2.", '=IF(L2<=H2,"Yes","No")'] },
  { col: "N", title: "Rate", f: "=VLOOKUP(F2,Rates!$A$2:$B$5,2,TRUE)", fmt: "money", req: ["VLOOKUP"], prompt: "Look up each shipment's rate by weight on the <b>Rates</b> sheet (the first column holds each tier's minimum weight).", hints: ["It's a tier lookup, so approximate match (TRUE).", "=VLOOKUP(F2,Rates!$A$2:$B$5,2,TRUE)"] },
  { sheet: "Summary", col: "B", title: "On-Time Shipments", f: '=COUNTIFS(Shipments!$C$2:$C$16,A2,Shipments!$M$2:$M$16,"Yes")', fmt: "int", rows: 4, req: ["COUNTIFS"], prompt: "On the <b>Summary</b> sheet, count each carrier's on-time shipments (Carrier matches A2 <b>and</b> On Time? is \"Yes\").", hints: ["Two criteria: the carrier column and the On Time? column on Shipments.", "Lock the Shipments ranges with $ (they're copied down to B5)."] },
  { sheet: "Summary", col: "C", title: "Avg Days in Transit", f: "=AVERAGEIFS(Shipments!$L$2:$L$16,Shipments!$C$2:$C$16,A2)", fmt: "dec2", rows: 4, req: ["AVERAGEIFS"], prompt: "On the <b>Summary</b> sheet, find each carrier's average Days in Transit.", hints: ["AVERAGEIFS(average range, criteria range, criteria)", "Average Shipments!L, where Shipments!C matches A2."] },
];

function scenarioProblems(prefix, steps, builder, mainSheet, intro) {
  steps.forEach((st, k) => {
    add(prefix + "-" + (k + 1), "Step " + (k + 1) + ": " + st.title, (R) => {
      const sc = builder(R);
      const sh = st.sheet || mainSheet;
      const rows = st.rows || sc.n;
      const prefill = steps.slice(0, k).map((p) => ({ sheet: p.sheet || mainSheet, cell: p.col + "2", rows: p.rows || sc.n, f: p.f }));
      return {
        prompt: `<span class="step-tag">${intro}</span> In <b>${st.col}2</b> (${st.title}): ${st.prompt}` + (k ? ` <span class="muted">Earlier columns are already filled in for you.</span>` : ""),
        sheets: sc.sheets, active: sh, prefill, target: tgt(sh, st.col + "2", rows, null, st.fmt),
        solution: st.f, require: st.req, forbid: st.forbid, strictType: st.strictType, hints: st.hints,
      };
    });
  });
}
level("Scenario 1: HR Workbook", `
<p>A full exam-style workbook for a regional outdoor-gear chain's HR team. Each step adds one column, the way exam questions build on each other.</p>
<p>Testing out is off here: work through every step. Use the sheet tabs at the bottom to look at the other sheets.</p>`, { noTestOut: true });
scenarioProblems("p7-hr", HR_STEPS, hrScenario, "Employees", "HR Workbook");
level("Scenario 2: Shipping Workbook", `
<p>A third-party logistics (3PL) company's shipment log needs cleaning up and summarizing: text functions, dates, IF, lookups, and a summary sheet.</p>
<p>Testing out is off here: work through every step.</p>`, { noTestOut: true });
scenarioProblems("p7-ship", SHIP_STEPS, shipScenario, "Shipments", "Shipping Workbook");

// Scenario 3: everything intertwined (sales commission workbook)
function salesScenario(R) {
  const n = 12, ppl = people(R.d, n), idl = ids(R.d, n, "R", 100, 999), t = todaySerial();
  const rows = idl.map((id, i) => [id, REGIONS[i % 4], R.d.int(15, 90) * 1000 + R.d.int(0, 999), R.d.int(15, 90) * 1000 + R.d.int(0, 999), t - R.d.int(20, 2500)]);
  rows[0][4] = t - 400; rows[1][4] = D(new Date().getFullYear(), 1, 1);
  const H = ["Rep ID", "Region", "Q1 Sales", "Q2 Sales", "Hire Date", "Rep Name", "Total Sales", "Commission", "Award", "Email"];
  const main = sheet("Sales", H, rows.map((r) => r.concat([null, null, null, null, null])), { "Q1 Sales": "money0", "Q2 Sales": "money0", "Hire Date": "date", "Total Sales": "money0", Commission: "money" });
  const reps = sheet("Reps", ["Name", "Rep ID"], R.d.shuffle(ppl.map((p, i) => [p.first + " " + p.last, idl[i]])));
  const rates = sheet("Rates", ["Min Sales", "Rate"], [[0, 0.02], [50000, 0.03], [100000, 0.045], [150000, 0.06]], { "Min Sales": "money0", Rate: "pct1" });
  const summary = sheet("Summary", ["Region", "Total Commission", "Award Winners"], REGIONS.map((r) => [r, null, null]));
  return { sheets: [main, reps, rates, summary], n };
}
const SALES_STEPS = [
  { col: "F", title: "Rep Name", f: "=INDEX(Reps!$A$2:$A$13,MATCH(A2,Reps!$B$2:$B$13,0))", req: ["INDEX", "MATCH"], forbid: ["VLOOKUP"],
    prompt: "Bring in each rep's name from the <b>Reps</b> sheet by Rep ID. (The ID is to the right of the name there.)", hints: ["MATCH the ID in Reps!$B$2:$B$13 with 0.", "INDEX Reps!$A$2:$A$13 at that position."] },
  { col: "G", title: "Total Sales", f: "=SUM(C2:D2)", fmt: "money0", prompt: "Add Q1 and Q2 sales.", hints: ["=SUM(C2:D2) or =C2+D2"] },
  { col: "H", title: "Commission", f: "=ROUND(G2*VLOOKUP(G2,Rates!$A$2:$B$5,2,TRUE),2)", fmt: "money", req: ["ROUND", "VLOOKUP"],
    prompt: "Commission is Total Sales × the rate for that sales tier (see <b>Rates</b>, which lists each tier's minimum sales), rounded to 2 decimals.", hints: ["The rate: VLOOKUP(G2, Rates!$A$2:$B$5, 2, TRUE).", "=ROUND(G2*VLOOKUP(...), 2)"] },
  { col: "I", title: "Award", f: '=IF(AND(G2>=AVERAGEIFS($G$2:$G$13,$B$2:$B$13,B2),YEAR(E2)<YEAR(TODAY())),"Award","")', req: ["IF", "AND", "AVERAGEIFS", "YEAR", "TODAY"],
    prompt: "A rep wins an <b>\"Award\"</b> if their Total Sales are at or above the <b>average Total Sales for their region</b>, <b>and</b> they were hired <b>before this year</b>. Otherwise blank (\"\").", hints: ["Region average: AVERAGEIFS($G$2:$G$13, $B$2:$B$13, B2). Lock the ranges!", "Hired before this year: YEAR(E2)<YEAR(TODAY()). Put both in AND inside IF."] },
  { col: "J", title: "Email", f: '=LOWER(LEFT(F2,1)&RIGHT(F2,LEN(F2)-SEARCH(" ",F2)))&"@crimsoncargo.com"', req: ["LOWER", "LEFT", "RIGHT", "LEN", "SEARCH"],
    prompt: "Build each email: first initial + last name, all lowercase, then <b>@crimsoncargo.com</b>. For example, Mary Smith becomes <b>msmith@crimsoncargo.com</b>.", hints: ['Initial: LEFT(F2,1). Last name: RIGHT(F2, LEN(F2)-SEARCH(" ",F2)).', 'LOWER(initial & last) & "@crimsoncargo.com"'] },
  { sheet: "Summary", col: "B", title: "Total Commission", f: "=SUMIFS(Sales!$H$2:$H$13,Sales!$B$2:$B$13,A2)", fmt: "money", rows: 4, req: ["SUMIFS"],
    prompt: "On <b>Summary</b>, add up the commission for each region.", hints: ["SUMIFS(Sales!$H$2:$H$13, Sales!$B$2:$B$13, A2)"] },
  { sheet: "Summary", col: "C", title: "Award Winners", f: '=COUNTIFS(Sales!$B$2:$B$13,A2,Sales!$I$2:$I$13,"Award")', fmt: "int", rows: 4, req: ["COUNTIFS"],
    prompt: "On <b>Summary</b>, count each region's award winners.", hints: ["Two criteria: region matches A2, and Award is \"Award\"."] },
];
level("Scenario 3: Everything Together", `
<p>A sales commission workbook where nearly every step mixes topics: lookups inside math, logic around statistics, and text built from text. This is the closest thing to how the exam feels.</p>
<p>Testing out is off here: work through every step.</p>`, { noTestOut: true });
scenarioProblems("p7-sales", SALES_STEPS, salesScenario, "Sales", "Sales Workbook");

// ################################################################## MORE SCENARIOS (added to every level)
// Different settings: airline, hospital, restaurant, plant, hotel, gym, car dealership, athletics.
const PART_TITLES = { 1: "Counting & Conditional Math", 2: "Logical Functions", 3: "Lookup Functions", 4: "Dates & Rounding", 5: "Text Functions", 6: "Mix It Up: Combined Formulas", 7: "Exam-Style Practice" };
function inLevel(num, name) { CUR.part = num; CUR.partTitle = PART_TITLES[num]; CUR.level = name; }

const AIRLINES = ["DL", "AA", "UA", "WN", "B6"];
const AIRPORTS = [["Birmingham", "AL", "BHM"], ["Atlanta", "GA", "ATL"], ["Dallas", "TX", "DFW"], ["Chicago", "IL", "ORD"], ["Denver", "CO", "DEN"], ["Charlotte", "NC", "CLT"], ["Orlando", "FL", "MCO"], ["Nashville", "TN", "BNA"], ["Houston", "TX", "IAH"], ["New Orleans", "LA", "MSY"]];
const CODES = AIRPORTS.map((a) => a[2]);
const DEPTS = ["Cardiology", "Orthopedics", "Pediatrics", "Neurology", "Emergency"];
const INSURERS = ["Medicare", "BCBS", "Aetna", "Self-Pay"];
const DIAG = ["Hip fracture", "Chest pain", "Wrist fracture", "Migraine", "Pneumonia", "Concussion", "Ankle fracture", "Arrhythmia", "Asthma"];
const ROOMS = ["Standard", "Deluxe", "Suite", "Family"];
const CHANNELS = ["Direct", "Expedia", "Booking.com", "Phone"];
const PLANS = ["Basic", "Plus", "Premium"];
const MODELS = [["Camry", "Toyota", 28400], ["RAV4", "Toyota", 31900], ["Civic", "Honda", 25100], ["CR-V", "Honda", 32400], ["F-150", "Ford", 41200], ["Explorer", "Ford", 38900], ["Tahoe", "Chevrolet", 58700], ["Malibu", "Chevrolet", 26300]];
const FL_H = ["Flight", "Origin", "Dest", "Date", "Delay (min)", "Passengers", "Seats"];
const HO_H = ["Patient ID", "Department", "Insurance", "Admit", "Discharge", "Charges", "Diagnosis"];
const HOT_H = ["Booking", "Room Type", "Check-in", "Nights", "Rate", "Channel"];
const PL_H = ["Line", "Shift", "Operator", "Units", "Defects", "Downtime (min)"];

function flights(d, n) {
  const t = todaySerial();
  return Array.from({ length: n }, () => {
    const o = d.pick(CODES); let de = d.pick(CODES); while (de === o) de = d.pick(CODES);
    const seats = d.pick([76, 128, 150, 180]);
    return [d.pick(AIRLINES) + d.int(100, 2999), o, de, t - d.int(1, 60), d.chance(0.6) ? d.int(-10, 15) : d.int(16, 200), d.int(Math.round(seats * 0.45), seats), seats];
  });
}
const FL_FMT = { Date: "date" };
function patients(d, n, alwaysDischarged) {
  const t = todaySerial(), idl = ids(d, n, "P", 10000, 99999);
  return idl.map((id, i) => {
    const adm = t - d.int(2, 60), stay = d.int(1, 12);
    const dis = adm + stay >= t && !alwaysDischarged ? null : (!alwaysDischarged && d.chance(0.15) ? null : adm + stay);
    return [id, i < 5 ? DEPTS[i] : d.pick(DEPTS), d.pick(INSURERS), adm, dis, d.int(18, 950) * 100 + d.int(0, 99), d.pick(DIAG)];
  });
}
const HO_FMT = { Admit: "date", Discharge: "date", Charges: "money" };
function hotel(d, n) {
  const t = todaySerial();
  return ids(d, n, "BK", 1000, 9999).map((id) => { const rt = d.pick(ROOMS); return [id, rt, t + d.int(-30, 90), d.int(1, 9), { Standard: 129, Deluxe: 169, Suite: 289, Family: 199 }[rt] + d.pick([0, 10, 20, -10]), d.pick(CHANNELS)]; });
}
const HOT_FMT = { "Check-in": "date", Rate: "money" };
function plant(d, n) {
  const ppl = people(d, 6);
  return Array.from({ length: n }, () => { const u = d.int(300, 1500); return ["L" + d.int(1, 5), d.pick(["1st", "2nd", "3rd"]), d.pick(ppl).first, u, Math.round(u * d.int(2, 60) / 1000), d.int(0, 90)]; });
}

// ---------------------------------------------------------------- Part 1
inLevel(1, "Warm-up: SUM, AVERAGE, COUNT, MAX");
add("m1-w-delay", "Longest flight delay", (R) => {
  const n = 15, s = sheet("Flights", FL_H, flights(R.d, n), FL_FMT, [["I2", "Longest Delay"]]);
  return { prompt: "An airline's operations desk wants the worst delay. In <b>J2</b>, find the longest <b>Delay (min)</b>.", sheets: [s], target: tgt("Flights", "J2", 1, null, "int"), solution: `=MAX(E2:E${n + 1})`, hints: ["Which function finds the biggest number?", `=MAX(E2:E${n + 1})`] };
}, { drill: "math" });
add("m1-w-defects", "Total defects", (R) => {
  const n = 15, s = sheet("Production", PL_H, plant(R.d, n), null, [["H2", "Total Defects"]]);
  return { prompt: "A manufacturing plant logs each production run. In <b>I2</b>, add up all of the <b>Defects</b>.", sheets: [s], target: tgt("Production", "I2", 1, null, "int"), solution: `=SUM(E2:E${n + 1})`, hints: ["Defects is column E.", `=SUM(E2:E${n + 1})`] };
}, { drill: "math" });
add("m1-w-charges", "Average patient charges", (R) => {
  const n = 15, s = sheet("Patients", HO_H, patients(R.d, n), HO_FMT, [["I2", "Avg Charges"]]);
  return { prompt: "A hospital's billing office wants the typical bill. In <b>J2</b>, find the average <b>Charges</b>.", sheets: [s], target: tgt("Patients", "J2", 1, null, "money"), solution: `=AVERAGE(F2:F${n + 1})`, hints: ["Charges is column F.", `=AVERAGE(F2:F${n + 1})`] };
}, { drill: "math" });

inLevel(1, "Absolute references ($)");
add("m1-abs-hotel", "Hotel total with tax", (R) => {
  const n = R.p.int(10, 13), rate = R.d.pick([0.09, 0.1, 0.11, 0.12, 0.13]);
  const s = sheet("Bookings", HOT_H.concat(["Total w/ Tax"]), hotel(R.d, n).map((r) => r.concat([null])), HOT_FMT, [["I1", "Occupancy Tax"], ["J1", rate, "pct"]]);
  return { prompt: "A hotel charges Nights × Rate, plus the occupancy tax in <b>J1</b>. In <b>G2</b>, find each booking's total: Nights × Rate × (1 + tax).", sheets: [s], target: tgt("Bookings", "G2", n, null, "money"), solution: "=D2*E2*(1+$J$1)", hints: ["Nights is D, Rate is E.", "Lock the tax cell: $J$1"] };
}, { drill: "math" });
add("m1-abs-tippool", "Tip pool share", (R) => {
  const n = R.p.int(6, 9), ppl = people(R.d, n), tot = n + 2;
  const s = sheet("Servers", ["Server", "Sales", "Tip Share"], ppl.map((p) => [p.first, R.d.int(400, 2200) + R.d.int(0, 99) / 100]), { Sales: "money" }, [["A" + tot, "Total"], ["F1", "Tip Pool"], ["G1", R.d.int(600, 1800), "money0"]]);
  return {
    prompt: `A restaurant splits tonight's tip pool (<b>G1</b>) by sales. B${tot} has total sales. In <b>C2</b>, find each server's share: their Sales ÷ total sales × the tip pool.`,
    sheets: [s], prefill: [{ sheet: "Servers", cell: "B" + tot, rows: 1, f: `=SUM(B2:B${n + 1})`, fmt: "money" }], target: tgt("Servers", "C2", n, null, "money"),
    solution: `=B2/$B$${tot}*$G$1`, hints: ["Two cells stay put: the total and the tip pool.", `=B2/$B$${tot}*$G$1`],
  };
}, { drill: "math" });

inLevel(1, "Mixed references");
add("m1-grid-tip", "Tip table", (R) => {
  const bills = R.d.sample([18.5, 24, 32.75, 41.2, 56, 63.4, 78.9, 95, 112.6], 6).sort((a, b) => a - b);
  const pc = R.p.pick([[0.15, 0.18, 0.2, 0.22], [0.1, 0.15, 0.2, 0.25]]);
  const s = sheet("Tips", ["Bill ↓  Tip % →"].concat(pc), bills.map((b) => [b, null, null, null, null]));
  for (let j = 1; j <= 4; j++) s.fmt["0:" + j] = "pct";
  for (let i = 1; i <= 6; i++) s.fmt[i + ":0"] = "money";
  return { prompt: "Make a tip table for a restaurant's servers. Bills are down column A and tip percentages are across row 1. In <b>B2</b>, write one formula (Bill × Tip %) for the whole grid B2:E7.", sheets: [s], target: tgt("Tips", "B2", 6, null, "money", 4), solution: "=$A2*B$1", hints: ["The bill is always in column A; the percent is always in row 1.", "=$A2*B$1"] };
}, { drill: "math" });
add("m1-grid-cal", "Calories burned grid", (R) => {
  const w = R.d.sample([120, 135, 150, 165, 180, 200, 220, 240], 6).sort((a, b) => a - b);
  const mins = R.p.pick([[15, 30, 45, 60], [10, 20, 30, 40]]);
  const s = sheet("Calories", ["Weight ↓  Minutes →"].concat(mins), w.map((x) => [x, null, null, null, null]), null, [["G1", "Cal per lb-min"], ["H1", R.d.pick([0.04, 0.05, 0.06, 0.075, 0.09])]]);
  return { prompt: "A gym wants a calories-burned chart: Weight × Minutes × the rate in <b>H1</b>. Weights are down column A and minutes are across row 1. In <b>B2</b>, write one formula for B2:E7.", sheets: [s], target: tgt("Calories", "B2", 6, null, "int", 4), solution: "=$A2*B$1*$H$1", hints: ["Three references: one locks the column, one locks the row, one locks both.", "=$A2*B$1*$H$1"] };
}, { drill: "math" });

inLevel(1, "COUNTIFS");
add("m1-cif-hosp", "Patients by department & insurer", (R) => {
  const n = 18, rows = patients(R.d, n), dep = R.p.pick(DEPTS), ins = R.p.pick(INSURERS);
  rows[6][1] = dep; rows[6][2] = ins; rows[7][1] = dep; rows[7][2] = ins;
  const s = sheet("Patients", HO_H, rows, HO_FMT, [["J1", "Department"], ["J2", dep], ["K1", "Insurance"], ["K2", ins], ["L1", "Patients"]]);
  return { prompt: `In <b>L2</b>, count patients in <b>${dep}</b> whose insurance is <b>${ins}</b>.`, sheets: [s], target: tgt("Patients", "L2", 1, null, "int"), solution: `=COUNTIFS(B2:B${n + 1},J2,C2:C${n + 1},K2)`, hints: ["Two pairs: Department (B) and Insurance (C).", `=COUNTIFS(B2:B${n + 1}, J2, C2:C${n + 1}, K2)`] };
}, { drill: "math" });
add("m1-cif-delay", "Delayed flights from an airport", (R) => {
  const n = 18, rows = flights(R.d, n), o = R.p.pick(CODES);
  rows[0][1] = o; rows[0][4] = 95; rows[1][1] = o; if (rows[1][2] === o) rows[1][2] = o === "BHM" ? "ATL" : "BHM";
  if (rows[0][2] === o) rows[0][2] = o === "BHM" ? "ATL" : "BHM";
  const s = sheet("Flights", FL_H, rows, FL_FMT, [["I1", "Origin"], ["I2", o], ["J1", "Delay Over"], ["J2", R.d.pick([15, 30, 45, 60])], ["K1", "Flights"]]);
  return { prompt: `In <b>K2</b>, count flights leaving <b>${o}</b> (I2) with a delay <b>greater than</b> the minutes in J2.`, sheets: [s], target: tgt("Flights", "K2", 1, null, "int"), solution: `=COUNTIFS(B2:B${n + 1},I2,E2:E${n + 1},">"&J2)`, hints: ["Origin is B, Delay is E.", `The second criteria is ">"&J2.`] };
}, { drill: "math" });
add("m1-cif-gym", "Members per plan (summary)", (R) => {
  const n = 18, ppl = people(R.d, n), t = todaySerial();
  const mem = sheet("Members", ["Member", "Plan", "Join Date", "Visits"], ppl.map((p, i) => [p.first + " " + p.last, i < 3 ? PLANS[i] : R.d.pick(PLANS), t - R.d.int(10, 2500), R.d.int(0, 25)]), { "Join Date": "date" });
  const sum = sheet("Summary", ["Plan", "Members"], PLANS.map((p) => [p]));
  return { prompt: "A gym wants a head count by plan. On <b>Summary</b>, in <b>B2</b>, count the members on the Members sheet with the plan in A2. It will be copied down to B4.", sheets: [sum, mem], target: tgt("Summary", "B2", 3, null, "int"), solution: `=COUNTIFS(Members!$B$2:$B$${n + 1},A2)`, hints: ["The range is Members!B2:B19.", "Lock it (or use Members!B:B) so it doesn't slide."] };
}, { drill: "math" });

inLevel(1, "SUMIFS & AVERAGEIFS");
add("m1-sif-tips", "Dinner tips by server (summary)", (R) => {
  const n = 18, srv = people(R.d, 4).map((p) => p.first);
  const rows = ids(R.d, n, "T", 100, 999).map((id, i) => { const sub = R.d.money(20, 180); return [id, srv[i % 4], i < 4 ? "Dinner" : R.d.pick(["Lunch", "Dinner"]), sub, Math.round(sub * R.d.int(10, 25)) / 100]; });
  const tk = sheet("Tickets", ["Ticket", "Server", "Shift", "Subtotal", "Tip"], rows, { Subtotal: "money", Tip: "money" });
  const sum = sheet("Summary", ["Server", "Dinner Tips"], srv.map((x) => [x]));
  return { prompt: "On <b>Summary</b>, in <b>B2</b>, add up each server's Tips for the <b>Dinner</b> shift only. It will be copied down to B5.", sheets: [sum, tk], target: tgt("Summary", "B2", 4, null, "money"), solution: `=SUMIFS(Tickets!$E$2:$E$${n + 1},Tickets!$B$2:$B$${n + 1},A2,Tickets!$C$2:$C$${n + 1},"Dinner")`, hints: ["Sum range: Tickets!E (Tip). Criteria: Server = A2 and Shift = \"Dinner\".", "Lock all three ranges."] };
}, { drill: "math" });
add("m1-aif-hosp", "Average charges by department", (R) => {
  const n = 18, p = sheet("Patients", HO_H, patients(R.d, n), HO_FMT);
  const sum = sheet("Summary", ["Department", "Avg Charges"], DEPTS.map((x) => [x]));
  return { prompt: "On <b>Summary</b>, in <b>B2</b>, find the average Charges for the department in A2. It will be copied down to B6.", sheets: [sum, p], target: tgt("Summary", "B2", 5, null, "money"), solution: `=AVERAGEIFS(Patients!$F$2:$F$${n + 1},Patients!$B$2:$B$${n + 1},A2)`, hints: ["Average Patients!F where Patients!B matches A2.", "Lock both ranges."] };
}, { drill: "math" });
add("m1-aif-flight", "Average delay on full flights", (R) => {
  const n = 18, rows = flights(R.d, n), dst = R.p.pick(CODES);
  rows[0][2] = dst; rows[0][5] = rows[0][6]; if (rows[0][1] === dst) rows[0][1] = dst === "BHM" ? "ATL" : "BHM";
  const s = sheet("Flights", FL_H, rows, FL_FMT, [["I1", "Destination"], ["I2", dst], ["J1", "Min Passengers"], ["J2", R.d.int(60, 90)], ["K1", "Avg Delay"]]);
  return { prompt: `In <b>K2</b>, find the average Delay for flights going to <b>${dst}</b> (I2) that carried at least the number of passengers in J2.`, sheets: [s], target: tgt("Flights", "K2", 1, null, "dec1"), solution: `=AVERAGEIFS(E2:E${n + 1},C2:C${n + 1},I2,F2:F${n + 1},">="&J2)`, hints: ["Average range first: Delay (E).", `Then Dest (C) = I2 and Passengers (F) ">="&J2.`] };
}, { drill: "math" });

inLevel(1, "Wildcards");
add("m1-wild-airline", "Flights for one airline", (R) => {
  const n = 18, rows = flights(R.d, n), al = R.p.pick(AIRLINES);
  rows[0][0] = al + "1234";
  const s = sheet("Flights", FL_H, rows, FL_FMT, [["I1", "Airline"], ["I2", al], ["J1", "Flights"]]);
  return { prompt: `Flight numbers start with a 2-letter airline code. In <b>J2</b>, count the flights for <b>${al}</b>.`, sheets: [s], target: tgt("Flights", "J2", 1, null, "int"), solution: `=COUNTIFS(A2:A${n + 1},I2&"*")`, hints: ["Starts with: the code followed by *.", `=COUNTIFS(A2:A${n + 1}, I2&"*")`] };
}, { drill: "math" });
add("m1-wild-fracture", "Charges for any fracture", (R) => {
  const n = 18, rows = patients(R.d, n); rows[0][6] = "Hip fracture";
  const s = sheet("Patients", HO_H, rows, HO_FMT, [["I1", "Keyword"], ["I2", "fracture"], ["J1", "Charges"]]);
  return { prompt: "In <b>J2</b>, add up the Charges for every patient whose Diagnosis <b>contains</b> the word in I2 (hip, wrist, ankle...).", sheets: [s], target: tgt("Patients", "J2", 1, null, "money"), solution: `=SUMIFS(F2:F${n + 1},G2:G${n + 1},"*"&I2&"*")`, hints: ["Contains: * on both sides.", `=SUMIFS(F2:F${n + 1}, G2:G${n + 1}, "*"&I2&"*")`] };
}, { drill: "math" });

inLevel(1, "SUMPRODUCT");
add("m1-sp-hotel", "Room revenue", (R) => {
  const n = 14, s = sheet("Bookings", HOT_H, hotel(R.d, n), HOT_FMT, [["H1", "Room Revenue"]]);
  return { prompt: "In <b>H2</b>, use SUMPRODUCT to find the hotel's total room revenue (Nights × Rate for every booking).", sheets: [s], target: tgt("Bookings", "H2", 1, null, "money"), solution: `=SUMPRODUCT(D2:D${n + 1},E2:E${n + 1})`, require: ["SUMPRODUCT"], hints: ["Nights is D, Rate is E.", `=SUMPRODUCT(D2:D${n + 1}, E2:E${n + 1})`] };
}, { drill: "math" });
add("m1-sp-gpa", "GPA from a transcript", (R) => {
  const courses = R.d.sample(["OM 305", "EC 110", "MKT 300", "FI 302", "ST 260", "MGT 300", "LGS 200", "AC 210", "EN 102"], 6);
  const s = sheet("Transcript", ["Course", "Credits", "Grade Points"], courses.map((c) => [c, R.d.pick([3, 3, 3, 4, 1]), R.d.pick([4, 4, 3, 3, 2, 3.7, 3.3])]), null, [["E1", "GPA"]]);
  return { prompt: "In <b>E2</b>, calculate the GPA: the sum of Credits × Grade Points (use SUMPRODUCT), divided by the total Credits.", sheets: [s], target: tgt("Transcript", "E2", 1, null, "dec2"), solution: "=SUMPRODUCT(B2:B7,C2:C7)/SUM(B2:B7)", require: ["SUMPRODUCT"], hints: ["Quality points: SUMPRODUCT(B2:B7, C2:C7).", "Divide by SUM(B2:B7)."] };
}, { drill: "math" });

// ---------------------------------------------------------------- Part 2
inLevel(2, "IF basics");
add("m2-if-delay", "Delayed or on time", (R) => {
  const n = R.p.int(10, 13), s = sheet("Flights", FL_H.concat(["Status"]), flights(R.d, n).map((r) => r.concat([null])), FL_FMT);
  return { prompt: "In <b>H2</b>, show <b>\"Delayed\"</b> if the Delay is <b>more than 15</b> minutes, otherwise <b>\"On Time\"</b>.", sheets: [s], target: tgt("Flights", "H2", n), solution: '=IF(E2>15,"Delayed","On Time")', hints: ["Delay is column E.", '=IF(E2>15, "Delayed", "On Time")'] };
}, { drill: "logic" });
add("m2-if-ot", "Overtime pay", (R) => {
  const n = R.p.int(10, 13), ppl = people(R.d, n);
  const s = sheet("Timecards", ["Employee", "Hourly Rate", "Hours", "OT Pay"], ppl.map((p) => [p.first + " " + p.last, R.d.pick([15, 16.5, 18, 19.25, 21, 24]), R.d.int(30, 52)]), { "Hourly Rate": "money" });
  s.rows[1][2] = 40;
  return { prompt: "Plant workers earn overtime for hours <b>over 40</b>: (Hours − 40) × Hourly Rate × 1.5. In <b>D2</b>, calculate OT Pay (0 if they worked 40 or fewer).", sheets: [s], target: tgt("Timecards", "D2", n, null, "money"), solution: "=IF(C2>40,(C2-40)*B2*1.5,0)", hints: ["The true part is math: (C2-40)*B2*1.5", "=IF(C2>40, (C2-40)*B2*1.5, 0)"] };
}, { drill: "logic" });

inLevel(2, "IF with AND / OR");
add("m2-and-allconf", "All-Conference", (R) => {
  const n = R.p.int(10, 13), ppl = people(R.d, n);
  const s = sheet("Players", ["Player", "Class", "Games", "PPG", "Award"], ppl.map((p) => [p.first + " " + p.last, R.d.pick(["FR", "SO", "JR", "SR"]), R.d.int(4, 30), R.d.int(20, 240) / 10]), { PPG: "dec1" }, [["G1", "Min Games"], ["H1", R.d.pick([10, 12, 15])], ["G2", "Min PPG"], ["H2", R.d.pick([12, 14, 15])]]);
  s.rows[1][2] = 25; s.rows[1][3] = 22.5;
  return { prompt: "A college basketball team's All-Conference nominees played at least the games in <b>H1</b> <b>and</b> averaged at least the points per game in <b>H2</b>. In <b>E2</b>, show <b>\"All-Conference\"</b> or leave it blank.", sheets: [s], target: tgt("Players", "E2", n), solution: '=IF(AND(C2>=$H$1,D2>=$H$2),"All-Conference","")', hints: ["Both conditions: AND(C2>=$H$1, D2>=$H$2)", "Lock H1 and H2."] };
}, { drill: "logic" });
add("m2-or-billing", "Billing review (OR)", (R) => {
  const n = R.p.int(10, 13), rows = patients(R.d, n); rows[0][2] = "Self-Pay";
  const s = sheet("Patients", HO_H.concat(["Billing Flag"]), rows.map((r) => r.concat([null])), HO_FMT, [["J1", "Review Over"], ["K1", R.d.pick([25000, 40000, 60000]), "money0"]]);
  return { prompt: "Billing reviews a patient's account if Charges are <b>over</b> the amount in <b>K1</b> <b>or</b> the insurance is <b>\"Self-Pay\"</b>. In <b>H2</b>, show <b>\"Review\"</b> or leave it blank.", sheets: [s], target: tgt("Patients", "H2", n), solution: '=IF(OR(F2>$K$1,C2="Self-Pay"),"Review","")', hints: ["Either is enough, so use OR.", '=IF(OR(F2>$K$1, C2="Self-Pay"), "Review", "")'] };
}, { drill: "logic" });

inLevel(2, "Nested IF");
add("m2-nest-room", "Free upgrade tier", (R) => {
  const n = R.p.int(10, 13), s = sheet("Bookings", HOT_H.concat(["Upgrade"]), hotel(R.d, n).map((r) => r.concat([null])), HOT_FMT);
  return { prompt: "A hotel upgrades long stays: <b>7 or more</b> nights gets <b>\"Suite\"</b>, <b>3 or more</b> gets <b>\"Deluxe\"</b>, and shorter stays get <b>\"Standard\"</b>. Fill in <b>G2</b>.", sheets: [s], target: tgt("Bookings", "G2", n), solution: '=IF(D2>=7,"Suite",IF(D2>=3,"Deluxe","Standard"))', hints: ["Start with the biggest tier: IF(D2>=7, \"Suite\", ...)", '=IF(D2>=7,"Suite",IF(D2>=3,"Deluxe","Standard"))'] };
}, { drill: "logic" });
add("m2-nest-delay", "Delay category", (R) => {
  const n = R.p.int(10, 13), s = sheet("Flights", FL_H.concat(["Category"]), flights(R.d, n).map((r) => r.concat([null])), FL_FMT);
  return { prompt: "In <b>H2</b>, label each flight's Delay: <b>0 or less</b> is <b>\"On Time\"</b>, up to <b>30</b> is <b>\"Minor\"</b>, up to <b>120</b> is <b>\"Major\"</b>, and anything longer is <b>\"Severe\"</b>.", sheets: [s], target: tgt("Flights", "H2", n), solution: '=IF(E2<=0,"On Time",IF(E2<=30,"Minor",IF(E2<=120,"Major","Severe")))', hints: ["Three IFs, smallest tier first with <=.", '=IF(E2<=0,"On Time",IF(E2<=30,"Minor",IF(E2<=120,"Major","Severe")))'] };
}, { drill: "logic" });

inLevel(2, "IS functions & IFERROR");
add("m2-ie-defect", "Defect rate (IFERROR)", (R) => {
  const n = R.p.int(10, 13), rows = plant(R.d, n); rows[1][3] = 0; rows[1][4] = 0;
  const s = sheet("Production", PL_H.concat(["Defect Rate"]), rows.map((r) => r.concat([null])));
  return { prompt: "In <b>G2</b>, calculate the defect rate (Defects ÷ Units). Some lines didn't run (0 units). Show <b>\"No Run\"</b> for those instead of an error.", sheets: [s], target: tgt("Production", "G2", n, null, "pct1"), solution: '=IFERROR(E2/D2,"No Run")', require: ["IFERROR"], hints: ["=IFERROR(E2/D2, \"No Run\")"] };
}, { drill: "logic" });
add("m2-isb-admit", "Still admitted? (ISBLANK)", (R) => {
  const n = R.p.int(10, 13), rows = patients(R.d, n); rows[0][4] = null;
  const s = sheet("Patients", HO_H.concat(["Status"]), rows.map((r) => r.concat([null])), HO_FMT);
  return { prompt: "In <b>H2</b>, show <b>\"Admitted\"</b> if the Discharge date is empty, otherwise <b>\"Discharged\"</b>. Use ISBLANK.", sheets: [s], target: tgt("Patients", "H2", n), solution: '=IF(ISBLANK(E2),"Admitted","Discharged")', require: ["ISBLANK"], hints: ["Discharge is column E.", '=IF(ISBLANK(E2), "Admitted", "Discharged")'] };
}, { drill: "logic" });

inLevel(2, "Logic in Excel Tables");
add("m2-tbl-late", "Late flights in a Table", (R) => {
  const n = R.p.int(10, 13);
  return { prompt: "The flights are a Table named <b>Flights</b>. In <b>H2</b>, show <b>\"Late\"</b> if the delay is over 15 minutes, otherwise <b>\"On Time\"</b>. Use a structured reference. The column name has a space, so it's <code>[@[Delay (min)]]</code>.", sheets: [sheet("Flights", FL_H.concat(["Status"]), flights(R.d, n).map((r) => r.concat([null])), FL_FMT)], tables: [{ name: "Flights", sheet: "Flights", top: 0, left: 0, ncols: 8, nrows: n }], target: tgt("Flights", "H2", n), requireStructured: true, solution: '=IF([@[Delay (min)]]>15,"Late","On Time")', hints: ["[@[Delay (min)]] is this row's delay.", '=IF([@[Delay (min)]]>15, "Late", "On Time")'] };
}, { drill: "logic" });
add("m2-tbl-vip", "VIP members in a Table", (R) => {
  const n = R.p.int(10, 13), ppl = people(R.d, n), t = todaySerial();
  const rows = ppl.map((p) => [p.first + " " + p.last, R.d.pick(PLANS), t - R.d.int(10, 2000), R.d.int(0, 25), null]); rows[0][1] = "Premium"; rows[0][3] = 20;
  return { prompt: "Table <b>Members</b>: a member is a <b>\"VIP\"</b> if their Plan is <b>\"Premium\"</b> <b>and</b> they have <b>12 or more</b> Visits. Otherwise leave it blank. Fill in <b>E2</b> with structured references.", sheets: [sheet("Members", ["Member", "Plan", "Join Date", "Visits", "VIP"], rows, { "Join Date": "date" })], tables: [{ name: "Members", sheet: "Members", top: 0, left: 0, ncols: 5, nrows: n }], target: tgt("Members", "E2", n), requireStructured: true, solution: '=IF(AND([@Plan]="Premium",[@Visits]>=12),"VIP","")', hints: ["[@Plan] and [@Visits]", '=IF(AND([@Plan]="Premium", [@Visits]>=12), "VIP", "")'] };
}, { drill: "logic" });

// ---------------------------------------------------------------- Part 3
inLevel(3, "VLOOKUP: exact match");
add("m3-vl-floor", "Department floor", (R) => {
  const n = R.p.int(10, 13), floors = R.p.sample([2, 3, 4, 5, 6, 7], 5);
  const dp = sheet("Depts", ["Department", "Floor", "Phone Ext"], DEPTS.map((d, i) => [d, floors[i], 4000 + i * 11]));
  const s = sheet("Patients", HO_H.concat(["Floor"]), patients(R.d, n).map((r) => r.concat([null])), HO_FMT);
  return { prompt: "In <b>H2</b>, look up the <b>Floor</b> for each patient's Department on the <b>Depts</b> sheet.", sheets: [s, dp], target: tgt("Patients", "H2", n), solution: "=VLOOKUP(B2,Depts!$A$2:$C$6,2,FALSE)", require: ["VLOOKUP"], hints: ["Look up B2 in Depts!$A$2:$C$6.", "Floor is column 2; use FALSE."] };
}, { drill: "lookup" });
add("m3-vl-car", "Sale price after discount", (R) => {
  const n = R.p.int(10, 13);
  const md = sheet("Models", ["Model", "Make", "MSRP"], MODELS.map((m) => [m[0], m[1], m[2] + R.d.int(0, 9) * 100]), { MSRP: "money0" });
  const inv = sheet("Inventory", ["Stock #", "Model", "Discount", "Sale Price"], ids(R.d, n, "S", 1000, 9999).map((id) => [id, R.d.pick(MODELS)[0], R.d.pick([0, 0.02, 0.03, 0.05, 0.08])]), { Discount: "pct" });
  return { prompt: "A car dealership prices each car at its MSRP (from the <b>Models</b> sheet) minus its Discount. In <b>D2</b>: VLOOKUP the MSRP × (1 − Discount).", sheets: [inv, md], target: tgt("Inventory", "D2", n, null, "money0"), solution: "=VLOOKUP(B2,Models!$A$2:$C$9,3,FALSE)*(1-C2)", require: ["VLOOKUP"], hints: ["MSRP: VLOOKUP(B2, Models!$A$2:$C$9, 3, FALSE)", "Multiply by (1-C2)."] };
}, { drill: "lookup" });

inLevel(3, "VLOOKUP: approximate match");
add("m3-va-latefee", "Late fee tiers", (R) => tierProblem(R, {
  tsheet: "Fees", theads: ["Days Late", "Fee"], tiers: [[0, 0], [1, 5], [15, 15], [30, 35]], tfmt: { Fee: "money0" },
  sheet: "Rentals", heads: ["Rental", "Days Late", "Late Fee"], fmt: {},
  row: (R2) => ["EQ-" + R2.d.int(100, 999), R2.d.pick([0, 0, R2.d.int(1, 45)]), null], cell: "C2", outfmt: "money0",
  prompt: "An equipment rental shop charges late fees by tier (see <b>Fees</b>, which lists each tier's starting number of days). In <b>C2</b>, look up each rental's fee.",
  solution: "=VLOOKUP(B2,Fees!$A$2:$B$5,2,TRUE)", hints: ["Approximate match on Days Late.", "=VLOOKUP(B2, Fees!$A$2:$B$5, 2, TRUE)"],
}), { drill: "lookup" });
add("m3-va-scholar", "Scholarship tier by GPA", (R) => tierProblem(R, {
  tsheet: "Tiers", theads: ["Min GPA", "Scholarship"], tiers: [[0, "None"], [2.5, "Bronze"], [3, "Silver"], [3.5, "Gold"]],
  sheet: "Athletes", heads: ["Athlete", "GPA", "Scholarship"], fmt: { GPA: "dec2" },
  row: (R2) => { const p = people(R2.d, 1)[0]; return [p.first + " " + p.last, R2.d.int(200, 400) / 100, null]; }, cell: "C2",
  prompt: "Athletic academic scholarships depend on GPA (see <b>Tiers</b>). In <b>C2</b>, look up each athlete's tier with an approximate-match VLOOKUP.",
  solution: "=VLOOKUP(B2,Tiers!$A$2:$B$5,2,TRUE)", hints: ["Look up B2 in Tiers!$A$2:$B$5.", "Column 2, TRUE."],
}), { drill: "lookup" });

inLevel(3, "INDEX / MATCH");
add("m3-im-city", "City from airport code", (R) => {
  const ap = sheet("Airports", ["City", "State", "Code"], R.d.shuffle(AIRPORTS));
  const n = R.p.int(10, 13);
  const fl = sheet("Flights", ["Flight", "Dest Code", "Dest City"], Array.from({ length: n }, () => [R.d.pick(AIRLINES) + R.d.int(100, 2999), R.d.pick(CODES), null]));
  return { prompt: "On the <b>Airports</b> sheet, the code is in the <b>last</b> column. In <b>C2</b>, use INDEX/MATCH to show each flight's destination city.", sheets: [fl, ap], target: tgt("Flights", "C2", n), require: ["INDEX", "MATCH"], forbid: ["VLOOKUP"], solution: "=INDEX(Airports!$A$2:$A$11,MATCH(B2,Airports!$C$2:$C$11,0))", hints: ["MATCH B2 in Airports!$C$2:$C$11 (exact, 0).", "INDEX Airports!$A$2:$A$11 at that spot."] };
}, { drill: "lookup" });
add("m3-im-doctor", "Doctor by ID", (R) => {
  const k = 8, docs = people(R.d, k), idl = ids(R.d, k, "D", 100, 999);
  const ds = sheet("Doctors", ["Doctor", "Specialty", "Doctor ID"], docs.map((p, i) => ["Dr. " + p.last, DEPTS[i % 5], idl[i]]));
  const n = R.p.int(10, 13);
  const pt = sheet("Visits", ["Patient ID", "Doctor ID", "Doctor"], ids(R.d, n, "P", 10000, 99999).map((id) => [id, R.d.pick(idl), null]));
  return { prompt: "In <b>C2</b>, use INDEX/MATCH to bring in the doctor's name from the <b>Doctors</b> sheet by Doctor ID.", sheets: [pt, ds], target: tgt("Visits", "C2", n), require: ["INDEX", "MATCH"], forbid: ["VLOOKUP"], solution: `=INDEX(Doctors!$A$2:$A$${k + 1},MATCH(B2,Doctors!$C$2:$C$${k + 1},0))`, hints: ["The ID is the last column of Doctors.", "=INDEX(Doctors!$A$2:$A$9, MATCH(B2, Doctors!$C$2:$C$9, 0))"] };
}, { drill: "lookup" });

inLevel(3, "INDEX / MATCH / MATCH");
add("m3-imm-season", "Room rate by season", (R) => {
  const seasons = ["Winter", "Spring", "Summer", "Fall"];
  const tbl = sheet("Rates", ["Room Type"].concat(seasons), ROOMS.map((r, i) => [r].concat(seasons.map((s, j) => 119 + i * 45 + [0, 20, 60, 15][j] + R.d.pick([0, 5, 10])))));
  for (let i = 1; i <= 4; i++) for (let j = 1; j <= 4; j++) tbl.fmt[i + ":" + j] = "money0";
  const n = R.p.int(10, 13);
  const bk = sheet("Bookings", ["Booking", "Room Type", "Season", "Nightly Rate"], ids(R.d, n, "BK", 1000, 9999).map((id) => [id, R.d.pick(ROOMS), R.d.pick(seasons), null]));
  return { prompt: "In <b>D2</b>, use INDEX/MATCH/MATCH to find each booking's nightly rate on the <b>Rates</b> sheet by Room Type (down) and Season (across).", sheets: [bk, tbl], target: tgt("Bookings", "D2", n, null, "money0"), require: ["INDEX", "MATCH"], solution: "=INDEX(Rates!$B$2:$E$5,MATCH(B2,Rates!$A$2:$A$5,0),MATCH(C2,Rates!$B$1:$E$1,0))", hints: ["Values: Rates!$B$2:$E$5.", "Match room type in A2:A5 and season in B1:E1."] };
}, { drill: "lookup" });
add("m3-imm-miles", "Distance between airports", (R) => {
  const c = R.p.sample(CODES, 5);
  const m = c.map((a, i) => [a].concat(c.map((b, j) => (i === j ? 0 : 150 + ((i + 1) * (j + 1) * 97 + (i + j) * 61) % 1300))));
  for (let i = 0; i < 5; i++) for (let j = 0; j < i; j++) m[i][j + 1] = m[j][i + 1];
  const tbl = sheet("Miles", ["From \\ To"].concat(c), m);
  const n = R.p.int(10, 13);
  const tr = sheet("Trips", ["Trip", "From", "To", "Miles"], Array.from({ length: n }, () => { const f = R.d.pick(c); let t = R.d.pick(c); while (t === f) t = R.d.pick(c); return ["TR" + R.d.int(100, 999), f, t, null]; }));
  return { prompt: "The <b>Miles</b> sheet is a distance chart. In <b>D2</b>, use INDEX/MATCH/MATCH to find each trip's miles (From down the side, To across the top).", sheets: [tr, tbl], target: tgt("Trips", "D2", n, null, "int"), require: ["INDEX", "MATCH"], solution: "=INDEX(Miles!$B$2:$F$6,MATCH(B2,Miles!$A$2:$A$6,0),MATCH(C2,Miles!$B$1:$F$1,0))", hints: ["Values: Miles!$B$2:$F$6", "MATCH(B2, Miles!$A$2:$A$6, 0) and MATCH(C2, Miles!$B$1:$F$1, 0)"] };
}, { drill: "lookup" });

inLevel(3, "Lookups with IF & IFERROR");
add("m3-ie-sold", "Sold or in stock", (R) => {
  const n = R.p.int(10, 13), stock = ids(R.d, n, "S", 1000, 9999), sold = R.d.sample(stock, 6);
  const ppl = people(R.d, 6);
  const sl = sheet("Sales", ["Stock #", "Buyer"], sold.map((s2, i) => [s2, ppl[i].first + " " + ppl[i].last]));
  const inv = sheet("Inventory", ["Stock #", "Model", "Buyer"], stock.map((s2) => [s2, R.d.pick(MODELS)[0], null]));
  return { prompt: "In <b>C2</b>, look up each car's Buyer on the <b>Sales</b> sheet. Cars that haven't sold aren't there, so show <b>\"In Stock\"</b> for those.", sheets: [inv, sl], target: tgt("Inventory", "C2", n), require: ["IFERROR", "VLOOKUP"], solution: '=IFERROR(VLOOKUP(A2,Sales!$A$2:$B$7,2,FALSE),"In Stock")', hints: ["VLOOKUP(A2, Sales!$A$2:$B$7, 2, FALSE)", 'Wrap it: =IFERROR(..., "In Stock")'] };
}, { drill: "lookup" });
add("m3-if-perk", "Loyalty perk", (R) => {
  const pk = sheet("Perks", ["Tier", "Perk"], [["Silver", "Late checkout"], ["Gold", "Free breakfast"], ["Platinum", "Suite upgrade"]]);
  const n = R.p.int(10, 13), ppl = people(R.d, n);
  const g = sheet("Guests", ["Guest", "Member?", "Tier", "Perk"], ppl.map((p) => { const m = R.d.chance(0.6); return [p.first + " " + p.last, m ? "Y" : "N", m ? R.d.pick(["Silver", "Gold", "Platinum"]) : null, null]; }));
  return { prompt: "If the guest is a loyalty member (B is <b>\"Y\"</b>), look up their tier's perk on the <b>Perks</b> sheet. Otherwise show <b>\"None\"</b>. Fill in <b>D2</b>.", sheets: [g, pk], target: tgt("Guests", "D2", n), require: ["IF", "VLOOKUP"], solution: '=IF(B2="Y",VLOOKUP(C2,Perks!$A$2:$B$4,2,FALSE),"None")', hints: ['=IF(B2="Y", VLOOKUP(...), "None")', "VLOOKUP(C2, Perks!$A$2:$B$4, 2, FALSE)"] };
}, { drill: "lookup" });

// ---------------------------------------------------------------- Part 4
inLevel(4, "YEAR, MONTH, DAY");
add("m4-fiscal", "Fiscal year", (R) => {
  const n = R.p.int(10, 13), base = D(2024, 1, 1);
  const s = sheet("Purchases", ["PO", "Date", "Amount", "Fiscal Year"], ids(R.d, n, "PO", 1000, 9999).map((id) => [id, base + R.d.int(0, 900), R.d.money(50, 5000)]), { Date: "date", Amount: "money" });
  s.rows[1][1] = D(2025, 10, 3);
  return { prompt: "The university's fiscal year starts <b>October 1</b>. Purchases in October–December count toward <b>next</b> year (October 2025 is fiscal year 2026). In <b>D2</b>, find each purchase's fiscal year.", sheets: [s], target: tgt("Purchases", "D2", n), require: ["MONTH", "YEAR"], solution: "=IF(MONTH(B2)>=10,YEAR(B2)+1,YEAR(B2))", hints: ["Test the month: MONTH(B2)>=10", "=IF(MONTH(B2)>=10, YEAR(B2)+1, YEAR(B2))"] };
}, { drill: "dates" });
add("m4-summer", "Summer sign-ups", (R) => {
  const n = R.p.int(10, 13), ppl = people(R.d, n), base = D(2023, 1, 1);
  const s = sheet("Members", ["Member", "Plan", "Join Date", "Promo"], ppl.map((p) => [p.first + " " + p.last, R.d.pick(PLANS), base + R.d.int(0, 1000)]), { "Join Date": "date" });
  s.rows[1][2] = D(2024, 7, 4);
  return { prompt: "A gym ran a summer promotion. In <b>D2</b>, show <b>\"Summer\"</b> if the member joined in June, July, or August (any year), otherwise leave it blank.", sheets: [s], target: tgt("Members", "D2", n), require: ["MONTH"], solution: '=IF(OR(MONTH(C2)=6,MONTH(C2)=7,MONTH(C2)=8),"Summer","")', hints: ["MONTH(C2) gives 1–12.", "OR(MONTH(C2)=6, MONTH(C2)=7, MONTH(C2)=8). (MONTH(C2)>=6 with AND <=8 works too.)"] };
}, { drill: "dates" });

inLevel(4, "TODAY");
add("m4-lot", "Days on the lot", (R) => {
  const n = R.p.int(10, 13), t = todaySerial();
  const s = sheet("Inventory", ["Stock #", "Model", "Arrived", "Days on Lot"], ids(R.d, n, "S", 1000, 9999).map((id) => [id, R.d.pick(MODELS)[0], t - R.d.int(1, 200)]), { Arrived: "date" });
  return { prompt: "In <b>D2</b>, find how many days each car has been on the dealership lot (today minus Arrived).", sheets: [s], target: tgt("Inventory", "D2", n, null, "int"), require: ["TODAY"], solution: "=TODAY()-C2", hints: ["Subtracting dates gives days.", "=TODAY()-C2"] };
}, { drill: "dates" });
add("m4-age", "Age this year", (R) => {
  const n = R.p.int(10, 13), ppl = people(R.d, n);
  const s = sheet("Members", ["Member", "Birth Date", "Age This Year"], ppl.map((p) => [p.first + " " + p.last, D(R.d.int(1950, 2007), R.d.int(1, 12), R.d.int(1, 28))]), { "Birth Date": "date" });
  return { prompt: "In <b>C2</b>, find the age each member turns <b>this year</b>: this year minus their birth year.", sheets: [s], target: tgt("Members", "C2", n), require: ["YEAR", "TODAY"], solution: "=YEAR(TODAY())-YEAR(B2)", hints: ["This year is YEAR(TODAY()).", "=YEAR(TODAY())-YEAR(B2)"] };
}, { drill: "dates" });
add("m4-renew", "Renew soon", (R) => {
  const n = R.p.int(10, 13), ppl = people(R.d, n), t = todaySerial();
  const s = sheet("Members", ["Member", "Expires", "Reminder"], ppl.map((p) => [p.first + " " + p.last, t + R.d.int(-20, 120)]), { Expires: "date" });
  s.rows[1][1] = t + 12;
  return { prompt: "In <b>C2</b>, show <b>\"Renew Soon\"</b> if the membership expires within <b>30 days</b> of today (or has already expired). That's when Expires − today is 30 or less. Otherwise leave it blank.", sheets: [s], target: tgt("Members", "C2", n), require: ["TODAY"], solution: '=IF(B2-TODAY()<=30,"Renew Soon","")', hints: ["Days left: B2-TODAY()", '=IF(B2-TODAY()<=30, "Renew Soon", "")'] };
}, { drill: "dates" });

inLevel(4, "ROUND");
add("m4-load", "Load factor", (R) => {
  const n = R.p.int(10, 13), f = flights(R.d, n);
  const s = sheet("Flights", ["Flight", "Passengers", "Seats", "Load Factor"], f.map((r) => [r[0], r[5], r[6]]));
  return { prompt: "An airline's load factor is Passengers ÷ Seats. In <b>D2</b>, calculate it rounded to <b>3 decimal places</b> (like 0.873).", sheets: [s], target: tgt("Flights", "D2", n, null, "pct1"), require: ["ROUND"], solution: "=ROUND(B2/C2,3)", hints: ["Put the division inside ROUND.", "=ROUND(B2/C2, 3)"] };
}, { drill: "dates" });
add("m4-perday", "Charges per day", (R) => {
  const n = R.p.int(10, 13), rows = patients(R.d, n, true);
  const s = sheet("Patients", ["Patient ID", "Admit", "Discharge", "Charges", "Per Day"], rows.map((r) => [r[0], r[3], r[4], r[5]]), { Admit: "date", Discharge: "date", Charges: "money" });
  return { prompt: "In <b>E2</b>, find each patient's charges per day: Charges ÷ (Discharge − Admit), rounded to <b>2 decimals</b>.", sheets: [s], target: tgt("Patients", "E2", n, null, "money"), require: ["ROUND"], solution: "=ROUND(D2/(C2-B2),2)", hints: ["Parentheses around the date subtraction.", "=ROUND(D2/(C2-B2), 2)"] };
}, { drill: "dates" });

// ---------------------------------------------------------------- Part 5
inLevel(5, "LEFT, RIGHT, LEN");
add("m5-airline", "Airline code", (R) => {
  const n = R.p.int(10, 13), f = flights(R.d, n);
  const s = sheet("Flights", ["Flight", "Origin", "Airline"], f.map((r) => [r[0], r[1], null]));
  return { prompt: "In <b>C2</b>, pull the 2-letter airline code off the front of each flight number.", sheets: [s], target: tgt("Flights", "C2", n), require: ["LEFT"], solution: "=LEFT(A2,2)", hints: ["The first 2 characters.", "=LEFT(A2, 2)"] };
}, { drill: "text" });
add("m5-vin", "Last 6 of the VIN", (R) => {
  const n = R.p.int(10, 13), ch = "ABCDEFGHJKLMNPRSTUVWXYZ0123456789";
  const vin = () => { let s = ""; for (let i = 0; i < 17; i++) s += ch[R.d.int(0, ch.length - 1)]; return s; };
  const s = sheet("Inventory", ["VIN", "Model", "Last 6"], Array.from({ length: n }, () => [vin(), R.d.pick(MODELS)[0], null]));
  return { prompt: "Dealerships often identify cars by the last 6 characters of the VIN. In <b>C2</b>, pull them out.", sheets: [s], target: tgt("Inventory", "C2", n), require: ["RIGHT"], solution: "=RIGHT(A2,6)", hints: ["Characters from the right.", "=RIGHT(A2, 6)"] };
}, { drill: "text" });

inLevel(5, "UPPER, LOWER, PROPER, TRIM");
add("m5-patient-name", "Clean patient names", (R) => {
  const n = R.p.int(10, 13), ppl = people(R.d, n), sp = () => R.d.pick(["", " ", "  "]);
  const s = sheet("Intake", ["Name (typed)", "Name"], ppl.map((p) => [sp() + p.first.toLowerCase() + R.d.pick([" ", "  ", "   "]) + p.last.toLowerCase() + sp(), null]));
  return { prompt: "Names typed at a hospital intake desk are lowercase with extra spaces. In <b>B2</b>, clean them up so they look like <b>Mary Smith</b>.", sheets: [s], target: tgt("Intake", "B2", n), require: ["PROPER", "TRIM"], solution: "=PROPER(TRIM(A2))", hints: ["TRIM removes the spaces; PROPER capitalizes.", "=PROPER(TRIM(A2))"] };
}, { drill: "text" });
add("m5-airport-upper", "Airport codes in caps", (R) => {
  const n = R.p.int(10, 13);
  const s = sheet("Bags", ["Bag Tag", "Dest (typed)", "Dest"], Array.from({ length: n }, () => [String(R.d.int(100000, 999999)), R.d.pick(CODES).split("").map((c) => (R.d.chance(0.7) ? c.toLowerCase() : c)).join(""), null]));
  return { prompt: "Baggage agents typed destination codes in mixed case. In <b>C2</b>, convert them to all capitals.", sheets: [s], target: tgt("Bags", "C2", n), require: ["UPPER"], solution: "=UPPER(B2)", hints: ["=UPPER(B2)"] };
}, { drill: "text" });

inLevel(5, "Joining text: CONCATENATE and &");
add("m5-route", "Route label", (R) => {
  const n = R.p.int(10, 13), f = flights(R.d, n);
  const s = sheet("Flights", ["Flight", "Origin", "Dest", "Route"], f.map((r) => [r[0], r[1], r[2], null]));
  return { prompt: "In <b>D2</b>, build a route label like <b>BHM-ATL</b> (origin, dash, destination).", sheets: [s], target: tgt("Flights", "D2", n), solution: '=B2&"-"&C2', hints: ['=B2&"-"&C2'] };
}, { drill: "text" });
add("m5-badge", "Name badge", (R) => {
  const n = R.p.int(10, 13), ppl = people(R.d, n);
  const s = sheet("Staff", ["Name", "Department", "Badge"], ppl.map((p) => [p.first + " " + p.last, R.d.pick(DEPTS), null]));
  return { prompt: "In <b>C2</b>, make a badge label like <b>Mary Smith (Cardiology)</b>.", sheets: [s], target: tgt("Staff", "C2", n), solution: '=A2&" ("&B2&")"', hints: ['The parentheses are text, so they go inside quotes: " (" and ")"', '=A2&" ("&B2&")"'] };
}, { drill: "text" });

inLevel(5, "SUBSTITUTE & VALUE");
add("m5-weight", "Weight text to number", (R) => {
  const n = R.p.int(10, 13);
  const s = sheet("Freight", ["Pallet", "Weight (text)", "Weight"], Array.from({ length: n }, () => ["PL" + R.d.int(100, 999), R.d.int(80, 4800).toLocaleString("en-US") + " lbs", null]));
  return { prompt: "Pallet weights were exported as text like <code>1,250 lbs</code>. In <b>C2</b>, remove <b>\" lbs\"</b> and turn the result into a real <b>number</b>.", sheets: [s], target: tgt("Freight", "C2", n, null, "int"), require: ["VALUE", "SUBSTITUTE"], strictType: true, solution: '=VALUE(SUBSTITUTE(B2," lbs",""))', hints: ['SUBSTITUTE(B2," lbs","") leaves "1,250".', "VALUE understands the comma."] };
}, { drill: "text" });
add("m5-pct", "Percent text to number", (R) => {
  const n = R.p.int(10, 13);
  const s = sheet("Survey", ["Store", "Satisfaction (text)", "Satisfaction"], Array.from({ length: n }, () => [R.d.pick(["Tuscaloosa", "Hoover", "Mobile", "Auburn", "Dothan", "Decatur"]) + " #" + R.d.int(1, 9), R.d.int(55, 99) + " pct", null]));
  return { prompt: "Scores came in as text like <code>87 pct</code>. In <b>C2</b>, turn each into a real percentage (0.87): remove <b>\" pct\"</b>, convert with VALUE, and divide by 100.", sheets: [s], target: tgt("Survey", "C2", n, null, "pct"), require: ["VALUE", "SUBSTITUTE"], strictType: true, solution: '=VALUE(SUBSTITUTE(B2," pct",""))/100', hints: ['VALUE(SUBSTITUTE(B2," pct","")) gives 87.', "Divide by 100."] };
}, { drill: "text" });

inLevel(5, "SEARCH");
add("m5-airport-city", "City from \"City (CODE)\"", (R) => {
  const n = R.p.int(10, 13);
  const s = sheet("Gates", ["Destination", "City"], Array.from({ length: n }, () => { const a = R.d.pick(AIRPORTS); return [a[0] + " (" + a[2] + ")", null]; }));
  return { prompt: "Gate screens show destinations like <code>New Orleans (MSY)</code>. In <b>B2</b>, pull out just the city: everything before the <b>\" (\"</b>.", sheets: [s], target: tgt("Gates", "B2", n), require: ["LEFT", "SEARCH"], solution: '=LEFT(A2,SEARCH(" (",A2)-1)', hints: ['SEARCH(" (", A2) finds the space before the parenthesis.', '=LEFT(A2, SEARCH(" (",A2)-1)'] };
}, { drill: "text" });
add("m5-domain", "Email domain", (R) => {
  const n = R.p.int(10, 13), c = customers(R, n);
  const s = sheet("Customers", ["Name", "Email", "Domain"], c.map((x) => [x.first + " " + x.last, x.email, null]));
  return { prompt: "In <b>C2</b>, pull out the email domain: everything <b>after</b> the @ (like gmail.com).", sheets: [s], target: tgt("Customers", "C2", n), require: ["RIGHT", "LEN", "SEARCH"], solution: '=RIGHT(B2,LEN(B2)-SEARCH("@",B2))', hints: ["Characters after the @: LEN(B2)-SEARCH(\"@\",B2)", '=RIGHT(B2, LEN(B2)-SEARCH("@",B2))'] };
}, { drill: "text" });

inLevel(5, "TEXT");
add("m5-checkin", "Check-in label", (R) => {
  const n = R.p.int(10, 13), h = hotel(R.d, n);
  const s = sheet("Bookings", ["Booking", "Check-in", "Label"], h.map((r) => [r[0], r[2], null]), { "Check-in": "date" });
  return { prompt: "In <b>C2</b>, build a label like <b>Check-in: Mar 5, 2026</b>.", sheets: [s], target: tgt("Bookings", "C2", n), require: ["TEXT"], solution: '="Check-in: "&TEXT(B2,"mmm d, yyyy")', hints: ['The format code is "mmm d, yyyy".', '="Check-in: "&TEXT(B2, "mmm d, yyyy")'] };
}, { drill: "text" });
add("m5-loadlabel", "Load factor label", (R) => {
  const n = R.p.int(10, 13), f = flights(R.d, n);
  const s = sheet("Flights", ["Flight", "Load Factor", "Label"], f.map((r) => [r[0], Math.round(r[5] / r[6] * 1000) / 1000, null]), { "Load Factor": "pct1" });
  return { prompt: "In <b>C2</b>, build a label like <b>DL1234: 87.3%</b> (flight, colon, space, then the load factor as a percent with 1 decimal).", sheets: [s], target: tgt("Flights", "C2", n), require: ["TEXT"], solution: '=A2&": "&TEXT(B2,"0.0%")', hints: ['TEXT(B2,"0.0%")', '=A2&": "&TEXT(B2, "0.0%")'] };
}, { drill: "text" });

// ---------------------------------------------------------------- Part 6
inLevel(6, "Two functions together");
add("m6-monthlabel", "Month-year label", (R) => {
  const n = R.p.int(10, 13), rows = patients(R.d, n);
  const s = sheet("Patients", ["Patient ID", "Admit", "Admit Month"], rows.map((r) => [r[0], r[3], null]), { Admit: "date" });
  return { prompt: "In <b>C2</b>, label each admission with its month and year, like <b>March 2026</b>.", sheets: [s], target: tgt("Patients", "C2", n), require: ["TEXT"], solution: '=TEXT(B2,"mmmm")&" "&YEAR(B2)', hints: ['TEXT(B2,"mmmm") gives the month name.', '=TEXT(B2,"mmmm")&" "&YEAR(B2)  (TEXT(B2,"mmmm yyyy") works too)'] };
}, { drill: "combo" });
add("m6-agedcar", "Aged inventory", (R) => {
  const n = R.p.int(10, 13), t = todaySerial();
  const s = sheet("Inventory", ["Stock #", "Model", "Arrived", "Action"], ids(R.d, n, "S", 1000, 9999).map((id) => [id, R.d.pick(MODELS)[0], t - R.d.int(5, 200)]), { Arrived: "date" });
  s.rows[1][2] = t - 150;
  return { prompt: "Cars on the lot for <b>more than 90 days</b> get marked down. In <b>D2</b>, show <b>\"Discount\"</b> for those, otherwise leave it blank.", sheets: [s], target: tgt("Inventory", "D2", n), require: ["IF", "TODAY"], solution: '=IF(TODAY()-C2>90,"Discount","")', hints: ["Days on lot: TODAY()-C2", '=IF(TODAY()-C2>90, "Discount", "")'] };
}, { drill: "combo" });

inLevel(6, "Three or more");
add("m6-pctdelayed", "Percent of flights delayed", (R) => {
  const n = 18, rows = flights(R.d, n), o = R.p.pick(CODES);
  for (const k of [0, 1, 2]) { rows[k][1] = o; if (rows[k][2] === o) rows[k][2] = o === "BHM" ? "ATL" : "BHM"; }
  rows[0][4] = 60;
  const s = sheet("Flights", FL_H, rows, FL_FMT, [["I1", "Origin"], ["I2", o], ["J1", "% Delayed"]]);
  return { prompt: `In <b>J2</b>, find the share of flights from <b>${o}</b> (I2) that were delayed <b>more than 15</b> minutes, rounded to <b>3 decimals</b>: delayed flights from ${o} ÷ all flights from ${o}.`, sheets: [s], target: tgt("Flights", "J2", 1, null, "pct1"), require: ["ROUND", "COUNTIFS"], solution: `=ROUND(COUNTIFS(B2:B${n + 1},I2,E2:E${n + 1},">15")/COUNTIFS(B2:B${n + 1},I2),3)`, hints: ["Two COUNTIFS divided: one with the delay condition, one without.", "Wrap the division in ROUND(..., 3)."] };
}, { drill: "combo" });
add("m6-los", "Length-of-stay band", (R) => {
  const n = R.p.int(10, 13), rows = patients(R.d, n); rows[0][4] = null;
  const bands = sheet("Bands", ["Min Days", "Stay"], [[0, "Short"], [3, "Medium"], [7, "Long"]]);
  const s = sheet("Patients", HO_H.concat(["Stay"]), rows.map((r) => r.concat([null])), HO_FMT);
  return { prompt: "In <b>H2</b>: if the patient hasn't been discharged (Discharge is empty), show <b>\"Still Admitted\"</b>. Otherwise look up their length of stay (Discharge − Admit) on the <b>Bands</b> sheet with an approximate match.", sheets: [s, bands], target: tgt("Patients", "H2", n), require: ["IF", "ISBLANK", "VLOOKUP"], solution: '=IF(ISBLANK(E2),"Still Admitted",VLOOKUP(E2-D2,Bands!$A$2:$B$4,2,TRUE))', hints: ["The lookup value is E2-D2.", '=IF(ISBLANK(E2), "Still Admitted", VLOOKUP(E2-D2, Bands!$A$2:$B$4, 2, TRUE))'] };
}, { drill: "combo" });

inLevel(6, "Boss formulas");
add("m6-flightlabel", "Departure board label", (R) => {
  const n = R.p.int(10, 13), rows = flights(R.d, n); rows[0][4] = 75; rows[1][4] = 3;
  const s = sheet("Flights", FL_H.concat(["Board"]), rows.map((r) => r.concat([null])), FL_FMT);
  return { prompt: "Build a departure-board label like <b>DL to ATL on Mar 5 (LATE)</b>: the airline code (first 2 characters of Flight), \" to \", Dest, \" on \", the Date formatted <code>\"mmm d\"</code>, then <b>\" (LATE)\"</b> only if the Delay is over 15 (otherwise nothing). Fill in <b>H2</b>.", sheets: [s], target: tgt("Flights", "H2", n), require: ["LEFT", "TEXT", "IF"], solution: '=LEFT(A2,2)&" to "&C2&" on "&TEXT(D2,"mmm d")&IF(E2>15," (LATE)","")', hints: ['Pieces: LEFT(A2,2), C2, TEXT(D2,"mmm d"), IF(E2>15," (LATE)","")', 'Glue with &" to "& and &" on "&.'] };
}, { drill: "combo" });
add("m6-gymfee", "Loyalty membership fee", (R) => {
  const prices = R.p.pick([[29, 49, 79], [25, 45, 70], [35, 55, 89]]);
  const pl = sheet("Plans", ["Plan", "Monthly Fee"], PLANS.map((p, i) => [p, prices[i]]), { "Monthly Fee": "money0" });
  const n = R.p.int(10, 13), ppl = people(R.d, n), y = new Date().getFullYear();
  const s = sheet("Members", ["Member", "Plan", "Join Date", "Fee"], ppl.map((p) => [p.first + " " + p.last, R.d.pick(PLANS), D(R.d.int(y - 10, y), R.d.int(1, 12), R.d.int(1, 28))]), { "Join Date": "date" });
  s.rows[1][2] = D(y - 8, 3, 3);
  return { prompt: "Each member pays their plan's monthly fee (see <b>Plans</b>), but members who joined <b>5 or more years ago</b> (this year minus join year ≥ 5) get <b>10% off</b>. Round the fee to <b>2 decimals</b>. Fill in <b>D2</b>.", sheets: [s, pl], target: tgt("Members", "D2", n, null, "money"), require: ["ROUND", "VLOOKUP", "IF", "YEAR", "TODAY"], solution: "=ROUND(VLOOKUP(B2,Plans!$A$2:$B$4,2,FALSE)*IF(YEAR(TODAY())-YEAR(C2)>=5,0.9,1),2)", hints: ["Fee: VLOOKUP(B2, Plans!$A$2:$B$4, 2, FALSE). Multiplier: IF(YEAR(TODAY())-YEAR(C2)>=5, 0.9, 1).", "Multiply them, then wrap in ROUND(..., 2)."] };
}, { drill: "combo" });

// ---------------------------------------------------------------- Part 7: two more workbooks
part(7, "Exam-Style Practice");
function hospScenario(R) {
  const n = 14, ppl = people(R.d, n), t = todaySerial(), codes = ["CAR", "ORT", "PED", "NEU", "EMR"];
  const rows = ppl.map((p, i) => { const adm = t - R.d.int(2, 40), stay = R.d.int(1, 12); return [p.last.toUpperCase() + ", " + p.first.toUpperCase(), "MRN" + R.d.int(10000, 99999), i < 5 ? codes[i] : R.d.pick(codes), adm, adm + stay < t && !R.d.chance(0.2) ? adm + stay : null, R.d.int(30, 900) * 100 + R.d.int(0, 99), R.d.pick(INSURERS)]; });
  rows[0][4] = null;
  const H = ["Patient (raw)", "MRN", "Dept Code", "Admit", "Discharge", "Charges", "Insurance", "Patient", "Department", "Length of Stay", "Insurance Pays", "Patient Owes", "Flag"];
  const main = sheet("Patients", H, rows.map((r) => r.concat([null, null, null, null, null, null])), { Admit: "date", Discharge: "date", Charges: "money", "Insurance Pays": "money", "Patient Owes": "money" }, [["O1", "Review If Owes Over"], ["P1", R.d.pick([15000, 20000, 25000]), "money0"]]);
  const dp = sheet("Depts", ["Code", "Department"], codes.map((c, i) => [c, DEPTS[i]]));
  const ins = sheet("Insurance", ["Insurer", "Covers"], [["Medicare", 0.8], ["BCBS", 0.85], ["Aetna", 0.75], ["Self-Pay", 0]], { Covers: "pct" });
  const sum = sheet("Summary", ["Department", "Patients", "Avg Stay"], DEPTS.map((d) => [d, null, null]));
  return { sheets: [main, dp, ins, sum], n };
}
const HOSP_STEPS = [
  { col: "H", title: "Patient", f: '=PROPER(RIGHT(A2,LEN(A2)-SEARCH(", ",A2)-1)&" "&LEFT(A2,SEARCH(",",A2)-1))', req: ["PROPER", "RIGHT", "LEFT", "SEARCH", "LEN"],
    prompt: "Names came in as <code>SMITH, MARY</code>. Rewrite each one as <b>Mary Smith</b>: the first name (after the comma and space), a space, the last name (before the comma), all in proper case.", hints: ['First name: RIGHT(A2, LEN(A2)-SEARCH(", ",A2)-1). Last name: LEFT(A2, SEARCH(",",A2)-1).', 'PROPER(first & " " & last)'] },
  { col: "I", title: "Department", f: "=VLOOKUP(C2,Depts!$A$2:$B$6,2,FALSE)", req: ["VLOOKUP"], prompt: "Look up the department name for each Dept Code on the <b>Depts</b> sheet.", hints: ["=VLOOKUP(C2, Depts!$A$2:$B$6, 2, FALSE)"] },
  { col: "J", title: "Length of Stay", f: "=IF(ISBLANK(E2),TODAY()-D2,E2-D2)", fmt: "int", req: ["IF", "ISBLANK", "TODAY"], prompt: "Days in the hospital: Discharge − Admit. For patients still admitted (no Discharge date), use today − Admit.", hints: ["Test ISBLANK(E2).", "=IF(ISBLANK(E2), TODAY()-D2, E2-D2)"] },
  { col: "K", title: "Insurance Pays", f: "=ROUND(F2*VLOOKUP(G2,Insurance!$A$2:$B$5,2,FALSE),2)", fmt: "money", req: ["ROUND", "VLOOKUP"], prompt: "Insurance pays Charges × the share its insurer covers (see the <b>Insurance</b> sheet), rounded to 2 decimals.", hints: ["Share: VLOOKUP(G2, Insurance!$A$2:$B$5, 2, FALSE)", "=ROUND(F2*VLOOKUP(...), 2)"] },
  { col: "L", title: "Patient Owes", f: "=F2-K2", fmt: "money", prompt: "What the patient owes: Charges minus what insurance pays.", hints: ["=F2-K2"] },
  { col: "M", title: "Flag", f: '=IF(OR(J2>7,L2>$P$1),"Review","")', prompt: "Show <b>\"Review\"</b> if the stay is <b>longer than 7 days</b> <b>or</b> the patient owes more than the amount in <b>P1</b>. Otherwise blank.", hints: ["OR(J2>7, L2>$P$1)", "Lock P1."] },
  { sheet: "Summary", col: "B", title: "Patients", f: "=COUNTIFS(Patients!$I$2:$I$15,A2)", fmt: "int", rows: 5, req: ["COUNTIFS"], prompt: "On <b>Summary</b>, count the patients in each department (use the Department column you built).", hints: ["=COUNTIFS(Patients!$I$2:$I$15, A2)"] },
  { sheet: "Summary", col: "C", title: "Avg Stay", f: "=ROUND(AVERAGEIFS(Patients!$J$2:$J$15,Patients!$I$2:$I$15,A2),1)", fmt: "dec1", rows: 5, req: ["ROUND", "AVERAGEIFS"], prompt: "On <b>Summary</b>, find each department's average Length of Stay, rounded to 1 decimal.", hints: ["AVERAGEIFS(Patients!$J$2:$J$15, Patients!$I$2:$I$15, A2)", "Wrap it in ROUND(..., 1)."] },
];
level("Scenario 4: Hospital Patient Log", `
<p>A hospital's patient log: clean up names, look up departments and insurance, work with admit and discharge dates, and summarize by department.</p>
<p>Testing out is off here: work through every step.</p>`, { noTestOut: true });
scenarioProblems("p7-hosp", HOSP_STEPS, hospScenario, "Patients", "Hospital Workbook");

function restScenario(R) {
  const n = 16, srv = people(R.d, 4), sids = ids(R.d, 4, "S", 10, 99), base = todaySerial() - 40;
  const rows = Array.from({ length: n }, (_, i) => { const sub = R.d.money(25, 220); return ["T" + (1000 + i), sids[i % 4], base + R.d.int(0, 35), sub, Math.round(sub * R.d.int(8, 28)) / 100, R.d.pick(["visa", "amex", "mastercard", "discover"]) + " ending " + R.d.int(1000, 9999)]; });
  for (let k = 0; k < 4; k++) { let d = rows[k][2]; while (XL.serialParts(d).dow !== 6) d++; rows[k][2] = d; }
  const H = ["Ticket", "Server ID", "Date", "Subtotal", "Tip", "Payment", "Server", "Day", "Weekend?", "Tip %", "Tip Rating", "Card"];
  const main = sheet("Tickets", H, rows.map((r) => r.concat([null, null, null, null, null, null])), { Date: "date", Subtotal: "money", Tip: "money", "Tip %": "pct1" });
  const sv = sheet("Servers", ["Name", "Server ID"], R.d.shuffle(srv.map((p, i) => [p.first + " " + p.last, sids[i]])));
  const sum = sheet("Summary", ["Server", "Weekend Tips", "Avg Tip %"], srv.map((p) => [p.first + " " + p.last, null, null]));
  return { sheets: [main, sv, sum], n };
}
const REST_STEPS = [
  { col: "G", title: "Server", f: "=INDEX(Servers!$A$2:$A$5,MATCH(B2,Servers!$B$2:$B$5,0))", req: ["INDEX", "MATCH"], forbid: ["VLOOKUP"], prompt: "Bring in the server's name from the <b>Servers</b> sheet by Server ID (the ID is to the right of the name there).", hints: ["=INDEX(Servers!$A$2:$A$5, MATCH(B2, Servers!$B$2:$B$5, 0))"] },
  { col: "H", title: "Day", f: '=TEXT(C2,"dddd")', req: ["TEXT"], prompt: "Show the day of the week for each ticket (like <b>Saturday</b>).", hints: ['=TEXT(C2, "dddd")'] },
  { col: "I", title: "Weekend?", f: '=IF(OR(H2="Saturday",H2="Sunday"),"Yes","No")', prompt: "Show <b>\"Yes\"</b> if the Day is Saturday or Sunday, otherwise <b>\"No\"</b>.", hints: ['OR(H2="Saturday", H2="Sunday")'] },
  { col: "J", title: "Tip %", f: "=ROUND(E2/D2,3)", fmt: "pct1", req: ["ROUND"], prompt: "Tip ÷ Subtotal, rounded to 3 decimals.", hints: ["=ROUND(E2/D2, 3)"] },
  { col: "K", title: "Tip Rating", f: '=IF(J2>=0.2,"Great",IF(J2>=0.15,"Good","Low"))', prompt: "Rate each tip: <b>20% or more</b> is <b>\"Great\"</b>, <b>15% or more</b> is <b>\"Good\"</b>, and anything less is <b>\"Low\"</b>.", hints: ["Nested IF, biggest tier first. 20% is 0.2."] },
  { col: "L", title: "Card", f: '=PROPER(LEFT(F2,SEARCH(" ",F2)-1))', req: ["PROPER", "LEFT", "SEARCH"], prompt: "Payments look like <code>visa ending 1234</code>. Pull out the card brand (everything before the first space) in proper case, like <b>Visa</b>.", hints: ['LEFT(F2, SEARCH(" ",F2)-1)', "Wrap it in PROPER."] },
  { sheet: "Summary", col: "B", title: "Weekend Tips", f: '=SUMIFS(Tickets!$E$2:$E$17,Tickets!$G$2:$G$17,A2,Tickets!$I$2:$I$17,"Yes")', fmt: "money", rows: 4, req: ["SUMIFS"], prompt: "On <b>Summary</b>, add up each server's Tips on weekend tickets only.", hints: ["Sum Tickets!E where G matches A2 and I is \"Yes\".", "Lock the ranges."] },
  { sheet: "Summary", col: "C", title: "Avg Tip %", f: "=ROUND(AVERAGEIFS(Tickets!$J$2:$J$17,Tickets!$G$2:$G$17,A2),3)", fmt: "pct1", rows: 4, req: ["ROUND", "AVERAGEIFS"], prompt: "On <b>Summary</b>, find each server's average Tip %, rounded to 3 decimals.", hints: ["AVERAGEIFS(Tickets!$J$2:$J$17, Tickets!$G$2:$G$17, A2)", "Wrap in ROUND(..., 3)."] },
];
level("Scenario 5: Restaurant Tickets", `
<p>A restaurant's ticket log: look up servers, work with days of the week, calculate and rate tips, clean up card text, and summarize by server.</p>
<p>Testing out is off here: work through every step.</p>`, { noTestOut: true });
scenarioProblems("p7-rest", REST_STEPS, restScenario, "Tickets", "Restaurant Workbook");

// keep each level's problems together, in level order
{
  const order = Object.keys(LEVEL_INFO);
  BANK.sort((a, b) => order.indexOf(a.part + "|" + a.level) - order.indexOf(b.part + "|" + b.level));
}

// ================================================================== drill categories
const DRILLS = [
  { key: "math", name: "Conditional Math", desc: "SUM, COUNTIFS, SUMIFS, AVERAGEIFS, SUMPRODUCT, $ references" },
  { key: "logic", name: "Logic", desc: "IF, AND, OR, nested IF, IS functions, IFERROR" },
  { key: "lookup", name: "Lookups", desc: "VLOOKUP (exact & approximate), INDEX/MATCH, INDEX/MATCH/MATCH" },
  { key: "dates", name: "Dates & Rounding", desc: "YEAR, MONTH, DAY, TODAY, ROUND" },
  { key: "text", name: "Text", desc: "LEFT, RIGHT, LEN, TRIM, PROPER, SUBSTITUTE, VALUE, SEARCH, TEXT, &" },
  { key: "combo", name: "Combined Formulas", desc: "Nested formulas that mix lookups, logic, text, dates, and math" },
  { key: "mixed", name: "Mixed Bag", desc: "A little of everything, like the exam" },
];

const Bank = { BANK, LEVEL_INFO, DRILLS, makeR, rand };
if (typeof module !== "undefined" && module.exports) module.exports = Bank;
else root.Bank = Bank;
})(typeof self !== "undefined" ? self : this);
