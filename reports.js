// Reports tab (English, like the "CashFlow Projection" sheet):
//  - Cash Flow: bank transactions (Cash Flow table) -> month x Category matrix, opening/closing balance, drill-down to names
//  - Cash Flow Outflow Uptoday: rows of category / month / type / amount -> the same matrix + the sheet's R:X annual panel
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const REPORTS = [
  { id: "cashflow", title: "Cash Flow", table: "Cash Flow", defaultType: "out" },
  { id: "outflow", title: "Cash Flow Outflow Uptoday", table: "Cash Flow Outflow Uptoday", defaultType: "out" },
];
// "Annual Sections Cost" (CashFlow Projection V:W). Category names are matched loosely (case, spaces, "(Test)").
const SECTIONS = [
  { label: "Cost of current worker combine", parts: [["wages"], ["tax"], ["butuach leumi"], ["pension", "pensia"]] },
  { label: "Food expenses total", parts: [["food yadhashmona"], ["credit card visa", "kredit card visa", 0.6]] }, // 60% of Visa is food
  { label: "Rent Facilities", parts: [["rent baptist village", "rent kfar habaptistim"]] },
  { label: "Insurance", parts: [["insurance"]] },
  { label: "Vehicle Cost", parts: [["insurance car"], ["vehilcles repair", "vehicles repair"], ["leasing"], ["gas"], ["roade 6", "road 6"], ["car license renawal fee", "car license renewal fee"]] },
  { label: "Bank fees and bookkeeping", parts: [["bank"], ["bookkeeping"]] },
];
const UNCAT = "Uncategorized", NONAME = "(no name)";
const rep = { id: "cashflow", year: null, open: new Set(), cache: {} };

const norm = (s) => String(s ?? "").toLowerCase().replace(/\(.*?\)/g, " ").replace(/[^a-z0-9֐-׿]+/g, " ").trim();
const money = (n) => (n == null || n === 0 ? "" : Math.round(n).toLocaleString("en-US"));
const exact = (n) => (n == null ? "" : n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
const toYM = (v) => {
  if (typeof v === "number" && v > 190000 && v < 300000) return `${Math.floor(v / 100)}-${String(v % 100).padStart(2, "0")}`;
  const m = String(v ?? "").match(/^(\d{4})-?(\d{2})/);
  return m ? `${m[1]}-${m[2]}` : null;
};
const typeOf = (v) => (/^\s*in|הכנס|inflow/i.test(String(v ?? "")) ? "in" : /^\s*out|הוצא|outflow/i.test(String(v ?? "")) ? "out" : null);

// Categories table -> order + Inflows/Outflows type per category
async function loadCategoryInfo() {
  const ds = state.datasets.find((d) => d.name === CATEGORY_COL.lookup);
  const info = new Map();
  if (!ds) return info;
  const nameCol = ds.columns.find((c) => /category|קטגורי|סיווג/i.test(c.label)) || ds.columns[0];
  const typeCol = ds.columns.find((c) => /type|סוג/i.test(c.label));
  (await loadRows(ds.id)).forEach((r, i) => {
    const name = r.data[nameCol?.key];
    if (name != null) info.set(norm(name), { name: String(name).trim(), type: typeCol ? typeOf(r.data[typeCol.key]) : null, order: i });
  });
  return info;
}

// bank balance before/after each transaction, in time order (Leumi lists the newest first)
function balanceTimeline(rows, dateCol, balCol, debit, credit) {
  const pts = rows.filter((r) => typeof r.data[balCol.key] === "number" && r.data[dateCol.key]);
  if (!pts.length) return null;
  // keep the bank's posting order (the balance runs in that order, even where a date is out of order);
  // Leumi lists the newest first, so reverse it
  pts.sort((a, b) => a.id - b.id);
  if (String(pts[0].data[dateCol.key]) > String(pts[pts.length - 1].data[dateCol.key])) pts.reverse();
  return pts.map((r) => {
    const close = r.data[balCol.key];
    const open = close - (Number(credit && r.data[credit.key]) || 0) + (Number(debit && r.data[debit.key]) || 0);
    return { day: String(r.data[dateCol.key]).slice(0, 10), open, close };
  });
}
// the bank's own balance at the start / end of a month (null when the bank data doesn't cover that month)
const tlRange = (tl) => tl.range || (tl.range = tl.reduce((r, p) => [p.day < r[0] ? p.day : r[0], p.day > r[1] ? p.day : r[1]], [tl[0].day, tl[0].day]));
const covers = (tl, ym) => !!tl && tlRange(tl)[0].slice(0, 7) <= ym && tlRange(tl)[1].slice(0, 7) >= ym;
const prevYM = (ym) => { const [y, m] = ym.split("-").map(Number); return m === 1 ? `${y - 1}-12` : `${y}-${String(m - 1).padStart(2, "0")}`; };
// balance after the last posting dated in (or before) this month
function closingAt(tl, ym) {
  if (!covers(tl, ym)) return null;
  for (let j = tl.length - 1; j >= 0; j--) if (tl[j].day <= `${ym}-99`) return tl[j].close;
  return null;
}
// = last month's closing, so the two rows always agree; the very first month starts from its first posting
function openingAt(tl, ym) {
  if (!covers(tl, ym)) return null;
  return closingAt(tl, prevYM(ym)) ?? tl.find((p) => p.day >= `${ym}-01`)?.open ?? null;
}

// table -> facts [{ym, cat, type, amt, name}] + optional balance timeline
async function loadFacts(r) {
  const ds = state.datasets.find((d) => d.name === r.table);
  if (!ds) return { error: `הטבלה "${r.table}" עדיין לא נטענה. טענו אותה מלשונית הטבלאות.` };
  const [rows, cats] = await Promise.all([loadRows(ds.id), loadCategoryInfo()]);
  const cols = ds.columns;
  const catCol = cols.find((c) => c.label === CATEGORY_COL.label && c.manual) || cols.find((c) => CAT_RE.test(c.label));
  const monthCol = cols.find((c) => /^\s*month\s*$|חודש/i.test(c.label));
  const amountCol = cols.find((c) => /amount|סכום/i.test(c.label));
  const typeCol = cols.find((c) => /^\s*type\s*$|סוג/i.test(c.label));
  const dateCol = cols.find((c) => c.type === "date");
  const debit = cols.find((c) => /חובה|debit/i.test(c.label));
  const credit = cols.find((c) => /זכות|credit/i.test(c.label));
  const balCol = cols.find((c) => /יתרה|balance/i.test(c.label) && c.type === "number");
  const nameCol = cols.find((c) => c.label === NAME_LABEL);
  const catName = (v) => (v == null || v === "" ? UNCAT : cats.get(norm(v))?.name || String(v).trim());
  const facts = [];
  const add = (ym, cat, type, amt, name) => { if (ym && amt) facts.push({ ym, cat, type, amt, name }); };
  for (const row of rows) {
    const d = row.data;
    const cat = catName(catCol && d[catCol.key]);
    const name = nameCol ? d[nameCol.key] ?? NONAME : null;
    const ym = toYM(monthCol ? d[monthCol.key] : dateCol && d[dateCol.key]);
    if (debit || credit) { // bank style: separate debit / credit columns
      add(ym, cat, "in", Number(credit && d[credit.key]) || 0, name);
      add(ym, cat, "out", Number(debit && d[debit.key]) || 0, name);
    } else if (amountCol) { // one amount column: the type comes from the row, else the Categories table, else the report default
      let amt = Number(d[amountCol.key]) || 0;
      let type = (typeCol && typeOf(d[typeCol.key])) || cats.get(norm(d[catCol?.key]))?.type || r.defaultType;
      if (amt < 0) { amt = -amt; type = type === "in" ? "out" : "in"; }
      add(ym, cat, type, amt, name);
    }
  }
  const tl = dateCol && balCol ? balanceTimeline(rows, dateCol, balCol, debit, credit) : null;
  return { facts, cats, tl, hasNames: !!nameCol, rows: rows.length };
}

// facts for one year -> matrix
function buildMatrix(f, year, opening0, tl) {
  const yms = MONTHS.map((_, i) => `${year}-${String(i + 1).padStart(2, "0")}`);
  const sec = { in: new Map(), out: new Map() };
  const blank = () => Array(12).fill(0);
  for (const x of f.facts) {
    const i = yms.indexOf(x.ym); if (i < 0) continue;
    const g = sec[x.type];
    if (!g.has(x.cat)) g.set(x.cat, { cat: x.cat, m: blank(), names: new Map() });
    const c = g.get(x.cat); c.m[i] += x.amt;
    if (x.name != null) { if (!c.names.has(x.name)) c.names.set(x.name, blank()); c.names.get(x.name)[i] += x.amt; }
  }
  const order = (a, b) => (a.cat === UNCAT) - (b.cat === UNCAT) || (f.cats.get(norm(a.cat))?.order ?? 1e9) - (f.cats.get(norm(b.cat))?.order ?? 1e9) || a.cat.localeCompare(b.cat);
  const list = (t) => [...sec[t].values()].sort(order);
  const sumCols = (arr) => blank().map((_, i) => arr.reduce((s, c) => s + c.m[i], 0));
  const ins = list("in"), outs = list("out");
  const tin = sumCols(ins), tout = sumCols(outs), net = tin.map((v, i) => v - tout[i]);
  const has = yms.map((ym) => f.facts.some((x) => x.ym === ym));
  const last = has.lastIndexOf(true);
  // opening / ending balance: the bank's real balance wherever the bank data covers the month,
  // otherwise carried forward from the previous month + this month's net (projection months)
  const open = Array(12).fill(null), close = Array(12).fill(null), mismatch = [];
  const lastBal = Math.max(last, ...yms.map((ym, i) => (covers(tl, ym) ? i : -1)));
  for (let i = 0; i <= lastBal; i++) {
    const bo = openingAt(tl, yms[i]), bc = closingAt(tl, yms[i]);
    open[i] = i === 0 && opening0 != null ? opening0 : bo ?? (i > 0 ? close[i - 1] : opening0);
    if (open[i] == null) continue;
    close[i] = bc ?? open[i] + net[i];
    if (f.tl && bo != null && bc != null && Math.abs(bo + net[i] - bc) > 0.5) mismatch.push({ i, diff: bc - (bo + net[i]) });
  }
  // a posting dated in one month but booked in the next shows as +x / -x in two months: that's timing, not missing rows
  if (Math.abs(mismatch.reduce((s, x) => s + x.diff, 0)) <= 0.5) mismatch.length = 0;
  const active = has.filter(Boolean).length || 1;
  const yearOf = (m) => m.reduce((s, v) => s + v, 0);
  // R:X annual panel
  const byNorm = new Map(outs.map((c) => [norm(c.cat), yearOf(c.m)]));
  const sections = SECTIONS.map((s) => ({
    label: s.label,
    value: s.parts.reduce((sum, p) => {
      const share = typeof p[p.length - 1] === "number" ? p[p.length - 1] : 1;
      const names = p.filter((x) => typeof x === "string");
      const hit = [...byNorm.keys()].find((k) => names.some((n) => k === n || k.startsWith(n + " ")));
      return sum + (hit ? byNorm.get(hit) * share : 0);
    }, 0),
  }));
  return { yms, ins, outs, tin, tout, net, open, close, has, last, yearOf, mismatch,
    annual: { avgIn: yearOf(tin) / active, avgOut: yearOf(tout) / active, months: has.filter(Boolean).length, sections } };
}

function openingFor(r, f, year) {
  try { const v = localStorage.getItem(`mb_open_${r.id}_${year}`); if (v !== null && v !== "") return { value: Number(v), manual: true }; } catch {}
  const tl = f.tl || rep.cache.cashflow?.tl; // the outflow sheet has no balance: take it from the bank (Cash Flow table)
  return { value: openingAt(tl, `${year}-01`), manual: false };
}

async function renderReports() {
  const el = $("view-reports");
  const r = REPORTS.find((x) => x.id === rep.id);
  const head = `<div class="toolbar"><h2>Reports</h2>
    ${REPORTS.map((x) => `<button class="btn ${x.id === rep.id ? "primary" : ""}" data-report="${x.id}">${esc(x.title)}</button>`).join("")}</div>`;
  if (!rep.cache[r.id]) {
    el.innerHTML = head + `<p class="muted">טוען…</p>`;
    try {
      if (r.id !== "cashflow" && !rep.cache.cashflow && state.datasets.some((d) => d.name === "Cash Flow")) rep.cache.cashflow = await loadFacts(REPORTS[0]);
      rep.cache[r.id] = await loadFacts(r);
    } catch (err) { console.error(err); el.innerHTML = head + `<p class="warn">שגיאה בטעינת הדוח</p>`; return; }
  }
  const f = rep.cache[r.id];
  if (f.error) { el.innerHTML = head + `<div class="card empty">${esc(f.error)}</div>`; return; }
  const years = [...new Set(f.facts.map((x) => x.ym.slice(0, 4)))].sort();
  if (!years.length) { el.innerHTML = head + `<div class="card empty">אין בטבלה נתונים עם תאריך/חודש וסכום.</div>`; return; }
  if (!years.includes(rep.year)) rep.year = years[years.length - 1];
  const op = openingFor(r, f, rep.year);
  const M = buildMatrix(f, rep.year, op.manual ? op.value : null, f.tl || rep.cache.cashflow?.tl);
  rep.last = { r, M, f };
  const uncats = [...M.ins, ...M.outs].filter((c) => c.cat === UNCAT).reduce((s, c) => s + M.yearOf(c.m), 0);

  const cells = (arr, cls = "") => arr.map((v, i) => `<td class="num ${cls}${M.has[i] ? "" : " future"}" title="${exact(v)}">${money(v)}</td>`).join("");
  const line = (label, arr, cls, total = true) => `<tr class="${cls}"><td class="lbl">${label}</td>${cells(arr)}<td class="num ytot">${total ? money(M.yearOf(arr.map((v) => v || 0))) : ""}</td></tr>`;
  const catRows = (list, type) => list.map((c) => {
    const key = `${type}:${c.cat}`, isOpen = rep.open.has(key), canOpen = f.hasNames && c.names.size;
    const names = isOpen ? [...c.names.entries()].sort((a, b) => M.yearOf(b[1]) - M.yearOf(a[1]))
      .map(([n, m]) => `<tr class="namerow"><td class="lbl">${esc(n)}</td>${cells(m)}<td class="num ytot">${money(M.yearOf(m))}</td></tr>`).join("") : "";
    return `<tr class="catrow${canOpen ? " drill" : ""}${c.cat === UNCAT ? " uncat" : ""}" ${canOpen ? `data-drill="${esc(key)}"` : ""}>
      <td class="lbl">${canOpen ? (isOpen ? "▾ " : "▸ ") : ""}${esc(c.cat)}</td>${cells(c.m)}<td class="num ytot">${money(M.yearOf(c.m))}</td></tr>${names}`;
  }).join("");
  const balRow = (label, arr) => `<tr class="bal"><td class="lbl">${label}</td>${arr.map((v, i) => `<td class="num" title="${exact(v)}">${v == null ? "" : Math.round(v).toLocaleString("en-US")}</td>`).join("")}<td></td></tr>`;

  el.innerHTML = head + `<div class="toolbar">
      <label>Year <select id="rep-year">${years.map((y) => `<option${y === rep.year ? " selected" : ""}>${y}</option>`).join("")}</select></label>
      <label title="יתרת הפתיחה של ינואר. ריק = לפי חשבון הבנק (טבלת Cash Flow)">Opening balance
        <input type="text" id="rep-open" dir="ltr" size="12" value="${op.value == null ? "" : Math.round(op.value * 100) / 100}" placeholder="${op.value == null ? "enter" : ""}">
        <span class="muted">${op.manual ? "(manual)" : op.value == null ? "" : "(from bank)"}</span></label>
      <span class="grow"></span>
      ${f.hasNames ? `<button class="btn" id="rep-expand">${rep.open.size ? "Collapse all" : "Expand all"}</button>` : ""}
      <button class="btn" id="rep-xlsx">⬇ Excel</button>
    </div>
    ${M.mismatch.length ? `<p class="warn">⚠ בחודשים ${M.mismatch.map((x) => `${MONTHS[x.i]} (הפרש ${money(x.diff)})`).join(", ")} התנועות בטבלה לא מסתכמות ליתרה של הבנק. ייתכן שיש שורות כפולות או חסרות (למשל קובץ שנטען פעמיים עם "הוספה בלבד"). היתרות בדוח לקוחות מהבנק עצמו.</p>` : ""}
    ${uncats ? `<p class="warn">⚠ ${money(uncats)} ₪ בשנה הזאת בלי Category (שורת Uncategorized). אפשר לסווג אותן בטבלה ${esc(r.table)}.</p>` : ""}
    <div class="report-grid" dir="ltr">
      <div class="table-wrap"><table class="data report"><thead><tr><th class="lbl">Category</th>${MONTHS.map((m) => `<th>${m}</th>`).join("")}<th>Total</th></tr></thead><tbody>
        ${balRow("Opening bank balance", M.open)}
        <tr class="sec"><td class="lbl">Inflows</td><td colspan="13"></td></tr>
        ${catRows(M.ins, "in")}
        ${line("Total inflows", M.tin, "total")}
        <tr class="sec"><td class="lbl">Outflows</td><td colspan="13"></td></tr>
        ${catRows(M.outs, "out")}
        ${line("Total outflows", M.tout, "total")}
        ${line("Net Cashflows", M.net, "total net")}
        ${balRow("Ending bank account balance", M.close)}
      </tbody></table></div>
      <aside class="annual card" dir="ltr">
        <h3>Annual Numbers ${rep.year}</h3>
        <div class="muted">${M.annual.months} months with data</div>
        <h4>Inflow</h4>
        <div class="kv"><span>Average Inflow Monthly</span><b>${money(M.annual.avgIn)}</b></div>
        ${M.ins.map((c) => `<div class="kv"><span>${esc(c.cat)}</span><b>${money(M.yearOf(c.m))}</b></div>`).join("")}
        <h4>Outflow</h4>
        <div class="kv"><span>Average Outflow Monthly</span><b>${money(M.annual.avgOut)}</b></div>
        ${M.outs.map((c) => `<div class="kv"><span>${esc(c.cat)}</span><b>${money(M.yearOf(c.m))}</b></div>`).join("")}
        <div class="kv total"><span>Total Outflow</span><b>${money(M.yearOf(M.tout))}</b></div>
        <h4>Annual Sections Cost</h4>
        ${M.annual.sections.map((s) => `<div class="kv"><span>${esc(s.label)}</span><b>${money(s.value) || "0"}</b></div>`).join("")}
      </aside>
    </div>`;
}

function reportToExcel() {
  const { r, M, f } = rep.last;
  const aoa = [[r.title, rep.year], [], ["Category", ...MONTHS, "Total"]];
  const nums = (arr) => arr.map((v) => (v == null || v === 0 ? null : Math.round(v * 100) / 100));
  const push = (label, arr, total = true) => aoa.push([label, ...nums(arr), total ? Math.round(M.yearOf(arr.map((v) => v || 0)) * 100) / 100 : null]);
  const cats = (list) => list.forEach((c) => {
    push(c.cat, c.m);
    if (f.hasNames) [...c.names.entries()].forEach(([n, m]) => push(`    ${n}`, m));
  });
  push("Opening bank balance", M.open, false);
  aoa.push(["Inflows"]); cats(M.ins); push("Total inflows", M.tin);
  aoa.push(["Outflows"]); cats(M.outs); push("Total outflows", M.tout);
  push("Net Cashflows", M.net); push("Ending bank account balance", M.close, false);
  aoa.push([], ["Annual Numbers"], ["Average Inflow Monthly", Math.round(M.annual.avgIn * 100) / 100], ["Average Outflow Monthly", Math.round(M.annual.avgOut * 100) / 100]);
  aoa.push(["Total Outflow", Math.round(M.yearOf(M.tout) * 100) / 100], [], ["Annual Sections Cost"]);
  M.annual.sections.forEach((s) => aoa.push([s.label, Math.round(s.value * 100) / 100]));
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 34 }, ...MONTHS.map(() => ({ wch: 11 })), { wch: 12 }];
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, r.title.slice(0, 31));
  XLSX.writeFile(wb, `${r.title} ${rep.year}.xlsx`);
}

$("view-reports").addEventListener("click", (e) => {
  const b = e.target.closest("[data-report]");
  if (b) { rep.id = b.dataset.report; rep.open.clear(); return renderReports(); }
  const d = e.target.closest("[data-drill]");
  if (d) { const k = d.dataset.drill; rep.open.has(k) ? rep.open.delete(k) : rep.open.add(k); return renderReports(); }
  if (e.target.closest("#rep-expand")) {
    const { M } = rep.last;
    if (rep.open.size) rep.open.clear();
    else { M.ins.forEach((c) => rep.open.add(`in:${c.cat}`)); M.outs.forEach((c) => rep.open.add(`out:${c.cat}`)); }
    return renderReports();
  }
  if (e.target.closest("#rep-xlsx")) return reportToExcel();
});
$("view-reports").addEventListener("change", (e) => {
  if (e.target.id === "rep-year") { rep.year = e.target.value; return renderReports(); }
  if (e.target.id === "rep-open") {
    const v = e.target.value.replace(/,/g, "").trim();
    try { if (v === "" || isNaN(Number(v))) localStorage.removeItem(`mb_open_${rep.id}_${rep.year}`); else localStorage.setItem(`mb_open_${rep.id}_${rep.year}`, v); } catch {}
    return renderReports();
  }
});
