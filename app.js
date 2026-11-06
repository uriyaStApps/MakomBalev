// MakomBalev: tables loaded from Excel into Supabase (schema makom_balev), edited on the site, downloadable as Excel.
// Google sign-in + allow-list (allowed_users); admins manage users.
const SUPABASE_URL = "https://pualpwrkztjzhgpaqudm.supabase.co";
const SUPABASE_KEY = "sb_publishable_TFMMqLxLWHatd3lXYrUlTQ_aO7MvaQ9";
const sb = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY, {
  db: { schema: "makom_balev" },
  auth: { detectSessionInUrl: false }, // handled in consumeAuthRedirect()
});

// one upload button per subject; each subject loads into its own table
const TOPICS = [
  { id: "cashflow", label: "טעינת cash flow", table: "Cash Flow", derive: deriveCashFlow },
  { id: "outflow", label: "טעינת cash flow outflow uptoday", table: "Cash Flow Outflow Uptoday", derive: deriveCashFlow },
];
// The Categories table is not uploaded: the site creates it once with this list, then it's edited on the site
// (add a row, ⬇ Excel to download). "type" is a dropdown.
const CATEGORIES_TABLE = "Categories";
const CATEGORY_TYPES = ["Inflows", "Outflows"];
const DEFAULT_CATEGORIES = [
  ["Partisipans Paymant", "Inflows"], ["Bituach Leumi", "Inflows"], ["Donations", "Inflows"], ["Other", "Inflows"],
  ["Wages", "Outflows"], ["Tax", "Outflows"], ["Butuach Leumi", "Outflows"], ["Pensia", "Outflows"], ["Food YadHashmona", "Outflows"],
  ["Food BV", "Outflows"], ["Rent Kfar Habaptistim", "Outflows"], ["Expense that Workers Pay", "Outflows"], ["Utility", "Outflows"],
  ["Office", "Outflows"], ["Administration", "Outflows"], ["Insurance", "Outflows"], ["Vehilcles Repair", "Outflows"], ["Gas", "Outflows"],
  ["Bank", "Outflows"], ["bookkeeping", "Outflows"], ["Development and Marketing Webside", "Outflows"], ["Lawer", "Outflows"],
  ["Media", "Outflows"], ["Project Yuval", "Outflows"], ["Kredit Card Visa", "Outflows"], ["Kredit Card Mastercard", "Outflows"],
];
async function ensureCategories() {
  const existing = state.datasets.find((d) => d.name === CATEGORIES_TABLE);
  if (existing) { // loaded earlier from a file: just make "type" a dropdown
    const typeCol = existing.columns.find((c) => /^\s*type\s*$/i.test(c.label));
    if (typeCol && !typeCol.options) {
      const columns = existing.columns.map((c) => (c === typeCol ? { ...c, options: CATEGORY_TYPES } : c));
      const { error } = await sb.from("datasets").update({ columns }).eq("id", existing.id);
      if (!error) existing.columns = columns;
    }
    return;
  }
  const columns = [{ key: "c1", label: "category", type: "text" }, { key: "c2", label: "type", type: "text", options: CATEGORY_TYPES }];
  const { data, error } = await sb.from("datasets").insert({ name: CATEGORIES_TABLE, columns, source_file: null }).select().single();
  if (error) { if (error.code !== "23505") console.error(error); return; } // 23505: another user just created it
  const { error: e2 } = await sb.from("dataset_rows").insert(DEFAULT_CATEGORIES.map(([c, t]) => ({ dataset_id: data.id, data: { c1: c, c2: t }, origin: "site" })));
  if (e2) console.error(e2);
  await loadDatasets();
}
// Cash Flow's "Category" is chosen by hand from the Categories table (manual = kept when the period is reloaded)
const CATEGORY_COL = { label: "Category", type: "text", lookup: "Categories", manual: true };

// ---------------- Cash Flow: who the money came from / went to ----------------
// "העברה מאת: דן שפר, מרינה שפ 10-843-003918627 תרומה"            -> "דן שפר"
// "העברה מאת: BANK OF AMERICA CORPOR 10-800-022004542"            -> "BANK OF AMERICA CORPOR"
// "מתאריך 28/10/26 12:46 בכרטיס המסתיים ב-4045 ב-בישולים ותבשילים בע"מ" -> "בישולים ותבשילים בע"מ"
function extractName(s) {
  if (typeof s !== "string") return null;
  let m = s.match(/(?:^|\s)(?:מאת|אל)\s*:\s*(.+)/); // whole word only ("ישראל:" is not "אל:")
  if (m) {
    let x = m[1].split(/\s+\d{1,3}-\d{2,3}-\d{3,}/)[0]; // stop at the bank account number
    x = x.split(/[,،]/)[0];                            // several account holders: keep the first
    x = x.replace(/[\s:.\-]+$/, "").trim();
    return x || null;
  }
  m = s.match(/המסתיים\s+ב-?\s*\d+\s+ב-?\s*(.+)/);   // card purchase: the business name
  if (m) return m[1].replace(/[\s:.]+$/, "").trim() || null;
  return null;
}
const NAME_LABEL = "שם";
function nameFromRow(row, columns) {
  // look in "תיאור מורחב" first (where Leumi puts these details), then in any other text column
  const first = (c) => /תיאור מורחב/.test(c.label);
  const order = [...columns.filter(first), ...columns.filter((c) => !first(c))];
  for (const c of order) {
    if (c.label === NAME_LABEL) continue;
    const n = extractName(row[c.key]);
    if (n) return n;
  }
  return null;
}
// adds/fills the "שם" column of a parsed table ({columns, rows}); fills only empty names
function deriveCashFlow(t) {
  if (!t.columns.length) return 0;
  // a file that already has its own category column (e.g. category / month / type / amount) keeps it as is
  const own = t.columns.find((c) => /category|קטגורי|סיווג/i.test(c.label) && !c.manual && c.label !== CATEGORY_COL.label);
  if (!own) {
    const cat = t.columns.find((c) => c.label === CATEGORY_COL.label);
    if (!cat) t.columns.push({ key: nextKey(t.columns), ...CATEGORY_COL });
    else Object.assign(cat, CATEGORY_COL);
  }
  let col = t.columns.find((c) => c.label === NAME_LABEL);
  const key = col ? col.key : nextKey(t.columns);
  let n = 0;
  for (const r of t.rows) {
    if (r[key] != null) continue;
    const name = nameFromRow(r, t.columns);
    if (name) { r[key] = name; n++; }
  }
  if (!col && n) t.columns.push({ key, label: NAME_LABEL, type: "text" }); // only when there is something to show
  return n;
}
const topicOf = (ds) => ds && TOPICS.find((tp) => tp.table === ds.name);

// ---------------- which file belongs where (blocks loading a file into the wrong subject) ----------------
const hasCol = (t, re, type) => t.columns.some((c) => re.test(c.label) && (!type || c.type === type));
const isBank = (t) => hasCol(t, /תאריך|date/i) && hasCol(t, /חובה|debit/i) && hasCol(t, /זכות|credit/i) && hasCol(t, /יתרה|balance/i);
const FILE_RULES = {
  cashflow: { test: isBank, need: "תאריך, בחובה, בזכות, יתרה בש\"ח (קובץ תנועות עו\"ש מהבנק)" },
  outflow: {
    test: (t) => !isBank(t) && (hasCol(t, /month|חודש/i) || t.columns.some((c) => c.type === "date")) && hasCol(t, /amount|סכום|חובה|זכות|outflow|הוצא/i),
    need: "חודש או תאריך, וסכום (למשל category / month / type / amountNIS)",
  },
};
const fitsOf = (t) => TOPICS.filter((tp) => FILE_RULES[tp.id]?.test(t)).map((tp) => tp.id); // run on the file as read, before derive
// null = OK, otherwise why this file can't go into that subject's table
function wrongPlace(t, topic) {
  if (!topic || !FILE_RULES[topic.id] || t.fits.includes(topic.id)) return null;
  const other = TOPICS.find((tp) => t.fits.includes(tp.id));
  if (other) return `הקובץ הזה נראה כמו קובץ של <b>${esc(other.table)}</b>, לא של ${esc(topic.table)}. השתמשו בכפתור "${esc(other.label)}".`;
  return `הקובץ לא מתאים ל-<b>${esc(topic.table)}</b>. צריך עמודות: ${esc(FILE_RULES[topic.id].need)}.<br>בקובץ יש: ${t.columns.map((c) => esc(c.label)).join(" · ") || "(אין עמודות)"}`;
}
const SHOW = 200;     // rows rendered at a time
const CHUNK = 500;    // rows per insert request
const state = { tab: "tables", datasets: [], counts: {}, ds: null, rows: [], q: "", sort: null, shown: SHOW, users: [], editing: null };
let user = null, isAdmin = false;

const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const toast = (msg, ms = 3000) => { const t = $("toast"); t.textContent = msg; t.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(() => (t.hidden = true), ms); };
const fmtDate = (iso) => iso ? new Date(iso).toLocaleString("he-IL", { dateStyle: "short", timeStyle: "short" }) : "";
const nextKey = (cols) => "c" + (Math.max(0, ...cols.map((c) => +c.key.slice(1) || 0)) + 1);

// ---------------- data access ----------------
async function fetchAll(query) { // PostgREST returns at most 1000 rows per request
  const out = [];
  for (let from = 0; ; from += 1000) {
    const { data, error } = await query().range(from, from + 999);
    if (error) throw error;
    out.push(...data);
    if (data.length < 1000) return out;
  }
}
async function loadDatasets() {
  const { data, error } = await sb.from("datasets").select("*").order("name");
  if (error) throw error;
  state.datasets = data;
  const counts = await Promise.all(data.map((d) => sb.from("dataset_rows").select("id", { count: "exact", head: true }).eq("dataset_id", d.id)));
  state.counts = Object.fromEntries(data.map((d, i) => [d.id, counts[i].count ?? 0]));
}
const loadRows = (dsId) => fetchAll(() => sb.from("dataset_rows").select("*").eq("dataset_id", dsId).order("id"));

async function insertRows(dsId, dataList, origin) {
  for (let i = 0; i < dataList.length; i += CHUNK) {
    const { error } = await sb.from("dataset_rows").insert(dataList.slice(i, i + CHUNK).map((data) => ({ dataset_id: dsId, data, origin })));
    if (error) throw error;
    $("upload-progress").textContent = `נטענו ${Math.min(i + CHUNK, dataList.length)} מתוך ${dataList.length} שורות…`;
  }
}

// ---------------- values ----------------
const pad = (n) => String(n).padStart(2, "0");
function cellValue(v) { // Excel cell -> JSON value
  if (v === null || v === undefined || v === "") return null;
  if (v instanceof Date) {
    v = new Date(Math.round(v.getTime() / 60000) * 60000); // Excel dates can come back a few seconds off
    const d = `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
    return v.getHours() || v.getMinutes() ? `${d} ${pad(v.getHours())}:${pad(v.getMinutes())}` : d;
  }
  if (typeof v === "string") {
    const s = v.trim();
    if (s === "") return null;
    // text from CSV files (bank exports): "1,234.56" -> number, "01/11/2026" -> date (day first, Israeli style)
    if (/^-?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$/.test(s) && !/^-?0\d/.test(s)) return Number(s.replace(/,/g, ""));
    const d = s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})$/);
    if (d && +d[2] >= 1 && +d[2] <= 12 && +d[1] >= 1 && +d[1] <= 31) return `${d[3]}-${pad(d[2])}-${pad(d[1])}`;
    return s;
  }
  return v;
}

// CSV files from Israeli banks are usually windows-1255, not UTF-8: decode them ourselves
async function readWorkbook(file) {
  const buf = await file.arrayBuffer();
  if (!/\.(csv|txt)$/i.test(file.name)) return XLSX.read(buf, { cellDates: true });
  let text;
  try { text = new TextDecoder("utf-8", { fatal: true }).decode(buf); }
  catch { text = new TextDecoder("windows-1255").decode(buf); }
  return XLSX.read(text.replace(/^﻿/, ""), { type: "string", raw: true }); // raw: we parse numbers/dates (day first) ourselves
}
function guessType(values) {
  const vals = values.filter((v) => v !== null);
  if (!vals.length) return "text";
  if (vals.every((v) => typeof v === "number")) return "number";
  if (vals.every((v) => typeof v === "string" && /^\d{4}-\d\d-\d\d( \d\d:\d\d)?$/.test(v))) return "date";
  return "text";
}
const showValue = (v, type) => v === null || v === undefined ? "" : type === "date" && /^\d{4}-\d\d-\d\d/.test(v) ? v.slice(0, 10).split("-").reverse().join("/") + v.slice(10) : String(v);
function parseInput(text, type) {
  const t = text.trim();
  if (!t) return null;
  if (type === "number") { const n = Number(t.replace(/,/g, "")); return isNaN(n) ? t : n; }
  if (type === "date") {
    const m = t.match(/^(\d{1,2})[./](\d{1,2})[./](\d{2,4})(?:\s+(\d{1,2}):(\d\d))?$/);
    if (m) return `${m[3].length === 2 ? "20" + m[3] : m[3]}-${pad(m[2])}-${pad(m[1])}` + (m[4] ? ` ${pad(m[4])}:${m[5]}` : "");
  }
  return t;
}

// ---------------- Excel in ----------------
// -> { columns, rows, skipped }. Bank exports start with title lines (report name, balance, dates...):
// the header is the first row that is about as wide as the table itself; everything above it is dropped.
function sheetToTable(ws) {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
  const filled = aoa.map((r) => r.filter((v) => cellValue(v) !== null).length);
  const widest = Math.max(0, ...filled);
  if (!widest) return { columns: [], rows: [], skipped: [] };
  const hi = filled.findIndex((n) => n >= Math.max(2, Math.ceil(widest * 0.6)));
  const skipped = aoa.slice(0, hi).map((r) => r.map(cellValue).filter((v) => v !== null).join(" ")).filter(Boolean);
  let header = aoa[hi].map((v) => (v === null || v === undefined ? null : String(v).trim() || null)); // headers stay text
  const body = aoa.slice(hi + 1).map((r) => r.map(cellValue)).filter((r) => r.some((v) => v !== null));
  // Leumi's CSV: the header row starts one cell later than the data rows -> every title lands one column off.
  // If the header has more empty cells in front than any data row, line it up with the data.
  const lead = (r) => { const i = r.findIndex((v) => v !== null && v !== undefined); return i < 0 ? Infinity : i; };
  const shift = lead(header) - Math.min(...body.map(lead));
  if (body.length && shift > 0 && shift < Infinity) header = header.slice(shift);
  const width = Math.max(header.length, ...body.map((r) => r.length));
  const columns = [], seen = {};
  for (let i = 0; i < width; i++) {
    if (header[i] == null && body.every((r) => (r[i] ?? null) === null)) continue; // empty column with no title
    let label = header[i] ?? `עמודה ${i + 1}`;
    if (seen[label]) label = `${label} (${++seen[label]})`; else seen[label] = 1;
    columns.push({ key: `c${i + 1}`, label, type: guessType(body.map((r) => r[i] ?? null)), i });
  }
  const rows = body.map((r) => Object.fromEntries(columns.map((c) => [c.key, r[c.i] ?? null]).filter(([, v]) => v !== null)));
  columns.forEach((c) => delete c.i);
  return { columns, rows, skipped };
}
// Hebrew that was already lost before the file reached us (saved as "????") can't be recovered
const lostHebrew = (t) => [...t.columns.map((c) => c.label), ...t.rows.slice(0, 50).flatMap((r) => Object.values(r))]
  .some((v) => typeof v === "string" && /\?{2,}/.test(v));

let upload = null; // { file, topic, tables: {sheetName: {columns, rows, skipped}} }
let uploadTopic = null; // the subject button that opened the dialog (null = general upload)
const dmy = (iso) => iso ? iso.slice(0, 10).split("-").reverse().join("/") : "";
const day = (v) => (typeof v === "string" && /^\d{4}-\d\d-\d\d/.test(v) ? v.slice(0, 10) : null);
function dateRange(values) {
  const ds = values.map(day).filter(Boolean).sort();
  return ds.length ? { from: ds[0], to: ds[ds.length - 1] } : null;
}
// file columns -> table columns, matched by header text; new headers become new columns
function mapToTable(ds, t) {
  const cols = ds.columns.map((c) => ({ ...c }));
  const map = {};
  for (const c of t.columns) {
    let hit = cols.find((x) => x.label === c.label);
    if (!hit) { hit = { ...c, key: nextKey(cols) }; cols.push(hit); }
    else if (c.lookup) Object.assign(hit, { lookup: c.lookup, manual: c.manual });
    map[c.key] = hit.key;
  }
  const rows = t.rows.map((r) => Object.fromEntries(Object.entries(r).map(([k, v]) => [map[k], v])));
  const fileDate = t.columns.find((c) => c.type === "date");
  // the bank's own columns identify a row; used to carry hand-picked values over when a period is reloaded
  const sigKeys = t.columns.filter((c) => !c.manual && c.label !== NAME_LABEL).map((c) => map[c.key]);
  return { cols, rows, dateKey: fileDate ? map[fileDate.key] : null, sigKeys };
}
const rowSig = (m, data) => JSON.stringify(m.sigKeys.map((k) => data[k] ?? null));
// before deleting rows for a reload: remember hand-filled values (e.g. Category) per identical row
async function rememberManual(ds, m, plan) {
  const manualCols = m.cols.filter((c) => c.manual);
  if (!manualCols.length) return null;
  const old = await fetchAll(() => {
    let q = sb.from("dataset_rows").select("id,data").eq("dataset_id", ds.id);
    if (plan.mode === "period") q = q.gte(`data->>${m.dateKey}`, plan.range.from).lt(`data->>${m.dateKey}`, plan.range.to + "~");
    return q.order("id");
  });
  const keep = new Map();
  for (const r of old) {
    const vals = Object.fromEntries(manualCols.filter((c) => r.data[c.key] != null).map((c) => [c.label, r.data[c.key]])); // by title: keys may change
    if (!Object.keys(vals).length) continue;
    const s = rowSig(m, r.data);
    if (!keep.has(s)) keep.set(s, []);
    keep.get(s).push(vals);
  }
  return keep;
}
function restoreManual(m, keep) {
  let n = 0;
  for (const r of m.rows) {
    const vals = keep.get(rowSig(m, r))?.shift();
    if (!vals) continue;
    for (const [label, v] of Object.entries(vals)) { const c = m.cols.find((x) => x.label === label); if (c) r[c.key] = v; }
    n++;
  }
  return n;
}
// the existing table that most likely receives this sheet: same name, or the same column headers
function guessTarget(name, t) {
  const byName = state.datasets.find((d) => d.name === name);
  if (byName) return byName;
  const labels = t.columns.filter((c) => !c.manual && c.label !== NAME_LABEL).map((c) => c.label); // "שם"/"Category" are ours, not the bank's
  const free = state.datasets.filter((d) => !TOPICS.some((x) => x.table === d.name)); // other subjects' tables are never guessed
  return labels.length ? free.find((d) => labels.every((l) => d.columns.some((c) => c.label === l))) : null;
}

$("upload-file").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  $("upload-sheets").innerHTML = ""; $("upload-go").disabled = true; $("upload-progress").textContent = "";
  if (!file) return;
  try {
    const wb = await readWorkbook(file);
    upload = { file, topic: uploadTopic, tables: Object.fromEntries(wb.SheetNames.map((n) => [n, sheetToTable(wb.Sheets[n])])) };
  } catch (err) { console.error(err); toast("לא הצלחתי לקרוא את הקובץ"); return; }
  Object.values(upload.tables).forEach((t) => { t.fits = fitsOf(t); });
  if (upload.topic?.derive) Object.values(upload.tables).forEach((t) => { if (t.fits.includes(upload.topic.id)) upload.topic.derive(t); });
  const opts = state.datasets.map((d) => `<option value="${d.id}">${esc(d.name)}</option>`).join("");
  $("upload-sheets").innerHTML = Object.entries(upload.tables).map(([name, t]) => {
    const dc = t.columns.find((c) => c.type === "date");
    const r = dc && dateRange(t.rows.map((x) => x[dc.key]));
    return `<div class="sheet-opt" data-sheet="${esc(name)}">
      <label><input type="checkbox" class="s-on" ${t.rows.length ? "checked" : ""}> <b>${esc(name)}</b> <span class="muted">${t.rows.length} שורות · ${t.columns.length} עמודות</span></label>
      ${r ? `<div>📅 תאריכים בקובץ: <b>${dmy(r.from)} – ${dmy(r.to)}</b> <span class="muted">(לפי העמודה "${esc(dc.label)}")</span></div>` : ""}
      ${t.columns.length ? `<div class="muted">עמודות: ${t.columns.map((c) => esc(c.label)).join(" · ")}</div>` : ""}
      ${t.skipped.length ? `<details class="muted"><summary>${t.skipped.length} שורות כותרת מעל הטבלה הושמטו</summary>${t.skipped.map((s) => `<div>${esc(s)}</div>`).join("")}</details>` : ""}
      ${lostHebrew(t) ? `<div class="warn">⚠ בקובץ מופיעים סימני שאלה במקום עברית. העברית כבר אבדה בקובץ עצמו (כנראה נשמר מחדש בתוכנה אחרת). כדאי להוריד אותו שוב מהבנק ולהעלות כמו שהוא, בלי לפתוח ולשמור.</div>` : ""}
      <label>לאן: <select class="s-target"><option value="new">טבלה חדשה</option>${opts}</select>
        <input type="text" class="s-name" value="${esc(name)}" placeholder="שם הטבלה"></label>
      <div class="s-plan"></div>
    </div>`;
  }).join("");
  const topic = upload.topic;
  // subject upload: pick the first sheet that fits the subject (a whole workbook may hold many sheets), else the first non-empty one
  const sheetNames = Object.keys(upload.tables);
  const squash = (s) => String(s).toLowerCase().replace(/[^a-z0-9֐-׿]/g, "");
  const fits = (n) => upload.tables[n].rows.length && upload.tables[n].fits.includes(topic.id);
  const named = (n) => squash(n).includes(squash(topic.table).slice(0, 6)); // "category" sheet for "Categories", "cashFlow" for the outflow table
  const pick = topic && (sheetNames.find((n) => fits(n) && named(n)) || sheetNames.find(fits) || sheetNames.find((n) => upload.tables[n].rows.length));
  document.querySelectorAll(".sheet-opt").forEach((el, i) => {
    const t = upload.tables[el.dataset.sheet];
    if (topic) {
      const on = el.dataset.sheet === pick;
      el.querySelector(".s-on").checked = on;
      const guess = state.datasets.find((d) => d.name === topic.table) || guessTarget(el.dataset.sheet, t);
      el.querySelector(".s-target").value = guess ? guess.id : "new";
      el.querySelector(".s-name").value = topic.table;
      el.querySelector(".s-target").closest("label").hidden = true;
      if (Object.keys(upload.tables).length === 1) el.querySelector(".s-on").closest("label").querySelector("input").hidden = true;
    } else {
      const guess = guessTarget(el.dataset.sheet, t);
      if (guess) el.querySelector(".s-target").value = guess.id;
    }
    const sync = () => planSheet(el, i, t);
    el.querySelector(".s-target").addEventListener("change", sync);
    el.querySelector(".s-on").addEventListener("change", sync); sync();
  });
  $("upload-go").disabled = false;
});

// Loading into a table that already has rows: compare date ranges and make the user choose what happens.
async function planSheet(el, i, t) {
  const target = el.querySelector(".s-target").value;
  const box = el.querySelector(".s-plan");
  el.querySelector(".s-name").hidden = target !== "new";
  el.plan = null;
  if (!el.querySelector(".s-on").checked) { box.innerHTML = ""; el.plan = { skip: true }; return; } // only check sheets that will load
  const ds = target === "new" ? null : state.datasets.find((d) => d.id === +target);
  // wrong place? (a subject button, or a subject's table picked by hand)
  const why = wrongPlace(t, upload.topic || topicOf(ds));
  if (why) { box.innerHTML = `<div class="blocked">⛔ ${why}<br>הטעינה חסומה.</div>`; el.plan = { blocked: true }; return; }
  if (!ds) {
    const looks = TOPICS.find((tp) => t.fits.includes(tp.id));
    box.innerHTML = !upload.topic && looks ? `<div class="warn">נראה כמו קובץ של ${esc(looks.table)}. כדי שייכנס לטבלה הנכונה, עדיף להשתמש בכפתור "${esc(looks.label)}".</div>` : "";
    el.plan = { mode: "new" };
    return;
  }
  const m = mapToTable(ds, t);
  // loading into an existing table: most of the file's columns must already be there
  const fileLabels = t.columns.filter((c) => !c.manual && c.label !== NAME_LABEL).map((c) => c.label);
  const shared = fileLabels.filter((l) => ds.columns.some((c) => c.label === l)).length;
  if (ds.columns.length && fileLabels.length && shared / fileLabels.length < 0.5) {
    box.innerHTML = `<div class="blocked">⛔ העמודות בקובץ לא תואמות לטבלה "${esc(ds.name)}" (רק ${shared} מתוך ${fileLabels.length} עמודות משותפות). כנראה בחרתם טבלה לא נכונה. בחרו טבלה אחרת או "טבלה חדשה".<br>הטעינה חסומה.</div>`;
    el.plan = { blocked: true };
    return;
  }
  box.innerHTML = `<span class="muted">בודק את הטבלה הקיימת…</span>`;
  let existing;
  try { existing = await fetchAll(() => sb.from("dataset_rows").select(m.dateKey ? `id,d:data->>${m.dateKey}` : "id").eq("dataset_id", ds.id).order("id")); }
  catch (err) { console.error(err); box.innerHTML = `<span class="warn">לא הצלחתי לבדוק את הטבלה הקיימת</span>`; return; }
  if (el.querySelector(".s-target").value !== target) return; // changed while loading
  const plan = { mode: null, ds, m, total: existing.length };
  el.plan = plan;
  if (!existing.length) { plan.mode = "append"; box.innerHTML = `<div class="muted">הטבלה "${esc(ds.name)}" ריקה, השורות ייטענו אליה.</div>`; return; }

  const fileR = m.dateKey && dateRange(m.rows.map((r) => r[m.dateKey]));
  const tableR = m.dateKey && dateRange(existing.map((r) => r.d));
  let info, periodOpt = "";
  if (fileR && tableR) {
    plan.range = fileR;
    plan.overlap = existing.filter((r) => { const d = day(r.d); return d && d >= fileR.from && d <= fileR.to; }).length;
    info = `בטבלה "${esc(ds.name)}" יש כבר <b>${existing.length}</b> שורות, מ-<b>${dmy(tableR.from)}</b> עד <b>${dmy(tableR.to)}</b>.<br>
      הקובץ: <b>${dmy(fileR.from)} – ${dmy(fileR.to)}</b>, ${m.rows.length} שורות.<br>
      ${plan.overlap ? `<b class="warn">${plan.overlap} שורות בטבלה נמצאות בתוך טווח התאריכים של הקובץ.</b>` : "אין חפיפה בתאריכים בין הקובץ לטבלה."}`;
    if (plan.overlap) periodOpt = `<label class="opt"><input type="radio" name="mode${i}" value="period">
      <span><b>החלפת התקופה ${dmy(fileR.from)} – ${dmy(fileR.to)}</b> (מומלץ): ${plan.overlap} השורות האלה יוחלפו בשורות מהקובץ, והשורות מחוץ לתקופה יישארו. שימו לב: גם שורות שנוספו באתר בתקופה הזאת יוחלפו.${m.cols.some((c) => c.manual) ? " ערכים שבחרתם ידנית (Category) יישמרו לשורות שמופיעות שוב בקובץ." : ""}</span></label>`;
  } else {
    info = `בטבלה "${esc(ds.name)}" יש כבר <b>${existing.length}</b> שורות${m.dateKey ? "" : " (אין עמודת תאריך להשוואה)"}.`;
  }
  box.innerHTML = `<div class="plan-box"><div>${info}</div><div class="muted">מה לעשות? (חובה לבחור)</div>
    ${periodOpt}
    <label class="opt"><input type="radio" name="mode${i}" value="append">
      <span><b>הוספה בלבד</b>: כל ${m.rows.length} שורות הקובץ יתווספו, בלי למחוק כלום${plan.overlap ? `. <span class="warn">שורות בתקופה החופפת יופיעו פעמיים.</span>` : "."}</span></label>
    <label class="opt danger-opt"><input type="radio" name="mode${i}" value="all">
      <span><b>⚠ דריסת כל הטבלה</b>: מחיקת כל ${existing.length} השורות הקיימות (כולל מה שנוסף באתר) והשארת תוכן הקובץ בלבד. אין דרך חזרה.${m.cols.some((c) => c.manual) ? " (Category שנבחר ידנית יישמר רק לשורות שמופיעות שוב בקובץ.)" : ""}</span></label>
  </div>`;
  box.querySelectorAll("input[type=radio]").forEach((r) => r.addEventListener("change", () => { plan.mode = r.value; }));
}

$("upload-go").addEventListener("click", async () => {
  const jobs = [...document.querySelectorAll(".sheet-opt")].filter((el) => el.querySelector(".s-on").checked);
  if (!jobs.length) return toast("לא נבחר אף גיליון");
  for (const el of jobs) {
    if (!el.plan) return toast("רגע, עדיין בודק את הטבלה הקיימת");
    if (el.plan.blocked) { el.scrollIntoView({ block: "center" }); return toast("הקובץ לא מתאים לטבלה הזאת, הטעינה חסומה", 4000); }
    if (!el.plan.mode) { el.scrollIntoView({ block: "center" }); return toast(`צריך לבחור מה לעשות עם הנתונים הקיימים בטבלה "${el.plan.ds.name}"`, 4000); }
    if (el.plan.mode === "all" && !confirm(`⚠ למחוק את כל ${el.plan.total} השורות בטבלה "${el.plan.ds.name}" ולהשאיר רק את תוכן הקובץ? אין דרך חזרה.`)) return;
  }
  $("upload-go").disabled = true;
  let lastId = null;
  try {
    for (const el of jobs) {
      const t = upload.tables[el.dataset.sheet];
      const plan = el.plan;
      if (plan.mode === "new") {
        const name = el.querySelector(".s-name").value.trim() || el.dataset.sheet;
        const { data, error } = await sb.from("datasets").insert({ name, columns: t.columns, source_file: upload.file.name }).select().single();
        if (error) throw error.code === "23505" ? new Error(`כבר קיימת טבלה בשם "${name}"`) : error;
        await insertRows(data.id, t.rows, "excel");
        lastId = data.id;
        continue;
      }
      const { ds } = plan;
      const keep = plan.mode === "all" || plan.mode === "period" ? await rememberManual(ds, plan.m, plan) : null;
      // overwriting everything also resets the column list to the file's columns (fixes a table loaded with shifted titles)
      const m = plan.mode === "all" ? mapToTable({ ...ds, columns: [] }, t) : plan.m;
      if (plan.mode === "all") {
        const { error } = await sb.from("dataset_rows").delete().eq("dataset_id", ds.id);
        if (error) throw error;
      } else if (plan.mode === "period") {
        $("upload-progress").textContent = `מוחק את השורות הקיימות בתקופה ${dmy(plan.range.from)} – ${dmy(plan.range.to)}…`;
        const { error } = await sb.from("dataset_rows").delete().eq("dataset_id", ds.id)
          .gte(`data->>${m.dateKey}`, plan.range.from).lt(`data->>${m.dateKey}`, plan.range.to + "~"); // "~" sorts after any time part
        if (error) throw error;
      }
      if (keep) restoreManual(m, keep);
      const patch = { columns: m.cols, source_file: upload.file.name };
      if (upload.topic && ds.name !== upload.topic.table) patch.name = upload.topic.table; // a matching table loaded before the subject buttons existed
      const { error } = await sb.from("datasets").update(patch).eq("id", ds.id);
      if (error) throw error;
      await insertRows(ds.id, m.rows, "excel");
      lastId = ds.id;
    }
    sb.from("usage_logs").insert({ visitor_name: user.email, action: `upload:${upload.file.name}` }).then(() => {});
    $("upload-dlg").close();
    toast("הקובץ נטען בהצלחה");
    await loadDatasets();
    if (jobs.length === 1 && lastId) await openDataset(lastId); else { state.ds = null; render(); }
  } catch (err) {
    console.error(err);
    toast(`הטעינה נכשלה: ${err.message}`, 6000);
  } finally { $("upload-go").disabled = false; }
});
$("upload-cancel").addEventListener("click", () => $("upload-dlg").close());
function openUpload(topic = null) {
  upload = null; uploadTopic = topic;
  $("upload-title").textContent = topic ? topic.label : "טעינת קובץ"; $("upload-file").value = ""; $("upload-sheets").innerHTML = ""; $("upload-progress").textContent = ""; $("upload-go").disabled = true;
  $("upload-dlg").showModal();
}

// ---------------- Excel out ----------------
function toSheet(ds, rows) {
  const aoa = [ds.columns.map((c) => c.label)];
  for (const r of rows) aoa.push(ds.columns.map((c) => {
    const v = r.data[c.key] ?? null;
    const m = c.type === "date" && typeof v === "string" && v.match(/^(\d{4})-(\d\d)-(\d\d)(?: (\d\d):(\d\d))?$/);
    if (m) return new Date(+m[1], m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0)); // local time, so Excel shows the same day
    return v;
  }));
  const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true, dateNF: "dd/mm/yyyy" });
  ws["!cols"] = ds.columns.map((c) => ({ wch: Math.min(40, Math.max(10, c.label.length + 2)) }));
  return ws;
}
const sheetName = (s, used) => { let n = s.replace(/[\\/?*[\]:]/g, " ").slice(0, 31) || "Sheet"; let i = 2; while (used.has(n)) n = `${n.slice(0, 28)} ${i++}`; used.add(n); return n; };
function saveWorkbook(sheets, fileName) {
  const wb = XLSX.utils.book_new(); const used = new Set();
  for (const [ds, rows] of sheets) XLSX.utils.book_append_sheet(wb, toSheet(ds, rows), sheetName(ds.name, used));
  wb.Workbook = { Views: [{ RTL: true }] };
  XLSX.writeFile(wb, fileName);
}
const stamp = () => new Date().toISOString().slice(0, 10);
async function downloadAll() {
  toast("מכין קובץ…");
  try {
    const sheets = [];
    for (const ds of state.datasets) sheets.push([ds, await loadRows(ds.id)]);
    saveWorkbook(sheets, `מקום בלב_${stamp()}.xlsx`);
  } catch (err) { console.error(err); toast("ההורדה נכשלה"); }
}

// ---------------- tables view ----------------
async function openDataset(id) {
  state.ds = state.datasets.find((d) => d.id === id);
  state.q = ""; state.sort = null; state.shown = SHOW; state.editing = null;
  $("status").textContent = "טוען…";
  try {
    state.rows = await loadRows(id);
    await loadLookups(state.ds);
    $("status").textContent = "";
  }
  catch (err) { console.error(err); $("status").textContent = "שגיאה בטעינת הטבלה"; state.rows = []; }
  render();
}

// dropdown options for lookup columns (e.g. Category <- the Categories table)
async function loadLookups(ds) {
  state.lookups = {};
  for (const c of ds.columns.filter((x) => x.lookup)) {
    const src = state.datasets.find((d) => d.name === c.lookup);
    if (!src) { state.lookups[c.lookup] = null; continue; }
    const col = src.columns.find((x) => /category|קטגורי|סיווג/i.test(x.label)) || src.columns[0];
    const rows = col ? await loadRows(src.id) : [];
    state.lookups[c.lookup] = [...new Set(rows.map((r) => r.data[col.key]).filter((v) => v != null).map(String))];
  }
}

function filteredRows() {
  const { ds, q, sort } = state;
  let rows = state.rows;
  if (q) { const needle = q.toLowerCase(); rows = rows.filter((r) => ds.columns.some((c) => showValue(r.data[c.key], c.type).toLowerCase().includes(needle))); }
  if (sort) {
    const { key, dir } = sort;
    rows = [...rows].sort((a, b) => {
      const x = a.data[key], y = b.data[key];
      if (x == null) return 1; if (y == null) return -1;
      return (typeof x === "number" && typeof y === "number" ? x - y : String(x).localeCompare(String(y), "he", { numeric: true })) * dir;
    });
  }
  return rows;
}

function renderTables() {
  const el = $("view-tables");
  if (!state.ds) {
    el.innerHTML = `<div class="toolbar"><h2>טבלאות</h2>
      ${TOPICS.map((tp) => `<button class="btn primary" data-topic="${tp.id}">⬆ ${esc(tp.label)}</button>`).join("")}
      <button class="btn" id="btn-upload">⬆ טעינת קובץ אחר</button>
      ${state.datasets.length ? `<button class="btn" id="btn-dl-all">⬇ הורדת הכול לאקסל</button>` : ""}</div>
      ${state.datasets.length ? `<div class="grid-cards">${state.datasets.map((d) => `
        <div class="card ds-card" data-open="${d.id}"><h3>${esc(d.name)}</h3>
        <div class="muted">${state.counts[d.id] ?? 0} שורות · ${d.columns.length} עמודות</div>
        <div class="muted">עודכן ${fmtDate(d.updated_at)}${d.source_file ? ` · מקור: ${esc(d.source_file)}` : ""}</div></div>`).join("")}</div>`
      : `<div class="card empty">עדיין אין טבלאות. לחצו על כפתור הטעינה של הנושא (למשל "${esc(TOPICS[0].label)}") כדי לטעון קובץ.</div>`}`;
    return;
  }
  const { ds } = state;
  const rows = filteredRows();
  const arrow = (k) => state.sort?.key === k ? (state.sort.dir > 0 ? " ▲" : " ▼") : "";
  el.innerHTML = `<div class="toolbar">
      <button class="btn" id="btn-back">→ כל הטבלאות</button>
      <h2>${esc(ds.name)}</h2>
      <input type="search" id="q" placeholder="חיפוש…" value="${esc(state.q)}">
      <button class="btn primary" id="btn-add-row">+ שורה</button>
      <button class="btn" id="btn-add-col">+ עמודה</button>
      <button class="btn" id="btn-dl">⬇ אקסל</button>
      ${topicOf(ds)?.derive ? `<button class="btn" id="btn-derive" title="ממלא את עמודת &quot;שם&quot; בשורות שבהן היא ריקה">🔄 חילוץ שמות</button>` : ""}
      <button class="btn" id="btn-rename">שינוי שם</button>
      ${isAdmin ? `<button class="btn danger" id="btn-del-ds">מחיקת טבלה</button>` : ""}
    </div>
    <p class="muted">${rows.length === state.rows.length ? `${rows.length} שורות` : `${rows.length} מתוך ${state.rows.length} שורות`} · לחיצה על תא לעריכה · לחיצה על כותרת למיון · מספר שורה בירוק = נוספה באתר</p>
    <div class="table-wrap"><table class="data"><thead><tr><th class="rowno">#</th>
      ${ds.columns.map((c) => `<th data-sort="${c.key}" title="לחיצה ימנית: שינוי שם עמודה">${esc(c.label)}${arrow(c.key)}</th>`).join("")}<th></th></tr></thead>
    <tbody>${rows.slice(0, state.shown).map((r, i) => `<tr data-row="${r.id}" class="${r.origin}"><td class="rowno">${i + 1}</td>
      ${ds.columns.map((c) => `<td class="cell${c.type === "number" ? " num" : ""}" data-col="${c.key}">${state.editing?.row === r.id && state.editing.col === c.key ? "" : esc(showValue(r.data[c.key], c.type))}</td>`).join("")}
      <td><button class="row-del" data-del="${r.id}" title="מחיקת שורה">🗑</button></td></tr>`).join("")}</tbody></table></div>
    ${rows.length > state.shown ? `<button class="btn more" id="btn-more">הצגת עוד ${Math.min(SHOW, rows.length - state.shown)}</button>` : ""}`;
  if (state.editing) startEdit(state.editing.row, state.editing.col);
}

function startEdit(rowId, col) {
  const td = document.querySelector(`tr[data-row="${rowId}"] td[data-col="${col}"]`);
  if (!td) { state.editing = null; return; }
  const row = state.rows.find((r) => r.id === rowId);
  const c = state.ds.columns.find((x) => x.key === col);
  state.editing = { row: rowId, col };
  if (c.lookup || c.options) {
    const opts = c.options || state.lookups?.[c.lookup];
    const cur = row.data[col] == null ? "" : String(row.data[col]);
    const list = opts ? [...opts] : [];
    if (cur && !list.includes(cur)) list.unshift(cur); // keep a value that was removed from the list
    td.innerHTML = `<select><option value="">${opts ? (opts.length ? "— בחירה —" : `— אין עדיין ערכים בטבלה ${esc(c.lookup)} —`) : `— צריך קודם לטעון את טבלת ${esc(c.lookup)} —`}</option>
      ${list.map((v) => `<option${v === cur ? " selected" : ""}>${esc(v)}</option>`).join("")}</select>`;
  } else {
    td.innerHTML = `<input type="text" value="${esc(showValue(row.data[col], c.type))}" ${c.type === "number" ? 'dir="ltr"' : ""} placeholder="${c.type === "date" ? "dd/mm/yyyy" : ""}">`;
  }
  const input = td.querySelector("input, select");
  input.focus(); if (input.select && input.tagName === "INPUT") input.select();
  if (c.lookup || c.options) input.addEventListener("change", () => input.blur());
  let done = false;
  const finish = async (save) => {
    if (done) return; done = true;
    state.editing = null;
    const val = parseInput(input.value, c.type);
    const old = row.data[col] ?? null;
    if (save && val !== old) {
      const data = { ...row.data };
      if (val === null) delete data[col]; else data[col] = val;
      const { data: upd, error } = await sb.from("dataset_rows").update({ data }).eq("id", row.id).select().single();
      if (error) { console.error(error); toast("השמירה נכשלה"); }
      else Object.assign(row, upd);
    }
    td.textContent = showValue(row.data[col], c.type);
  };
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); input.blur(); }
    if (e.key === "Escape") { finish(false); }
  });
  input.addEventListener("blur", () => finish(true));
}

$("view-tables").addEventListener("click", async (e) => {
  const t = e.target;
  if (t.closest("#btn-upload")) return openUpload();
  const tp = t.closest("[data-topic]"); if (tp) return openUpload(TOPICS.find((x) => x.id === tp.dataset.topic));
  if (t.closest("#btn-dl-all")) return downloadAll();
  const card = t.closest("[data-open]"); if (card) return openDataset(+card.dataset.open);
  if (t.closest("#btn-back")) { state.ds = null; await loadDatasets(); return render(); }
  if (t.closest("#btn-more")) { state.shown += SHOW; return render(); }
  if (t.closest("#btn-dl")) return saveWorkbook([[state.ds, filteredRows()]], `${state.ds.name}_${stamp()}.xlsx`);
  const th = t.closest("th[data-sort]");
  if (th) { const k = th.dataset.sort; state.sort = state.sort?.key === k ? (state.sort.dir > 0 ? { key: k, dir: -1 } : null) : { key: k, dir: 1 }; return render(); }
  const td = t.closest("td.cell");
  if (td && !td.querySelector("input")) return startEdit(+td.parentElement.dataset.row, td.dataset.col);
  const del = t.closest("[data-del]");
  if (del) {
    if (!confirm("למחוק את השורה?")) return;
    const id = +del.dataset.del;
    const { error } = await sb.from("dataset_rows").delete().eq("id", id);
    if (error) return toast("המחיקה נכשלה");
    state.rows = state.rows.filter((r) => r.id !== id); return render();
  }
  if (t.closest("#btn-add-row")) {
    const { data, error } = await sb.from("dataset_rows").insert({ dataset_id: state.ds.id, data: {}, origin: "site" }).select().single();
    if (error) { console.error(error); return toast("ההוספה נכשלה"); }
    state.rows.push(data); state.q = ""; state.sort = null; state.shown = Math.max(state.shown, state.rows.length);
    state.editing = { row: data.id, col: state.ds.columns[0]?.key };
    render();
    document.querySelector(`tr[data-row="${data.id}"]`)?.scrollIntoView({ block: "center" });
    return;
  }
  if (t.closest("#btn-add-col")) {
    const label = prompt("שם העמודה החדשה:")?.trim(); if (!label) return;
    if (state.ds.columns.some((c) => c.label === label)) return toast("כבר יש עמודה בשם הזה");
    return saveColumns([...state.ds.columns, { key: nextKey(state.ds.columns), label, type: "text" }]);
  }
  if (t.closest("#btn-derive")) return deriveExisting(t.closest("#btn-derive"));
  if (t.closest("#btn-rename")) {
    const name = prompt("שם חדש לטבלה:", state.ds.name)?.trim(); if (!name || name === state.ds.name) return;
    const { error } = await sb.from("datasets").update({ name }).eq("id", state.ds.id);
    if (error) return toast(error.code === "23505" ? "כבר יש טבלה בשם הזה" : "השינוי נכשל");
    state.ds.name = name; return render();
  }
  if (t.closest("#btn-del-ds")) {
    if (prompt(`מחיקת הטבלה "${state.ds.name}" וכל ${state.rows.length} השורות שלה. כדי לאשר, הקלידו את שם הטבלה:`) !== state.ds.name) return;
    const { error } = await sb.from("datasets").delete().eq("id", state.ds.id);
    if (error) return toast("המחיקה נכשלה");
    state.ds = null; await loadDatasets(); toast("הטבלה נמחקה"); return render();
  }
});
$("view-tables").addEventListener("contextmenu", (e) => { // right-click a header: rename the column
  const th = e.target.closest("th[data-sort]"); if (!th) return;
  e.preventDefault();
  const col = state.ds.columns.find((c) => c.key === th.dataset.sort);
  const label = prompt("שם חדש לעמודה:", col.label)?.trim(); if (!label || label === col.label) return;
  saveColumns(state.ds.columns.map((c) => (c.key === col.key ? { ...c, label } : c)));
});
$("view-tables").addEventListener("input", (e) => {
  if (e.target.id !== "q") return;
  state.q = e.target.value; state.shown = SHOW; render();
  const q = $("q"); q.focus(); q.setSelectionRange(q.value.length, q.value.length);
});
// run the subject's derive rule on rows already in the table (e.g. loaded before the rule existed)
async function deriveExisting(btn) {
  const ds = state.ds, tp = topicOf(ds);
  const t = { columns: ds.columns.map((c) => ({ ...c })), rows: state.rows.map((r) => ({ ...r.data })) };
  const found = tp.derive(t);
  btn.disabled = true;
  try {
    if (JSON.stringify(t.columns) !== JSON.stringify(ds.columns)) { // e.g. the "שם" / "Category" columns were just added
      const { error } = await sb.from("datasets").update({ columns: t.columns }).eq("id", ds.id);
      if (error) throw error;
      ds.columns = t.columns;
      await loadLookups(ds);
    }
    if (!found) { btn.disabled = false; render(); return toast("לא נמצאו שמות חדשים לחילוץ"); }
    const changed = state.rows.map((r, i) => [r, t.rows[i]]).filter(([r, d]) => JSON.stringify(r.data) !== JSON.stringify(d));
    for (let i = 0; i < changed.length; i += 10) {
      await Promise.all(changed.slice(i, i + 10).map(async ([r, data]) => {
        const { error } = await sb.from("dataset_rows").update({ data }).eq("id", r.id);
        if (error) throw error;
        r.data = data;
      }));
      $("status").textContent = `מעדכן שמות… ${Math.min(i + 10, changed.length)}/${changed.length}`;
    }
    $("status").textContent = "";
    toast(`נמצאו שמות ל-${changed.length} שורות`);
  } catch (err) { console.error(err); $("status").textContent = ""; toast("העדכון נכשל באמצע, אפשר ללחוץ שוב"); }
  render();
}
async function saveColumns(columns) {
  const { error } = await sb.from("datasets").update({ columns }).eq("id", state.ds.id);
  if (error) return toast("השמירה נכשלה");
  state.ds.columns = columns; render();
}

// ---------------- users (admins) ----------------
async function loadUsers() {
  const { data, error } = await sb.from("allowed_users").select("*").order("added_at");
  if (!error) state.users = data;
}
function renderUsers() {
  const self = (u) => u.email.toLowerCase() === user.email.toLowerCase();
  $("view-users").innerHTML = `<form class="add-user" id="add-user">
      <input type="email" name="email" placeholder="כתובת Gmail" required dir="ltr">
      <input type="text" name="full_name" placeholder="שם (לא חובה)">
      <label><input type="checkbox" name="admin"> מנהל/ת</label>
      <button class="btn primary">הוספת משתמש/ת</button></form>
    <table class="users-table"><thead><tr><th>מייל</th><th>שם</th><th>הרשאה</th><th>נוסף ע"י</th><th></th></tr></thead><tbody>
    ${state.users.map((u) => `<tr><td dir="ltr">${esc(u.email)}</td><td>${esc(u.full_name || "")}</td>
      <td>${u.is_admin ? "מנהל/ת" : "משתמש/ת"}</td><td dir="ltr">${esc(u.added_by || "")}</td>
      <td>${self(u) ? "" : `<button class="link-btn" data-toggle-admin="${esc(u.email)}">${u.is_admin ? "הסרת ניהול" : "הפיכה למנהל/ת"}</button> · <button class="link-btn" data-del-user="${esc(u.email)}">הסרה</button>`}</td></tr>`).join("")}
    </tbody></table>`;
}
$("view-users").addEventListener("submit", async (e) => {
  e.preventDefault();
  const f = new FormData(e.target);
  const email = f.get("email").trim().toLowerCase();
  const { error } = await sb.from("allowed_users").insert({ email, full_name: f.get("full_name").trim() || null, is_admin: f.has("admin") });
  if (error) return toast(error.code === "23505" ? "המשתמש/ת כבר ברשימה" : "ההוספה נכשלה");
  await loadUsers(); renderUsers(); toast(`${email} נוסף/ה`);
});
$("view-users").addEventListener("click", async (e) => {
  const del = e.target.closest("[data-del-user]");
  if (del) {
    if (!confirm(`להסיר את ${del.dataset.delUser}?`)) return;
    const { error } = await sb.from("allowed_users").delete().eq("email", del.dataset.delUser);
    if (error) return toast("ההסרה נכשלה");
  }
  const tog = e.target.closest("[data-toggle-admin]");
  if (tog) {
    const u = state.users.find((x) => x.email === tog.dataset.toggleAdmin);
    const { error } = await sb.from("allowed_users").update({ is_admin: !u.is_admin }).eq("email", u.email);
    if (error) return toast("השינוי נכשל");
  }
  if (del || tog) { await loadUsers(); renderUsers(); }
});

// ---------------- tabs / render ----------------
document.querySelector(".tabs").addEventListener("click", (e) => {
  const b = e.target.closest(".tab"); if (!b) return;
  state.tab = b.dataset.tab;
  if (state.tab === "reports" && typeof rep !== "undefined") rep.cache = {}; // always report on the latest data
  render();
});
function render() {
  document.querySelectorAll(".tab").forEach((b) => b.classList.toggle("active", b.dataset.tab === state.tab));
  for (const v of ["tables", "reports", "users"]) $(`view-${v}`).hidden = state.tab !== v;
  if (state.tab === "tables") renderTables();
  else if (state.tab === "reports") renderReports();
  else renderUsers();
}

// ---------------- sign-in ----------------
const redirectTo = location.origin + location.pathname;
$("google").addEventListener("click", () => sb.auth.signInWithOAuth({ provider: "google", options: { redirectTo } }));
document.addEventListener("click", (e) => { if (e.target.closest("#logout")) sb.auth.signOut().then(() => location.reload()); });

async function consumeAuthRedirect() {
  const hash = new URLSearchParams(location.hash.slice(1));
  const query = new URLSearchParams(location.search);
  let err = hash.get("error_description") || query.get("error_description");
  try {
    if (hash.get("access_token") && hash.get("refresh_token")) {
      const { error } = await sb.auth.setSession({ access_token: hash.get("access_token"), refresh_token: hash.get("refresh_token") });
      if (error) err = error.message;
    } else if (query.get("code")) {
      const { error } = await sb.auth.exchangeCodeForSession(query.get("code"));
      if (error) err = error.message;
    }
  } catch (e) { err = e.message; }
  if (location.hash || location.search) history.replaceState(null, "", location.pathname);
  return err;
}

function showLogin(html) {
  $("status").textContent = "";
  $("login").hidden = false;
  $("google").hidden = !!html;
  $("noaccess").hidden = !html;
  if (html) $("noaccess").innerHTML = html;
}

async function start() {
  const authErr = await consumeAuthRedirect();
  const { data: { session } } = await sb.auth.getSession();
  user = session?.user || null;
  if (!user) { showLogin(); if (authErr) { $("noaccess").hidden = false; $("noaccess").textContent = `הכניסה נכשלה: ${authErr}`; } return; }
  const { data: me, error: meErr } = await sb.from("allowed_users").select("email,is_admin").ilike("email", user.email).maybeSingle();
  if (meErr && ["PGRST205", "PGRST106", "42P01", "3F000"].includes(meErr.code)) {
    return showLogin(`מחובר/ת בתור <b dir="ltr">${esc(user.email)}</b> ✓<br>המערכת עדיין לא הוגדרה: צריך להריץ את supabase_setup.sql ב-Supabase ולהוסיף את הסכמה makom_balev ל-Exposed schemas.<br><button class="link-btn" id="logout">התנתקות</button>`);
  }
  if (meErr || !me) {
    return showLogin(`מחובר/ת בתור <b dir="ltr">${esc(user.email)}</b>, אבל הכתובת עדיין לא ברשימת המורשים. בקשו ממנהל להוסיף אותה.<br><button class="link-btn" id="logout">התחברות עם חשבון אחר</button>`);
  }
  isAdmin = me.is_admin;
  document.querySelector('[data-tab="users"]').hidden = !isAdmin;
  document.querySelector(".tabs").hidden = false;
  $("who").innerHTML = `${esc(user.user_metadata?.full_name || user.email)} · <button class="link-btn" id="logout">התנתקות</button>`;
  try {
    await Promise.all([loadDatasets(), isAdmin ? loadUsers() : null]);
    await ensureCategories();
    $("status").textContent = "";
    render();
    sb.from("usage_logs").insert({ visitor_name: user.email, action: "page_view" }).then(() => {});
  } catch (err) {
    console.error(err);
    $("status").textContent = "שגיאה בטעינת הנתונים, נסו לרענן";
  }
}
start();
