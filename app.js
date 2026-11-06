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
  { id: "cashflow", label: "טעינת תנועות בנק", table: "Cash Flow", derive: deriveCashFlow, kind: "bank" },
  // the bank shows one total charge per card; the card files hold the details -> separate tables, never mixed into Cash Flow
  { id: "visa", label: "טעינת לאומי ויזה", table: "לאומי ויזה", derive: deriveCashFlow, kind: "card" },
  { id: "mastercard", label: "טעינת ל.מאסטרקארד", table: "ל.מאסטרקארד", derive: deriveCashFlow, kind: "card" },
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
// ---------------- accounts: the bank account + the 2 credit cards, each with the number its files must carry ----------------
// Kept in a small "Accounts" table (editable on the site). The number is learned on the first confirmed upload.
const ACCOUNTS_TABLE = "Accounts";
const ACCOUNT_COLS = [{ key: "c1", label: "table", type: "text" }, { key: "c2", label: "name", type: "text" }, { key: "c3", label: "number", type: "text" }];
async function ensureAccounts() {
  if (state.datasets.some((d) => d.name === ACCOUNTS_TABLE)) return;
  const { data, error } = await sb.from("datasets").insert({ name: ACCOUNTS_TABLE, columns: ACCOUNT_COLS, source_file: null }).select().single();
  if (error) { if (error.code !== "23505") console.error(error); return; }
  const rows = TOPICS.filter((tp) => tp.kind).map((tp) => ({ c1: tp.table, c2: tp.kind === "bank" ? "חשבון בנק לאומי" : tp.table }));
  const { error: e2 } = await sb.from("dataset_rows").insert(rows.map((d) => ({ dataset_id: data.id, data: d, origin: "site" })));
  if (e2) console.error(e2);
  await loadDatasets();
}
async function loadAccounts() {
  const ds = state.datasets.find((d) => d.name === ACCOUNTS_TABLE);
  state.accounts = ds ? (await loadRows(ds.id)).map((r) => ({ id: r.id, data: r.data, table: r.data.c1, name: r.data.c2, number: r.data.c3 == null ? "" : String(r.data.c3) })) : [];
  return state.accounts;
}
const accountOf = (topic) => state.accounts?.find((a) => a.table === topic.table) || { table: topic.table, name: topic.table, number: "" };
async function saveAccountNumber(acc, number) {
  if (!acc.id) return;
  const data = { ...acc.data, c3: number };
  const { error } = await sb.from("dataset_rows").update({ data }).eq("id", acc.id);
  if (!error) { acc.data = data; acc.number = number; }
}
const digits = (s) => String(s ?? "").replace(/\D/g, "");
// the account / card number written in a file (header lines, columns, details)
function detectNumber(t, kind) {
  if (kind === "bank") {
    const br = t.columns.find((c) => /סניף|branch/i.test(c.label)), ac = t.columns.find((c) => /^\s*(מספר\s*)?חשבון\s*$|account/i.test(c.label));
    if (ac) {
      const count = {};
      for (const r of t.rows) { const v = r[ac.key]; if (v == null) continue; const k = (br && r[br.key] != null ? r[br.key] + "-" : "") + v; count[k] = (count[k] || 0) + 1; }
      const top = Object.entries(count).sort((a, b) => b[1] - a[1])[0];
      if (top) return top[0];
    }
    const m = t.skipped.join(" ").match(/(\d{2,3}[-\s]\d{4,}(?:\/\d{2})?)/);
    return m ? m[1].replace(/\s/, "-") : null;
  }
  // card: "המסתיים ב-1234", "כרטיס 1234", "****1234", or a column of last-4 digits
  const text = [...t.skipped, ...t.columns.map((c) => c.label)].join(" ");
  const m = text.match(/(?:המסתיים\s*ב-?|מסתיים\s*ב-?|כרטיס[^\d]{0,20}|card[^\d]{0,20}|\*{2,})(\d{4})(?!\d)/i);
  if (m) return m[1];
  const col = t.columns.find((c) => /4 ספרות|ספרות אחרונות|מספר כרטיס|last 4|card/i.test(c.label));
  if (col) {
    const v = t.rows.map((r) => digits(r[col.key]).slice(-4)).filter((x) => x.length === 4);
    if (v.length) return v.sort((a, b) => v.filter((x) => x === b).length - v.filter((x) => x === a).length)[0];
  }
  return null;
}
const sameNumber = (kind, a, b) => kind === "card" ? digits(a).slice(-4) === digits(b).slice(-4) : digits(a) === digits(b);
const numberLabel = (kind, n) => (kind === "card" ? `כרטיס המסתיים ב-${digits(n).slice(-4)}` : `חשבון ${n}`);
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
// a file's own category column: "category", "קטגוריה", "סיווג", or "outflows" (as in the expenses / recieve sheets)
const CAT_RE = /category|קטגורי|סיווג|^\s*outflows?\s*$/i;
function nameFromRow(row, columns) {
  // card statements name the business in its own column
  const biz = columns.find((c) => /בית\s*ה?עסק|merchant|שם העסק/i.test(c.label));
  if (biz && typeof row[biz.key] === "string" && row[biz.key].trim()) return row[biz.key].trim();
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
  const own = t.columns.find((c) => CAT_RE.test(c.label) && !c.manual && c.label !== CATEGORY_COL.label);
  if (own) addPickers(t.columns);
  else {
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
// rows typed on the site (e.g. Cash Flow Outflow Uptoday): category = dropdown from Categories, type = Inflows/Outflows.
// Returns true when a column changed.
function addPickers(columns) {
  let changed = false;
  for (const c of columns) {
    if (CAT_RE.test(c.label) && !c.manual && c.label !== CATEGORY_COL.label && !c.lookup) { c.lookup = CATEGORIES_TABLE; changed = true; }
    if (/^\s*type\s*$/i.test(c.label) && !c.options) { c.options = CATEGORY_TYPES; changed = true; }
  }
  return changed;
}
async function ensurePickers(ds) { // tables loaded before the pickers existed
  if (!topicOf(ds)?.derive) return;
  const columns = ds.columns.map((c) => ({ ...c }));
  if (!addPickers(columns)) return;
  const { error } = await sb.from("datasets").update({ columns }).eq("id", ds.id);
  if (!error) ds.columns = columns;
}

// ---------------- which file belongs where (blocks loading a file into the wrong subject) ----------------
const hasCol = (t, re, type) => t.columns.some((c) => re.test(c.label) && (!type || c.type === type));
const isBank = (t) => hasCol(t, /תאריך|date/i) && hasCol(t, /חובה|debit/i) && hasCol(t, /זכות|credit/i) && hasCol(t, /יתרה|balance/i);
// credit card statement: dates + amounts (charge/transaction), no running bank balance and no debit/credit pair
const isCard = (t) => !isBank(t) && t.columns.some((c) => c.type === "date") && hasCol(t, /סכום|חיוב|amount|charge/i)
  && !hasCol(t, /^\s*month\s*$|חודש/i) && !hasCol(t, CAT_RE);
const CARD_RULE = { test: isCard, need: "תאריך עסקה וסכום חיוב (קובץ פירוט עסקאות של כרטיס אשראי, בלי עמודת יתרה)" };
const FILE_RULES = {
  cashflow: { test: isBank, need: "תאריך, בחובה, בזכות, יתרה בש\"ח (קובץ תנועות עו\"ש מהבנק)" },
  visa: CARD_RULE,
  mastercard: CARD_RULE,
  outflow: {
    test: (t) => !isBank(t) && hasCol(t, CAT_RE) && (hasCol(t, /month|חודש/i) || t.columns.some((c) => c.type === "date")) && hasCol(t, /amount|סכום|חובה|זכות|outflow|הוצא/i),
    need: "category, חודש או תאריך, וסכום (category / month / type / amountNIS)",
  },
};
const KIND_NAME = { bank: "תנועות חשבון בנק", card: "פירוט כרטיס אשראי" };
const fitsOf = (t) => TOPICS.filter((tp) => FILE_RULES[tp.id]?.test(t)).map((tp) => tp.id); // run on the file as read, before derive
// null = OK, otherwise why this file can't go into that subject's table
function wrongPlace(t, topic) {
  if (!topic || !FILE_RULES[topic.id] || t.fits.includes(topic.id)) return null;
  // bank file into a card, or a card file into the bank: say it in those words
  if (topic.kind && t.kind && t.kind !== topic.kind) {
    const other = TOPICS.find((tp) => tp.kind && tp.kind !== topic.kind);
    return `זה קובץ של <b>${KIND_NAME[topic.kind === "card" ? "bank" : "card"]}</b>, ולא של ${KIND_NAME[topic.kind]}. טענו אותו בריבוע של ${esc(other.table)}${topic.kind === "bank" ? " או של הכרטיס השני" : ""}.`;
  }
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
    const d = s.match(/^(\d{1,2})([./-])(\d{1,2})\2(\d{4}|\d{2})$/); // 01/11/2026, 3.6.25, 10-02-2025 (day first)
    if (d && +d[3] >= 1 && +d[3] <= 12 && +d[1] >= 1 && +d[1] <= 31) return `${d[4].length === 2 ? "20" + d[4] : d[4]}-${pad(d[3])}-${pad(d[1])}`;
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
  // a plain number typed into a text column is stored as a number (so sums work); "050..." stays text
  if (/^-?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$/.test(t) && !/^-?0\d/.test(t)) return Number(t.replace(/,/g, ""));
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
  const hi = filled.findIndex((n) => n >= Math.max(Math.min(2, widest), Math.ceil(widest * 0.6))); // a one-column list is fine too
  if (hi < 0) return { columns: [], rows: [], skipped: [] };
  const skipped = aoa.slice(0, hi).map((r) => r.map(cellValue).filter((v) => v !== null).join(" ")).filter(Boolean);
  let header = aoa[hi].map((v) => (v === null || v === undefined ? null : String(v).trim() || null)); // headers stay text
  const body = aoa.slice(hi + 1).map((r) => r.map(cellValue)).filter((r) => r.some((v) => v !== null));
  // Leumi's CSV: the header row starts one cell later than the data rows -> every title lands one column off.
  // If the header has more empty cells in front than any data row, line it up with the data.
  const lead = (r) => { const i = r.findIndex((v) => v !== null && v !== undefined); return i < 0 ? Infinity : i; };
  const shift = lead(header) - Math.min(...body.map(lead));
  if (body.length && shift > 0 && shift < Infinity) header = header.slice(shift);
  let width = body.reduce((w, r) => Math.max(w, r.length), header.length);
  // a summary block next to the data (e.g. a monthly matrix to the side) is not part of the table:
  // the table ends at the first column that is completely empty (no title, no values)
  const first = header.findIndex((v) => v != null);
  for (let j = first + 1; j < width; j++) if (header[j] == null && body.every((r) => (r[j] ?? null) === null)) { width = j; break; }
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
    else {
      if (c.lookup) Object.assign(hit, { lookup: c.lookup, manual: c.manual });
      if (c.options) hit.options = c.options;
    }
    map[c.key] = hit.key;
  }
  const rows = t.rows.map((r) => Object.fromEntries(Object.entries(r).map(([k, v]) => [map[k], v])));
  // card statements: the charge date is the same for the whole monthly file, so it marks the file's period exactly
  const fileDate = t.columns.find((c) => c.type === "date" && /חיוב|charge|billing/i.test(c.label)) || t.columns.find((c) => c.type === "date");
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
    const safe = (n) => { try { return sheetToTable(wb.Sheets[n]); } catch (err) { console.error(n, err); return { columns: [], rows: [], skipped: [], broken: true }; } };
    upload = { file, topic: uploadTopic, tables: Object.fromEntries(wb.SheetNames.map((n) => [n, safe(n)])) };
  } catch (err) { console.error(err); toast("לא הצלחתי לקרוא את הקובץ"); return; }
  Object.values(upload.tables).forEach((t) => { t.fits = fitsOf(t); t.kind = isBank(t) ? "bank" : isCard(t) ? "card" : null; }); // before our own columns are added
  // name + Category columns for any file that fits a subject with a derive rule (the account question may still change the subject)
  if (upload.topic?.derive) Object.values(upload.tables).forEach((t) => { if (t.fits.some((id) => TOPICS.find((x) => x.id === id)?.derive)) deriveCashFlow(t); });
  const opts = state.datasets.map((d) => `<option value="${d.id}">${esc(d.name)}</option>`).join("");
  $("upload-sheets").innerHTML = Object.entries(upload.tables).map(([name, t]) => {
    const dc = t.columns.find((c) => c.type === "date" && /חיוב|charge|billing/i.test(c.label)) || t.columns.find((c) => c.type === "date");
    const r = dc && dateRange(t.rows.map((x) => x[dc.key]));
    return `<div class="sheet-opt" data-sheet="${esc(name)}">
      <label><input type="checkbox" class="s-on" ${t.rows.length ? "checked" : ""}> <b>${esc(name)}</b> <span class="muted">${t.rows.length} שורות · ${t.columns.length} עמודות</span></label>
      ${r ? `<div>📅 תאריכים בקובץ: <b>${dmy(r.from)} – ${dmy(r.to)}</b> <span class="muted">(לפי העמודה "${esc(dc.label)}")</span></div>` : ""}
      ${t.columns.length ? `<div class="muted">עמודות: ${t.columns.map((c) => esc(c.label)).join(" · ")}</div>` : ""}
      ${t.skipped.length ? `<details class="muted"><summary>${t.skipped.length} שורות כותרת מעל הטבלה הושמטו</summary>${t.skipped.map((s) => `<div>${esc(s)}</div>`).join("")}</details>` : ""}
      ${lostHebrew(t) ? `<div class="warn">⚠ בקובץ מופיעים סימני שאלה במקום עברית. העברית כבר אבדה בקובץ עצמו (כנראה נשמר מחדש בתוכנה אחרת). כדאי להוריד אותו שוב מהבנק ולהעלות כמו שהוא, בלי לפתוח ולשמור.</div>` : ""}
      <label>לאן: <select class="s-target"><option value="new">טבלה חדשה</option>${opts}</select>
        <input type="text" class="s-name" value="${esc(name)}" placeholder="שם הטבלה"></label>
      <div class="s-acc"></div>
      <div class="s-plan"></div>
    </div>`;
  }).join("");
  const topic = upload.topic;
  // subject upload: pick the first sheet that fits the subject (a whole workbook may hold many sheets), else the first non-empty one
  const sheetNames = Object.keys(upload.tables);
  const squash = (s) => String(s).toLowerCase().replace(/[^a-z0-9֐-׿]/g, "");
  const fits = (n) => upload.tables[n].rows.length && upload.tables[n].fits.includes(topic.id);
  const named = (n) => squash(n).includes(squash(topic.table).slice(0, 6)); // "category" sheet for "Categories", "cashFlow" for the outflow table
  // the sheet that probably fits: only a hint in the sheet question, never chosen for the user
  const hint = topic ? (sheetNames.find((n) => fits(n) && named(n)) || sheetNames.find(fits)) : null;
  const multi = sheetNames.length > 1;
  document.querySelectorAll(".sheet-opt").forEach((el, i) => {
    const t = upload.tables[el.dataset.sheet];
    el.querySelector(".s-on").hidden = true; // one sheet per load: chosen in the question above
    el.querySelector(".s-on").checked = !multi;
    el.hidden = multi;
    if (topic) {
      const guess = state.datasets.find((d) => d.name === topic.table) || guessTarget(el.dataset.sheet, t);
      el.querySelector(".s-target").value = guess ? guess.id : "new";
      el.querySelector(".s-name").value = topic.table;
      el.querySelector(".s-target").closest("label").hidden = true;
    } else {
      const guess = guessTarget(el.dataset.sheet, t);
      if (guess) el.querySelector(".s-target").value = guess.id;
    }
    el.sync = () => planSheet(el, i, t);
    el.querySelector(".s-target").addEventListener("change", el.sync);
    el.sync();
  });
  // Uriya: every file with several sheets asks which sheet is meant
  if (multi) {
    $("upload-sheets").insertAdjacentHTML("afterbegin", `<label class="sheet-pick">באיזה גיליון (Sheet) מדובר?
      <select id="upload-sheet"><option value="">— בחרו גיליון —</option>
      ${sheetNames.map((n) => `<option value="${esc(n)}">${esc(n)} (${upload.tables[n].rows.length} שורות)${n === hint ? " ← מתאים לכפתור הזה" : ""}</option>`).join("")}</select></label>`);
    $("upload-sheet").addEventListener("change", (e) => {
      document.querySelectorAll(".sheet-opt").forEach((el) => {
        const on = el.dataset.sheet === e.target.value;
        el.hidden = !on; el.querySelector(".s-on").checked = on; el.sync();
      });
    });
  }
  // Uriya: every bank / card upload asks which account the file is: bank account Z, card X or card Y.
  // The account whose number appears in the file is marked; the answer decides the table.
  if (topic?.kind) {
    upload.chosen = null;
    const accTopics = TOPICS.filter((tp) => tp.kind);
    const inFile = (tp) => { const a = accountOf(tp); return !!a.number && Object.values(upload.tables).some((t) => { const n = detectNumber(t, tp.kind); return n && sameNumber(tp.kind, a.number, n); }); };
    const html = `<div class="sheet-pick acct-pick">לאיזה חשבון / כרטיס שייך הקובץ?
      ${accTopics.map((tp) => { const a = accountOf(tp); return `<label class="opt"><input type="radio" name="upload-acct" value="${tp.id}">
        <span>${tp.kind === "bank" ? "🏦" : "💳"} <b>${esc(a.name || tp.table)}</b> · ${a.number ? esc(numberLabel(tp.kind, a.number)) : `<span class="muted">מספר עוד לא ידוע</span>`}${inFile(tp) ? ` <b class="found">✓ זה המספר שבקובץ</b>` : ""}</span></label>`; }).join("")}</div>`;
    const sheetQ = $("upload-sheet")?.closest("label");
    if (sheetQ) sheetQ.insertAdjacentHTML("afterend", html); else $("upload-sheets").insertAdjacentHTML("afterbegin", html);
    document.querySelectorAll('input[name="upload-acct"]').forEach((r) => r.addEventListener("change", () => {
      const tp = TOPICS.find((x) => x.id === r.value);
      upload.topic = upload.chosen = tp;
      $("upload-title").textContent = tp.label;
      const d = state.datasets.find((x) => x.name === tp.table);
      document.querySelectorAll(".sheet-opt").forEach((el) => { el.querySelector(".s-target").value = d ? d.id : "new"; el.querySelector(".s-name").value = tp.table; el.sync(); });
    }));
    document.querySelectorAll(".sheet-opt").forEach((el) => el.sync());
  }
  $("upload-go").disabled = false;
});

// Bank / card uploads: the number in the file must be the tile's number. Renders into .s-acc and returns
// null (not an account table) | {blocked} | {acc, ok} | {acc, save: number} (first time: confirm, then remember) | {acc} (no number found: confirm)
function verifyAccount(el, t) {
  const box = el.querySelector(".s-acc");
  const topic = upload.topic || topicOf(state.datasets.find((d) => d.id === +el.querySelector(".s-target").value));
  if (!topic?.kind) { box.innerHTML = ""; return null; }
  const kind = topic.kind, acc = accountOf(topic), name = esc(acc.name || topic.table);
  const what = kind === "card" ? "כרטיס" : "חשבון";
  const found = detectNumber(t, kind);
  if (found) {
    const owner = (state.accounts || []).find((a) => a.table !== topic.table && a.number && TOPICS.find((tp) => tp.table === a.table)?.kind === kind && sameNumber(kind, a.number, found));
    if (owner) { box.innerHTML = `<div class="blocked">⛔ בקובץ מופיע ${numberLabel(kind, found)}, וזה המספר של <b>${esc(owner.name)}</b>, לא של ${name}. טענו אותו בריבוע של ${esc(owner.name)}. הטעינה חסומה.</div>`; return { blocked: true }; }
    if (acc.number && sameNumber(kind, acc.number, found)) { box.innerHTML = `<div class="ok">✓ הקובץ שייך ל-${name} (${numberLabel(kind, found)})</div>`; return { acc, ok: true }; }
    if (acc.number) { box.innerHTML = `<div class="blocked">⛔ בקובץ מופיע ${numberLabel(kind, found)}, אבל ${name} הוא ${numberLabel(kind, acc.number)}. כנראה זה קובץ של ${what} אחר. הטעינה חסומה.<br><span class="muted">אם המספר השתנה באמת, עדכנו אותו בטבלה ${ACCOUNTS_TABLE}.</span></div>`; return { blocked: true }; }
    box.innerHTML = `<div class="ok">זוהה בקובץ <b>${numberLabel(kind, found)}</b>. המספר יישמר ל-${name}, ומעכשיו הבדיקה תהיה אוטומטית.</div>`;
    return { acc, save: found };
  }
  box.innerHTML = `<div class="warn">לא מצאתי בקובץ מספר ${what}${acc.number ? ` (מצופה: ${numberLabel(kind, acc.number)})` : ""}. הקובץ ייטען ל-${name} לפי הבחירה שלכם.</div>`;
  return { acc };
}

// Loading into a table that already has rows: compare date ranges and make the user choose what happens.
async function planSheet(el, i, t) {
  const target = el.querySelector(".s-target").value;
  const box = el.querySelector(".s-plan");
  el.querySelector(".s-name").hidden = target !== "new";
  el.plan = null;
  el.querySelector(".s-acc").innerHTML = "";
  if (!el.querySelector(".s-on").checked) { box.innerHTML = ""; el.plan = { skip: true }; return; } // only check sheets that will load
  if (upload.topic?.kind && !upload.chosen) { box.innerHTML = `<div class="warn">בחרו למעלה לאיזה חשבון / כרטיס שייך הקובץ.</div>`; el.plan = { needAcct: true }; return; }
  const ds = target === "new" ? null : state.datasets.find((d) => d.id === +target);
  // wrong place? (a subject button, or a subject's table picked by hand)
  const why = wrongPlace(t, upload.topic || topicOf(ds));
  if (why) { box.innerHTML = `<div class="blocked">⛔ ${why}<br>הטעינה חסומה.</div>`; el.plan = { blocked: true }; return; }
  // the bank account / card number in the file must be the one this tile is for
  const acc = verifyAccount(el, t);
  if (acc?.blocked) { box.innerHTML = ""; el.plan = { blocked: true }; return; }
  if (!ds) {
    const looks = TOPICS.find((tp) => t.fits.includes(tp.id));
    box.innerHTML = !upload.topic && looks ? `<div class="warn">נראה כמו קובץ של ${esc(looks.table)}. כדי שייכנס לטבלה הנכונה, עדיף להשתמש בכפתור "${esc(looks.label)}".</div>` : "";
    el.plan = { mode: "new", acc };
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
  const plan = { mode: null, ds, m, total: existing.length, acc };
  el.plan = plan;
  if (!existing.length) { plan.mode = "append"; box.innerHTML = `<div class="muted">הטבלה "${esc(ds.name)}" ריקה, השורות ייטענו אליה.</div>`; return; }

  // Uriya: card files come one month at a time and are a lot of work to reload -> never offer wiping a card table
  const noWipe = (upload.topic || topicOf(ds))?.kind === "card";
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
    ${noWipe ? `<div class="muted">🔒 בטבלת כרטיס אשראי אין אפשרות לדרוס את כל הטבלה: כל קובץ הוא חודש אחד, ורק החודש שלו מתעדכן.</div>`
      : `<label class="opt danger-opt"><input type="radio" name="mode${i}" value="all">
      <span><b>⚠ דריסת כל הטבלה</b>: מחיקת כל ${existing.length} השורות הקיימות (כולל מה שנוסף באתר) והשארת תוכן הקובץ בלבד. אין דרך חזרה.${m.cols.some((c) => c.manual) ? " (Category שנבחר ידנית יישמר רק לשורות שמופיעות שוב בקובץ.)" : ""}</span></label>`}
  </div>`;
  box.querySelectorAll("input[type=radio]").forEach((r) => r.addEventListener("change", () => { plan.mode = r.value; }));
}

$("upload-go").addEventListener("click", async () => {
  const jobs = [...document.querySelectorAll(".sheet-opt")].filter((el) => el.querySelector(".s-on").checked);
  if (!jobs.length) { $("upload-sheet")?.focus(); return toast("צריך לבחור באיזה גיליון (Sheet) מדובר"); }
  for (const el of jobs) {
    if (!el.plan) return toast("רגע, עדיין בודק את הטבלה הקיימת");
    if (el.plan.needAcct) { document.querySelector('input[name="upload-acct"]')?.focus(); return toast("צריך לבחור לאיזה חשבון / כרטיס שייך הקובץ", 4000); }
    if (el.plan.blocked) { el.scrollIntoView({ block: "center" }); return toast("הקובץ לא מתאים לטבלה הזאת, הטעינה חסומה", 4000); }
    if (!el.plan.mode) { el.scrollIntoView({ block: "center" }); return toast(`צריך לבחור מה לעשות עם הנתונים הקיימים בטבלה "${el.plan.ds.name}"`, 4000); }
    if (el.plan.mode === "all" && topicOf(el.plan.ds)?.kind === "card") return toast("בטבלת כרטיס אשראי אי אפשר לדרוס את כל הטבלה", 4000);
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
        if (plan.acc?.save) await saveAccountNumber(plan.acc.acc, plan.acc.save);
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
      if (plan.acc?.save) await saveAccountNumber(plan.acc.acc, plan.acc.save);
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
    await ensurePickers(state.ds);
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
      ${TOPICS.filter((tp) => !tp.kind).map((tp) => `<button class="btn primary" data-topic="${tp.id}">⬆ ${esc(tp.label)}</button>`).join("")}
      <button class="btn" id="btn-upload">⬆ טעינת קובץ אחר</button>
      ${state.datasets.length ? `<button class="btn" id="btn-dl-all">⬇ הורדת הכול לאקסל</button>` : ""}</div>
      <div class="acct-tiles">${TOPICS.filter((tp) => tp.kind).map((tp) => {
        const acc = accountOf(tp), d = state.datasets.find((x) => x.name === tp.table);
        return `<div class="card acct k-${tp.kind}">
          <div class="acct-kind">${tp.kind === "bank" ? "🏦 תנועות בנק" : "💳 כרטיס אשראי"}</div>
          <h3>${esc(acc.name || tp.table)}</h3>
          <div class="acct-num">${acc.number ? esc(numberLabel(tp.kind, acc.number)) : `<span class="muted">המספר יזוהה בטעינה הבאה</span>`}</div>
          <div class="muted">${d ? `${state.counts[d.id] ?? 0} שורות · עודכן ${fmtDate(d.updated_at)}` : "עדיין לא נטען"}</div>
          <div class="acct-btns"><button class="btn primary" data-topic="${tp.id}">⬆ טעינה</button>${d ? `<button class="btn" data-open="${d.id}">פתיחת הטבלה</button>` : ""}</div>
        </div>`;
      }).join("")}</div>
      ${state.datasets.length ? `<div class="grid-cards">${state.datasets.filter((d) => !TOPICS.some((tp) => tp.kind && tp.table === d.name)).map((d) => `
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
  if (td && !td.querySelector("input, select")) return startEdit(+td.parentElement.dataset.row, td.dataset.col); // a click inside the open dropdown is not a new edit
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
    const typed = prompt(`מחיקת הטבלה "${state.ds.name}" וכל ${state.rows.length} השורות שלה. כדי לאשר, הקלידו את שם הטבלה:`);
    if (typed === null) return; // cancelled
    if (typed.trim() !== state.ds.name) return toast(`השם שהוקלד ("${typed.trim()}") לא תואם לשם הטבלה, ולכן הטבלה לא נמחקה.`, 5000);
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
    await ensureAccounts();
    await loadAccounts();
    $("status").textContent = "";
    render();
    sb.from("usage_logs").insert({ visitor_name: user.email, action: "page_view" }).then(() => {});
  } catch (err) {
    console.error(err);
    $("status").textContent = "שגיאה בטעינת הנתונים, נסו לרענן";
  }
}
start();
