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
  { id: "cashflow", label: "טעינת cash flow", table: "Cash Flow" },
];
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
  const header = aoa[hi].map((v) => (v === null || v === undefined ? null : String(v).trim() || null)); // headers stay text
  const body = aoa.slice(hi + 1).map((r) => r.map(cellValue)).filter((r) => r.some((v) => v !== null));
  const width = Math.max(header.length, ...body.map((r) => r.length));
  const columns = [], seen = {};
  for (let i = 0; i < width; i++) {
    if (header[i] === null && body.every((r) => (r[i] ?? null) === null)) continue; // empty column with no title
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
    if (!hit) { hit = { key: nextKey(cols), label: c.label, type: c.type }; cols.push(hit); }
    map[c.key] = hit.key;
  }
  const rows = t.rows.map((r) => Object.fromEntries(Object.entries(r).map(([k, v]) => [map[k], v])));
  const fileDate = t.columns.find((c) => c.type === "date");
  return { cols, rows, dateKey: fileDate ? map[fileDate.key] : null };
}
// the existing table that most likely receives this sheet: same name, or the same column headers
function guessTarget(name, t) {
  const byName = state.datasets.find((d) => d.name === name);
  if (byName) return byName;
  const labels = t.columns.map((c) => c.label);
  return labels.length ? state.datasets.find((d) => labels.every((l) => d.columns.some((c) => c.label === l))) : null;
}

$("upload-file").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  $("upload-sheets").innerHTML = ""; $("upload-go").disabled = true; $("upload-progress").textContent = "";
  if (!file) return;
  try {
    const wb = await readWorkbook(file);
    upload = { file, topic: uploadTopic, tables: Object.fromEntries(wb.SheetNames.map((n) => [n, sheetToTable(wb.Sheets[n])])) };
  } catch (err) { console.error(err); toast("לא הצלחתי לקרוא את הקובץ"); return; }
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
  let first = true;
  document.querySelectorAll(".sheet-opt").forEach((el, i) => {
    const t = upload.tables[el.dataset.sheet];
    if (topic) {
      // subject upload: the (first non-empty) sheet goes into the subject's table, created on the first load
      const on = first && t.rows.length > 0; if (on) first = false;
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
    el.querySelector(".s-target").addEventListener("change", sync); sync();
  });
  $("upload-go").disabled = false;
});

// Loading into a table that already has rows: compare date ranges and make the user choose what happens.
async function planSheet(el, i, t) {
  const target = el.querySelector(".s-target").value;
  const box = el.querySelector(".s-plan");
  el.querySelector(".s-name").hidden = target !== "new";
  el.plan = null;
  if (target === "new") { box.innerHTML = ""; el.plan = { mode: "new" }; return; }
  const ds = state.datasets.find((d) => d.id === +target);
  const m = mapToTable(ds, t);
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
      <span><b>החלפת התקופה ${dmy(fileR.from)} – ${dmy(fileR.to)}</b> (מומלץ): ${plan.overlap} השורות האלה יוחלפו בשורות מהקובץ, והשורות מחוץ לתקופה יישארו. שימו לב: גם שורות שנוספו באתר בתקופה הזאת יוחלפו.</span></label>`;
  } else {
    info = `בטבלה "${esc(ds.name)}" יש כבר <b>${existing.length}</b> שורות${m.dateKey ? "" : " (אין עמודת תאריך להשוואה)"}.`;
  }
  box.innerHTML = `<div class="plan-box"><div>${info}</div><div class="muted">מה לעשות? (חובה לבחור)</div>
    ${periodOpt}
    <label class="opt"><input type="radio" name="mode${i}" value="append">
      <span><b>הוספה בלבד</b>: כל ${m.rows.length} שורות הקובץ יתווספו, בלי למחוק כלום${plan.overlap ? `. <span class="warn">שורות בתקופה החופפת יופיעו פעמיים.</span>` : "."}</span></label>
    <label class="opt danger-opt"><input type="radio" name="mode${i}" value="all">
      <span><b>⚠ דריסת כל הטבלה</b>: מחיקת כל ${existing.length} השורות הקיימות (כולל מה שנוסף באתר) והשארת תוכן הקובץ בלבד. אין דרך חזרה.</span></label>
  </div>`;
  box.querySelectorAll("input[type=radio]").forEach((r) => r.addEventListener("change", () => { plan.mode = r.value; }));
}

$("upload-go").addEventListener("click", async () => {
  const jobs = [...document.querySelectorAll(".sheet-opt")].filter((el) => el.querySelector(".s-on").checked);
  if (!jobs.length) return toast("לא נבחר אף גיליון");
  for (const el of jobs) {
    if (!el.plan) return toast("רגע, עדיין בודק את הטבלה הקיימת");
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
      const { ds, m } = plan;
      if (plan.mode === "all") {
        const { error } = await sb.from("dataset_rows").delete().eq("dataset_id", ds.id);
        if (error) throw error;
      } else if (plan.mode === "period") {
        $("upload-progress").textContent = `מוחק את השורות הקיימות בתקופה ${dmy(plan.range.from)} – ${dmy(plan.range.to)}…`;
        const { error } = await sb.from("dataset_rows").delete().eq("dataset_id", ds.id)
          .gte(`data->>${m.dateKey}`, plan.range.from).lt(`data->>${m.dateKey}`, plan.range.to + "~"); // "~" sorts after any time part
        if (error) throw error;
      }
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
  try { state.rows = await loadRows(id); $("status").textContent = ""; }
  catch (err) { console.error(err); $("status").textContent = "שגיאה בטעינת הטבלה"; state.rows = []; }
  render();
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
  td.innerHTML = `<input type="text" value="${esc(showValue(row.data[col], c.type))}" ${c.type === "number" ? 'dir="ltr"' : ""} placeholder="${c.type === "date" ? "dd/mm/yyyy" : ""}">`;
  const input = td.querySelector("input");
  input.focus(); input.select();
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
  state.tab = b.dataset.tab; render();
});
function render() {
  document.querySelectorAll(".tab").forEach((b) => b.classList.toggle("active", b.dataset.tab === state.tab));
  $("view-tables").hidden = state.tab !== "tables";
  $("view-users").hidden = state.tab !== "users";
  if (state.tab === "tables") renderTables(); else renderUsers();
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
    $("status").textContent = "";
    render();
    sb.from("usage_logs").insert({ visitor_name: user.email, action: "page_view" }).then(() => {});
  } catch (err) {
    console.error(err);
    $("status").textContent = "שגיאה בטעינת הנתונים, נסו לרענן";
  }
}
start();
