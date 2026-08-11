// Master physician roster — read/write via download-edit-upload with `exceljs`.
// We avoid Microsoft Graph's /workbook API because it goes through Office Online (WAC) and
// app-only credentials hit "Could not obtain a WAC access token" on this OneDrive. exceljs runs
// purely in-memory on Vercel's Node runtime — no Office services involved.
const ExcelJS = require("exceljs");
const { accessToken, encPath, GRAPH } = require("./graph");

// The master roster moved during the 2026-06 OneDrive reorg: the old "Compliance" folder was
// archived to "..ZCompliance" and the live master now lives under "..WCGTX Master Rosters".
// Env-overridable so a future move is a Vercel env change, not a code deploy.
const ROSTER_PATH = process.env.MS_ROSTER_PATH ||
  "WCGTX Phyicians_04.08.2020/..WCGTX Master Rosters/WCGTX Physician Roster.xlsx";
const SHEET_ACTIVE = "WCGTX Credentials";
const SHEET_INACTIVE = "Inactive Providers";
const DRIVE_ID = process.env.MS_DRIVE_ID || "b!hICmGNzaFEiC8Z6vebrpNWzB937MR0tFsLlTxA2x3Z9-nxsW_blJTrLUhaL3IsBm";

function rosterRoot() { return GRAPH + "/drives/" + DRIVE_ID + "/root:/" + encPath(ROSTER_PATH); }

// Dashboard credential category -> 1-based column in the "WCGTX Credentials" sheet.
// ONLY columns that hold an expiry DATE are listed. Deliberately excluded: Board Certified
// (holds a board name), and the Yes/No documentation columns (Initial App, CV, Diploma, ECFMG,
// Residency, Peer Refs, Photo, DOP, VOE, COI, Odessa) — auto-filling those would corrupt them.
const CRED_COLUMNS = {
  "ACLS Certification": 3,
  "ATLS Certification": 4,
  "PALS Certification": 5,
  "BLS Certification": 6,
  "State Medical License": 8,
  "Medical License Verify (annual)": 9,
  "Individual DEA Registration": 10,
  "DEA Verify (annual)": 11,
  "CME (20 hrs / 2 yrs)": 12,
  "Influenza Vaccination": 13,
  "TB Screening": 14,
  "Driver's License": 15,
  "NPDB Query (2 yrs)": 16,
  "OIG / SAM Exclusion Check": 17,
  "TSCA Documents": 19,
};
// Pull plain text out of any exceljs cell value (string, number, {text}, {richText:[{text}]}, {result}, Date, etc.)
function cellText(v) {
  if (v == null) return "";
  if (typeof v === "string") return v;
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === "object") {
    if (Array.isArray(v.richText)) return v.richText.map(r => r && r.text ? r.text : "").join("");
    if (typeof v.text === "string") return v.text;
    if (typeof v.result === "string" || typeof v.result === "number") return String(v.result);
    if (typeof v.formula === "string") return "";
  }
  return String(v);
}
function normName(s) { return cellText(s).replace(/[\*,()]+/g, "").replace(/\s+/g, " ").trim().toLowerCase(); }

// ---- Header-based column detection ---------------------------------------------------------
// The master roster occasionally gains a leading tracking column (e.g. a "THHS Email Sent"
// status column), which shifts EVERY column to the right and silently breaks the old fixed
// "col A = last, col B = first, credential dates at columns 3..19" assumptions. Detecting the
// columns from the header row (row 1) is robust to that: the schema follows the labels, not the
// position. All detectors take a header row as an array of cell values and return 0-based indexes.
function _normHdr(h) { return cellText(h).replace(/\s+/g, " ").trim().toLowerCase(); }
function detectNameCols(headerArr) {
  let lastIdx = -1, firstIdx = -1;
  (headerArr || []).forEach((h, i) => {
    const n = _normHdr(h);
    if (lastIdx < 0 && n === "last name") lastIdx = i;
    if (firstIdx < 0 && n === "first name") firstIdx = i;
  });
  if (lastIdx < 0) lastIdx = 0;                       // fall back to the historical layout
  // The first name sits immediately right of the last name in every version of this sheet.
  // The old fallback used column 0 whenever lastIdx wasn't 0, which on a shifted roster read the
  // first name out of the leading tracking column (and wrote it back there too).
  if (firstIdx < 0) firstIdx = lastIdx + 1;
  return { lastIdx, firstIdx };
}
// Map a header label to its dashboard credential category — ONLY for columns that hold an
// expiry DATE. Returns null for Yes/No columns (Board Certified, NPI Verify, DOP, Peer Refs,
// Initial App, etc.) so we never mistake those for dates.
function credCatForHeader(n) {
  if (!n) return null;
  if (n.includes("verif")) {                          // disambiguate the two "verify" columns
    if (n.includes("dea")) return "DEA Verify (annual)";
    if (n.includes("lic")) return "Medical License Verify (annual)";
    return null;                                       // e.g. "NPI Verify" is Yes/No, not a date
  }
  if (n.startsWith("acls")) return "ACLS Certification";
  if (n.startsWith("atls")) return "ATLS Certification";
  if (n.startsWith("pals")) return "PALS Certification";
  if (n.startsWith("bls")) return "BLS Certification";
  if (n.startsWith("med lic") || n.startsWith("med license")) return "State Medical License";
  if (n.startsWith("dea")) return "Individual DEA Registration";
  if (n.startsWith("cme")) return "CME (20 hrs / 2 yrs)";
  if (n.startsWith("flu")) return "Influenza Vaccination";
  if (n.startsWith("tb")) return "TB Screening";
  if (n.includes("driver")) return "Driver's License";
  if (n.startsWith("npdb")) return "NPDB Query (2 yrs)";
  if (n.startsWith("oig")) return "OIG / SAM Exclusion Check";
  if (n.startsWith("tsca")) return "TSCA Documents";
  return null;
}
function detectCredCols(headerArr) {
  const map = {};
  (headerArr || []).forEach((h, i) => { const cat = credCatForHeader(_normHdr(h)); if (cat && !(cat in map)) map[cat] = i; });
  return map;
}
// Worksheet variant: read row 1 and return 1-based { lastCol, firstCol } for getCell() callers.
function detectNameColsWs(ws) {
  const arr = [];
  const hr = ws.getRow(1);
  for (let c = 1; c <= ws.columnCount; c++) arr.push(hr.getCell(c).value);
  const { lastIdx, firstIdx } = detectNameCols(arr);
  return { lastCol: lastIdx + 1, firstCol: firstIdx + 1 };
}

async function downloadWorkbook(token) {
  const r = await fetch(rosterRoot() + ":/content", { headers: { Authorization: "Bearer " + token } });
  if (!r.ok) throw new Error("download roster: " + r.status + " " + (await r.text()).slice(0, 200));
  const buf = Buffer.from(await r.arrayBuffer());
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  return wb;
}

async function uploadWorkbook(token, wb) {
  const buf = await wb.xlsx.writeBuffer();
  // Files this size (~few MB) upload fine in a single PUT. If the file is locked (HTTP 423 —
  // someone has it open in Excel desktop/online, or an autosave is mid-flight), retry a few
  // times; brief autosave locks usually clear within a couple of seconds.
  let detail = "";
  for (let attempt = 0; attempt < 3; attempt++) {
    if (attempt) await new Promise(r => setTimeout(r, 1800));
    const r = await fetch(rosterRoot() + ":/content", {
      method: "PUT",
      headers: {
        Authorization: "Bearer " + token,
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      },
      body: buf
    });
    if (r.ok) return;
    detail = (await r.text()).slice(0, 200);
    if (r.status === 423 || /resourceLocked|is locked/i.test(detail)) continue;   // transient lock — retry
    throw new Error("upload roster: " + r.status + " " + detail);
  }
  // Still locked after retries — almost always a human has the file open in Excel.
  throw new Error("ROSTER_LOCKED: The master roster (WCGTX Physician Roster.xlsx) is open in Excel right now, so the app can't save to it. Close the file everywhere — the Excel desktop app AND any browser tab — then try again.");
}

function getSheet(wb, name) {
  const ws = wb.getWorksheet(name);
  if (!ws) throw new Error("sheet not found: " + name);
  return ws;
}

// readSheet returns { values, rowCount, columnCount } — same shape as the old workbook-API caller expected.
async function readSheet(token, sheetName) {
  const wb = await downloadWorkbook(token);
  const ws = getSheet(wb, sheetName);
  const values = [];
  let maxCol = 0;
  ws.eachRow({ includeEmpty: true }, (row) => {
    const arr = [];
    for (let c = 1; c <= ws.columnCount; c++) {
      const v = row.getCell(c).value;
      arr.push(v == null ? null : (typeof v === "object" && v.text ? v.text : v));
    }
    if (arr.length > maxCol) maxCol = arr.length;
    values.push(arr);
  });
  return { values, rowCount: values.length, columnCount: maxCol };
}

function _findRowIndex(ws, lastName, firstName) {
  const tLast = normName(lastName), tFirst = normName(firstName);
  const { lastCol, firstCol } = detectNameColsWs(ws);   // header-based; survives a shifted layout
  // 1st pass: exact last-name + first-name (or prefix).
  for (let r = 2; r <= ws.rowCount; r++) {
    const last = normName(ws.getRow(r).getCell(lastCol).value);
    const first = normName(ws.getRow(r).getCell(firstCol).value);
    if (last && last === tLast && (!tFirst || first === tFirst || first.startsWith(tFirst) || tFirst.startsWith(first))) return r;
  }
  // 2nd pass: maybe columns are swapped in this row (some entries are typed "First, Last").
  for (let r = 2; r <= ws.rowCount; r++) {
    const a = normName(ws.getRow(r).getCell(lastCol).value);
    const b = normName(ws.getRow(r).getCell(firstCol).value);
    if (a === tFirst && b === tLast) return r;
  }
  // 3rd pass: last-name only when first name didn't match — handles roster typos / suffixes.
  if (tLast) {
    for (let r = 2; r <= ws.rowCount; r++) {
      const last = normName(ws.getRow(r).getCell(lastCol).value);
      if (last === tLast) return r;
    }
  }
  return null;
}

async function findRow(token, sheetName, lastName, firstName) {
  const wb = await downloadWorkbook(token);
  const ws = getSheet(wb, sheetName);
  const idx = _findRowIndex(ws, lastName, firstName);
  if (!idx) return null;
  const arr = [];
  for (let c = 1; c <= ws.columnCount; c++) arr.push(ws.getRow(idx).getCell(c).value);
  return { rowIndex: idx, values: arr };
}

async function findAnywhere(token, lastName, firstName) {
  const wb = await downloadWorkbook(token);
  const ai = _findRowIndex(getSheet(wb, SHEET_ACTIVE), lastName, firstName);
  if (ai) return { sheet: SHEET_ACTIVE, rowIndex: ai };
  const ii = _findRowIndex(getSheet(wb, SHEET_INACTIVE), lastName, firstName);
  if (ii) return { sheet: SHEET_INACTIVE, rowIndex: ii };
  return null;
}

// Slug a (last, first) pair the same way generate_data.py and the JS dashboard do, so we
// can find an Excel row by the dashboard's entityKey without guessing how to split a display name.
function slugKey(last, first) {
  return (String(last || "") + "-" + String(first || ""))
    .replace(/[\*,()]+/g, "")
    .replace(/[^A-Za-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .toLowerCase();
}
async function findRowByEntityKey(token, sheetName, entityKey) {
  const wb = await downloadWorkbook(token);
  const ws = getSheet(wb, sheetName);
  const { lastCol, firstCol } = detectNameColsWs(ws);
  for (let r = 2; r <= ws.rowCount; r++) {
    const last = cellText(ws.getRow(r).getCell(lastCol).value).replace(/[\*,()]+/g, "").trim();
    const first = cellText(ws.getRow(r).getCell(firstCol).value).trim();
    if (!last) continue;
    if (slugKey(last, first) === entityKey) return { rowIndex: r, last, first };
  }
  return null;
}
// Find a row by the provider's DISPLAY name ("First Last", i.e. how the SharePoint folder is
// named) without guessing where the surname starts. "Maria De La Cruz" cannot be split by
// assuming the last token is the surname — that produced last="Cruz", first="Maria De La", which
// matched no existing row and caused a duplicate to be appended. Comparing the assembled
// "first last" against every row needs no guess at all.
async function findRowByFullName(token, fullName) {
  const target = normName(fullName);
  if (!target) return null;
  const wb = await downloadWorkbook(token);
  for (const sheetName of [SHEET_ACTIVE, SHEET_INACTIVE]) {
    const ws = getSheet(wb, sheetName);
    const { lastCol, firstCol } = detectNameColsWs(ws);
    for (let r = 2; r <= ws.rowCount; r++) {
      const last = cellText(ws.getRow(r).getCell(lastCol).value).replace(/[\*,()]+/g, "").trim();
      const first = cellText(ws.getRow(r).getCell(firstCol).value).trim();
      if (!last) continue;
      if (normName((first ? first + " " : "") + last) === target) return { sheet: sheetName, rowIndex: r, last, first };
    }
  }
  return null;
}

async function findByEntityKeyAnywhere(token, entityKey) {
  const a = await findRowByEntityKey(token, SHEET_ACTIVE, entityKey);
  if (a) return { sheet: SHEET_ACTIVE, ...a };
  const i = await findRowByEntityKey(token, SHEET_INACTIVE, entityKey);
  if (i) return { sheet: SHEET_INACTIVE, ...i };
  return null;
}

async function appendRow(token, sheetName, rowValues) {
  const wb = await downloadWorkbook(token);
  const ws = getSheet(wb, sheetName);
  const newRow = ws.addRow(rowValues);
  await uploadWorkbook(token, wb);
  return { rowIndex: newRow.number };
}

// Append a provider placing last/first into the sheet's ACTUAL name columns (detected from the
// header), so a shifted roster layout doesn't drop the name into a tracking column.
async function appendProviderRow(token, sheetName, last, first) {
  const wb = await downloadWorkbook(token);
  const ws = getSheet(wb, sheetName);
  const { lastCol, firstCol } = detectNameColsWs(ws);
  // Build a DENSE array (explicit nulls, no holes). A sparse array like [ , "Smith", "Jane"]
  // gets compacted by exceljs's addRow, which silently shifts the values one column LEFT — on a
  // roster with a leading tracking column that wrote the surname into that tracking column and
  // left First Name empty, i.e. exactly the corruption this function exists to prevent.
  // Verified against the real workbook, shifted and unshifted.
  const width = Math.max(ws.columnCount || 0, lastCol, firstCol);
  const arr = new Array(width).fill(null);
  arr[lastCol - 1] = last;
  arr[firstCol - 1] = first;
  const newRow = ws.addRow(arr);
  // Belt-and-braces: set the cells directly too, so the row is correct even if addRow's
  // array handling ever changes again.
  newRow.getCell(lastCol).value = last;
  newRow.getCell(firstCol).value = first;
  await uploadWorkbook(token, wb);
  return { rowIndex: newRow.number };
}

async function deleteRow(token, sheetName, rowIndex) {
  const wb = await downloadWorkbook(token);
  const ws = getSheet(wb, sheetName);
  ws.spliceRows(rowIndex, 1);
  await uploadWorkbook(token, wb);
}

// Move a row between sheets in ONE download-upload cycle (atomic).
async function moveRow(token, fromSheet, fromRowIndex, toSheet) {
  const wb = await downloadWorkbook(token);
  const src = getSheet(wb, fromSheet);
  const dst = getSheet(wb, toSheet);
  const row = src.getRow(fromRowIndex);
  const vals = [];
  for (let c = 1; c <= src.columnCount; c++) vals.push(row.getCell(c).value);
  dst.addRow(vals);
  src.spliceRows(fromRowIndex, 1);
  await uploadWorkbook(token, wb);
}

// HARD DELETE: remove rows matching entityKey (or last/first) from BOTH Credentials and
// Inactive sheets in one download-upload cycle. Returns the deleted rows' values for
// logging into the recycle bin so they can be restored.
async function hardDelete(token, entityKey, lastName, firstName) {
  const wb = await downloadWorkbook(token);
  const removed = [];
  for (const sheetName of [SHEET_ACTIVE, SHEET_INACTIVE]) {
    const ws = getSheet(wb, sheetName);
    const { lastCol, firstCol } = detectNameColsWs(ws);
    // Read every row's identity once, so matching logic below is explicit about what it targets.
    const rows = [];
    for (let r = 2; r <= ws.rowCount; r++) {
      const last = cellText(ws.getRow(r).getCell(lastCol).value).replace(/[\*,()]+/g, "").trim();
      const first = cellText(ws.getRow(r).getCell(firstCol).value).trim();
      if (!last) continue;
      const ek = (last + "-" + (first || "")).replace(/[\*,()]+/g, "").replace(/[^A-Za-z0-9]+/g, "-").replace(/^-+|-+$/g, "").toLowerCase();
      rows.push({ r, last, first, ek });
    }
    // An entityKey is AUTHORITATIVE: match on it ALONE. The old code OR'd in a name match even
    // when a key was supplied, so deleting one "Patel" could delete every Patel on the roster.
    // With no key, fall back to name — but require the first name to disambiguate.
    let matched;
    if (entityKey) {
      matched = rows.filter(x => x.ek === entityKey);
    } else if (lastName) {
      const tLast = normName(lastName), tFirst = normName(firstName);
      matched = rows.filter(x => normName(x.last) === tLast && (tFirst ? normName(x.first) === tFirst : true));
      // Refuse an ambiguous surname-only delete rather than removing several people at once.
      if (matched.length > 1 && !tFirst) {
        throw new Error("AMBIGUOUS_DELETE: " + matched.length + " providers share the last name \"" +
          lastName + "\" (" + matched.map(x => (x.first || "?") + " " + x.last).join(", ") +
          "). Nothing was deleted — open the specific provider and delete from their page so the exact record is targeted.");
      }
    } else matched = [];
    matched.sort((a, b) => b.r - a.r);   // descending so spliceRows doesn't shift indices we need
    for (const m of matched) {
      const row = ws.getRow(m.r);
      const vals = [];
      for (let c = 1; c <= ws.columnCount; c++) vals.push(cellText(row.getCell(c).value));
      // Carry the header-detected name out with the row so callers don't have to guess which
      // array position holds it — that guess is what left orphaned SharePoint folders behind.
      removed.push({ sheet: sheetName, rowIndex: m.r, values: vals, last: m.last, first: m.first });
      ws.spliceRows(m.r, 1);
    }
  }
  if (removed.length) await uploadWorkbook(token, wb);
  return removed;
}

// Restore a single trashed row back to the Credentials sheet.
async function restoreRow(token, sheetName, values) {
  const wb = await downloadWorkbook(token);
  const ws = getSheet(wb, sheetName || SHEET_ACTIVE);
  ws.addRow(values);
  await uploadWorkbook(token, wb);
}

// Backup the current workbook to _Sentinel/roster_backups/.
// Returns { ok, method, name } on success or { ok:false, error } on failure so callers
// can surface the actual reason instead of a silent false.
async function snapshotWorkbook(token, label) {
  const ts = new Date().toISOString().replace(/[:T.]/g, "-").slice(0, 19);
  const name = "Roster_" + ts + "_" + (label || "auto") + ".xlsx";
  try {
    const { ensureFolder } = require("./graph");
    // Graph's content-PUT doesn't reliably create intermediate folders on this drive;
    // create the backups folder explicitly so the first ever backup works.
    await ensureFolder(token, "_Sentinel/roster_backups");

    // Prefer server-side copy (no download/upload bytes, works at any file size).
    const copyR = await fetch(rosterRoot() + ":/copy", {
      method: "POST",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify({
        parentReference: { driveId: DRIVE_ID, path: "/drives/" + DRIVE_ID + "/root:/_Sentinel/roster_backups" },
        name
      })
    });
    if (copyR.ok || copyR.status === 202) return { ok: true, method: "copy", name };
    const copyText = await copyR.text();

    // Copy failed — fall back to download/upload PUT (works for files < ~4MB).
    const dl = await fetch(rosterRoot() + ":/content", { headers: { Authorization: "Bearer " + token } });
    if (!dl.ok) return { ok: false, error: "copy " + copyR.status + " then download " + dl.status, detail: copyText.slice(0, 160) };
    const buf = Buffer.from(await dl.arrayBuffer());
    const path = "_Sentinel/roster_backups/" + name;
    const up = await fetch(GRAPH + "/drives/" + DRIVE_ID + "/root:/" + encPath(path) + ":/content", {
      method: "PUT",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
      body: buf
    });
    if (!up.ok) {
      const upText = await up.text();
      return { ok: false, error: "copy " + copyR.status + ", put " + up.status, detail: upText.slice(0, 160), sizeBytes: buf.length };
    }
    return { ok: true, method: "put", name, sizeBytes: buf.length };
  } catch (e) { return { ok: false, error: String(e.message || e).slice(0, 200), name }; }
}

// Compute the dashboard entityKey for an Excel row (col A = last, col B = first).
function _ekOfRow(ws, r) {
  const { lastCol, firstCol } = detectNameColsWs(ws);
  const last = cellText(ws.getRow(r).getCell(lastCol).value).replace(/[\*,()]+/g, "").trim();
  const first = cellText(ws.getRow(r).getCell(firstCol).value).trim();
  if (!last) return null;
  return (last + "-" + (first || "")).replace(/[\*,()]+/g, "").replace(/[^A-Za-z0-9]+/g, "-").replace(/^-+|-+$/g, "").toLowerCase();
}

// Fill expiry dates into the Credentials sheet, ONLY into cells that are currently empty
// (never overwrites a human-entered value). updates: [{ entityKey, category, date:"YYYY-MM-DD" }].
// Returns { written:[...], skipped:[...] } so callers can report exactly what happened.
async function fillEmptyDates(token, updates) {
  const written = [], skipped = [];
  const jobs = [];
  for (const u of (updates || [])) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(u.date || ""));
    // Validate the date + key up front; the actual column is resolved from the live header below.
    if (!u.category || !m || !u.entityKey) { skipped.push({ category: u.category, reason: (!m ? "bad-date" : "no-entityKey") }); continue; }
    jobs.push({ entityKey: String(u.entityKey).toLowerCase(), y: +m[1], mo: +m[2], d: +m[3], category: u.category, date: u.date });
  }
  if (!jobs.length) return { written, skipped };
  const wb = await downloadWorkbook(token);
  const ws = getSheet(wb, SHEET_ACTIVE);
  // Resolve each credential category to its column FROM THE LIVE HEADER (1-based), so a shifted
  // roster layout never causes a date to be written into the wrong column.
  const hdr = []; { const hr = ws.getRow(1); for (let c = 1; c <= ws.columnCount; c++) hdr.push(hr.getCell(c).value); }
  const credIdx = detectCredCols(hdr);   // { category: 0-based idx }
  // NO fallback to the old fixed CRED_COLUMNS map. If a header was renamed so we can't identify
  // its column, the honest answer is "skip this one" — falling back to a hardcoded position on a
  // shifted sheet wrote dates into whatever column happened to sit there (a flu date landing in
  // CME, a TSCA date landing in the Yes/No "NPI Verify" column), silently and unrecoverably.
  const colOfCat = (cat) => (cat in credIdx) ? credIdx[cat] + 1 : null;
  const rowOf = {};
  for (let r = 2; r <= ws.rowCount; r++) { const ek = _ekOfRow(ws, r); if (ek && !(ek in rowOf)) rowOf[ek] = r; }
  let dirty = 0;
  for (const j of jobs) {
    const col = colOfCat(j.category);
    if (!col) { skipped.push({ entityKey: j.entityKey, category: j.category, reason: "not-a-date-column" }); continue; }
    const r = rowOf[j.entityKey];
    if (!r) { skipped.push({ entityKey: j.entityKey, category: j.category, reason: "provider-not-in-active-roster" }); continue; }
    const cell = ws.getRow(r).getCell(col);
    const cur = cell.value == null ? "" : cellText(cell.value).trim();
    if (cur !== "") { skipped.push({ entityKey: j.entityKey, category: j.category, reason: "cell-already-filled" }); continue; }
    cell.value = new Date(j.y, j.mo - 1, j.d);
    cell.numFmt = "m/d/yyyy";
    written.push({ entityKey: j.entityKey, category: j.category, col, date: j.date });
    dirty++;
  }
  if (dirty) {
    try { await snapshotWorkbook(token, "before-autofill"); } catch (e) {}
    await uploadWorkbook(token, wb);
  }
  return { written, skipped };
}

// Parse an expiry date out of a Credentials cell. Cells are a mix of real Date objects,
// "M/D/YYYY" strings (sometimes with trailing text like "4/2/2026 Rqstd"), named-month
// strings ("Fri Sep 18 2026..."), and non-dates ("Requested", "Yes", "N/A", ""). Returns
// "YYYY-MM-DD" or null.
function cellDate(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date) return isNaN(v) ? null : v.toISOString().slice(0, 10);
  const s = String(v).trim();
  if (!s) return null;
  let m = s.match(/\b(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})\b/);
  if (m) { let y = +m[3]; if (y < 100) y += 2000; const mo = +m[1], d = +m[2]; if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) return y + "-" + String(mo).padStart(2, "0") + "-" + String(d).padStart(2, "0"); }
  m = s.match(/\b(\d{4})-(\d{2})-(\d{2})\b/);
  if (m) return m[0];
  const d2 = new Date(s);
  if (!isNaN(d2) && d2.getFullYear() > 1990 && d2.getFullYear() < 2100) return d2.toISOString().slice(0, 10);
  return null;
}

// Read expiry dates straight out of the Credentials sheet values (from readSheet), so the
// dashboard reflects dates a human types directly into Excel. Returns [{entityKey,category,expires}].
function expiryDatesFromValues(values) {
  const out = [];
  const header = (values && values[0]) || [];
  const { lastIdx, firstIdx } = detectNameCols(header);       // 0-based, header-driven
  const credIdx = detectCredCols(header);                     // { category: 0-based idx }
  const cols = Object.keys(credIdx).map(cat => [cat, credIdx[cat]]);
  (values || []).slice(1).forEach(r => {
    const last = String(r[lastIdx] || "").replace(/[\*,()]+/g, "").trim();
    const first = String(r[firstIdx] || "").trim();
    if (!last) return;
    const ek = (last + "-" + (first || "")).replace(/[\*,()]+/g, "").replace(/[^A-Za-z0-9]+/g, "-").replace(/^-+|-+$/g, "").toLowerCase();
    for (const [cat, idx] of cols) { const dt = cellDate(r[idx]); if (dt) out.push({ entityKey: ek, category: cat, expires: dt }); }
  });
  return out;
}

// ---- Live staff sync: read the CHER/Frisco RN+FD credentialing workbooks server-side and
//      return scope:"staff" items (credentials + training), same shape as the local merge. ----
function _slug() { return [].join.call(arguments, "-").replace(/[^a-zA-Z0-9]+/g, "-").replace(/^-+|-+$/g, "").toLowerCase(); }
const _SR = "https://wcgtx.sharepoint.com/sites/CorporateArchivesDirectory/Shared%20Documents/Sama%20Farooqui/Sentinel/State%20Readiness";
const STAFF_FILES = [
  { file: "WCGTX Phyicians_04.08.2020/..WCGTX Master Rosters/CHER RN and FD Roster and Credentialing log.xlsx",
    facility: "Castle Hills ER", folder: _SR + "/Castle%20Hills",
    sheets: { "CHER RN": ["RN", true, "cred"], "CHER FD": ["Front Desk", true, "cred"], "Inactive RN": ["RN", false, "cred"], "InactiveFD": ["Front Desk", false, "cred"], "Training RN": ["RN", true, "training"], "Training FD": ["Front Desk", true, "training"] } },
  { file: "WCGTX Phyicians_04.08.2020/..WCGTX Master Rosters/Frisco ER RN and FD Roster and Credentialing.xlsx",
    facility: "Frisco ER", folder: _SR + "/Frisco",
    sheets: { "FriscoER RN": ["RN", true, "cred"], "FriscoER FD": ["Front Desk", true, "cred"], "Inactive RNs": ["RN", false, "cred"], "Inactive FD": ["Front Desk", false, "cred"], "Training RN": ["RN", true, "training"], "Training FD": ["Front Desk", true, "training"] } },
];
function _normCat(h) {
  const s = String(h || "").replace(/\s+/g, " ").trim(), l = s.toLowerCase();
  if (l.startsWith("nursys")) return "Nursys Verification";
  if (l.indexOf("nursing") >= 0 && (l.indexOf("diploma") >= 0 || l.indexOf("degree") >= 0)) return "Nursing Diploma";
  if (l.indexOf("driver") >= 0) return "Driver's License";
  if (l.indexOf("references") === 0) return "Peer References";
  if (l.indexOf("social security") >= 0) return "Social Security Card";
  return s;
}
function _parseStaffCell(v) {
  if (v == null || v === "") return ["skip", ""];
  const dt = cellDate(v); if (dt) return ["date", dt];
  const s = cellText(v).trim();
  if (!s || /^(n\/?a|no|none|x|-|\.)$/i.test(s)) return ["skip", ""];
  if (/^(yes|y|true|done|complete|completed|on file|received|signed)$/i.test(s)) return ["onfile", "On file"];
  if (/request|rqst|need|sent|pending|waiting|reminder|assigned/i.test(s)) return ["pending", s];
  return ["onfile", s];
}
async function downloadWorkbookAt(token, path) {
  const r = await fetch(GRAPH + "/drives/" + DRIVE_ID + "/root:/" + encPath(path) + ":/content", { headers: { Authorization: "Bearer " + token } });
  if (!r.ok) return null;
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(Buffer.from(await r.arrayBuffer())); return wb;
}
async function buildStaffItems(token) {
  const out = [];
  for (const F of STAFF_FILES) {
    const wb = await downloadWorkbookAt(token, F.file); if (!wb) continue;
    for (const sheet in F.sheets) {
      const spec = F.sheets[sheet], role = spec[0], active = spec[1], kind = spec[2];
      const ws = wb.getWorksheet(sheet); if (!ws) continue;
      const colCount = ws.columnCount;   // O(rows) getter in exceljs — read it once, not per row
      const header = []; const hr = ws.getRow(1);
      for (let c = 1; c <= colCount; c++) header[c] = cellText(hr.getCell(c).value).trim();
      // Find the name columns by HEADER, same as the physician roster. These workbooks are
      // maintained by hand too, so they can gain a leading tracking column exactly like the
      // physician roster did — and staff items have no corruption guard downstream, so a
      // misparse here would wholesale replace good staff data on the board.
      const hdr0 = header.slice(1);                       // 0-based view for the detector
      const nameCols = detectNameCols(hdr0);
      const lastCol = nameCols.lastIdx + 1, firstCol = nameCols.firstIdx + 1;
      const cols = [];
      // Credential columns are everything to the RIGHT of the name pair (was hardcoded to 3).
      for (let c = Math.max(lastCol, firstCol) + 1; c <= colCount; c++) {
        const h = header[c]; if (!h || h.toLowerCase() === "status") continue;
        cols.push([c, kind === "training" ? "Training: " + h.replace(/\s*\(.*$/, "").trim() : _normCat(h)]);
      }
      for (let r = 2; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const last = cellText(row.getCell(lastCol).value).replace(/\(.*?\)/g, "").trim(); if (!last) continue;
        const first = cellText(row.getCell(firstCol).value).trim();
        const name = (first + " " + last).trim();
        const ek = _slug(F.facility.split(" ")[0], role.replace(/ /g, ""), last, first);
        for (const cc of cols) {
          const parsed = _parseStaffCell(row.getCell(cc[0]).value); if (parsed[0] === "skip") continue;
          const item = { id: _slug(ek, cc[1]), scope: "staff", entity: name, entityKey: ek, category: cc[1], authority: "", role: role, facility: F.facility, staffFacility: F.facility, renewalLeadDays: kind === "training" ? 30 : 60, training: kind === "training", owner: name, fileLink: F.folder, folderLink: F.folder, isFile: false, centralProof: true, active: active, number: "", issued: null };
          if (parsed[0] === "date") item.expires = parsed[1];
          else if (parsed[0] === "onfile") { item.expires = null; item.permanent = true; item.docStatus = "On file"; if (parsed[1] !== "On file") item.notes = parsed[1]; }
          else if (parsed[0] === "pending") { item.expires = null; item.pending = true; item.notes = parsed[1]; }
          out.push(item);
        }
      }
    }
  }
  return out;
}

module.exports = {
  ROSTER_PATH, SHEET_ACTIVE, SHEET_INACTIVE, CRED_COLUMNS,
  readSheet, appendRow, appendProviderRow, deleteRow, moveRow, findRow, findAnywhere,
  findRowByEntityKey, findByEntityKeyAnywhere, findRowByFullName,
  hardDelete, restoreRow, snapshotWorkbook, fillEmptyDates,
  cellDate, expiryDatesFromValues, buildStaffItems,
  detectNameCols, detectCredCols, detectNameColsWs
};
