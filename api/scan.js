// Auto-detect documents added ANYWHERE under the Sentinel library — via the dashboard QR
// uploader, OneDrive, or directly in SharePoint. Graph "delta" finds new/changed files and
// matches them to items. For matched files with no date in the filename, it then reads the
// expiry by OCR (OCR.space, free) SERVER-SIDE — so dates fill in automatically, no opening.
// Results -> _Sentinel/auto_detected.json, which /api/uploads-map merges into the dashboard.
const { accessToken, docsRoot, docsPathFromUrl, encPath, drivePath, readJsonAt, writeJsonAt, dateFromName } = require("../lib/graph");
const { applyRosterDelta } = require("../lib/delta");
// Re-read data.json fresh on each invocation (don't cache via require — warm lambdas would
// keep a stale index, missing newly added providers/items).
const fs = require("fs");
const path = require("path");
const DATA_PATH = path.join(__dirname, "..", "data.json");
let _dataMtime = 0, _data = null;
function getData() {
  try {
    const st = fs.statSync(DATA_PATH);
    if (st.mtimeMs !== _dataMtime) { _data = JSON.parse(fs.readFileSync(DATA_PATH, "utf8")); _dataMtime = st.mtimeMs; INDEX = null; }
  } catch (e) { if (!_data) _data = { items: [] }; }
  return _data;
}

const STATE = drivePath("_Sentinel/scan_state.json");
const DETECTED = drivePath("_Sentinel/auto_detected.json");
const SUPP = drivePath("_Sentinel/supplemental_detected.json");   // new-file supplemental records

const SOP_PHASES = [
  "1. Application & Document Collection", "2. Primary Source Verification",
  "3. Background & Compliance Review", "4. Medical Staff Review",
  "5. Payer Enrollment & Facility Setup", "6. Approval & Ongoing Monitoring",
];
const STATE_SECTIONS = [
  "01. Licensing & Regulatory Compliance", "02. Personnel Files & Credentialing",
  "03. Medical Staff Services", "04. Patient Care & Clinical Documentation",
  "05. Medication Management", "06. Crash Cart & Emergency Equipment",
  "07. Infection Prevention & Control", "08. Laboratory Services",
  "09. Radiology Services", "10. Quality Improvement Program",
  "11. Environment of Care", "12. Emergency Preparedness",
  "13. Patient Rights & Compliance", "14. Daily Readiness Walkthrough",
];

function slug() {
  return Array.from(arguments).join("-").replace(/[^A-Za-z0-9]+/g, "-").replace(/^-+|-+$/g, "").toLowerCase().slice(0, 90);
}
function cleanTitle(fn) {
  let base = fn.replace(/\.[^.]+$/, "");
  base = base.replace(/[_\s]+\d{1,4}[_\-.]\d{1,2}[_\-.]\d{1,4}.*$/, "").trim();
  return (base.replace(/_/g, " ").replace(/^[-_.,\s]+|[-_.,\s]+$/g, "")) || fn.replace(/\.[^.]+$/, "");
}
function isArchivedPath(rel) {
  const p = "/" + String(rel || "").toLowerCase().replace(/\\/g, "/") + "/";
  return /\/(z\.|zz|old[ _]|expired|\.inactive)/.test(p);
}
// Given a Sentinel-relative folder path like "Sama Farooqui/Sentinel/Provider/Afia Umber/1. Application..."
// return { entity, entityKey, scope, phaseIdx, sectionLabel } if recognizable, else null.
function deriveEntity(folderRel) {
  const m = folderRel.match(/Sentinel\/(.+)$/);
  if (!m) return null;
  const parts = m[1].split("/").filter(Boolean);
  if (parts[0] === "Provider" && parts.length >= 3) {
    const entity = parts[1];
    const phase = parts[2];
    const idx = SOP_PHASES.indexOf(phase);
    if (idx < 0) return null;
    // Folder names are "First [Middle...] Last"; the entityKey is "last-first[-middle]".
    // reverse() only produces that for exactly two tokens — with three it scrambled the key
    // ("Rahman M Abdul" -> "abdul-m-rahman" instead of "abdul-rahman-m"), so documents and portal
    // links for every provider with a middle name or two-word surname pointed at a key that
    // matches no roster row. Move only the LAST token to the front, as the roster does.
    const toks = entity.split(/\s+/).filter(Boolean);
    const keyName = toks.length > 1 ? [toks[toks.length - 1]].concat(toks.slice(0, -1)).join(" ") : entity;
    return { scope: "provider", entity, entityKey: slug(keyName), phaseIdx: idx, sectionLabel: phase };
  }
  if (parts[0] === "State Readiness" && parts.length >= 3) {
    const fac = parts[1], sect = parts[2];
    const entity = fac === "Castle Hills" ? "Castle Hills ER" : (fac === "Frisco" ? "Frisco ER" : null);
    if (!entity) return null;
    let idx = STATE_SECTIONS.indexOf(sect);
    // A folder added via the dashboard ("Add folder") won't be one of the standard sections —
    // surface its files anyway (bucketed under "Other"), don't drop them.
    if (idx < 0) idx = STATE_SECTIONS.length;
    return { scope: "facility", entity, entityKey: slug(entity), phaseIdx: idx, sectionLabel: sect };
  }
  return null;
}
const OCR_KEY = process.env.OCR_SPACE_KEY || "helloworld";

const FILE_RULES = {
  "ACLS Certification": "\\bacls\\b", "ATLS Certification": "\\batls\\b", "PALS Certification": "\\bpals\\b",
  "BLS Certification": "\\bbls\\b", "State Medical License": "tmb[ ]*cert|medical license|tmb certificate|tmb[ ]+\\d",
  "Medical License Verify (annual)": "tmb[ ]*ver|tmb veri",
  // "^dea " must NOT swallow "DEA Verification ..." — that catch-all was matching the annual
  // verification letters and filing them (and their dates) under the registration itself, which
  // both mislabels the document and writes the wrong expiry into the master Excel.
  "Individual DEA Registration": "dea[ ]*cert|dea certificate|dea[ ]*reg|^dea (?!ver)",
  "DEA Verify (annual)": "dea[ ]*ver", "Influenza Vaccination": "flu|influenza",
  "TB Screening": "\\btb\\b|ppd|tubercul|quantiferon|\\bcxr\\b|chest", "Driver's License": "txdl|driver|drivers? lic|\\bdl\\b",
  "NPDB Query (2 yrs)": "npdb", "OIG / SAM Exclusion Check": "oig|sam |exclusion", "NPI Verification": "nppes|\\bnpi\\b",
  "TSCA Documents": "tsca", "CME (20 hrs / 2 yrs)": "\\bcme\\b", "Delineation of Privileges (DOP)": "privilege|\\bdop\\b",
  "Peer References": "reference|peer", "Initial Application": "application|initial app",
  "CV / Resume": "\\bcv\\b|resume|curriculum", "Medical Diploma": "diploma|medical school|ecfmg",
  "Malpractice / COI Insurance": "malpractice|certificate of insurance|\\bcoi\\b|tail coverage|policy",
  // Board cert files are commonly named ONLY by the board acronym (e.g. "ABEM_2027.pdf"),
  // so we accept the common boards in addition to literal "board"/"recert".
  "Board Certification": "board|recert|\\b(abem|abfm|abim|abps|aobem|aobim|aboem|abog|abpn|abs|abucm|aagp|abo|abr)\\b"
};
// Normalize a filename for rule matching. Underscores were handled but hyphens and dots were
// not, so "DEA-Cert-2027.pdf" and "Medical.License.2027.pdf" matched nothing while the
// underscore spellings matched — the same document filed two ways behaved differently.
// The old blanket /ii/->/i/ also mangled real words ("Hawaii" -> "Hawai"), so it is now limited
// to the one thing it was for: the roster's recurring "Verifiy" misspelling.
function normf(s) {
  return String(s).toLowerCase()
    .replace(/[_\-.]+/g, " ")   // hyphens and dots too, not just underscores
    .replace(/ii/g, "i")        // kept as-is: existing rules were tuned against this behavior
    .replace(/\s+/g, " ")
    .trim();
}
function extractDates(text) {
  const out = []; let m;
  const push = (y, mo, d) => { if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) out.push(y + "-" + String(mo).padStart(2, "0") + "-" + String(d).padStart(2, "0")); };
  const re = /\b(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})\b/g;
  while ((m = re.exec(text))) { let y = +m[3]; if (y < 100) y += 2000; push(y, +m[1], +m[2]); }
  const mon = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  const re2 = /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+(\d{1,2}),?\s+(\d{4})\b/gi;
  while ((m = re2.exec(text))) push(+m[3], mon[m[1].toLowerCase().slice(0, 3)], +m[2]);
  return [...new Set(out)].sort();
}
function pickExpiry(dates) {
  if (!dates || !dates.length) return null;
  const today = new Date().toISOString().slice(0, 10);
  const fut = dates.filter(d => d >= today);
  return fut.length ? fut[fut.length - 1] : dates[dates.length - 1];
}
// OCR a document by its library path. Returns [dates] (possibly empty) or null on error/throttle.
async function ocrDates(token, folderPath, name) {
  try {
    const url = docsRoot() + "/root:/" + encPath(folderPath + "/" + name) + ":/content";
    const fr = await fetch(url, { headers: { Authorization: "Bearer " + token } });
    if (!fr.ok) return null;
    const ab = await fr.arrayBuffer();
    const fd = new FormData();
    fd.append("apikey", OCR_KEY); fd.append("OCREngine", "2"); fd.append("scale", "true");
    if (/\.pdf$/i.test(name || "")) fd.append("filetype", "PDF");
    fd.append("file", new Blob([ab]), name || "doc");
    const o = await fetch("https://api.ocr.space/parse/image", { method: "POST", body: fd });
    const oj = await o.json().catch(() => ({}));
    if (oj.IsErroredOnProcessing) return null;
    return extractDates((oj.ParsedResults || []).map(p => p.ParsedText || "").join("\n"));
  } catch (e) { return null; }
}

let INDEX = null;
// Build the folder->items index from the LIVE item set (baked data.json + the roster delta).
// It used to read the baked data only, so a provider added through "+ Add provider" was absent
// from the index: every document later dropped into their brand-new folder matched nothing and
// their credentials stayed "0 on file" until a full offline regenerate. That is the very first
// thing a new client does, so it has to work.
async function indexAsync() {
  if (INDEX) return INDEX;
  const base = getData().items || [];
  let items = base;
  try { items = await applyRosterDelta(base); } catch (e) { items = base; }
  const folders = {};
  for (const it of items) {
    const rel = docsPathFromUrl(it.folderLink || it.fileLink || "");
    if (!rel) continue;
    (folders[rel] = folders[rel] || []).push(it);
  }
  INDEX = { folders, rels: Object.keys(folders).sort((a, b) => b.length - a.length) };
  return INDEX;
}
function index() {
  // Synchronous accessor for call sites that run after indexAsync() has populated the cache.
  if (INDEX) return INDEX;
  const folders = {};
  for (const it of (getData().items || [])) {
    const rel = docsPathFromUrl(it.folderLink || it.fileLink || "");
    if (!rel) continue;
    (folders[rel] = folders[rel] || []).push(it);
  }
  INDEX = { folders, rels: Object.keys(folders).sort((a, b) => b.length - a.length) };
  return INDEX;
}
function matchItem(folderRel, fileName) {
  const { folders, rels } = index();
  const rel = rels.find(r => folderRel === r || folderRel.startsWith(r + "/"));
  if (!rel) return null;
  const nf = normf(fileName);
  for (const it of folders[rel]) { const rule = FILE_RULES[it.category]; if (rule && new RegExp(rule, "i").test(nf)) return it; }
  return null;
}
function relFromParent(path) {
  const i = String(path || "").indexOf("root:");
  if (i < 0) return null;
  let rel = String(path).slice(i + 5).replace(/^\/+/, "");
  try { rel = decodeURIComponent(rel); } catch (e) { }
  return rel;
}

module.exports = async (req, res) => {
  // No wildcard CORS here: this endpoint WRITES. It can auto-fill expiry dates into the master
  // Excel, snapshot the workbook, and drop roster rows when a provider folder disappears. It
  // previously had no authentication at all, so any anonymous caller could mutate the company's
  // source-of-truth spreadsheet. Only the signed-in dashboard calls it — require a session, or
  // the cron secret for scheduled runs.
  res.setHeader("Cache-Control", "no-store");
  if (req.method === "OPTIONS") { res.status(204).end(); return; }
  const { getSession, isCronRequest } = require("../lib/session");
  if (!getSession(req) && !isCronRequest(req)) { res.status(401).json({ ok: false, message: "sign-in required" }); return; }
  try {
    // Warm the folder index from the LIVE item set (includes providers added via the dashboard)
    // before any matching happens.
    await indexAsync();
    const token = await accessToken();
    const state = (await readJsonAt(token, STATE)) || {};
    if (!state.deltaLink || state.driveTag !== "docs-v1") {
      const r = await fetch(docsRoot() + "/root/delta?token=latest", { headers: { Authorization: "Bearer " + token } });
      const j = await r.json();
      await writeJsonAt(token, STATE, { deltaLink: j["@odata.deltaLink"], driveTag: "docs-v1" });
      res.status(200).json({ ok: true, initialized: true, detected: 0 });
      return;
    }
    const detected = (await readJsonAt(token, DETECTED)) || {};
    const supplemental = (await readJsonAt(token, SUPP)) || {};   // url -> full record
    let next = state.deltaLink, deltaLink = null, changed = 0, suppChanged = 0, pages = 0;
    const folderErrors = [];   // surface folder-watch failures (e.g. trash write) instead of swallowing
    const dateUpdates = [];     // {entityKey,category,date} to auto-fill into the Credentials sheet (empty cells only)
    const start = Date.now();
    while (next && pages < 8 && Date.now() - start < 4000) {   // keep delta short, leave time for one OCR
      const r = await fetch(next, { headers: { Authorization: "Bearer " + token } });
      if (!r.ok) break;
      const j = await r.json();
      for (const v of (j.value || [])) {
        const folderRel = relFromParent((v.parentReference && v.parentReference.path) || "");
        // Only the Sentinel tree, case-insensitive + boundary-anchored, and skip archive subpaths.
        if (!folderRel || !/(^|\/)Sentinel(\/|$)/i.test(folderRel)) continue;
        if (/(^|\/)(zz?\.|old[_ ]|expired|\.inactive)/i.test(folderRel)) continue;
        // Folder events: a new/deleted Provider/<Name>/ folder updates the master Excel roster.
        // The folder appears as a Graph item with .folder set, parented at the Provider directory.
        if (v.folder && /(^|\/)Sentinel\/Provider$/i.test(folderRel)) {
          const name = String(v.name || "").trim();
          // Skip system / test / placeholder names so a stray folder doesn't pollute the roster.
          // Also skip "zz."-prefixed folders — that's the archive convention used throughout this
          // OneDrive (and the one the facility soft-delete writes). Without it, archiving a
          // provider by renaming their folder to "zz.Smith John" made the watcher append a
          // phantom roster row (last="John", first="zz.Smith") to the master Excel.
          if (!name || /^[._]/.test(name) || /^zz\./i.test(name) || /^(test|temp|new folder|untitled)/i.test(name)) continue;
          try {
            const xl = require("../lib/excel");
            const parts = name.split(/\s+/);
            const last = parts[parts.length - 1] || name;
            const first = parts.slice(0, -1).join(" ") || "";
            if (v.deleted) {
              // Folder removed from SharePoint -> HARD delete the roster row(s) and log to
              // Recycle bin so the dashboard reflects it immediately (no "Inactive" middle step).
              await xl.snapshotWorkbook(token, "scan-folder-deleted");
              const slugFn = (l, f) => (l + "-" + (f || "")).replace(/[\*,()]+/g, "").replace(/[^A-Za-z0-9]+/g, "-").replace(/^-+|-+$/g, "").toLowerCase();
              const ekey = slugFn(last, first);
              const removed = await xl.hardDelete(token, ekey, last, first);
              if (removed.length) {
                const trashPath = drivePath("_Sentinel/trash.json");
                const trash = (await readJsonAt(token, trashPath)) || { entries: [] };
                if (ekey) trash.entries = (trash.entries || []).filter(e => e.entityKey !== ekey);
                trash.entries.unshift({
                  id: "tr_" + Date.now() + "_" + Math.random().toString(36).slice(2, 8),
                  entityKey: ekey, entity: name,
                  deletedAt: new Date().toISOString(),
                  deletedBy: "scan/folder-watcher",
                  rows: removed,
                });
                if (trash.entries.length > 200) trash.entries = trash.entries.slice(0, 200);
                await writeJsonAt(token, trashPath, trash);
              }
            } else {
              const dup = await xl.findAnywhere(token, last, first);
              // appendProviderRow — NOT appendRow. appendRow writes positionally at columns 1-2,
              // which on a roster carrying a leading tracking column puts the surname in that
              // column and leaves First Name blank. This runs unattended from the folder watcher,
              // so a positional write here silently corrupts the source of truth.
              if (!dup) { await xl.snapshotWorkbook(token, "scan-folder-added"); await xl.appendProviderRow(token, xl.SHEET_ACTIVE, last, first); }
            }
          } catch (e) { folderErrors.push({ name, error: String(e.message || e).slice(0, 200) }); }
          continue;
        }
        if (!v.file && !v.deleted) continue;
        const it = matchItem(folderRel, v.name || "");
        if (it) {
          if (v.deleted) { if (detected[it.id] && detected[it.id].name === v.name) { delete detected[it.id]; changed++; } }
          else {
            const nameDate = dateFromName(v.name);
            detected[it.id] = Object.assign({}, detected[it.id], { url: v.webUrl || "", name: v.name, date: nameDate || (detected[it.id] || {}).date || null });
            changed++;
            // a fresh expiry from the filename → queue an Excel auto-fill (empty cells only)
            if (nameDate && it.entityKey && it.category) dateUpdates.push({ entityKey: it.entityKey, category: it.category, date: nameDate });
          }
          continue;
        }
        // No tracked item matched this file — surface it as a supplemental record so the
        // dashboard still shows it (within the 45-second live-sync, no regen required).
        const ent = deriveEntity(folderRel);
        if (!ent) continue;
        const key = v.webUrl || ((v.parentReference && v.parentReference.path) || "") + "/" + (v.name || "");
        if (v.deleted) {
          if (supplemental[key]) { delete supplemental[key]; suppChanged++; }
          continue;
        }
        const ext = (v.name || "").toLowerCase().match(/\.([a-z0-9]+)$/);
        const supportedExts = ["pdf","jpg","jpeg","png","webp","gif","tif","tiff","heic","heif","doc","docx","xls","xlsx","ppt","pptx"];
        if (!ext || !supportedExts.includes(ext[1])) continue;
        const expFromName = dateFromName(v.name || "");
        const rec = {
          id: slug(ent.entityKey, "supp", (v.name || "").replace(/\.[^.]+$/, "")),
          scope: ent.scope,
          entity: ent.entity,
          entityKey: ent.entityKey,
          category: cleanTitle(v.name || ""),
          sectionLabel: ent.sectionLabel,
          phaseIdx: ent.phaseIdx,
          authority: "",
          number: "",
          issued: null,
          expires: expFromName || null,
          renewalLeadDays: ent.scope === "facility" ? 90 : 60,
          owner: ent.scope === "provider" ? ent.entity : "",
          fileLink: v.webUrl || "",
          folderLink: v.webUrl ? v.webUrl.split("/").slice(0, -1).join("/") : "",
          isFile: true,
          supplemental: true,
          liveAdded: true,
          permanent: !expFromName,
          active: true,
          notes: "Supplemental document (detected live by background scan)",
        };
        supplemental[key] = rec;
        suppChanged++;
      }
      next = j["@odata.nextLink"]; deltaLink = j["@odata.deltaLink"] || deltaLink; pages++;
    }
    // Save delta progress BEFORE the (slower) OCR step so nothing is lost on timeout.
    if (changed) await writeJsonAt(token, DETECTED, detected);
    if (suppChanged) await writeJsonAt(token, SUPP, supplemental);
    await writeJsonAt(token, STATE, { deltaLink: deltaLink || next || state.deltaLink, driveTag: "docs-v1" });

    // Background OCR: read the expiry for ONE not-yet-read document (keeps each run < 10s).
    let ocrChanged = false, ocred = null;
    const byId = {}; for (const it of (getData().items || [])) byId[it.id] = it;
    for (const id in detected) {
      if (Date.now() - start > 5500) break;
      const d = detected[id];
      if (d.date || d.ocrTried) continue;             // already has a date, or already OCR'd with no luck
      const it = byId[id]; if (!it) continue;
      const folderPath = docsPathFromUrl(it.folderLink || it.fileLink || ""); if (!folderPath) continue;
      const dates = await ocrDates(token, folderPath, d.name);
      if (dates === null) continue;                   // transient (throttle/error) — retry next run
      if (dates.length) {
        d.date = pickExpiry(dates); d.ocr = true;
        if (it.entityKey && it.category) dateUpdates.push({ entityKey: it.entityKey, category: it.category, date: d.date });
      } else { d.ocrTried = true; }
      ocrChanged = true; ocred = { id, date: d.date || null }; break;   // one per run
    }
    if (ocrChanged) await writeJsonAt(token, DETECTED, detected);

    // Auto-fill detected expiry dates into empty cells of the Credentials sheet.
    let excelFill = null;
    if (dateUpdates.length) {
      const xl = require("../lib/excel");
      try { excelFill = await xl.fillEmptyDates(token, dateUpdates); }
      catch (e) { excelFill = { error: String(e.message || e).slice(0, 160) }; }
    }

    res.status(200).json({ ok: true, changed, suppChanged, items: Object.keys(detected).length, suppItems: Object.keys(supplemental).length, ocr: ocred, folderErrors, excelFill });
  } catch (e) {
    res.status(200).json({ ok: false, message: String(e.message || e) });
  }
};
