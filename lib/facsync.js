// Facility document sync engine.
//
// The facility documents the dashboard reads live under
//   Sama Farooqui/Sentinel/State Readiness/<Facility>/<NN. Section>/
// and are a COPY of a folder somebody else maintains. Measured 2026-09-11: the whole copy
// matched the legacy "..State Requirements_<SITE>" tree byte-for-byte, while the live COLA
// accreditation certificate was two years newer than the one the board was showing. Nobody did
// anything wrong — there was simply no mechanism to carry updates across.
//
// This module is the ONE implementation, used by both api/data.js (the button) and
// api/digest.js (the nightly cron). digest.js previously carried a private copy of the roster
// regen and kept an outdated reader after the shared one was fixed, silently corrupting the
// board every night. Not repeating that.
//
// Everything here is COPY-ONLY at the source. The one operation that moves a file is
// reorganize(), and it only ever moves a document that is already inside Sentinel's own tree
// into a different Sentinel section.
const G = require("./graph");
const FAC = require("./facility");

const FAC_DIR = { "Castle Hills ER": "Castle Hills", "Frisco ER": "Frisco" };
const SOURCES = G.drivePath("_Sentinel/facility_sources.json");
const baseFor = (f) => "Sama Farooqui/Sentinel/State Readiness/" + FAC_DIR[f];

function isFacility(f) { return !!FAC_DIR[f]; }
function facilities() { return Object.keys(FAC_DIR); }

async function readSources(token) {
  return (await G.readJsonAt(token, SOURCES)) || { facilities: {} };
}
async function saveSource(token, facility, src, who) {
  const cur = await readSources(token);
  cur.facilities = cur.facilities || {};
  cur.facilities[facility] = { src: String(src).trim(), savedAt: new Date().toISOString(), savedBy: who || "unknown" };
  await G.writeJsonAt(token, SOURCES, cur);
  return cur;
}

// Resolve a "Copy link" share URL (or a plain SharePoint URL) to a drive folder.
async function resolveSource(token, link) {
  const shareId = "u!" + Buffer.from(String(link), "utf8").toString("base64")
    .replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  const r = await fetch(G.GRAPH + "/shares/" + shareId + "/driveItem?$select=id,name,folder,parentReference,webUrl",
    { headers: { Authorization: "Bearer " + token } });
  if (!r.ok) {
    return { error: "Could not open that link (HTTP " + r.status + "). Check it is a link to the FOLDER, and that the Sentinel app has access to that drive.",
      detail: (await r.text()).slice(0, 200) };
  }
  const item = await r.json();
  if (!item.folder) return { error: "That link points at a file, not a folder. Share the folder that holds the documents." };
  return { driveId: item.parentReference && item.parentReference.driveId, id: item.id, name: item.name, webUrl: item.webUrl };
}

async function children(token, driveId, itemId) {
  const out = [];
  let u = G.GRAPH + "/drives/" + driveId + "/items/" + itemId +
    "/children?$select=id,name,file,folder,size,lastModifiedDateTime&$top=200";
  while (u) {
    const r = await fetch(u, { headers: { Authorization: "Bearer " + token } });
    if (!r.ok) break;
    const j = await r.json();
    (j.value || []).forEach(x => out.push(x));
    u = j["@odata.nextLink"] || null;
  }
  return out;
}

// Walk the source tree. Their real trees nest five deep
// (QAPI Meetings_CHER / QAPI_2025 / QAPI_1st Qtr_CHER_2025 / QAPI_January 2026_CHER), so the
// one-level walk the provider import uses would have missed most of it.
async function walk(token, driveId, itemId, rel, depth, acc, deadline) {
  if (depth > 6 || acc.files.length >= 900 || Date.now() > deadline) return;
  for (const c of await children(token, driveId, itemId)) {
    if (c.folder) {
      if (FAC.isArchivedSegment(c.name)) { acc.skippedFolders.push((rel ? rel + "/" : "") + c.name); continue; }
      await walk(token, driveId, c.id, (rel ? rel + "/" : "") + c.name, depth + 1, acc, deadline);
    } else if (c.file) {
      acc.files.push({ id: c.id, name: c.name, size: c.size || 0, rel: rel });
    }
  }
}

// Everything already inside the facility's Sentinel tree, keyed by lowercased filename, so a
// re-run copies nothing twice. Without this the nightly autosync would duplicate the entire
// library every night.
// Walks the WHOLE section, to any depth. Measured 2026-09-11, the real tree runs seven path
// segments deep (Frisco/08. Laboratory Services/QAPI Meetings/QAPI_2025/QAPI_1st Qtr/<file>) —
// 100 of its 200 documents sit below the second level. A two-level look missed every one of
// them, which would have made the dedupe blind and had the nightly autosync re-copy half the
// library every night, and left the misfiled-document check reporting far less than is there.
async function existingIn(token, facility) {
  const map = {};
  const recurse = async (itemId, sect, rel, depth) => {
    if (depth > 6) return;
    for (const c of await children(token, G.DOCS_DRIVE_ID, itemId)) {
      if (c.folder) {
        if (FAC.isArchivedSegment(c.name)) continue;
        await recurse(c.id, sect, rel ? rel + "/" + c.name : c.name, depth + 1);
      } else if (c.file) {
        map[String(c.name).toLowerCase()] = { section: sect, sub: rel, id: c.id, size: c.size || 0, name: c.name };
      }
    }
  };
  for (const sect of FAC.STATE_SECTIONS) {
    const r = await fetch(G.docsRoot() + "/root:/" + G.encPath(baseFor(facility) + "/" + sect) + "?$select=id",
      { headers: { Authorization: "Bearer " + token } });
    if (!r.ok) continue;                       // section folder not created yet — nothing in it
    await recurse((await r.json()).id, sect, "", 0);
  }
  return map;
}

// `relPath` is a section ("10. Quality Improvement Program") or a section plus the subfolders
// to keep underneath it ("10. Quality Improvement Program/QAPI Meetings/QAPI_2025").
// ensureFolderIn creates each missing segment in turn, so either form is safe.
async function destFolderId(token, facility, relPath, cache) {
  if (cache && cache[relPath]) return cache[relPath];
  const p = baseFor(facility) + "/" + relPath;
  try { await G.ensureFolderIn(token, G.docsRoot(), p); } catch (e) {}
  const r = await fetch(G.docsRoot() + "/root:/" + G.encPath(p) + "?$select=id", { headers: { Authorization: "Bearer " + token } });
  if (!r.ok) return null;
  const id = (await r.json()).id;
  if (cache) cache[relPath] = id;
  return id;
}

// Build the filing plan for one facility. Never writes anything.
async function plan(token, facility, src, deadline) {
  const source = await resolveSource(token, src);
  if (source.error) return { facility, error: source.error, detail: source.detail };
  const acc = { files: [], skippedFolders: [] };
  await walk(token, source.driveId, source.id, "", 0, acc, deadline);
  const have = await existingIn(token, facility);

  const toCopy = [], already = [], datedInThePast = [];
  for (const f of acc.files) {
    const c = FAC.classifySection(f.rel, f.name);
    const section = c ? c.section : FAC.STATE_SECTIONS[0];
    const hit = have[String(f.name).toLowerCase()];
    if (hit) { already.push({ name: f.name, section: hit.section, wanted: section, misfiled: hit.section !== section }); continue; }
    // Their filename convention puts the EXPIRY date in the name ("COLA_01_17_2026_CHER.pdf"
    // is the certificate issued 17 Jan 2024 that runs out two years later). An incoming file
    // dated in the past is therefore either genuinely expired, or — far more often — named
    // with the date printed on the certificate instead of the date it runs out. Worth saying
    // out loud rather than silently importing something that will read as expired.
    if (!FAC.isNonExpiring(f.name)) {
      const d = G.dateFromName(f.name);
      if (d && new Date(d) < new Date()) datedInThePast.push({ name: f.name, date: d });
    }
    toCopy.push({ id: f.id, name: f.name, size: f.size, from: f.rel || "(top level)", section });
  }
  return { facility, source, files: acc.files, skippedFolders: acc.skippedFolders, toCopy, already, datedInThePast };
}

// Copy the planned files in. Idempotent (anything already present by name is skipped) and
// resumable (stops on the deadline and reports what is left).
async function copyIn(token, facility, p, deadline, cap) {
  const cache = {};
  let copying = 0; const errors = [];
  const list = p.toCopy.slice().sort((a, b) => (b.name > a.name ? 1 : -1));
  for (const f of list) {
    if (Date.now() > deadline || copying >= (cap || 120)) break;
    const destId = await destFolderId(token, facility, f.section, cache);
    if (!destId) { errors.push({ name: f.name, error: "could not open destination section" }); continue; }
    const r = await fetch(G.GRAPH + "/drives/" + p.source.driveId + "/items/" + f.id + "/copy", {
      method: "POST", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify({
        parentReference: { driveId: G.DOCS_DRIVE_ID, id: destId },
        name: f.name, "@microsoft.graph.conflictBehavior": "rename",
      }),
    });
    if (r.status === 202 || r.ok) copying++;
    else errors.push({ name: f.name, error: "HTTP " + r.status + " " + (await r.text()).slice(0, 120) });
  }
  return { copying, failed: errors.length, errors: errors.slice(0, 10), remaining: Math.max(0, p.toCopy.length - copying - errors.length) };
}

// Documents already inside Sentinel that sit in a section which does not match what they are.
// 50 QAPI meeting packs were filed under "08. Laboratory Services" when this was written.
async function misfiled(token, facility) {
  const have = await existingIn(token, facility);
  const moves = [];
  for (const lname in have) {
    const cur = have[lname];
    const c = FAC.classifySection(cur.sub || "", cur.name || lname);
    if (!c || c.section === cur.section) continue;
    moves.push({
      facility, name: cur.name || lname, id: cur.id,
      from: cur.section + (cur.sub ? "/" + cur.sub : ""),
      to: c.section,
      // Keep whatever structure the document already sits in ("QAPI Meetings/QAPI_2025/
      // QAPI_1st Qtr_CHER_2025"). Refiling it into the right section should not flatten the
      // year/quarter folders somebody built and still navigates by.
      sub: cur.sub || "",
    });
  }
  return moves;
}

async function reorganize(token, moves, deadline) {
  const cache = {};
  let moved = 0; const errors = [];
  for (const m of moves) {
    if (Date.now() > deadline) break;
    const destId = await destFolderId(token, m.facility, m.to + (m.sub ? "/" + m.sub : ""), cache);
    if (!destId) { errors.push({ name: m.name, error: "could not open destination section" }); continue; }
    const r = await fetch(G.GRAPH + "/drives/" + G.DOCS_DRIVE_ID + "/items/" + m.id, {
      method: "PATCH", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify({ parentReference: { id: destId } }),
    });
    if (r.ok) moved++;
    else errors.push({ name: m.name, error: "HTTP " + r.status + " " + (await r.text()).slice(0, 120) });
  }
  return { moved, failed: errors.length, errors: errors.slice(0, 10), remaining: Math.max(0, moves.length - moved - errors.length) };
}

// Nightly entry point. Silent no-op until somebody has saved a source folder, so this is safe
// to leave wired into the cron before the feature is ever used.
async function syncSavedFacilities(opts) {
  const budgetMs = (opts && opts.budgetMs) || 40000;
  const deadline = Date.now() + budgetMs;
  const token = await G.accessToken();
  const saved = await readSources(token);
  const names = Object.keys(saved.facilities || {});
  if (!names.length) return { ran: 0, note: "no facility source folders saved yet" };
  const out = [];
  for (const f of names) {
    if (Date.now() > deadline) { out.push({ facility: f, skipped: "ran out of time this run" }); continue; }
    if (!isFacility(f)) { out.push({ facility: f, error: "unknown facility" }); continue; }
    const p = await plan(token, f, saved.facilities[f].src, deadline);
    if (p.error) { out.push({ facility: f, error: p.error }); continue; }
    const r = await copyIn(token, f, p, deadline, 80);
    out.push({ facility: f, found: p.files.length, copying: r.copying, alreadyThere: p.already.length,
      failed: r.failed, remaining: r.remaining, datedInThePast: p.datedInThePast.length });
  }
  return { ran: out.length, facilities: out };
}

module.exports = {
  FAC_DIR, SOURCES, baseFor, isFacility, facilities,
  readSources, saveSource, resolveSource, existingIn,
  plan, copyIn, misfiled, reorganize, syncSavedFacilities,
};
