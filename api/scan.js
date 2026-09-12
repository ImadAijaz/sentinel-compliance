// Auto-detect documents added ANYWHERE under the Sentinel library — via the dashboard QR
// uploader, OneDrive, or directly in SharePoint. Graph "delta" finds new/changed files and
// matches them to items. Expiry dates are read only from filenames by the shared date parser.
// Results -> _Sentinel/auto_detected.json, which /api/uploads-map merges into the dashboard.
const { accessToken, docsRoot, docsPathFromUrl, drivePath, readJsonAt, writeJsonAt, dateFromName } = require("../lib/graph");
const { applyRosterDelta } = require("../lib/delta");
const FAC = require("../lib/facility");
const evidence = require("../lib/evidence");
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
// Archive-folder test now lives in lib/facility.js and is shared with the facility sync, so the
// provider tree ("z.Expired Docs") and the facility tree ("Z_Expired Documents_FriscoER") are
// judged by the same rule. The old private copy here only knew the dotted spelling, was never
// actually called, and the live skip below used a third regex that missed the facility form —
// which is how a superseded 2025 facility licence reached the board as an expired credential.
const isArchivedPath = FAC.isArchivedPath;
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
    // The entityKey is "last-first[-middle]". A folder named "Last, First" already has the
    // surname first, so it needs no reordering — only the "First Last" form does.
    let keyName;
    if (entity.indexOf(",") >= 0) {
      keyName = entity.replace(/,/g, " ");
    } else {
      const toks = entity.split(/\s+/).filter(Boolean);
      keyName = toks.length > 1 ? [toks[toks.length - 1]].concat(toks.slice(0, -1)).join(" ") : entity;
    }
    return { scope: "provider", entity, entityKey: slug(keyName), phaseIdx: idx, sectionLabel: phase };
  }
  // Staff documents, filed by the master sync as Staff/<Site>/<Person>/<Category>/<file>.
  // The sync brought 914 of these across for 64 people at four sites; without this they land
  // in SharePoint correctly but never reach the dashboard, because deriveEntity only knew
  // about providers and facilities.
  if (parts[0] === "Staff" && parts.length >= 3) {
    const person = parts[2];
    if (!person || /^[._]/.test(person)) return null;
    // Staff folders are "Last, First_ROLE" — keep the role out of the displayed name.
    const bare = person.replace(/_[^_]*$/, "").trim() || person;
    const keyName = bare.indexOf(",") >= 0 ? bare.replace(/,/g, " ") : bare;
    return {
      scope: "staff", entity: bare, entityKey: slug(keyName),
      phaseIdx: 0, sectionLabel: parts[3] || "Documents", site: parts[1],
    };
  }
  if (parts[0] === "State Readiness" && parts.length >= 3) {
    const fac = parts[1], sect = parts[2];
    // Was a hardcoded two-way map, so a site the sync newly files documents for (Urgent Care
    // Ennis, Urgent Care Plano) returned null and every one of its documents was dropped.
    // An ER site keeps its " ER" suffix; anything else is used as named.
    const entity = fac === "Castle Hills" ? "Castle Hills ER"
      : fac === "Frisco" ? "Frisco ER"
      : (/^[A-Za-z0-9][A-Za-z0-9 .,'&-]{1,60}$/.test(fac) ? fac : null);
    if (!entity) return null;
    let idx = STATE_SECTIONS.indexOf(sect);
    // A folder added via the dashboard ("Add folder") won't be one of the standard sections —
    // surface its files anyway (bucketed under "Other"), don't drop them.
    if (idx < 0) idx = STATE_SECTIONS.length;
    return { scope: "facility", entity, entityKey: slug(entity), phaseIdx: idx, sectionLabel: sect };
  }
  return null;
}
const { FILE_RULES, normf } = require("../lib/filerules");
let INDEX = null;
function itemFolderRel(it) {
  if (it && it.scope === "staff") {
    const site = it.staffFacility || it.facility || "Unassigned";
    const siteDir = site === "Castle Hills ER" ? "Castle Hills" : site === "Frisco ER" ? "Frisco" : site;
    const toks = String(it.entity || "").trim().split(/\s+/).filter(Boolean);
    const last = toks.length ? toks[toks.length - 1] : "Staff";
    const first = toks.slice(0, -1).join(" ");
    const role = /front desk/i.test(it.role || "") ? "FD" : (it.role || "Staff");
    return "Sama Farooqui/Sentinel/Staff/" + siteDir + "/" + last + (first ? ", " + first : "") + "_" + role;
  }
  return docsPathFromUrl((it && (it.folderLink || it.fileLink)) || "");
}
function ownerRoot(folder) {
  const m = String(folder || "").match(/^(.*?Sentinel\/(?:Provider\/[^/]+|Staff\/[^/]+\/[^/]+|State Readiness\/[^/]+))(?:\/|$)/i);
  return m ? m[1].toLowerCase() : folder;
}
// Build the folder->items index from the LIVE item set (baked data.json + the roster delta).
// It used to read the baked data only, so a provider added through "+ Add provider" was absent
// from the index: every document later dropped into their brand-new folder matched nothing and
// their credentials stayed "0 on file" until a full offline regenerate. That is the very first
// thing a new client does, so it has to work.
async function indexAsync() {
  const base = getData().items || [];
  let items = base;
  try { items = await applyRosterDelta(base); } catch (e) { items = base; }
  const folders = {};
  for (const it of items) {
    const rel = ownerRoot(itemFolderRel(it));
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
    const rel = ownerRoot(itemFolderRel(it));
    if (!rel) continue;
    (folders[rel] = folders[rel] || []).push(it);
  }
  INDEX = { folders, rels: Object.keys(folders).sort((a, b) => b.length - a.length) };
  return INDEX;
}
function matchItems(folderRel, fileName) {
  const { folders, rels } = index();
  const root = ownerRoot(folderRel);
  const rel = rels.find(r => root === r);
  if (!rel) return [];
  const found = [];
  for (const it of folders[rel]) {
    if (it.scope === "other" || it.supplemental) continue;
    const rule = FILE_RULES[it.category]; if (rule && new RegExp(rule, "i").test(normf(fileName))) found.push(it);
  }
  return found;
}
function matchRecurringItems(folderRel, fileName) {
  const { folders, rels } = index();
  const rel = rels.find(r => ownerRoot(folderRel) === r);
  if (!rel) return [];
  return FAC.matchRecurringObligations(folders[rel], folderRel + "/" + fileName, fileName);
}
function relFromParent(path) {
  const i = String(path || "").indexOf("root:");
  if (i < 0) return null;
  let rel = String(path).slice(i + 5).replace(/^\/+/, "");
  try { rel = decodeURIComponent(rel); } catch (e) { }
  return rel;
}

module.exports = async (req, res) => {
  // Writes Sentinel's evidence cache only. Never edits the roster or source documents.
  res.setHeader("Cache-Control", "no-store");
  if (req.method === "OPTIONS") { res.status(204).end(); return; }
  const { getSession, isCronRequest } = require("../lib/session");
  if (!getSession(req) && !isCronRequest(req)) { res.status(401).json({ ok: false, message: "sign-in required" }); return; }
  try {
    // Warm the folder index from the LIVE item set (includes providers added via the dashboard)
    // before any matching happens.
    await indexAsync();
    const token = await accessToken();
    const release = await require("../lib/master-intake").lock(token, "scan");
    if (!release) { res.status(200).json({ ok: true, busy: true, moreToScan: true }); return; }
    try {
    const state = (await readJsonAt(token, STATE)) || {};
    // ?rescan=1 forces a FULL re-crawl of the document tree from the beginning, instead of only
    // watching for changes from now on. This is the recovery path when folders that already exist
    // were never picked up — the normal first-run below starts at "latest", which by design skips
    // everything that was already there.
    const forceRescan = new URL(req.url, "http://localhost").searchParams.get("rescan") === "1" || state.driveTag !== "docs-v2";
    if (forceRescan || !state.deltaLink) {
      // Bound a full crawl to Sentinel, not every unrelated file in Corporate Archives.
      // Capture the change cursor BEFORE walking, so changes during the walk are not lost.
      const G = require("../lib/graph");
      const [rootResponse, cursorResponse] = await Promise.all([
        fetch(docsRoot() + "/root:/" + G.encPath("Sama Farooqui/Sentinel") + "?$select=id,folder", { headers: { Authorization: "Bearer " + token } }),
        fetch(docsRoot() + "/root/delta?token=latest", { headers: { Authorization: "Bearer " + token } }),
      ]);
      if (!rootResponse.ok || !cursorResponse.ok) throw new Error("Could not initialize the full Sentinel document scan");
      const root = await rootResponse.json(), cursor = await cursorResponse.json();
      if (!root.folder || !cursor["@odata.deltaLink"]) throw new Error("Microsoft did not return the document folder or change cursor");
      Object.assign(state, { deltaLink: cursor["@odata.deltaLink"], driveTag: "docs-v2", walkQueue: [{ id: root.id, rel: "Sama Farooqui/Sentinel" }], fullRescanStartedAt: new Date().toISOString(), pending: [], parentPaths: {}, inspected: 0, inSentinel: 0 });
      await writeJsonAt(token, STATE, state);
    }
    const detected = (await readJsonAt(token, DETECTED)) || {};
    const supplemental = (await readJsonAt(token, SUPP)) || {};   // url -> full record
    const walking = Array.isArray(state.walkQueue);
    const walkQueue = state.walkQueue || [];
    let next = walking ? null : state.deltaLink;
    let deltaLink = null, changed = 0, suppChanged = 0, pages = 0, resynced = false, fetchError = null;
    let pending = forceRescan ? [] : (state.pending || []);
    const parentPaths = forceRescan ? {} : (state.parentPaths || {});
    let inspected = 0, inSentinel = 0;
    async function parentPath(id, depth = 0) {
      if (!id || depth > 30) throw new Error("Cannot resolve a document's parent folder");
      if (Object.prototype.hasOwnProperty.call(parentPaths, id)) return parentPaths[id];
      const r = await fetch(docsRoot() + "/items/" + encodeURIComponent(id) + "?$select=id,name,parentReference,root", { headers: { Authorization: "Bearer " + token }, signal: AbortSignal.timeout(10000) });
      if (!r.ok) throw new Error("Resolve document parent HTTP " + r.status);
      const p = await r.json();
      if (p.root || !p.parentReference) return (parentPaths[id] = "");
      const parent = relFromParent(p.parentReference.path) ?? await parentPath(p.parentReference.id, depth + 1);
      return (parentPaths[id] = [parent, p.name].filter(Boolean).join("/"));
    }
    const folderErrors = [];   // surface folder-watch failures (e.g. trash write) instead of swallowing
    const start = Date.now();
    while ((next || pending.length || walkQueue.length) && pages < 12 && Date.now() - start < 15000) {
      if (!pending.length) {
      const directory = walking ? walkQueue[0] : null;
      const pageUrl = directory ? (directory.next || docsRoot() + "/items/" + encodeURIComponent(directory.id) + "/children?$select=id,name,file,folder,parentReference,webUrl&$top=500") : next;
      const r = await fetch(pageUrl, { headers: { Authorization: "Bearer " + token }, signal: AbortSignal.timeout(10000) });
      if (!r.ok) {
        const body = await r.text().catch(() => "");
        // A delta token eventually expires; Graph answers 410 Gone / resyncRequired. The old code
        // just broke out of the loop and then wrote the SAME dead link back, so the scan was
        // permanently and SILENTLY dead — no new folder or file was ever detected again, while
        // still reporting ok:true. Recover by restarting the crawl from the beginning.
        if (!walking && (r.status === 410 || /resyncRequired|resync/i.test(body))) {
          state.driveTag = "expired";
          fetchError = "Change cursor expired; a full Sentinel scan will restart on the next pass";
          break;
        }
        fetchError = "delta HTTP " + r.status + " " + body.slice(0, 140);
        break;
      }
      const j = await r.json();
      pending = j.value || [];
      if (directory) {
        walkQueue.shift();
        if (j["@odata.nextLink"]) walkQueue.push({ ...directory, next: j["@odata.nextLink"] });
        for (const item of pending) {
          item.parentReference = { ...(item.parentReference || {}), path: "root:/" + directory.rel };
          if (item.folder) {
            const rel = directory.rel + "/" + item.name;
            const top = directory.rel === "Sama Farooqui/Sentinel";
            if ((!top || ["Provider", "Staff", "State Readiness"].includes(item.name)) && !require("../lib/master-filing").excluded(rel)) walkQueue.push({ id: item.id, rel });
          }
        }
      } else { next = j["@odata.nextLink"] || null; deltaLink = j["@odata.deltaLink"] || deltaLink; }
      pages++;
      }
      while (pending.length && Date.now() - start < 15000) {
        const v = pending[0];
        if (!v.file || v.deleted) { pending.shift(); continue; }
        let folderRel;
        try { folderRel = relFromParent((v.parentReference && v.parentReference.path) || "") ?? await parentPath(v.parentReference && v.parentReference.id); }
        catch (e) { fetchError = String(e.message || e); break; }
        pending.shift(); inspected++;
        // Only the Sentinel tree, case-insensitive + boundary-anchored, and skip archive subpaths.
        if (!folderRel || !/(^|\/)Sentinel(\/|$)/i.test(folderRel)) continue;
        if (isArchivedPath(folderRel)) continue;
        if (require("../lib/master-filing").excluded(folderRel)) continue;
        inSentinel++;
        // A SharePoint folder is document storage, not roster authority. Ignore provider-folder
        // events here: additions and removals must come from the master Excel roster. In
        // particular, deleting or archiving a folder must never hard-delete a provider row.
        if (v.folder && /(^|\/)Sentinel\/Provider$/i.test(folderRel)) {
          continue;
        }
        if (!v.file && !v.deleted) continue;
        const matched = matchItems(folderRel, v.name || "");
        const recurring = matchRecurringItems(folderRel, v.name || "");
        let matchedTracked = false;
        for (const it of matched) {
          if (v.deleted) { if (detected[it.id] && detected[it.id].name === v.name) { delete detected[it.id]; changed++; } }
          else {
            const nameDate = dateFromName(v.name);
            const selected = evidence.choose(detected[it.id], { url: v.webUrl || "", name: v.name, date: nameDate || null });
            if (JSON.stringify(selected) !== JSON.stringify(detected[it.id])) { detected[it.id] = selected; changed++; }
          }
          matchedTracked = true;
        }
        for (const rec of recurring) {
          const rit = rec.item;
          if (v.deleted) {
            if (detected[rit.id] && detected[rit.id].name === v.name) { delete detected[rit.id]; changed++; }
          } else if (rec.recordDate) {
            const prev = detected[rit.id] || {};
            // Delta events can arrive out of order during a full crawl.  Keep the newest event
            // evidence so an older inspection can never pull a recurring due date backwards.
            if (!prev.recordDate || rec.recordDate >= prev.recordDate) {
              detected[rit.id] = {
                url: v.webUrl || "", name: v.name,
                date: rec.nextDue || null, recordDate: rec.recordDate,
                recurring: true, cadenceMonths: rec.cadenceMonths || null,
                cadenceNote: rec.note || null,
              };
              changed++;
            }
          }
          matchedTracked = true;
        }
        if (matchedTracked) continue;
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
        const nameDate2 = dateFromName(v.name || "");
        // Minutes, agendas, inspection reports, service calls, "as of" snapshots, emails and
        // incident reports carry the date the thing HAPPENED, not a date they stop being valid.
        // Treating that as an expiry turned every set of meeting minutes red the day after the
        // meeting: 54 of the 62 dated facility documents on the board were records of this kind.
        // They keep their date — shown as a dated record — but no longer run an expiry clock.
        const eventOnly = FAC.isNonExpiring(v.name || "");
        const expFromName = eventOnly ? null : nameDate2;
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
          recordDate: eventOnly ? nameDate2 : null,
          datedRecord: eventOnly && !!nameDate2,
          notes: eventOnly
            ? (nameDate2 ? "Dated record — " + nameDate2 + ". This kind of document does not expire." : "Record document — no expiry.")
            : "Supplemental document (detected live by background scan)",
        };
        supplemental[key] = rec;
        suppChanged++;
      }
      if (fetchError) break;
    }
    // Save delta progress before any workbook update so nothing is lost on timeout.
    if (changed) await writeJsonAt(token, DETECTED, detected);
    if (suppChanged) await writeJsonAt(token, SUPP, supplemental);
    // Never write back a link we know is dead. `deltaLink` is the fresh cursor from a completed
    // page, `next` is the nextLink mid-crawl; only fall back to the previous cursor if we have
    // neither AND the previous one didn't just fail.
    const nextCursor = deltaLink || next || state.deltaLink;
    const moreToScan = pending.length > 0 || walkQueue.length > 0 || (!!next && !deltaLink);
    await writeJsonAt(token, STATE, {
      deltaLink: nextCursor, driveTag: state.driveTag, pending, parentPaths,
      ...(walking && moreToScan ? { walkQueue } : {}),
      inspected: (forceRescan ? 0 : state.inspected || 0) + inspected,
      inSentinel: (forceRescan ? 0 : state.inSentinel || 0) + inSentinel,
      lastPassAt: new Date().toISOString(),
      completedAt: !moreToScan && !fetchError ? new Date().toISOString() : state.completedAt,
    });

    // Report the crawl state, so a stalled or recovering scan is visible instead of silent.
    res.status(200).json({
      ok: true, changed, suppChanged, items: Object.keys(detected).length,
      suppItems: Object.keys(supplemental).length, folderErrors,
      pages, resynced, inspected, inSentinel, moreToScan, error: fetchError || undefined,
    });
    } finally { await release(); }
  } catch (e) {
    res.status(200).json({ ok: false, message: String(e.message || e) });
  }
};
