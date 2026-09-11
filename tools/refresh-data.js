#!/usr/bin/env node
// Refresh the dashboard's baked data from the documents actually on disk.
//
// WHY THIS EXISTS
// data.json carries an expiry per credential, baked on 2026-06-23. The live Graph delta scan
// is supposed to notice a newer certificate and push the new date onto the card, but that scan
// only runs while somebody has a dashboard tab open, works through a backlog a page at a time,
// and depends on SharePoint having caught up. After the master sync filed 4,092 documents, the
// board still showed COLA expiring 17 Jan 2026 while the certificate on the wall ran to
// 26 Feb 2028 — and waiting on the cloud to notice is not an answer.
//
// Every one of those documents is on this machine. This reads them directly, matches each to
// the credential it is evidence for using the SAME rules as the scan (lib/filerules.js), and
// writes the newer expiry into data.json. Deterministic, immediate, and verifiable here.
//
//   node tools/refresh-data.js            # dry run, prints what would change
//   node tools/refresh-data.js --apply    # write data.json
//
// SAFETY: a date only ever moves FORWARD. A credential is never given an earlier expiry than
// it already has, so a stray old copy of a document cannot pull a current credential backwards
// into looking expired.

const fs = require("fs");
const path = require("path");
const FAC = require("../lib/facility");
const { matchByRules } = require("../lib/filerules");

const HOME = process.env.USERPROFILE || process.env.HOME || "";
const SENTINEL = process.env.SENTINEL_DEST ||
  path.join(HOME, "Wellness & Care Group of Texas Inc", "Corporate Archives Directory - Documents", "Sama Farooqui", "Sentinel");
const DATA = path.join(__dirname, "..", "data.json");
const APPLY = process.argv.includes("--apply");
const GREP = (process.argv.find(a => a.startsWith("--grep=")) || "").replace("--grep=", "");

// Date reading comes from lib/graph.js — the SAME function the live scan uses. A private copy
// here was wrong within minutes: it anchored the pattern with , and "_" is a word character,
// so "COLA_02_26_2028_CHER_Khan_M.pdf" had no boundary before the "02" and parsed as no date
// at all. The credential this whole exercise was about silently stayed on its old expiry.
const { dateFromName } = require("../lib/graph");

// Every live document under the Sentinel tree, keyed by the folder it sits in.
function indexDisk() {
  const byFolder = new Map();
  const walk = (dir, rel, depth) => {
    if (depth > 8) return;
    let entries;
    try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch (e) { return; }
    for (const e of entries) {
      if (e.isDirectory()) {
        if (FAC.isArchivedSegment(e.name)) continue;      // retired / superseded: not evidence
        walk(path.join(dir, e.name), rel ? rel + "/" + e.name : e.name, depth + 1);
      } else if (e.isFile() && /\.(pdf|jpg|jpeg|png|webp|tif|tiff|heic|doc|docx|xls|xlsx)$/i.test(e.name)) {
        if (!byFolder.has(rel)) byFolder.set(rel, []);
        byFolder.get(rel).push(e.name);
      }
    }
  };
  walk(SENTINEL, "", 0);
  return byFolder;
}

// data.json stores a full SharePoint URL; reduce it to the Sentinel-relative folder so it can
// be compared with what is on disk.
function relOf(url) {
  const u = decodeURIComponent(String(url || ""));
  const m = u.match(/\/Sentinel\/(.+)$/);
  return m ? m[1] : null;
}

function run() {
  if (!fs.existsSync(SENTINEL)) { console.error("Sentinel folder not found: " + SENTINEL); process.exit(2); }
  const data = JSON.parse(fs.readFileSync(DATA, "utf8"));
  const disk = indexDisk();
  const folders = Array.from(disk.keys()).sort((a, b) => b.length - a.length);

  // Group the tracked credentials by the folder they belong to, exactly as the scan's index does.
  const itemsByFolder = new Map();
  for (const it of data.items || []) {
    if (it.supplemental) continue;                        // supplementals ARE the document
    const rel = relOf(it.folderLink || it.fileLink);
    if (!rel) continue;
    if (!itemsByFolder.has(rel)) itemsByFolder.set(rel, []);
    itemsByFolder.get(rel).push(it);
  }

  const changes = [];
  for (const folder of folders) {
    // The credentials that own this folder: the longest item folder that prefixes it.
    let owner = null;
    for (const rel of Array.from(itemsByFolder.keys()).sort((a, b) => b.length - a.length)) {
      if (folder === rel || folder.startsWith(rel + "/")) { owner = rel; break; }
    }
    if (!owner) continue;
    const items = itemsByFolder.get(owner);
    for (const name of disk.get(folder)) {
      const hit = matchByRules(items, name);
      if (process.env.RD_DEBUG && new RegExp(process.env.RD_DEBUG, "i").test(name)) {
        console.log("DEBUG " + name);
        console.log("   folder = " + folder);
        console.log("   owner  = " + owner);
        console.log("   hit    = " + (hit ? hit.id + "  expires=" + hit.expires : "none"));
        console.log("   date   = " + dateFromName(name));
      }
      if (!hit) continue;
      const d = dateFromName(name);
      if (!d) continue;
      // Forward only. A newer document may arrive before an older one is tidied away, and a
      // credential must never be dragged backwards into looking expired by a stale copy.
      if (hit.expires && d <= hit.expires) continue;
      changes.push({ item: hit, from: hit.expires || "(none)", to: d, file: name, folder });
    }
  }

  // Keep only the furthest-forward date per credential.
  const best = new Map();
  for (const c of changes) {
    const cur = best.get(c.item.id);
    if (!cur || c.to > cur.to) best.set(c.item.id, c);
  }
  const final = Array.from(best.values());

  console.log("Refresh dashboard data from the documents on disk" + (APPLY ? "  [APPLY]" : "  [DRY RUN]"));
  console.log("  documents indexed : " + Array.from(disk.values()).reduce((a, b) => a + b.length, 0));
  console.log("  credentials moved forward : " + final.length);
  if (GREP) {
    const rx = new RegExp(GREP, "i");
    console.log("  --grep=" + GREP);
    final.filter(c => rx.test(c.item.category) || rx.test(c.file))
      .forEach(c => console.log("     " + c.item.entity + " / " + c.item.category + "   " + c.from + "  ->  " + c.to + "   (" + c.file + ")"));
  }
  final.slice(0, 25).forEach(c =>
    console.log("     " + (c.item.entity + " / " + c.item.category).padEnd(52).slice(0, 52) + c.from + "  ->  " + c.to));
  if (final.length > 25) console.log("     ...and " + (final.length - 25) + " more");

  if (!APPLY) { console.log("\n  Dry run - data.json not written. Re-run with --apply."); return; }

  const byId = new Map(final.map(c => [c.item.id, c]));
  let n = 0;
  for (const it of data.items || []) {
    const c = byId.get(it.id);
    if (!c) continue;
    it.expires = c.to;
    it.fileLink = "https://wcgtx.sharepoint.com/sites/CorporateArchivesDirectory/Shared%20Documents/Sama%20Farooqui/Sentinel/" +
      c.folder.split("/").map(encodeURIComponent).join("/") + "/" + encodeURIComponent(c.file);
    it.isFile = true;
    n++;
  }
  data.refreshedFromDisk = new Date().toISOString();
  fs.writeFileSync(DATA, JSON.stringify(data));
  console.log("\n  data.json updated - " + n + " credentials now carry the date on their current document.");
}

run();
