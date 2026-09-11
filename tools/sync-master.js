#!/usr/bin/env node
// Sentinel master-folder sync.
//
// WHY THIS RUNS ON THE PC AND NOT IN THE CLOUD
// The master folder is ~28,000 files / 30 GB, and it needs to be checked every 15 minutes.
// Vercel's Hobby plan allows one cron run a DAY and caps the project at 12 serverless
// functions (already at the ceiling). But both folders are OneDrive-synced onto this machine,
// so the whole job is local file I/O — no Graph API, no rate limits, no timeouts, nothing
// pushed through the cloud. OneDrive uploads the organized result to SharePoint on its own.
//
//   SOURCE (read only, never modified):
//     ...\Sama Farooqui - WCGTX Phyicians_04.08.2020\
//   DESTINATION (the tree the Sentinel dashboard reads):
//     ...\Corporate Archives Directory - Documents\Sama Farooqui\Sentinel\
//
// MIRROR, NOT PILE-UP. A document that has been superseded at the source does not linger in
// the destination: it is moved into "z.Superseded" beside where it was. Nothing is ever hard
// deleted — a compliance library should not lose a document because a script decided to.
//
//   node tools/sync-master.js              # dry run: prints the plan, writes nothing
//   node tools/sync-master.js --apply      # do it
//   node tools/sync-master.js --apply --only=facility,provider,staff
//   node tools/sync-master.js --apply --include-hr    # see HR_NOTE below
//
// HR_NOTE: folders named "HR ONLY" / "HR File" / "z.HR ONLY" are EXCLUDED by default. They
// hold WCGTX initial applications, verifications of employment and physician agreements —
// employment contracts with compensation terms. The destination is a shared team library that
// feeds a dashboard several people can open. Whoever named those folders "HR ONLY" meant it,
// so this script does not quietly widen who can read them. --include-hr overrides it.

const fs = require("fs");
const path = require("path");
const FACILITY = require("../lib/facility");

// ----------------------------------------------------------------------------- configuration
const HOME = process.env.USERPROFILE || process.env.HOME || "";
const SRC_ROOT = process.env.SENTINEL_MASTER ||
  path.join(HOME, "Wellness & Care Group of Texas Inc", "Sama Farooqui - WCGTX Phyicians_04.08.2020");
const DST_ROOT = process.env.SENTINEL_DEST ||
  path.join(HOME, "Wellness & Care Group of Texas Inc", "Corporate Archives Directory - Documents", "Sama Farooqui", "Sentinel");
const STATE_FILE = path.join(DST_ROOT, "_sync", "sync-state.json");
const LOG_FILE = path.join(DST_ROOT, "_sync", "sync-log.txt");

const ARGS = process.argv.slice(2);
const APPLY = ARGS.includes("--apply");
const INCLUDE_HR = ARGS.includes("--include-hr");
const QUIET = ARGS.includes("--quiet");
const ONLY = (ARGS.find(a => a.startsWith("--only=")) || "").replace("--only=", "").split(",").filter(Boolean);
const want = (fam) => !ONLY.length || ONLY.includes(fam);

// Which site folders in the master map to which dashboard facility.
const SITES = {
  "..Frisco ER": "Frisco",
  "..Castle Hills ER": "Castle Hills",
  "..Urgent Care Ennis": "Urgent Care Ennis",
  "..Urgent Care Plano": "Urgent Care Plano",
};
const PROVIDER_ROOT = "..WCGTX Master Physician File";
const SOP_PHASES = [
  "1. Application & Document Collection", "2. Primary Source Verification",
  "3. Background & Compliance Review", "4. Medical Staff Review",
  "5. Payer Enrollment & Facility Setup", "6. Approval & Ongoing Monitoring",
];

// The destination already uses these category names; near-duplicates from the source are
// folded in so the tree does not sprout "Licenses" beside "License".
const CATEGORY_ALIAS = {
  "z.health": "Health", "licenses": "License", "cme": "CMEs",
  "certificates": "Certifications", "miscellaneous": "Misc", "cois": "COI",
  "hr file": "HR File", "hr only": "HR ONLY", "z.hr only": "HR ONLY",
};
const HR_FOLDER = /^(z\.)?hr\b/i;
const TEMPLATE_FOLDER = /^\.{0,2}(folder template|z\.folder template)$/i;
const JUNK_FILE = /^(~\$|\.ds_store$|thumbs\.db$|desktop\.ini$)/i;

// ------------------------------------------------------------------------------------ helpers
const log = [];
function say(line) { log.push(line); if (!QUIET) console.log(line); }
// Filenames the master has explicitly retired into an archive folder, kept PER ENTITY.
// A global set was tried first and is unsafe: generic names ("Photo.pdf", "CV.pdf",
// "Vendor List.pdf") recur across people, so one provider archiving theirs would have retired
// everybody else's copy. The key is the destination folder the documents belong to.
const archivedNames = new Set();                 // fallback only
const archivedByOwner = new Map();               // owner key -> Set(filename lowercased)
function archiveSinkFor(ownerKey) {
  if (!archivedByOwner.has(ownerKey)) archivedByOwner.set(ownerKey, new Set());
  return archivedByOwner.get(ownerKey);
}
// "Provider/Jane Doe/..." -> "provider/jane doe"; "State Readiness/Frisco/..." -> that site;
// "Staff/Frisco/Doe, Jane_RN/..." -> that person.
function ownerOfRel(rel) {
  const seg = String(rel).split(/[\\/]+/).filter(Boolean);
  if (!seg.length) return "";
  if (seg[0] === "Staff") return seg.slice(0, 3).join("/").toLowerCase();
  return seg.slice(0, 2).join("/").toLowerCase();
}
function collectNames(dir, into, depth) {
  if (depth > 6) return;
  let entries;
  try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch (e) { return; }
  for (const e of entries) {
    if (e.isDirectory()) collectNames(path.join(dir, e.name), into, depth + 1);
    else if (e.isFile() && !JUNK_FILE.test(e.name)) into.add(e.name.toLowerCase());
  }
}
function walk(dir, rel, out, skipped, depth, archiveSink) {
  if (depth > 8) return;
  let entries;
  try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch (e) { return; }
  for (const e of entries) {
    const full = path.join(dir, e.name);
    const r = rel ? rel + "/" + e.name : e.name;
    if (e.isDirectory()) {
      if (FACILITY.isArchivedSegment(e.name)) {
        skipped.archive.push(r);
        // Note WHAT was archived, not just that a folder was skipped. When the coworker moves
        // a superseded certificate into "z.Expired Docs" they are making an explicit judgment
        // that it is finished — the strongest possible signal, and far safer than inferring
        // supersession from filename similarity. A destination copy of one of these names is
        // then retired too, which is how the old COLA (01_17_2026) gets out of the live
        // section once COLA_02_26_2028 arrives, even though the names are not a fuzzy match.
        collectNames(full, archiveSink || archivedNames, 0);
        continue;
      }
      if (TEMPLATE_FOLDER.test(e.name)) { skipped.template.push(r); continue; }
      if (HR_FOLDER.test(e.name) && !INCLUDE_HR) { skipped.hr.push(r); continue; }
      walk(full, r, out, skipped, depth + 1, archiveSink);
    } else if (e.isFile()) {
      if (JUNK_FILE.test(e.name)) continue;
      let st; try { st = fs.statSync(full); } catch (x) { continue; }
      out.push({ abs: full, rel: r, name: e.name, size: st.size, mtime: st.mtimeMs });
    }
  }
}
// "Abdul, Rahman M MD" -> "Rahman M Abdul"; matches the folder naming the dashboard already
// uses and the entityKey the roster builds ("abdul-rahman-m").
function providerFolderName(raw) {
  let n = String(raw).trim()
    .replace(/_credential(ing)?$/i, "")
    .replace(/\s*\b(MD|DO|PA|PA-?C|NP|FNP-?C|APRN|DDS|PHD)\b\.?\s*$/i, "")
    .trim();
  if (n.indexOf(",") >= 0) {
    const seg = n.split(",");
    const last = seg[0].trim(), first = seg.slice(1).join(",").trim();
    return (first + " " + last).replace(/\s+/g, " ").trim();
  }
  return n;
}
function cleanCategory(raw) {
  const k = String(raw).trim().toLowerCase();
  return CATEGORY_ALIAS[k] || String(raw).trim();
}
// Which SOP phase a provider document belongs in. Convention confirmed with Afia Umber:
// primary-source verifications sit in phase 2, OIG/NPDB/sanction work in phase 3, and
// everything else stays in phase 1 under its own category name.
function providerPhase(category, fileName) {
  const c = String(category).toLowerCase(), f = String(fileName).toLowerCase();
  if (c === "sanction checks" || /\boig\b|npdb|\bsam\b|exclusion/.test(f)) return 2;      // 3.
  if (/verif|primary source/.test(f)) return 1;                                            // 2.
  return 0;                                                                                // 1.
}

// Known providers — from the baked roster AND from the folders the dashboard already has. The
// master's provider root also holds non-people ("Zz..Agreements", "ZZZ Archive",
// "..Forms for Providers"), and guessing by folder-name shape would either sweep those in as
// fake providers or drop a real one. Matching against who actually exists is unambiguous.
const normKey = (s) => String(s).toLowerCase().replace(/[^a-z0-9]+/g, "");
// Credential suffixes appear stuck to the surname in the roster ("Nguyen DO"), which is why a
// plain surname comparison failed for several real providers.
const stripCred = (s) => String(s).replace(/\b(MD|DO|PA|PA-?C|NP|FNP-?C|APRN|DDS|PHD|RN)\b\.?/gi, "").trim();

// Who the dashboard knows. Three sources, because no single one is complete:
//   data.json          — the baked roster
//   Sentinel/Provider  — folders that already exist (prefer these names, so the sync never
//                        creates a second folder beside a provider who already has one)
//   the live roster    — the workbook is synced locally now, so it can be read directly
// Matching is exact-name first, then surname + first-name token, then surname + first initial.
// Measured on the real data: exact-name alone left 12 active providers unmatched, including
// "Couch, Chris" (roster says Christopher), "Zahid, Abdul" (roster says Abdul Rauf),
// "Nguyen DO, Lien" (credential glued to the surname) and "Mohiuddin, Mohammed Amer" (first
// and last reversed in the roster). Their documents would simply never have been filed.
function buildProviderIndex() {
  const exact = new Map(), loose = new Map(), initial = new Map();
  const put = (name, canonical, strong) => {
    const n = stripCred(name).trim(); if (!n) return;
    const toks = n.split(/\s+/).filter(Boolean); if (!toks.length) return;
    const last = toks[toks.length - 1], first = toks[0];
    exact.set(normKey(n), canonical);
    const lk = normKey(last + first);
    if (strong || !loose.has(lk)) loose.set(lk, canonical);
    const ik = normKey(last) + "|" + normKey(first).slice(0, 1);
    if (strong || !initial.has(ik)) initial.set(ik, canonical);
  };
  // Weakest first, strongest last, so a real folder name wins the key.
  try {
    const d = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "data.json"), "utf8"));
    (d.items || []).forEach(i => { if (i.scope === "provider" && i.entity) put(i.entity, i.entity, false); });
  } catch (e) {}
  try {
    const ExcelJS = require("exceljs");
    const wbPath = path.join(SRC_ROOT, "..WCGTX Master Rosters", "WCGTX Physician Roster.xlsx");
    if (fs.existsSync(wbPath)) {
      const wb = new ExcelJS.Workbook();
      // Synchronous read is fine here: this is a local one-off at startup, not a server path.
      const buf = fs.readFileSync(wbPath);
      const XL = require("../lib/excel");
      const done = wb.xlsx.load(buf);
      done.then(() => {}).catch(() => {});
      providerIndexPending = done.then(() => {
        for (const ws of wb.worksheets) {
          let cols; try { cols = XL.detectNameColsWs(ws); } catch (e) { continue; }
          if (!cols || cols.lastCol < 1 || cols.firstCol < 1) continue;
          ws.eachRow((row, rn) => {
            if (rn === 1) return;
            const cv = (c) => { const v = row.getCell(c).value; return String(v && v.text ? v.text : (v == null ? "" : v)).trim(); };
            const last = stripCred(cv(cols.lastCol)), first = stripCred(cv(cols.firstCol));
            if (!last || last.length > 40 || /^(last|name)$/i.test(last)) return;
            const display = (first ? first + " " + last : last).replace(/\s+/g, " ").trim();
            put(display, display, false);
          });
        }
      }).catch(() => {});
    }
  } catch (e) {}
  safeDirs(path.join(DST_ROOT, "Provider")).forEach(n => put(n, n, true));
  return { exact, loose, initial };
}
let providerIndexPending = Promise.resolve();
const PIDX = buildProviderIndex();
function resolveProvider(folderName) {
  const n = stripCred(providerFolderName(folderName)).trim();
  if (!n) return null;
  const hit = PIDX.exact.get(normKey(n));
  if (hit) return hit;
  const toks = n.split(/\s+/).filter(Boolean);
  if (toks.length < 2) return null;
  const last = toks[toks.length - 1], first = toks[0];
  return PIDX.loose.get(normKey(last + first))
      || PIDX.initial.get(normKey(last) + "|" + normKey(first).slice(0, 1))
      || null;
}

// Strip dates, version numbers and trailing initials so two filings of the SAME document line
// up: "COLA_01_17_2026_CHER.pdf" and "COLA_02_26_2028_CHER_Khan_M.pdf" both key to "cola cher".
// This is what makes "replace the old one with the new one" mean a real replacement rather
// than "delete anything I did not happen to see in the master this run".
function docKey(name) {
  return String(name)
    .replace(/\.[^.]+$/, "")
    .replace(/\b\d{1,4}[_.\-]\d{1,2}[_.\-]\d{2,4}\b/g, " ")   // 01_17_2026 / 2026-01-17
    .replace(/\b(19|20)\d{2}\b/g, " ")                         // bare years
    .replace(/\b\d+\b/g, " ")
    .replace(/[_\-.,]+/g, " ")
    .replace(/\b[a-z]\b/gi, " ")                               // trailing single initials
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}
function dateFromName(name) {
  let best = null, m;
  const re = /\b(\d{1,4})[_.\-](\d{1,2})[_.\-](\d{2,4})\b/g;
  while ((m = re.exec(String(name)))) {
    let a = +m[1], b = +m[2], c = +m[3];
    let y, mo, d;
    if (a > 31) { y = a; mo = b; d = c; } else { mo = a; d = b; y = c < 100 ? 2000 + c : c; }
    if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
      const iso = y + "-" + String(mo).padStart(2, "0") + "-" + String(d).padStart(2, "0");
      if (!best || iso > best) best = iso;
    }
  }
  return best;
}

// ------------------------------------------------------------------------- build the file plan
const plan = [];      // { src, dst, family, reason }
const skipped = { archive: [], template: [], hr: [], notAProvider: [] };
const counts = {};
const bump = (k) => { counts[k] = (counts[k] || 0) + 1; };

function addSite(siteDir, facility) {
  // --- facility compliance -> Sentinel/State Readiness/<Facility>/<NN. Section>/
  if (want("facility")) {
    for (const sub of safeDirs(path.join(SRC_ROOT, siteDir))) {
      if (!/^\.{0,2}state requirements/i.test(sub)) continue;
      const files = [];
      walk(path.join(SRC_ROOT, siteDir, sub), "", files, skipped, 0,
        archiveSinkFor(("State Readiness/" + facility).toLowerCase()));
      for (const f of files) {
        const i = f.rel.lastIndexOf("/");
        const c = FACILITY.classifySection(i < 0 ? "" : f.rel.slice(0, i), f.name);
        const section = c ? c.section : FACILITY.STATE_SECTIONS[0];
        plan.push({ src: f, family: "facility",
          dst: path.join(DST_ROOT, "State Readiness", facility, section, f.name) });
        bump("facility");
      }
    }
  }
  // --- staff -> Sentinel/Staff/<Facility>/<Person>/<Category>/
  if (want("staff")) {
    for (const group of safeDirs(path.join(SRC_ROOT, siteDir))) {
      if (/^\.{0,2}state requirements/i.test(group)) continue;
      if (FACILITY.isArchivedSegment(group) || TEMPLATE_FOLDER.test(group)) continue;
      for (const person of safeDirs(path.join(SRC_ROOT, siteDir, group))) {
        if (FACILITY.isArchivedSegment(person) || TEMPLATE_FOLDER.test(person)) continue;
        const files = [];
        walk(path.join(SRC_ROOT, siteDir, group, person), "", files, skipped, 0,
          archiveSinkFor(("Staff/" + facility + "/" + person.trim()).toLowerCase()));
        for (const f of files) {
          const i = f.rel.lastIndexOf("/");
          const segs = i < 0 ? [] : f.rel.slice(0, i).split("/");
          const cat = cleanCategory(segs.length ? segs[0] : "Misc");
          // Keep any deeper structure the source has. Flattening to the category alone made two
          // documents that share a name under different sub-folders map to one destination:
          // each run copied both, the loser always differed, and the sync never settled.
          plan.push({ src: f, family: "staff", deeper: segs.slice(1),
            base: path.join(DST_ROOT, "Staff", facility, person.trim(), cat),
            dst: path.join(DST_ROOT, "Staff", facility, person.trim(), cat, f.name) });
          bump("staff");
        }
      }
    }
  }
}
function safeDirs(p) {
  try { return fs.readdirSync(p, { withFileTypes: true }).filter(d => d.isDirectory()).map(d => d.name); }
  catch (e) { return []; }
}

function addProviders() {
  if (!want("provider")) return;
  for (const folder of safeDirs(path.join(SRC_ROOT, PROVIDER_ROOT))) {
    // "..Forms for Providers", "..OnBoarding_New Physicians", ".HOSPITAL_EHMC Specialists" are
    // reference material, not a person. The master marks them with a leading dot.
    if (folder.startsWith(".")) { skipped.template.push(PROVIDER_ROOT + "/" + folder); continue; }
    if (FACILITY.isArchivedSegment(folder)) { skipped.archive.push(PROVIDER_ROOT + "/" + folder); continue; }
    // Only file documents for somebody the dashboard or the live roster actually knows. An
    // unrecognised folder is reported, never invented as a new provider — a misspelt folder
    // silently becoming a second "Mohammad Khan" beside the real one is how a roster forks.
    const person = resolveProvider(folder);
    if (!person) { skipped.notAProvider.push(folder + "  (reads as: " + providerFolderName(folder) + ")"); continue; }
    const files = [];
    walk(path.join(SRC_ROOT, PROVIDER_ROOT, folder), "", files, skipped, 0,
      archiveSinkFor(("Provider/" + person).toLowerCase()));
    for (const f of files) {
      const i = f.rel.lastIndexOf("/");
      const segs = i < 0 ? [] : f.rel.slice(0, i).split("/");
      const rawCat = segs.length ? segs[0] : "Misc";
      const cat = cleanCategory(rawCat);
      const phase = SOP_PHASES[providerPhase(cat, f.name)];
      plan.push({ src: f, family: "provider", deeper: segs.slice(1),
        base: path.join(DST_ROOT, "Provider", person, phase, cat),
        dst: path.join(DST_ROOT, "Provider", person, phase, cat, f.name) });
      bump("provider");
    }
  }
}

// ------------------------------------------------------------- decide copy / replace / skip
function loadState() {
  try { return JSON.parse(fs.readFileSync(STATE_FILE, "utf8")); } catch (e) { return { files: {} }; }
}
function sameFile(dst, src) {
  let st; try { st = fs.statSync(dst); } catch (e) { return false; }
  // Size plus a two-second mtime tolerance. OneDrive and FAT-derived filesystems round mtimes,
  // so an exact comparison would re-copy the whole library on every run.
  return st.size === src.size && Math.abs(st.mtimeMs - src.mtime) < 2000;
}
function ensureDir(d) { fs.mkdirSync(d, { recursive: true }); }

async function run() {
  const t0 = Date.now();
  // The live roster loads asynchronously; the provider index is incomplete until it lands.
  await providerIndexPending;
  say("Sentinel master sync  " + new Date().toISOString() + (APPLY ? "  [APPLY]" : "  [DRY RUN — nothing will be written]"));
  say("  source      " + SRC_ROOT);
  say("  destination " + DST_ROOT);
  if (!fs.existsSync(SRC_ROOT)) { say("  !! source folder not found — is OneDrive synced?"); process.exit(2); }
  if (!fs.existsSync(DST_ROOT)) { say("  !! destination folder not found — is the team library synced?"); process.exit(2); }

  for (const dir in SITES) addSite(dir, SITES[dir]);
  addProviders();

  // Two source documents can carry the same name under different sub-folders of one category.
  // Filed flat they fight over a single destination: every run copies both, the loser always
  // differs, and the sync never settles (measured: 60 files churning forever). Only the
  // colliding ones get their sub-folder back, so the tree stays flat everywhere else.
  const remapOrphans = [];
  {
    const byDst = new Map();
    plan.forEach(p => { const k = p.dst.toLowerCase(); (byDst.get(k) || byDst.set(k, []).get(k)).push(p); });
    let fixed = 0;
    for (const [, group] of byDst) {
      if (group.length < 2) continue;
      for (const p of group) {
        if (!p.base || !p.deeper || !p.deeper.length) continue;
        // An earlier run may already have filed this flat. That copy is now in the wrong place,
        // so retire it rather than leaving a stale duplicate beside the correct one.
        if (fs.existsSync(p.dst)) remapOrphans.push(p.dst);
        p.dst = path.join(p.base, ...p.deeper, p.src.name);
        fixed++;
      }
    }
    if (fixed) say("  (" + fixed + " documents kept their sub-folder to avoid a name clash)");
  }

  const toCopy = [], toReplace = [];
  for (const p of plan) {
    if (!fs.existsSync(p.dst)) { toCopy.push(p); continue; }
    if (!sameFile(p.dst, p.src)) toReplace.push(p);
  }

  // ---- what counts as "superseded" -------------------------------------------------------
  // NOT "absent from the master". The master physician file holds 98 provider folders while
  // the dashboard tracks 182 — 100 of those people have no folder in the master at all, and
  // their documents were gathered from elsewhere. A first draft of this treated absence as
  // supersession and proposed archiving 6,919 live credentials, including ACLS, ATLS and BLS
  // certificates. Absence from one folder is not evidence that a document is obsolete.
  //
  // A document is superseded only when the master brings in a NEWER FILING OF THE SAME
  // DOCUMENT: same destination folder, same docKey, later date in the name (or a later
  // modified time when neither carries a date). That is what "replace the old one with the
  // new one" actually means, and it cannot touch anything the master has no opinion about.
  const superseded = [];
  const plannedDst = new Set(plan.map(p => p.dst.toLowerCase()));
  // Every member of a colliding group nominates the same flat path, so this list repeats.
  // Without the dedupe the first move succeeds and the rest report ENOENT on a file that has
  // already been filed correctly — 60 alarming "failures" that were nothing of the kind.
  for (const abs of Array.from(new Set(remapOrphans))) {
    if (plannedDst.has(abs.toLowerCase())) continue;
    superseded.push({ abs, name: path.basename(abs), rel: path.relative(DST_ROOT, abs),
      by: "refiled into its sub-folder to resolve a name clash" });
  }
  // (a) The master itself retired this exact filename into an archive folder.
  for (const family of ["Provider", "State Readiness", "Staff"]) {
    const root = path.join(DST_ROOT, family);
    if (!fs.existsSync(root)) continue;
    const have = [];
    walk(root, "", have, { archive: [], template: [], hr: [] }, 0);
    for (const f of have) {
      if (f.rel.indexOf("z.Superseded") >= 0) continue;
      // A filename can sit in BOTH a live folder and an archive folder in the master (the same
      // document filed twice). Retiring one this run is also filing in meant copy -> retire ->
      // copy, 14 files churning on every single run and never settling.
      if (plannedDst.has(f.abs.toLowerCase())) continue;
      const ownerSet = archivedByOwner.get(ownerOfRel(path.relative(DST_ROOT, f.abs)));
      if (!ownerSet || !ownerSet.has(f.name.toLowerCase())) continue;
      superseded.push({ abs: f.abs, name: f.name, rel: path.relative(DST_ROOT, f.abs),
        by: "the master moved it to an archive folder" });
    }
  }
  // (b) A newer filing of the same document arrived.
  const incomingByDir = {};
  for (const p of toCopy.concat(toReplace)) {
    const dir = path.dirname(p.dst);
    (incomingByDir[dir] = incomingByDir[dir] || []).push(p);
  }
  for (const dir in incomingByDir) {
    let existing = [];
    try { existing = fs.readdirSync(dir, { withFileTypes: true }).filter(d => d.isFile()).map(d => d.name); }
    catch (e) { continue; }                       // folder does not exist yet — nothing to supersede
    for (const p of incomingByDir[dir]) {
      const key = docKey(p.src.name);
      if (!key) continue;
      const newDate = dateFromName(p.src.name);
      for (const old of existing) {
        if (old.toLowerCase() === p.src.name.toLowerCase()) continue;   // same file, handled as a replace
        if (docKey(old) !== key) continue;
        const oldDate = dateFromName(old);
        let isOlder;
        // A date in the filename is the only evidence trusted here. Modified-time was tried and
        // dropped: it would have archived "Photo_Panzarella_R.pdf" because a similarly named
        // file happened to be saved later, and a file's timestamp says when someone touched it,
        // not which version of a credential it is. When neither filing carries a date, both are
        // kept and a human can decide.
        if (newDate && oldDate) isOlder = oldDate < newDate;
        else if (newDate && !oldDate) isOlder = true;                    // undated old, dated new
        else isOlder = false;
        if (!isOlder) continue;
        const abs = path.join(dir, old);
        if (superseded.some(s => s.abs === abs)) continue;
        superseded.push({ abs, name: old, rel: path.relative(DST_ROOT, abs), by: p.src.name });
      }
    }
  }

  say("");
  say("  PLAN");
  say("    new documents to copy in .............. " + toCopy.length);
  say("    updated documents to replace .......... " + toReplace.length);
  say("    already current, untouched ............ " + (plan.length - toCopy.length - toReplace.length));
  say("    replaced by a newer filing -> z.Superseded  " + superseded.length);
  say("    ----");
  Object.keys(counts).sort().forEach(k => say("    " + k.padEnd(10) + " documents seen in master: " + counts[k]));
  say("    skipped: " + skipped.archive.length + " archive folders, " + skipped.template.length +
      " template/reference folders, " + skipped.hr.length + " HR-restricted folders" +
      (INCLUDE_HR ? " (INCLUDED by --include-hr)" : " (use --include-hr to include)") +
      ", " + skipped.notAProvider.length + " folders that are not a known provider");
  if (skipped.notAProvider.length) {
    say("    not recognised as a provider (nothing filed for these):");
    skipped.notAProvider.slice(0, 20).forEach(n => say("       " + n));
    if (skipped.notAProvider.length > 20) say("       ...and " + (skipped.notAProvider.length - 20) + " more");
  }

  const FOCUS = (ARGS.find(a => a.startsWith("--grep=")) || "").replace("--grep=", "");
  if (FOCUS) {
    const rx = new RegExp(FOCUS, "i");
    say("");
    say("  --grep=" + FOCUS);
    toCopy.filter(p => rx.test(p.src.name)).forEach(p => say("    COPY IN    " + path.relative(DST_ROOT, p.dst)));
    toReplace.filter(p => rx.test(p.src.name)).forEach(p => say("    REPLACE    " + path.relative(DST_ROOT, p.dst)));
    superseded.filter(f => rx.test(f.name)).forEach(f => say("    RETIRE     " + f.rel + "   (" + f.by + ")"));
  }
  if (!APPLY) {
    say("");
    say("  Sample of what would be filed:");
    toCopy.slice(0, 12).forEach(p => say("    " + p.family.padEnd(9) + path.relative(DST_ROOT, p.dst)));
    if (toCopy.length > 12) say("    …and " + (toCopy.length - 12) + " more");
    if (superseded.length) {
      say("");
      say("  Sample of what would move to z.Superseded (each replaced by a newer filing):");
      superseded.slice(0, 8).forEach(f => { say("    " + f.rel); say("         replaced by: " + f.by); });
      if (superseded.length > 8) say("    …and " + (superseded.length - 8) + " more");
    }
    say("");
    say("  Dry run — nothing written. Re-run with --apply to do it.");
    writeLog();
    return;
  }

  let copied = 0, replaced = 0, moved = 0, failed = 0;
  for (const p of toCopy.concat(toReplace)) {
    try {
      ensureDir(path.dirname(p.dst));
      fs.copyFileSync(p.src.abs, p.dst);
      fs.utimesSync(p.dst, new Date(), new Date(p.src.mtime));   // keep mtime so the next run skips it
      if (fs.existsSync(p.dst)) (toCopy.includes(p) ? copied++ : replaced++);
    } catch (e) { failed++; say("    !! " + p.src.name + " — " + e.message); }
  }
  const seenMove = new Set();
  for (const f of superseded) {
    try {
      if (seenMove.has(f.abs.toLowerCase())) continue;
      seenMove.add(f.abs.toLowerCase());
      if (!fs.existsSync(f.abs)) continue;          // already moved earlier in this same run
      const dir = path.dirname(f.abs);
      const dest = path.join(dir, "z.Superseded", path.basename(f.abs));
      ensureDir(path.dirname(dest));
      fs.renameSync(f.abs, dest);
      moved++;
    } catch (e) { failed++; say("    !! could not archive " + f.name + " — " + e.message); }
  }

  const state = loadState();
  state.lastRun = new Date().toISOString();
  state.lastResult = { copied, replaced, moved, failed, seen: plan.length };
  try { ensureDir(path.dirname(STATE_FILE)); fs.writeFileSync(STATE_FILE, JSON.stringify(state, null, 2)); } catch (e) {}

  say("");
  say("  DONE in " + Math.round((Date.now() - t0) / 1000) + "s — " + copied + " copied, " +
      replaced + " replaced, " + moved + " superseded, " + failed + " failed.");
  say("  OneDrive will upload these to SharePoint; the dashboard picks them up on its next scan.");
  writeLog();
}

function writeLog() {
  try {
    ensureDir(path.dirname(LOG_FILE));
    fs.appendFileSync(LOG_FILE, log.join("\n") + "\n\n");
  } catch (e) {}
}

run().catch(e => { console.error("sync failed: " + (e && e.message ? e.message : e)); process.exit(1); });
