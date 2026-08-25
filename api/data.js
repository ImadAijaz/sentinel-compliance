// The roster — served ONLY to a signed-in staff member, FILTERED to the tabs they're allowed.
// Also: ?file=<url>  -> same-origin bytes proxy (for in-browser tools)
//       ?ocr=<url>   -> server-side OCR (OCR.space, free) that returns any dates it reads.
const { getSession } = require("../lib/session");
const data = require("../data.json");

function extractDates(text) {
  const out = [];
  const push = (y, mo, d) => { if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) out.push(y + "-" + String(mo).padStart(2, "0") + "-" + String(d).padStart(2, "0")); };
  let m;
  const re = /\b(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})\b/g;
  while ((m = re.exec(text))) { let y = +m[3]; if (y < 100) y += 2000; push(y, +m[1], +m[2]); }
  const mon = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  const re2 = /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+(\d{1,2}),?\s+(\d{4})\b/gi;
  while ((m = re2.exec(text))) push(+m[3], mon[m[1].toLowerCase().slice(0, 3)], +m[2]);
  return [...new Set(out)].sort();
}
async function resolveDownloadUrl(token, GRAPH, fileUrl) {
  const shareId = "u!" + Buffer.from(fileUrl, "utf8").toString("base64").replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  const di = await fetch(GRAPH + "/shares/" + shareId + "/driveItem", { headers: { Authorization: "Bearer " + token } });
  if (!di.ok) return { error: "resolve " + di.status, detail: (await di.text()).slice(0, 160) };
  const item = await di.json();
  const dl = item["@microsoft.graph.downloadUrl"] || item["@content.downloadUrl"];
  if (!dl) return { error: "no download url" };
  return { dl, item };
}

module.exports = async (req, res) => {
  // Never let the edge/CDN cache a roster or trash response — a stale empty
  // ?roster=trash read was making the Recycle bin look empty after a delete.
  res.setHeader("Cache-Control", "no-store, max-age=0");
  const url = new URL(req.url, "http://localhost");
  // Vercel crons (no user session) are allowed to hit ?regen=1 only. Verified with the
  // CRON_SECRET bearer token — the old check tested the User-Agent string, which any caller can
  // set, so `curl -A vercel-cron` was treated as the scheduler and could rewrite the cache
  // unauthenticated and unthrottled.
  const isCron = require("../lib/session").isCronRequest(req);
  const s = getSession(req);
  if (!s && !(isCron && url.searchParams.get("regen") === "1")) { res.status(401).json({ error: "sign-in required" }); return; }

  // ---- hourly cron regen: pull the live Excel + write roster_delta.json to OneDrive.
  //      /api/data merges the delta on every read, so changes propagate within seconds
  //      of the cron firing (not a full Python regen — those need data-entry tools we
  //      can't run server-side without Python). Schedule lives in vercel.json. ----
  if (url.searchParams.get("regen") === "1") {
    // ANY signed-in user may trigger a refresh: this only writes the app's own cache files
    // (roster_delta.json / staff_delta.json) and never touches the user's Excel. That's what
    // makes "open the dashboard = fresh data" work for staff, not just admins.
    // A short global throttle stops a room full of people opening the app at 8am from
    // stampeding Graph with workbook downloads — the first open does the work, the rest reuse
    // it. Admins can bypass with &force=1 (that's what the manual "Sync from Excel" sends).
    if (!isCron && !s) { res.status(401).json({ error: "sign-in required" }); return; }
    const force = url.searchParams.get("force") === "1" && s && s.admin;
    try {
      const xl = require("../lib/excel");
      const { accessToken, drivePath, writeJsonAt, readJsonAt } = require("../lib/graph");
      const token = await accessToken();
      if (!force && !isCron) {
        let prev = null;
        try { prev = await readJsonAt(token, drivePath("_Sentinel/roster_delta.json")); } catch (e) { prev = null; }
        const stamp = prev && prev.generatedAt ? new Date(prev.generatedAt).getTime() : NaN;
        const age = isNaN(stamp) ? Infinity : Date.now() - stamp;
        // Only skip on a *recent* stamp. A future-dated or unparseable stamp falls through and regenerates.
        if (age >= 0 && age < 60000) { res.status(200).json({ ok: true, skipped: "fresh", ageMs: age }); return; }
      }
      // Shared implementation (lib/regen.js) — the daily cron in api/digest.js calls the SAME
      // function, so the two can never drift apart again. It refuses to write a delta that would
      // wipe the board, and reports that refusal instead of silently succeeding.
      const { regenerateRoster } = require("../lib/regen");
      const r = await regenerateRoster(token);
      if (r.refused) { res.status(200).json({ ok: false, refused: true, error: r.reason, providersRead: r.providersRead }); return; }
      res.status(200).json({ ok: true, delta: { newProviders: r.newProviders, inactivated: r.inactivated, removed: r.removed, dates: r.dates, staff: r.staff, providersRead: r.providersRead } });
    } catch (e) { res.status(500).json({ error: String(e.message || e) }); }
    return;
  }

  // ---- SELF-TEST (admin only): prove the whole Excel + OneDrive round trip actually works. ----
  //   ?selftest=diagnose  READ-ONLY. Reports the live roster's real column layout, whether the
  //                       column-shift fix is engaging, provider counts, sample parsed names,
  //                       delta health, and whether both drives are reachable. Changes nothing.
  //   ?selftest=full      Everything above, then a REAL round trip: add a clearly-marked test
  //                       provider ("Sentinel ZZSelftest") -> read it back out of Excel ->
  //                       create its OneDrive folder + 6 SOP subfolders -> verify they exist ->
  //                       confirm it would appear on the dashboard -> DELETE the row and the
  //                       folder -> verify both are gone. Self-cleaning: it removes everything
  //                       it created, and it snapshots the workbook first. Uses a fixed test
  //                       name (no timestamp) so a half-finished run is cleaned up by the next.
  // ---- Read contact details out of a provider's own documents --------------------------------
  // POST /api/data?contact=<entityKey>   (admin)
  // The roster carries no phone number anywhere, and only ~half the providers have an email on
  // the COI sheet, so the remaining details only exist inside their CV / application PDFs.
  // Text is extracted IN THIS FUNCTION with pdf-parse — deliberately NOT via the third-party OCR
  // service, because these are personnel documents and that service is on a shared public key.
  // Scanned (image-only) PDFs yield no text; that is reported rather than guessed at.
  const contactKey = url.searchParams.get("contact");
  if (contactKey) {
    if (!s || !s.admin) { res.status(403).json({ error: "admins only" }); return; }
    try {
      const G = require("../lib/graph");
      const { applyRosterDelta } = require("../lib/delta");
      const all = await applyRosterDelta((data.items || []).filter(i => i.scope === "provider"));
      const mine = all.filter(i => i.entityKey === contactKey);
      if (!mine.length) { res.status(404).json({ error: "unknown provider" }); return; }

      // Prefer the documents most likely to carry contact details, best first.
      const fname = it => { try { return decodeURIComponent(String(it.fileLink || "").split("/").pop() || ""); } catch (e) { return ""; } };
      const rank = it => {
        const t = ((it.category || "") + " " + fname(it)).toLowerCase();
        if (/\bcv\b|resume|curriculum/.test(t)) return 0;
        if (/initial app|application/.test(t)) return 1;
        if (/tsca|peer|reference/.test(t)) return 2;
        return 3;
      };
      const cands = mine.filter(i => i.isFile && /\.pdf(\?|#|$)/i.test(String(i.fileLink || "")))
        .sort((a, b) => rank(a) - rank(b)).slice(0, 4);
      if (!cands.length) { res.status(200).json({ ok: false, message: "No PDF documents on file to read." }); return; }

      const token = await G.accessToken();
      // Require the inner module, NOT the package entry point. pdf-parse's index.js runs a debug
      // block when `module.parent` is falsy, which reads a test PDF relative to the working
      // directory — in a bundled serverless function that throws at require time and takes the
      // whole endpoint down. lib/pdf-parse.js is the actual parser with no such side effect.
      // Loaded lazily and guarded: if the PDF library is unavailable in the deployed bundle this
      // one feature must degrade with a clear message, never break /api/data (which serves the
      // whole dashboard).
      let pdf;
      try { pdf = require("pdf-parse/lib/pdf-parse.js"); }
      catch (le) { res.status(200).json({ ok: false, message: "The PDF reader is not available in this deployment, so documents can't be read here." }); return; }
      // Measured against the real roster: 76% of CVs are readable text, and of those 84% name a
      // specialty and 28% carry a licence number — so both are worth pulling. NPI/DEA appear in
      // only ~4% of CVs (they live in the registry/certificate documents instead), but cost
      // nothing to look for. Everything is returned as CANDIDATES for a human to pick: a CV lists
      // training rotations and past employers, so the first specialty mentioned is often not the
      // provider's own.
      const found = { emails: [], phones: [], degree: "", licenses: [], specialties: [], npi: [], dea: [], from: [], scanned: [] };
      for (const it of cands) {
        try {
          const shareId = "u!" + Buffer.from(it.fileLink, "utf8").toString("base64").replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
          const di = await fetch(G.GRAPH + "/shares/" + shareId + "/driveItem", { headers: { Authorization: "Bearer " + token } });
          if (!di.ok) continue;
          const item = await di.json();
          const dl = item["@microsoft.graph.downloadUrl"] || item["@content.downloadUrl"];
          if (!dl) continue;
          const fr = await fetch(dl);
          if (!fr.ok) continue;
          const buf = Buffer.from(await fr.arrayBuffer());
          if (buf.length > 12 * 1024 * 1024) continue;                 // skip very large scans
          const parsed = await pdf(buf).catch(() => null);
          const text = (parsed && parsed.text) || "";
          const label = fname(it) || it.category;
          if (text.replace(/\s/g, "").length < 40) { found.scanned.push(label); continue; }
          found.from.push(label);
          (text.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g) || []).forEach(e => {
            const v = e.toLowerCase().replace(/[.,;:]+$/, "");
            // Ignore the practice's own addresses — we want the provider's.
            if (!/wcgtx\.com$|example\.|noreply|no-reply/.test(v) && found.emails.indexOf(v) < 0) found.emails.push(v);
          });
          // US phone numbers, tolerant of (214) 555-1234 / 214.555.1234 / +1 214 555 1234.
          (text.match(/(?:\+?1[\s.-]?)?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}\b/g) || []).forEach(raw => {
            const d = raw.replace(/\D/g, "").replace(/^1(?=\d{10}$)/, "");
            if (d.length !== 10) return;
            if (/^(\d)\1{9}$/.test(d)) return;                         // 0000000000 etc
            if (/^(19|20)\d{8}$/.test(d)) return;                      // date-like runs
            const pretty = "(" + d.slice(0, 3) + ") " + d.slice(3, 6) + "-" + d.slice(6);
            if (found.phones.indexOf(pretty) < 0) found.phones.push(pretty);
          });
          if (!found.degree) {
            const m = text.match(/\b(M\.?D\.?|D\.?O\.?|P\.?A\.?-?C?|N\.?P\.?|MBBS)\b/);
            if (m) found.degree = m[1].replace(/\./g, "").toUpperCase();
          }
          const push = (arr, v, cap) => { const s = String(v || "").trim(); if (s && arr.indexOf(s) < 0 && arr.length < (cap || 6)) arr.push(s); };
          // Texas Medical Board licence numbers look like L3659 / K6266 — a letter or two then
          // 4-6 digits. Only take them when the text actually labels them as a licence, otherwise
          // any old reference number would qualify.
          const licRe = /\b(?:licen[cs]e|lic\.?)\s*(?:no\.?|number|#|:)?\s*([A-Z]{1,2}\d{4,6})\b/ig;
          let lm; while ((lm = licRe.exec(text))) push(found.licenses, lm[1].toUpperCase());
          const specRe = /\b(emergency medicine|internal medicine|family (?:medicine|practice)|pediatrics|general surgery|anesthesiolog\w+|psychiatry|radiolog\w+|urgent care|obstetrics|orthopaedics|orthopedics)\b/ig;
          let sm; while ((sm = specRe.exec(text))) push(found.specialties, sm[1].replace(/\b\w/g, c => c.toUpperCase()));
          const npiRe = /\bNPI\b[^\n]{0,25}?\b(\d{10})\b/ig;
          let nm; while ((nm = npiRe.exec(text))) push(found.npi, nm[1], 3);
          const deaRe = /\bDEA\b[^\n]{0,25}?\b([A-Z]{2}\d{7})\b/ig;
          let dm; while ((dm = deaRe.exec(text))) push(found.dea, dm[1].toUpperCase(), 3);
        } catch (fe) { /* one unreadable document must not stop the rest */ }
      }
      res.status(200).json({
        ok: true, entityKey: contactKey,
        emails: found.emails.slice(0, 5), phones: found.phones.slice(0, 5), degree: found.degree || "",
        licenses: found.licenses, specialties: found.specialties, npi: found.npi, dea: found.dea,
        readFrom: found.from, unreadable: found.scanned,
        note: found.from.length
          ? "Read from the documents listed. Check these before relying on them — a CV lists training rotations and past employers, so the first specialty or number mentioned is not always the provider's own."
          : "These documents are scanned images (pictures of paper), so there is no text in them to read. Reading details out of them would need OCR, which this does not do.",
      });
      return;
    } catch (e) { res.status(500).json({ error: String(e.message || e) }); return; }
  }

  // ---- Import existing documents into a provider's Sentinel folder ---------------------------
  // Providers often already have a document folder somewhere else on the drive, outside the
  // Sentinel/Provider tree the dashboard reads. This copies those files in, so they get picked
  // up, without anyone hand-dragging them in OneDrive.
  //   GET  ?import=preview&src=<folder link>&entityKey=<key>   -> what WOULD be copied
  //   POST ?import=run&src=<folder link>&entityKey=<key>&phase=<0-5>
  // Always a COPY — never a move or delete. The originals are left exactly where they are.
  const imp = url.searchParams.get("import");
  if (imp) {
    if (!s || !s.admin) { res.status(403).json({ error: "admins only" }); return; }
    const PHASES = [
      "1. Application & Document Collection", "2. Primary Source Verification",
      "3. Background & Compliance Review", "4. Medical Staff Review",
      "5. Payer Enrollment & Facility Setup", "6. Approval & Ongoing Monitoring",
    ];
    try {
      const G = require("../lib/graph");
      const token = await G.accessToken();
      const src = String(url.searchParams.get("src") || "").trim();
      const entityKey = String(url.searchParams.get("entityKey") || "").trim();
      let provName = String(url.searchParams.get("name") || "").trim();
      const phaseIdx = Math.min(5, Math.max(0, parseInt(url.searchParams.get("phase") || "0", 10) || 0));
      if (!src) { res.status(400).json({ error: "paste the folder's OneDrive/SharePoint link" }); return; }

      // Work out the provider's display name (folder names use "First Last").
      if (!provName && entityKey) {
        const { applyRosterDelta } = require("../lib/delta");
        const all = await applyRosterDelta((data.items || []).filter(i => i.scope === "provider"));
        const hit = all.find(i => i.entityKey === entityKey);
        if (hit) provName = hit.entity;
      }
      if (!provName) { res.status(400).json({ error: "could not work out which provider to import into" }); return; }
      // Folder names come from spreadsheet content, so strip anything path-significant.
      const safeName = provName.replace(/[\/\\:*?"<>|]/g, "").replace(/\s+/g, " ").trim();

      // Resolve the source. Accepts a share link ("...:f:/p/...") or a plain SharePoint URL.
      const shareId = "u!" + Buffer.from(src, "utf8").toString("base64").replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
      const sr = await fetch(G.GRAPH + "/shares/" + shareId + "/driveItem?$select=id,name,folder,file,parentReference,webUrl",
        { headers: { Authorization: "Bearer " + token } });
      if (!sr.ok) {
        res.status(502).json({ error: "Could not open that link (HTTP " + sr.status + "). Make sure it is a link to the FOLDER and that it is shared with the Sentinel app.", detail: (await sr.text()).slice(0, 200) });
        return;
      }
      const srcItem = await sr.json();
      const srcDrive = srcItem.parentReference && srcItem.parentReference.driveId;
      if (!srcItem.folder) { res.status(400).json({ error: "That link points at a file, not a folder. Share the folder that holds the documents." }); return; }

      // List the files in it (one level of subfolders too — documents are often filed in one).
      const listChildren = async (driveId, itemId) => {
        const out = [];
        let u = G.GRAPH + "/drives/" + driveId + "/items/" + itemId + "/children?$select=id,name,file,folder,size&$top=200";
        while (u) {
          const r = await fetch(u, { headers: { Authorization: "Bearer " + token } });
          if (!r.ok) break;
          const j = await r.json();
          (j.value || []).forEach(x => out.push(x));
          u = j["@odata.nextLink"] || null;
        }
        return out;
      };
      const top = await listChildren(srcDrive, srcItem.id);
      const files = top.filter(x => x.file).map(x => ({ id: x.id, name: x.name, size: x.size, from: srcItem.name }));
      for (const sub of top.filter(x => x.folder).slice(0, 12)) {
        const kids = await listChildren(srcDrive, sub.id);
        kids.filter(x => x.file).forEach(x => files.push({ id: x.id, name: x.name, size: x.size, from: srcItem.name + "/" + sub.name }));
      }

      const destPath = "Sama Farooqui/Sentinel/Provider/" + safeName + "/" + PHASES[phaseIdx];
      if (imp === "preview") {
        res.status(200).json({
          ok: true, mode: "preview", provider: safeName, source: srcItem.name, sourceUrl: srcItem.webUrl,
          destination: destPath, phase: PHASES[phaseIdx], fileCount: files.length,
          files: files.slice(0, 100).map(f => ({ name: f.name, kb: Math.round((f.size || 0) / 1024), from: f.from })),
          note: "Nothing has been copied yet. These files will be COPIED — the originals stay where they are.",
        });
        return;
      }
      if (imp !== "run") { res.status(400).json({ error: "import must be preview or run" }); return; }
      if (!files.length) { res.status(400).json({ error: "no files found in that folder" }); return; }

      // Make sure the provider's Sentinel folder + all 6 phase subfolders exist.
      const base = "Sama Farooqui/Sentinel/Provider/" + safeName;
      await G.ensureFolderIn(token, G.docsRoot(), base);
      for (const p of PHASES) { try { await G.ensureFolderIn(token, G.docsRoot(), base + "/" + p); } catch (e) {} }

      // Destination folder id, needed as the copy target.
      const dr = await fetch(G.docsRoot() + "/root:/" + G.encPath(destPath) + "?$select=id,parentReference", { headers: { Authorization: "Bearer " + token } });
      if (!dr.ok) { res.status(500).json({ error: "could not open the destination folder after creating it (HTTP " + dr.status + ")" }); return; }
      const destItem = await dr.json();

      // Copy each file. Graph copy is asynchronous and returns 202 — we record acceptance and let
      // it finish server-side rather than holding the request open (and risking a timeout).
      const results = [];
      for (const f of files.slice(0, 60)) {
        const r = await fetch(G.GRAPH + "/drives/" + srcDrive + "/items/" + f.id + "/copy", {
          method: "POST",
          headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
          body: JSON.stringify({
            parentReference: { driveId: G.DOCS_DRIVE_ID, id: destItem.id },
            name: f.name,
            "@microsoft.graph.conflictBehavior": "rename",
          }),
        });
        results.push({ name: f.name, ok: r.status === 202 || r.ok, status: r.status, error: (r.status === 202 || r.ok) ? undefined : (await r.text()).slice(0, 160) });
      }
      const started = results.filter(r => r.ok).length;
      res.status(200).json({
        ok: started > 0, mode: "run", provider: safeName, destination: destPath,
        started, failed: results.length - started, skipped: Math.max(0, files.length - 60),
        results,
        note: "Copies run in the background on Microsoft's side and usually land within a minute. The originals were NOT moved or deleted. Click 'Sync from Excel', then reopen the provider.",
      });
      return;
    } catch (e) { res.status(500).json({ error: String(e.message || e) }); return; }
  }

  // ---- "Why can't I see this provider?" ------------------------------------------------------
  // GET/POST /api/data?whois=<name>   (admin) — traces one name through every stage:
  //   live Excel (Credentials / Inactive)  ->  baked data.json  ->  cached delta  ->  dashboard.
  // Answers the actual question, instead of leaving the user to guess which stage dropped them.
  const whois = url.searchParams.get("whois");
  if (whois) {
    if (!s || !s.admin) { res.status(403).json({ error: "admins only" }); return; }
    const q = String(whois).trim().toLowerCase();
    if (q.length < 2) { res.status(400).json({ error: "give at least 2 characters" }); return; }
    try {
      const xl = require("../lib/excel");
      const G = require("../lib/graph");
      const token = await G.accessToken();
      const slug = (l, f) => (l + "-" + (f || "")).replace(/[^A-Za-z0-9]+/g, "-").replace(/^-+|-+$/g, "").toLowerCase();
      const rowsOf = async (sheet) => {
        const sh = await xl.readSheet(token, sheet);
        const n = xl.detectNameCols((sh.values || [])[0] || []);
        return (sh.values || []).slice(1).map((r, i) => ({
          row: i + 2,
          last: String(r[n.lastIdx] == null ? "" : r[n.lastIdx]).replace(/[\*,()]+/g, "").trim(),
          first: String(r[n.firstIdx] == null ? "" : r[n.firstIdx]).trim(),
        })).filter(x => x.last);
      };
      const [actRows, inactRows] = await Promise.all([rowsOf(xl.SHEET_ACTIVE), rowsOf(xl.SHEET_INACTIVE)]);
      // Accept the way people actually type names: "Crespo, Jose", "Jose Crespo", or just
      // "crespo". Punctuation and extra spaces are collapsed on both sides before comparing.
      const norm = t => String(t || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
      const qn = norm(q);
      const qTokens = qn.split(" ").filter(Boolean);
      const match = arr => arr.filter(x => {
        const a = norm(x.first + " " + x.last), b = norm(x.last + " " + x.first);
        if (a.includes(qn) || b.includes(qn)) return true;
        // Every typed word appears somewhere in the name, in any order.
        return qTokens.length > 0 && qTokens.every(t => a.includes(t));
      });
      const inActive = match(actRows), inInactive = match(inactRows);

      const baked = (data.items || []).filter(i => i.scope === "provider");
      const nameHit = t => { const a = norm(t); return a.includes(qn) || (qTokens.length > 0 && qTokens.every(w => a.includes(w))); };
      const bakedHit = [...new Set(baked.filter(i => nameHit(i.entity) || String(i.entityKey || "").includes(qTokens.join("-"))).map(i => i.entityKey))];

      let delta = null; try { delta = await G.readJsonAt(token, G.drivePath("_Sentinel/roster_delta.json")); } catch (e) {}
      const dNew = ((delta && delta.newProviders) || []).filter(p => {
        const a = norm((p.first || "") + " " + (p.last || ""));
        return a.includes(qn) || (qTokens.length > 0 && qTokens.every(t => a.includes(t)));
      });
      const { applyRosterDelta } = require("../lib/delta");
      const merged = await applyRosterDelta(baked);
      const visible = [...new Set(merged.filter(i => i.scope === "provider" && nameHit(i.entity)).map(i => i.entityKey + (i.active === false ? " (inactive)" : "")))];

      // Plain-English verdict.
      let verdict;
      if (!inActive.length && !inInactive.length) {
        verdict = "NOT IN THE MASTER EXCEL. The dashboard only ever shows people who are in the roster workbook, so this provider has to be added first — use '+ Add provider', which also creates their OneDrive folder.";
      } else if (inInactive.length && !inActive.length) {
        verdict = "In the roster but on the 'Inactive Providers' sheet. Inactive people are hidden until you tick 'Show inactive' on the Providers tab.";
      } else if (visible.length) {
        verdict = "Present and visible on the dashboard.";
      } else if (dNew.length) {
        verdict = "In the active roster and in the cached delta, but not rendering. Click 'Sync from Excel' once, then reload.";
      } else {
        verdict = "In the active roster but NOT in the cached delta, so the dashboard has not picked them up yet. Click 'Sync from Excel' once — that rebuilds the cache from the workbook.";
      }
      res.status(200).json({
        query: whois, verdict,
        liveRoster: { credentialsSheet: inActive, inactiveSheet: inInactive },
        inBakedData: bakedHit,
        inCachedDelta: { newProviders: dNew, deltaGeneratedAt: (delta && delta.generatedAt) || null },
        visibleOnDashboard: visible,
        counts: { liveActive: actRows.length, liveInactive: inactRows.length, bakedProviders: new Set(baked.map(i => i.entityKey)).size },
      });
      return;
    } catch (e) { res.status(500).json({ error: String(e.message || e) }); return; }
  }

  const selftest = url.searchParams.get("selftest");
  if (selftest) {
    if (!s || !s.admin) { res.status(403).json({ error: "admins only" }); return; }
    const steps = [];
    const step = (name, pass, detail) => { steps.push({ name, pass: !!pass, detail: detail === undefined ? "" : detail }); return !!pass; };
    const TEST_LAST = "ZZSelftest", TEST_FIRST = "Sentinel";
    const TEST_NAME = TEST_FIRST + " " + TEST_LAST;
    const TEST_KEY = (TEST_LAST + "-" + TEST_FIRST).toLowerCase();
    let cleanupNote = "nothing to clean";
    try {
      const xl = require("../lib/excel");
      const G = require("../lib/graph");
      const token = await G.accessToken();
      step("Microsoft Graph token", true, "app-only auth OK");

      // --- 1. Read the live roster and report its ACTUAL column layout -----------------------
      const act = await xl.readSheet(token, xl.SHEET_ACTIVE);
      const header = (act.values || [])[0] || [];
      const labels = header.map(h => (h == null ? "" : String(h && h.text ? h.text : h))).map(x => x.replace(/\s+/g, " ").trim());
      const { lastIdx, firstIdx } = xl.detectNameCols(header);
      const credMap = xl.detectCredCols(header);
      step("Roster downloaded", true, xl.SHEET_ACTIVE + ": " + act.values.length + " rows, " + labels.length + " columns");
      const colLetter = i => String.fromCharCode(65 + i);
      const shifted = lastIdx !== 0;
      step("Name columns detected by header", labels[lastIdx] && /last/i.test(labels[lastIdx]),
        "Last Name = column " + colLetter(lastIdx) + " (\"" + labels[lastIdx] + "\"), First Name = column " + colLetter(firstIdx) + " (\"" + labels[firstIdx] + "\")" +
        (shifted ? "  <-- ROSTER IS SHIFTED: extra column(s) before the names. The header-detection fix IS what's keeping this working." : "  (standard A/B layout)"));
      if (shifted) step("Leading extra column(s)", true, "column A is \"" + labels[0] + "\" — not a name column. Old code read this as the surname; that was the bug.");
      step("Credential date columns mapped", Object.keys(credMap).length >= 10, Object.keys(credMap).length + " of 15 expiry columns found by header name");
      const parsed = (act.values || []).slice(1)
        .map(r => ({ last: String(r[lastIdx] == null ? "" : r[lastIdx]).replace(/[\*,()]+/g, "").trim(), first: String(r[firstIdx] == null ? "" : r[firstIdx]).trim() }))
        .filter(x => x.last && x.last.toLowerCase() !== TEST_LAST.toLowerCase());
      const junky = parsed.filter(p => /\b(thhs|email sent|requested|verify|yrly)\b/i.test(p.last)).length;
      step("Provider names parse cleanly", parsed.length > 0 && junky === 0,
        parsed.length + " providers read; " + junky + " unusable names. Samples: " +
        parsed.slice(0, 3).map(p => "\"" + p.first + " " + p.last + "\"").join(", "));
      const dates = xl.expiryDatesFromValues(act.values);
      step("Expiry dates read from Excel", dates.length > 0, dates.length + " dates parsed out of the credential columns");

      // --- 2. Delta health ------------------------------------------------------------------
      let rd = null;
      try { rd = await G.readJsonAt(token, G.drivePath("_Sentinel/roster_delta.json")); } catch (e) { rd = null; }
      if (rd) {
        const nl = rd.newProviders || [];
        const junkNew = nl.filter(p => /\b(thhs|email sent|requested|verify|yrly)\b/i.test(String(p.last) + " " + String(p.first))).length;
        const bakedProviderKeys = new Set((data.items || []).filter(i => i.scope === "provider").map(i => i.entityKey));
        const removedHit = (rd.removed || []).filter(k => bakedProviderKeys.has(k)).length;
        const corrupt = (nl.length >= 2 && junkNew > nl.length * 0.3) || (bakedProviderKeys.size > 0 && removedHit > bakedProviderKeys.size * 0.3);
        step("Cached roster delta is healthy", !corrupt,
          "generated " + (rd.generatedAt || "?") + " — " + nl.length + " new (" + junkNew + " junk), " + removedHit + " of " + bakedProviderKeys.size + " baked providers marked removed" +
          (corrupt ? ".  CORRUPT -> the guard is suppressing it and showing clean baked data. Click 'Sync from Excel' to rewrite it." : ""));
      } else step("Cached roster delta is healthy", true, "no delta cached yet (clean baseline)");

      // --- 3. Both drives reachable ----------------------------------------------------------
      const provRootPath = "Sama Farooqui/Sentinel/Provider";
      const pr = await fetch(G.docsRoot() + "/root:/" + G.encPath(provRootPath) + ":/children?$select=name&$top=1", { headers: { Authorization: "Bearer " + token } });
      step("Documents drive reachable (SharePoint)", pr.ok, pr.ok ? provRootPath + " is readable" : "HTTP " + pr.status + " — check MS_DOCS_DRIVE_ID / permissions");
      let staffOk = 0;
      for (const sf of ["CHER RN and FD Roster and Credentialing log.xlsx", "Frisco ER RN and FD Roster and Credentialing.xlsx"]) {
        const p = "WCGTX Phyicians_04.08.2020/..WCGTX Master Rosters/" + sf;
        const r = await fetch(G.driveRoot() + "/root:/" + G.encPath(p), { headers: { Authorization: "Bearer " + token } }).catch(() => ({ ok: false, status: 0 }));
        if (r && r.ok) staffOk++;
      }
      step("Staff workbooks reachable", staffOk === 2, staffOk + " of 2 RN/FD workbooks found");

      // --- Every provider in the live roster should be visible on the dashboard ---------------
      // This is the "I can't see Dr X" check, done for the whole roster at once.
      const slugP = (l, f) => (l + "-" + (f || "")).replace(/[^A-Za-z0-9]+/g, "-").replace(/^-+|-+$/g, "").toLowerCase();
      const bakedProv = (data.items || []).filter(i => i.scope === "provider");
      const { applyRosterDelta: merge } = require("../lib/delta");
      const mergedProv = await merge(bakedProv);
      const visibleKeys = new Set(mergedProv.filter(i => i.scope === "provider").map(i => i.entityKey));
      const liveKeys = parsed.map(p => ({ key: slugP(p.last, p.first), name: (p.first ? p.first + " " : "") + p.last }));
      const notShowing = liveKeys.filter(k => !visibleKeys.has(k.key));
      step("Every active roster provider is on the dashboard", notShowing.length === 0,
        notShowing.length === 0
          ? parsed.length + " active providers in the workbook, all present on the dashboard"
          : notShowing.length + " of " + parsed.length + " are in the workbook but NOT on the dashboard: " +
            notShowing.slice(0, 20).map(k => k.name).join(", ") + (notShowing.length > 20 ? " …" : "") +
            "  — click 'Sync from Excel' once; if they still don't appear, use 'Find a provider' for the reason.");

      if (selftest !== "full") {
        const passed = steps.filter(x => x.pass).length;
        res.status(200).json({ ok: passed === steps.length, mode: "diagnose (read-only, nothing changed)", passed, total: steps.length, rosterShifted: shifted, steps });
        return;
      }

      // --- 4. WRITE ROUND TRIP --------------------------------------------------------------
      // Clean up any leftover test row from an interrupted earlier run, so this is repeatable.
      const pre = await xl.hardDelete(token, TEST_KEY, TEST_LAST, TEST_FIRST);
      if (pre.length) cleanupNote = "removed " + pre.length + " leftover test row(s) before starting";
      const snap = await xl.snapshotWorkbook(token, "before-selftest");
      step("Workbook backed up first", snap && snap.ok, snap && snap.ok ? ("saved " + snap.name + " to _Sentinel/roster_backups") : ("backup failed: " + ((snap && snap.error) || "?") + " — ABORTING the write test"));
      if (!snap || !snap.ok) {
        res.status(200).json({ ok: false, mode: "full (aborted before any write)", passed: steps.filter(x => x.pass).length, total: steps.length, steps });
        return;
      }

      // 4a. Add the row the same way "+ Add provider" does.
      const added = await xl.appendProviderRow(token, xl.SHEET_ACTIVE, TEST_LAST, TEST_FIRST);
      step("Test provider written to Excel", !!(added && added.rowIndex), "added \"" + TEST_NAME + "\" at row " + (added && added.rowIndex));

      // 4b. Read it back — proves the write landed in the REAL name columns, not a shifted one.
      const back = await xl.findRowByEntityKey(token, xl.SHEET_ACTIVE, TEST_KEY);
      step("Row reads back correctly from Excel", !!(back && back.last === TEST_LAST && back.first === TEST_FIRST),
        back ? ("row " + back.rowIndex + " -> last=\"" + back.last + "\" first=\"" + back.first + "\"" + (back.last === TEST_LAST ? " (matches)" : " (MISMATCH — wrote into the wrong column)")) : "not found after write");

      // 4c. Create the OneDrive folder + 6 SOP subfolders, exactly like the Add flow.
      const PHASES = ["1. Application & Document Collection", "2. Primary Source Verification", "3. Background & Compliance Review", "4. Medical Staff Review", "5. Payer Enrollment & Facility Setup", "6. Approval & Ongoing Monitoring"];
      const base = provRootPath + "/" + TEST_NAME;
      await G.ensureFolderIn(token, G.docsRoot(), base);
      for (const p of PHASES) { try { await G.ensureFolderIn(token, G.docsRoot(), base + "/" + p); } catch (e) {} }
      const chk = await fetch(G.docsRoot() + "/root:/" + G.encPath(base) + ":/children?$select=name&$top=50", { headers: { Authorization: "Bearer " + token } });
      const kids = chk.ok ? (((await chk.json()).value) || []).map(x => x.name) : [];
      step("OneDrive folder created", chk.ok, chk.ok ? ("Provider/" + TEST_NAME + "/ exists") : ("could not read it back: HTTP " + chk.status));
      const missing = PHASES.filter(p => !kids.includes(p));
      step("All 6 SOP subfolders created", missing.length === 0, missing.length === 0 ? "all 6 present" : ("missing: " + missing.join(", ")));

      // 4d. Would it actually show on the dashboard? Rebuild the cache with the SAME shared regen
      // the app itself uses, then merge for real — so this tests the real code path, not a copy.
      const { regenerateRoster } = require("../lib/regen");
      const seedKeys2 = new Set((data.items || []).filter(i => i.scope === "provider").map(i => i.entityKey));
      const regen1 = await regenerateRoster(token);
      step("Roster cache rebuilt from Excel", !!regen1.ok, regen1.ok
        ? (regen1.providersRead + " providers read, " + regen1.newProviders + " new, " + regen1.dates + " dates, staff " + regen1.staff)
        : ("refused: " + regen1.reason));
      const { applyRosterDelta } = require("../lib/delta");
      const merged = await applyRosterDelta((data.items || []).filter(i => i.scope === "provider"));
      const testItems = merged.filter(i => i.entityKey === TEST_KEY);
      const realStill = new Set(merged.filter(i => i.scope === "provider" && i.entityKey !== TEST_KEY).map(i => i.entityKey)).size;
      step("Appears on the dashboard", testItems.length > 0, testItems.length + " credential rows generated for the test provider");
      step("Real providers still intact", realStill >= seedKeys2.size * 0.9, realStill + " real providers present after the merge (baked baseline " + seedKeys2.size + ")");

      // 4e. Delete both sides and verify.
      const del = await xl.hardDelete(token, TEST_KEY, TEST_LAST, TEST_FIRST);
      const folderDel = await G.deleteProviderFolder(token, TEST_NAME);
      step("Test row deleted from Excel", del.length > 0, "removed " + del.length + " row(s)");
      step("Test OneDrive folder deleted", folderDel && folderDel.ok, "status " + (folderDel && folderDel.status) + (folderDel && folderDel.note ? " (" + folderDel.note + ")" : ""));
      const goneA = await xl.findRowByEntityKey(token, xl.SHEET_ACTIVE, TEST_KEY);
      const goneI = await xl.findRowByEntityKey(token, xl.SHEET_INACTIVE, TEST_KEY);
      step("Verified gone from both sheets", !goneA && !goneI, (!goneA && !goneI) ? "no trace left in Credentials or Inactive" : "STILL PRESENT — remove row for \"" + TEST_NAME + "\" by hand");
      const chk2 = await fetch(G.docsRoot() + "/root:/" + G.encPath(base), { headers: { Authorization: "Bearer " + token } });
      step("Verified folder gone from OneDrive", chk2.status === 404, chk2.status === 404 ? "folder removed (recoverable from the SharePoint recycle bin)" : ("still there: HTTP " + chk2.status + " — delete Provider/" + TEST_NAME + " by hand"));

      // 4f. Rebuild the cache properly now the test row is gone, via the SAME shared regen the app
      // uses. (Hand-rolling it here wrote inactivated:[], which silently reactivated every
      // provider who had been marked inactive.)
      const restored = await regenerateRoster(token);
      step("Roster cache restored", !!restored.ok, restored.ok
        ? ("rebuilt from Excel — " + restored.providersRead + " providers, " + restored.dates + " dates; test provider gone")
        : ("could not rebuild: " + (restored.reason || "unknown") + " — click 'Sync from Excel' once"));
      cleanupNote = "all test data removed (Excel row + OneDrive folder + cache entry)";

      const passed = steps.filter(x => x.pass).length;
      res.status(200).json({ ok: passed === steps.length, mode: "full round trip", passed, total: steps.length, rosterShifted: shifted, cleanup: cleanupNote, backup: snap.name, steps });
      return;
    } catch (e) {
      const msg = String(e.message || e);
      steps.push({ name: "ERROR — test stopped here", pass: false, detail: /ROSTER_LOCKED|423|resourceLocked/i.test(msg) ? "The master Excel is open in Excel right now. Close it everywhere and re-run." : msg.slice(0, 300) });
      res.status(200).json({ ok: false, mode: selftest === "full" ? "full (incomplete)" : "diagnose", passed: steps.filter(x => x.pass).length, total: steps.length, cleanup: cleanupNote, hint: "If a test row or folder for \"" + TEST_NAME + "\" was left behind, re-run the test — it cleans up leftovers on start.", steps });
      return;
    }
  }

  // ---- roster ops: add/remove a provider in the master Excel (admin-only) ----
  // GET  /api/data?roster=list           -> { active:[{last,first,row}], inactive:[...] }
  // POST /api/data?roster=add            { last, first, email? }
  // POST /api/data?roster=remove         { last, first }       (moves Credentials -> Inactive)
  const rosterAction = url.searchParams.get("roster");
  if (rosterAction) {
    if (!s.admin && rosterAction !== "list") { res.status(403).json({ error: "admins only" }); return; }
    try {
      const xl = require("../lib/excel");
      const { accessToken } = require("../lib/graph");
      const token = await accessToken();
      if (rosterAction === "list") {
        const [a, i] = await Promise.all([xl.readSheet(token, xl.SHEET_ACTIVE), xl.readSheet(token, xl.SHEET_INACTIVE)]);
        const flat = sh => {
          const { lastIdx, firstIdx } = xl.detectNameCols((sh.values || [])[0] || []);
          return (sh.values || []).slice(1).map((row, idx) => ({ last: row[lastIdx] || "", first: row[firstIdx] || "", row: idx + 2 }))
            .filter(r => String(r.last).trim());
        };
        res.status(200).json({ active: flat(a), inactive: flat(i) });
        return;
      }
      let body = ""; await new Promise(r => { req.on("data", c => body += c); req.on("end", r); });
      let b = {}; try { b = JSON.parse(body || "{}"); } catch (e) {}
      const last = String(b.last || "").trim();
      const first = String(b.first || "").trim();
      // 'trash' is a read (no body) and 'restore' is keyed by {id} — neither needs a last name.
      // This guard was bouncing the Recycle bin read with "last name required" before it ever
      // reached the trash branch, which is why the bin always looked empty.
      if (!last && rosterAction !== "trash" && rosterAction !== "restore") { res.status(400).json({ error: "last name required" }); return; }
      if (rosterAction === "add") {
        // Only block dups in the ACTIVE sheet — if the same name exists in Inactive Providers,
        // they were deactivated and the user can legitimately re-add. Force=true bypasses
        // even the active-sheet check (for true duplicate-named providers).
        if (!b.force) {
          const dup = await xl.findRow(token, xl.SHEET_ACTIVE, last, first);
          if (dup) {
            res.status(409).json({
              error: "already present in active roster",
              row: dup.rowIndex,
              hint: "Send {force:true} to add a second row with the same name."
            });
            return;
          }
        }
        const snap = await xl.snapshotWorkbook(token, "before-add");
        // Place the name into the sheet's real last/first columns (header-detected), not a fixed
        // col A/B — the roster may carry a leading tracking column that shifts the layout.
        const result = await xl.appendProviderRow(token, xl.SHEET_ACTIVE, last, first);
        // Also append email to COI roster so reminders reach them
        if (b.email && b.email.includes("@")) {
          try { await xl.appendRow(token, "WCGTX COI Roster", [last, first, null, String(b.email).trim()]); } catch (e) {}
        }
        // ALSO create the SharePoint folder Sentinel/Provider/<First Last>/ so uploads/QR work
        // and it shows up in OneDrive immediately. Symmetric with the delete handler which
        // removes the folder. Failure here doesn't roll back the Excel write — folder will be
        // created later on first file upload via ensureFolderIn anyway.
        let folderCreate = null;
        try {
          const { ensureFolderIn, docsRoot } = require("../lib/graph");
          const folderName = ((first || "") + " " + (last || "")).trim();
          if (folderName) {
            const base = "Sama Farooqui/Sentinel/Provider/" + folderName;
            await ensureFolderIn(token, docsRoot(), base);
            // Build the same 6 SOP-phase subfolders the dashboard shows, so the OneDrive folder
            // mirrors the dashboard structure and files land in the right phase.
            const PHASES = [
              "1. Application & Document Collection", "2. Primary Source Verification",
              "3. Background & Compliance Review", "4. Medical Staff Review",
              "5. Payer Enrollment & Facility Setup", "6. Approval & Ongoing Monitoring",
            ];
            for (const p of PHASES) { try { await ensureFolderIn(token, docsRoot(), base + "/" + p); } catch (e) {} }
            folderCreate = { ok: true, name: folderName, subfolders: PHASES.length };
          }
        } catch (e) { folderCreate = { ok: false, error: String(e.message || e).slice(0, 200) }; }
        res.status(200).json({ ok: true, action: "added", sheet: xl.SHEET_ACTIVE, rowIndex: result.rowIndex, snapshot: snap, folder: folderCreate });
        return;
      }
      if (rosterAction === "trash") {
        // GET-style action (also accept POST) — list currently trashed providers
        const { readJsonAt, drivePath } = require("../lib/graph");
        const trash = (await readJsonAt(token, drivePath("_Sentinel/trash.json"))) || { entries: [] };
        res.status(200).json(trash);
        return;
      }
      if (rosterAction === "restore") {
        const id = String(b.id || "").trim();
        if (!id) { res.status(400).json({ error: "id required" }); return; }
        const { readJsonAt, writeJsonAt, drivePath } = require("../lib/graph");
        const trashPath = drivePath("_Sentinel/trash.json");
        const trash = (await readJsonAt(token, trashPath)) || { entries: [] };
        const entry = (trash.entries || []).find(e => e.id === id);
        if (!entry) { res.status(404).json({ error: "trash entry not found" }); return; }
        await xl.snapshotWorkbook(token, "before-restore");
        // Restore each row to the sheet it CAME FROM. hardDelete records `sheet` per row; this
        // used to always restore into the active Credentials sheet, so a terminated provider
        // deleted out of "Inactive Providers" came back as an ACTIVE provider counting against
        // compliance — and anyone with rows in both sheets came back twice.
        for (const r of (entry.rows || [])) await xl.restoreRow(token, r.sheet || xl.SHEET_ACTIVE, r.values);
        trash.entries = trash.entries.filter(e => e.id !== id);
        await writeJsonAt(token, trashPath, trash);
        res.status(200).json({ ok: true, action: "restored", entity: entry.entity });
        return;
      }
      if (rosterAction === "delete") {
        // HARD DELETE: remove from BOTH Credentials and Inactive, ALSO delete the provider's
        // SharePoint folder (Sentinel/Provider/<name>/), log to trash.json for recovery.
        const entityKey = String(b.entityKey || "").trim();
        const snap = await xl.snapshotWorkbook(token, "before-delete");
        // Never destroy rows we couldn't back up first.
        if (!snap || !snap.ok) { res.status(503).json({ error: "Could not back up the roster before deleting, so nothing was deleted. " + ((snap && snap.error) || ""), snapshot: snap }); return; }
        let removed;
        try { removed = await xl.hardDelete(token, entityKey, last, first); }
        catch (de) {
          const m = String(de.message || de);
          // hardDelete refuses an ambiguous surname-only delete rather than removing several people.
          if (/^AMBIGUOUS_DELETE/.test(m)) { res.status(409).json({ error: m.replace(/^AMBIGUOUS_DELETE:\s*/, "") }); return; }
          throw de;
        }
        if (!removed.length) { res.status(404).json({ error: "not found in roster", tried: { last, first, entityKey } }); return; }
        // Folder name comes from the HEADER-DETECTED last/first that hardDelete returns — not from
        // guessed array positions. Guessing produced names like "Smith Sent" on a shifted roster,
        // the delete 404'd, and graph.js reports a 404 as success — so the real folder and every
        // document in it were left behind while the app said "deleted".
        const r0 = removed[0];
        const folderName = ((r0.first ? r0.first + " " : "") + (r0.last || "")).trim();
        const { deleteProviderFolder } = require("../lib/graph");
        const folderDel = await deleteProviderFolder(token, folderName.trim());
        // Log to trash so the user can recover from "Recycle bin".
        const { readJsonAt, writeJsonAt, drivePath } = require("../lib/graph");
        const trashPath = drivePath("_Sentinel/trash.json");
        let trashWriteError = null;
        let trashEntryCount = 0;
        try {
          const trash = (await readJsonAt(token, trashPath)) || { entries: [] };
          const id = "tr_" + Date.now() + "_" + Math.random().toString(36).slice(2, 8);
          // Only dedupe by entityKey when it's non-empty (an empty key would match every empty key).
          if (entityKey) trash.entries = (trash.entries || []).filter(e => e.entityKey !== entityKey);
          else trash.entries = trash.entries || [];
          trash.entries.unshift({
            id, entityKey: entityKey || null,
            entity: ((first || "") + " " + (last || "")).trim() || folderName,
            deletedAt: new Date().toISOString(),
            deletedBy: s ? s.email : "system",
            rows: removed,
          });
          if (trash.entries.length > 200) trash.entries = trash.entries.slice(0, 200);
          await writeJsonAt(token, trashPath, trash);
          trashEntryCount = trash.entries.length;
          // Read the file straight back so the toast reports what actually persisted, not just
          // what we tried to write. If this comes back 0 the write isn't sticking (path/perm).
          let trashVerified = null;
          try { const rb = await readJsonAt(token, trashPath); trashVerified = ((rb && rb.entries) || []).length; } catch (ve) { trashVerified = "read-back failed: " + String(ve.message || ve).slice(0, 120); }
          res.status(200).json({ ok: true, action: "deleted", removedRows: removed.length, trashId: id, trashEntries: trashEntryCount, trashVerified, trashPath, snapshot: snap, folder: folderDel });
          return;
        } catch (te) {
          trashWriteError = String(te.message || te).slice(0, 200);
          res.status(200).json({ ok: true, action: "deleted", removedRows: removed.length, warning: "trash log failed: " + trashWriteError, trashEntries: trashEntryCount, snapshot: snap, folder: folderDel });
          return;
        }
      }
      if (rosterAction === "remove") {
        // Prefer entityKey when sent (no name-splitting ambiguity); fall back to last/first.
        const entityKey = String(b.entityKey || "").trim();
        let found = null;
        if (entityKey) found = await xl.findRowByEntityKey(token, xl.SHEET_ACTIVE, entityKey);
        if (!found) found = await xl.findRow(token, xl.SHEET_ACTIVE, last, first);
        if (!found) {
          let inact = null;
          if (entityKey) inact = await xl.findRowByEntityKey(token, xl.SHEET_INACTIVE, entityKey);
          if (!inact) inact = await xl.findRow(token, xl.SHEET_INACTIVE, last, first);
          if (inact) { res.status(409).json({ error: "already inactive" }); return; }
          const sh = await xl.readSheet(token, xl.SHEET_ACTIVE);
          const sample = (sh.values || []).slice(1, 12).map(r => ({ last: r[0], first: r[1] })).filter(x => x.last);
          res.status(404).json({
            error: "not found in roster",
            tried: { last, first, entityKey },
            sample,
            hint: "Check that the provider's name appears in the first two columns of the WCGTX Credentials sheet."
          });
          return;
        }
        await xl.snapshotWorkbook(token, "before-remove");
        await xl.moveRow(token, xl.SHEET_ACTIVE, found.rowIndex, xl.SHEET_INACTIVE);
        res.status(200).json({ ok: true, action: "moved to inactive", fromRow: found.rowIndex });
        return;
      }
      res.status(400).json({ error: "roster action must be list, add, or remove" });
      return;
    } catch (e) {
      res.status(500).json({ error: String(e.message || e) });
      return;
    }
  }

  // ---- facility folder ops (admin): create / soft-delete / restore / list folders under a
  //      facility's State Readiness tree in SharePoint. Soft-delete = rename with a "zz." prefix
  //      (the scan ignores zz.* folders, and it matches the user's own archive naming), so the
  //      folder + ALL its files are preserved and fully restorable. ----
  const facilityAction = url.searchParams.get("facility");
  if (facilityAction) {
    if (!s.admin) { res.status(403).json({ error: "admins only" }); return; }
    try {
      const { accessToken, docsRoot, encPath, ensureFolderIn, readJsonAt, writeJsonAt, drivePath } = require("../lib/graph");
      const token = await accessToken();
      const TRASH = drivePath("_Sentinel/facility_trash.json");
      const FAC_DIR = { "Castle Hills ER": "Castle Hills", "Frisco ER": "Frisco" };
      const baseFor = (fac) => "Sama Farooqui/Sentinel/State Readiness/" + FAC_DIR[fac];
      const cleanName = (n) => String(n || "").replace(/[\/\\:*?"<>|]/g, "").replace(/\s+/g, " ").trim();

      if (facilityAction === "trash") {
        const t = (await readJsonAt(token, TRASH)) || { entries: [] };
        res.status(200).json(t); return;
      }

      let body = ""; await new Promise(r => { req.on("data", c => body += c); req.on("end", r); });
      let b = {}; try { b = JSON.parse(body || "{}"); } catch (e) {}
      const fac = String(b.facility || "").trim();
      if (!FAC_DIR[fac]) { res.status(400).json({ error: "facility must be 'Castle Hills ER' or 'Frisco ER'", got: fac }); return; }

      if (facilityAction === "list") {
        const r = await fetch(docsRoot() + "/root:/" + encPath(baseFor(fac)) + ":/children?$select=name,folder&$top=400", { headers: { Authorization: "Bearer " + token } });
        if (!r.ok) { res.status(r.status).json({ error: "list HTTP " + r.status, detail: (await r.text()).slice(0, 160) }); return; }
        const all = ((await r.json()).value || []).filter(x => x.folder).map(x => x.name);
        // hide the soft-deleted (zz.*) ones from the live list
        res.status(200).json({ facility: fac, folders: all.filter(n => !/^zz\./i.test(n)).sort() });
        return;
      }

      if (facilityAction === "add") {
        const name = cleanName(b.name);
        if (!name) { res.status(400).json({ error: "folder name required" }); return; }
        if (/^zz\./i.test(name)) { res.status(400).json({ error: "name can't start with 'zz.'" }); return; }
        await ensureFolderIn(token, docsRoot(), baseFor(fac) + "/" + name);
        res.status(200).json({ ok: true, action: "created", facility: fac, name }); return;
      }

      if (facilityAction === "delete") {
        const name = cleanName(b.name);
        if (!name) { res.status(400).json({ error: "folder name required" }); return; }
        const hidden = "zz." + name;
        const r = await fetch(docsRoot() + "/root:/" + encPath(baseFor(fac) + "/" + name), {
          method: "PATCH", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
          body: JSON.stringify({ name: hidden })
        });
        if (!r.ok) { res.status(r.status === 404 ? 404 : 500).json({ error: "delete (rename) failed " + r.status, detail: (await r.text()).slice(0, 160) }); return; }
        const t = (await readJsonAt(token, TRASH)) || { entries: [] };
        const id = "fac_" + Date.now() + "_" + Math.random().toString(36).slice(2, 8);
        t.entries.unshift({ id, facility: fac, name, hiddenName: hidden, deletedAt: new Date().toISOString(), deletedBy: s.email });
        if (t.entries.length > 200) t.entries = t.entries.slice(0, 200);
        await writeJsonAt(token, TRASH, t);
        res.status(200).json({ ok: true, action: "deleted", facility: fac, name, trashId: id, trashEntries: t.entries.length }); return;
      }

      if (facilityAction === "restore") {
        const id = String(b.id || "").trim();
        const t = (await readJsonAt(token, TRASH)) || { entries: [] };
        const entry = (t.entries || []).find(e => e.id === id);
        if (!entry) { res.status(404).json({ error: "trash entry not found" }); return; }
        const r = await fetch(docsRoot() + "/root:/" + encPath(baseFor(entry.facility) + "/" + entry.hiddenName), {
          method: "PATCH", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
          body: JSON.stringify({ name: entry.name })
        });
        if (!r.ok) { res.status(500).json({ error: "restore (rename) failed " + r.status, detail: (await r.text()).slice(0, 160) }); return; }
        t.entries = t.entries.filter(e => e.id !== id);
        await writeJsonAt(token, TRASH, t);
        res.status(200).json({ ok: true, action: "restored", facility: entry.facility, name: entry.name }); return;
      }

      res.status(400).json({ error: "facility action must be add, delete, list, trash, or restore" });
      return;
    } catch (e) { res.status(500).json({ error: String(e.message || e) }); return; }
  }

  // ---- onboarding status board: any signed-in user reads; admins toggle steps ----
  const onboardAction = url.searchParams.get("onboard");
  if (onboardAction) {
    try {
      const { accessToken, readJsonAt, writeJsonAt, drivePath } = require("../lib/graph");
      const token = await accessToken();
      const OB = drivePath("_Sentinel/onboarding.json");
      if (onboardAction === "all") {
        res.status(200).json((await readJsonAt(token, OB)) || {});
        return;
      }
      if (onboardAction === "set") {
        if (!s.admin) { res.status(403).json({ error: "admins only" }); return; }
        let body = ""; await new Promise(r => { req.on("data", c => body += c); req.on("end", r); });
        let b = {}; try { b = JSON.parse(body || "{}"); } catch (e) {}
        const ek = String(b.entityKey || "").trim();
        const step = String(b.step || "").trim();
        if (!ek || !step) { res.status(400).json({ error: "entityKey and step required" }); return; }
        const data = (await readJsonAt(token, OB)) || {};
        const rec = data[ek] || { entity: b.entity || ek, steps: {} };
        if (b.entity) rec.entity = b.entity;
        rec.steps = rec.steps || {};
        if (b.done) rec.steps[step] = { done: true, at: new Date().toISOString(), by: s ? s.email : "" };
        else delete rec.steps[step];
        data[ek] = rec;
        await writeJsonAt(token, OB, data);
        res.status(200).json({ ok: true, record: rec }); return;
      }
      res.status(400).json({ error: "onboard action must be all or set" }); return;
    } catch (e) { res.status(500).json({ error: String(e.message || e) }); return; }
  }

  // ---- server-side OCR (free, OCR.space): read dates off a scanned doc, PDF or image ----
  const ocrUrl = url.searchParams.get("ocr");
  if (ocrUrl) {
    try {
      const { accessToken, GRAPH } = require("../lib/graph");
      const token = await accessToken();
      const r = await resolveDownloadUrl(token, GRAPH, ocrUrl);
      if (r.error) { res.status(502).json({ error: r.error, detail: r.detail }); return; }
      const fileResp = await fetch(r.dl);
      if (!fileResp.ok) { res.status(502).json({ error: "download " + fileResp.status }); return; }
      const ab = await fileResp.arrayBuffer();
      const isPdf = /\.pdf$/i.test(r.item.name || "");
      const key = process.env.OCR_SPACE_KEY || "helloworld";
      const fd = new FormData();
      fd.append("apikey", key); fd.append("OCREngine", "2"); fd.append("scale", "true"); fd.append("isOverlayRequired", "false");
      if (isPdf) fd.append("filetype", "PDF");
      fd.append("file", new Blob([ab], { type: (r.item.file && r.item.file.mimeType) || "application/octet-stream" }), r.item.name || "document");
      const o = await fetch("https://api.ocr.space/parse/image", { method: "POST", body: fd });
      const oj = await o.json().catch(() => ({}));
      if (oj.IsErroredOnProcessing) { res.status(502).json({ error: "ocr: " + (Array.isArray(oj.ErrorMessage) ? oj.ErrorMessage.join("; ") : (oj.ErrorMessage || "failed")) }); return; }
      const text = (oj.ParsedResults || []).map(p => p.ParsedText || "").join("\n");
      res.status(200).json({ ok: true, dates: extractDates(text), chars: text.length });
    } catch (e) { res.status(500).json({ error: String(e.message || e) }); }
    return;
  }

  // ---- same-origin file bytes proxy ----
  const fileUrl = url.searchParams.get("file");
  if (fileUrl) {
    try {
      const { accessToken, GRAPH } = require("../lib/graph");
      const token = await accessToken();
      const r = await resolveDownloadUrl(token, GRAPH, fileUrl);
      if (r.error) { res.status(502).json({ error: r.error, detail: r.detail }); return; }
      const f = await fetch(r.dl);
      if (!f.ok) { res.status(502).json({ error: "download " + f.status }); return; }
      const buf = Buffer.from(await f.arrayBuffer());
      res.setHeader("Content-Type", (r.item.file && r.item.file.mimeType) || f.headers.get("content-type") || "application/octet-stream");
      res.status(200).send(buf);
    } catch (e) { res.status(500).json({ error: String(e.message || e) }); }
    return;
  }

  res.setHeader("Cache-Control", "no-store");
  const tabs = (s.tabs && s.tabs.length) ? s.tabs : ["provider", "facility", "other"];
  let items = (data.items || []).filter(i => tabs.includes(i.scope));
  let deltaSuppressed = null;
  // Merge the live roster delta — surfaces Excel-roster changes without a code redeploy.
  // Placeholders carry a SharePoint folderLink so the QR/upload flow works immediately.
  if (tabs.includes("provider")) {
    const { applyRosterDelta } = require("../lib/delta");
    items = await applyRosterDelta(items);
    // Carry the "we ignored the live roster cache" warning through to the client, so a stale
    // board announces itself instead of quietly looking normal.
    deltaSuppressed = items.deltaSuppressed || null;
  }
  const keys = new Set(items.map(i => i.entityKey));
  const entityFiles = {};
  for (const k in (data.entityFiles || {})) if (keys.has(k)) entityFiles[k] = data.entityFiles[k];
  const contacts = tabs.includes("facility") ? (data.contacts || []) : [];
  res.status(200).json(Object.assign({}, data, { items, entityFiles, contacts, allowedTabs: tabs, deltaSuppressed }));
};
