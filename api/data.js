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
  // Vercel crons (no user session) are allowed to hit ?regen=1 only.
  const isCron = /vercel-cron/i.test(req.headers["user-agent"] || "");
  const s = getSession(req);
  if (!s && !(isCron && url.searchParams.get("regen") === "1")) { res.status(401).json({ error: "sign-in required" }); return; }

  // ---- hourly cron regen: pull the live Excel + write roster_delta.json to OneDrive.
  //      /api/data merges the delta on every read, so changes propagate within seconds
  //      of the cron firing (not a full Python regen — those need data-entry tools we
  //      can't run server-side without Python). Schedule lives in vercel.json. ----
  if (url.searchParams.get("regen") === "1") {
    if (!isCron && (!s || !s.admin)) { res.status(403).json({ error: "admins only" }); return; }
    try {
      const xl = require("../lib/excel");
      const { accessToken, drivePath, writeJsonAt } = require("../lib/graph");
      const token = await accessToken();
      const [act, inact] = await Promise.all([
        xl.readSheet(token, xl.SHEET_ACTIVE),
        xl.readSheet(token, xl.SHEET_INACTIVE),
      ]);
      const slug = (l, f) => (l + "-" + f).replace(/[^A-Za-z0-9]+/g, "-").replace(/^-+|-+$/g, "").toLowerCase();
      const namesFrom = sh => (sh.values || []).slice(1)
        .map(r => ({ last: String((r[0] || "")).replace(/[\*,()]+/g, "").trim(), first: String((r[1] || "")).trim() }))
        .filter(x => x.last);
      const liveActive = namesFrom(act);
      const liveInactive = namesFrom(inact);
      const liveActiveKeys = new Set(liveActive.map(p => slug(p.last, p.first)));
      const liveInactiveKeys = new Set(liveInactive.map(p => slug(p.last, p.first)));
      // Stub records for providers that exist in the live Excel but aren't yet in the baked data.json.
      // The dashboard will show them as "pending — awaiting roster data" until the next full regen.
      const seedKeys = new Set((data.items || []).filter(i => i.scope === "provider").map(i => i.entityKey));
      const newProviders = liveActive.filter(p => !seedKeys.has(slug(p.last, p.first)));
      const inactivated = (data.items || []).filter(i => i.scope === "provider" && i.active && liveInactiveKeys.has(i.entityKey)).map(i => i.entityKey);
      const removed = (data.items || []).filter(i => i.scope === "provider" && i.active && !liveActiveKeys.has(i.entityKey) && !liveInactiveKeys.has(i.entityKey)).map(i => i.entityKey);
      const dates = xl.expiryDatesFromValues(act.values);
      const delta = { generatedAt: new Date().toISOString(), newProviders, inactivated: [...new Set(inactivated)], removed: [...new Set(removed)], dates };
      await writeJsonAt(token, drivePath("_Sentinel/roster_delta.json"), delta);
      // Live staff sync: read the CHER/Frisco RN+FD workbooks and refresh staff_delta.json.
      // Best-effort — never fail the roster sync if the staff workbooks are slow/locked.
      let staffCount = null;
      try { const staff = await xl.buildStaffItems(token); await writeJsonAt(token, drivePath("_Sentinel/staff_delta.json"), { generatedAt: new Date().toISOString(), items: staff }); staffCount = staff.length; }
      catch (se) { staffCount = "staff sync skipped: " + String(se.message || se).slice(0, 100); }
      res.status(200).json({ ok: true, delta: { newProviders: newProviders.length, inactivated: delta.inactivated.length, removed: delta.removed.length, dates: dates.length, staff: staffCount } });
    } catch (e) { res.status(500).json({ error: String(e.message || e) }); }
    return;
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
        const flat = sh => (sh.values || []).slice(1).map((row, idx) => ({ last: row[0] || "", first: row[1] || "", row: idx + 2 }))
          .filter(r => String(r.last).trim());
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
        const result = await xl.appendRow(token, xl.SHEET_ACTIVE, [last, first]);
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
            await ensureFolderIn(token, docsRoot(), "Sama Farooqui/Sentinel/Provider/" + folderName);
            folderCreate = { ok: true, name: folderName };
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
        // Restore to the Credentials sheet using the saved row values.
        for (const r of (entry.rows || [])) await xl.restoreRow(token, xl.SHEET_ACTIVE, r.values);
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
        const removed = await xl.hardDelete(token, entityKey, last, first);
        if (!removed.length) { res.status(404).json({ error: "not found in roster", tried: { last, first, entityKey } }); return; }
        // Delete the SharePoint Provider/<name>/ folder. Build the folder name from the row we
        // just removed (col 1 = last, col 2 = first) — matches whatever's in the Excel exactly.
        const r0 = removed[0].values;
        const folderName = ((r0 && r0[1]) ? r0[1] + " " : "") + (r0 ? r0[0] : "");
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
            entity: ((first || "") + " " + (last || "")).trim() || (removed[0].values && [removed[0].values[1], removed[0].values[0]].filter(Boolean).join(" ")),
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
  // Merge the live roster delta — surfaces Excel-roster changes without a code redeploy.
  // Placeholders carry a SharePoint folderLink so the QR/upload flow works immediately.
  if (tabs.includes("provider")) {
    const { applyRosterDelta } = require("../lib/delta");
    items = await applyRosterDelta(items);
  }
  const keys = new Set(items.map(i => i.entityKey));
  const entityFiles = {};
  for (const k in (data.entityFiles || {})) if (keys.has(k)) entityFiles[k] = data.entityFiles[k];
  const contacts = tabs.includes("facility") ? (data.contacts || []) : [];
  res.status(200).json(Object.assign({}, data, { items, entityFiles, contacts, allowedTabs: tabs }));
};
