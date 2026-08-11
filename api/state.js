// Shared overlay (edits, notes, tasks, verify marks, watchlist, deletions) stored in
// OneDrive (_Sentinel/app_state.json) — replaces Supabase. Session-gated.
const { getSession } = require("../lib/session");
const { accessToken, readJsonAt, writeJsonAt, drivePath } = require("../lib/graph");
const PATH = drivePath("_Sentinel/app_state.json");

module.exports = async (req, res) => {
  res.setHeader("Cache-Control", "no-store");
  const s = getSession(req);
  if (!s) { res.status(401).json({ error: "sign-in required" }); return; }
  try {
    const token = await accessToken();
    if (req.method === "GET") {
      res.status(200).json((await readJsonAt(token, PATH)) || {});
      return;
    }
    if (req.method === "POST") {
      // Read-only users must not be able to write the shared overlay at all.
      if (s.readonly) { res.status(403).json({ error: "view-only access" }); return; }
      let raw = ""; await new Promise(r => { req.on("data", c => raw += c); req.on("end", r); });
      let body = null; try { body = JSON.parse(raw || "null"); } catch (e) { body = null; }
      // This is a WHOLE-FILE overwrite of the shared compliance overlay: every note, task,
      // renewal log, verify mark and the audit trail for all users. It accepted any body at all,
      // so a single empty POST silently erased the lot. Require a plausibly-shaped object, and
      // refuse a write that would drop the existing audit trail on the floor.
      if (!body || typeof body !== "object" || Array.isArray(body)) { res.status(400).json({ error: "expected a state object" }); return; }
      const KNOWN = ["edits", "added", "deleted", "logs", "watch", "audit", "leads", "tasks", "snapshots", "verified"];
      if (!KNOWN.some(k => k in body)) { res.status(400).json({ error: "state object has none of the expected keys" }); return; }
      const prev = (await readJsonAt(token, PATH)) || {};
      const prevAudit = Array.isArray(prev.audit) ? prev.audit.length : 0;
      const nextAudit = Array.isArray(body.audit) ? body.audit.length : 0;
      if (prevAudit > 20 && nextAudit < prevAudit * 0.5) {
        res.status(409).json({ error: "refused: this would discard " + (prevAudit - nextAudit) + " of " + prevAudit + " change-log entries. Reload the dashboard and try again." });
        return;
      }
      await writeJsonAt(token, PATH, body);
      res.status(200).json({ ok: true });
      return;
    }
    res.status(405).json({ error: "GET or POST only" });
  } catch (e) {
    res.status(200).json({ ok: false, message: String(e.message || e) });
  }
};
