// What the dashboard reads to show attached documents. Merges:
//   _Sentinel/auto_detected.json  (files dropped straight into OneDrive — from /api/scan)
//   _Sentinel/uploads.json        (files sent via the QR upload page — these win)
// GET returns the merged map; POST adds a QR-upload entry. All app-only, no Supabase.
const { accessToken, readJsonAt, writeJsonAt, drivePath, dateFromName } = require("../lib/graph");

const UPLOADS = drivePath("_Sentinel/uploads.json");
const DETECTED = drivePath("_Sentinel/auto_detected.json");
const SUPP = drivePath("_Sentinel/supplemental_detected.json");

module.exports = async (req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  if (req.method === "OPTIONS") { res.status(204).end(); return; }

  try {
    const token = await accessToken();
    if (req.method === "GET") {
      const [detected, uploads, suppMap] = await Promise.all([readJsonAt(token, DETECTED), readJsonAt(token, UPLOADS), readJsonAt(token, SUPP)]);
      // QR uploads win for url/name, but keep a date from EITHER source (so a server-OCR'd
      // expiry in auto_detected isn't lost when an older uploads.json entry has no date).
      const merged = Object.assign({}, detected || {});
      const u = uploads || {};
      for (const id in u) {
        const prev = merged[id] || {};
        merged[id] = Object.assign({}, prev, u[id]);
        if (!merged[id].date && prev.date) merged[id].date = prev.date;
      }
      // Supplemental new-file records (no matching tracked item) — full records, the
      // dashboard merges them into DATA at runtime so they appear without a regenerate.
      const supp = Object.values(suppMap || {});
      res.status(200).json({ attachments: merged, supplemental: supp });
      return;
    }
    if (req.method === "POST") {
      let raw = ""; await new Promise(r => { req.on("data", c => raw += c); req.on("end", r); });
      let b = {}; try { b = JSON.parse(raw || "{}"); } catch (e) {}
      if (!b.item_id || !b.url) { res.status(400).json({ ok: false, message: "need item_id and url" }); return; }
      // This endpoint records "document X is now on file for credential Y", and it has to stay
      // reachable without a login for the QR upload flow. Two checks stop it being used to
      // falsify compliance evidence, which it previously allowed outright:
      //   1. item_id must be a credential we actually track (no inventing records).
      //   2. the URL must live in OUR OWN SharePoint/OneDrive — never an arbitrary external
      //      link, which anyone could otherwise present as an official credential document.
      const data = require("../data.json");
      const { applyRosterDelta } = require("../lib/delta");
      const known = await applyRosterDelta(data.items || []);
      if (!known.some(i => i.id === b.item_id)) { res.status(404).json({ ok: false, message: "unknown item_id" }); return; }
      let host = "";
      try { host = new URL(String(b.url)).hostname.toLowerCase(); } catch (e) { host = ""; }
      const ok = /(^|\.)sharepoint\.com$/.test(host) || /(^|\.)wcgtx\.sharepoint\.com$/.test(host) ||
                 /(^|\.)onedrive\.com$/.test(host) || /(^|\.)1drv\.ms$/.test(host);
      if (!ok) { res.status(400).json({ ok: false, message: "url must point to the company SharePoint/OneDrive" }); return; }
      const map = (await readJsonAt(token, UPLOADS)) || {};
      map[b.item_id] = { url: String(b.url), name: String(b.name || "").slice(0, 300), date: dateFromName(b.name || "") };
      await writeJsonAt(token, UPLOADS, map);
      res.status(200).json({ ok: true });
      return;
    }
    res.status(405).json({ ok: false, message: "GET or POST only" });
  } catch (e) {
    res.status(200).json({ ok: false, message: String(e.message || e) });
  }
};
