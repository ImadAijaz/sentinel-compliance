// Attached-document map used by the dashboard and the private provider portal.
// Every response is filtered to the signed-in staff member's allowed tabs, or to the single
// provider named by a private provider cookie. Nobody can read or link another person's files.
const { accessToken, readJsonAt, writeJsonAt, drivePath, dateFromName } = require("../lib/graph");
const { getSession, getProviderSession } = require("../lib/session");
const { applyRosterDelta } = require("../lib/delta");
const data = require("../data.json");

const UPLOADS = drivePath("_Sentinel/uploads.json");
const DETECTED = drivePath("_Sentinel/auto_detected.json");
const SUPP = drivePath("_Sentinel/supplemental_detected.json");

function allowedItem(item, staff, provider) {
  if (!item) return false;
  if (provider) return item.scope === "provider" && item.entityKey === provider.entityKey;
  if (!staff) return false;
  return !!(staff.admin || (Array.isArray(staff.tabs) && staff.tabs.includes(item.scope)));
}
function pickMap(map, allowedIds) {
  const out = {};
  Object.keys(map || {}).forEach(id => { if (allowedIds.has(id)) out[id] = map[id]; });
  return out;
}
async function bodyOf(req) {
  if (req.body && typeof req.body === "object") return req.body;
  if (typeof req.body === "string") { try { return JSON.parse(req.body); } catch (e) { return {}; } }
  let raw = "";
  await new Promise(resolve => { req.on("data", c => raw += c); req.on("end", resolve); });
  try { return JSON.parse(raw || "{}"); } catch (e) { return {}; }
}

module.exports = async (req, res) => {
  res.setHeader("Cache-Control", "no-store");
  res.setHeader("X-Content-Type-Options", "nosniff");
  if (req.method === "OPTIONS") { res.status(204).end(); return; }

  const staff = getSession(req);
  const provider = getProviderSession(req);
  if (!staff && !provider) { res.status(401).json({ ok: false, message: "private link or sign-in required" }); return; }

  try {
    const token = await accessToken();
    const [detected, uploads, suppMap, roster] = await Promise.all([
      readJsonAt(token, DETECTED), readJsonAt(token, UPLOADS), readJsonAt(token, SUPP),
      applyRosterDelta(data.items || []),
    ]);
    const supplementals = Object.values(suppMap || {});
    const all = roster.concat(supplementals);
    const allowed = all.filter(i => allowedItem(i, staff, provider));
    const byId = new Map(allowed.map(i => [i.id, i]));
    const allowedIds = new Set(byId.keys());

    if (req.method === "GET") {
      // App uploads win for URL/name, but keep a date from either source.
      const merged = Object.assign({}, pickMap(detected, allowedIds));
      const selectedUploads = pickMap(uploads, allowedIds);
      for (const id in selectedUploads) {
        const prev = merged[id] || {};
        merged[id] = Object.assign({}, prev, selectedUploads[id]);
        if (!merged[id].date && prev.date) merged[id].date = prev.date;
      }
      res.status(200).json({
        attachments: merged,
        supplemental: supplementals.filter(i => allowedItem(i, staff, provider)),
      });
      return;
    }

    if (req.method === "POST") {
      const b = await bodyOf(req);
      if (!b.item_id || !b.url) { res.status(400).json({ ok: false, message: "need item_id and url" }); return; }
      const target = byId.get(String(b.item_id));
      if (!target) { res.status(403).json({ ok: false, message: "that credential is outside this private access" }); return; }

      // Only a file stored in WCGTX's own SharePoint tenant can become compliance proof.
      const TENANT_HOST = (process.env.MS_SHAREPOINT_HOST || "wcgtx.sharepoint.com").toLowerCase();
      let host = "";
      try { host = new URL(String(b.url)).hostname.toLowerCase(); } catch (e) { host = ""; }
      if (host !== TENANT_HOST) { res.status(400).json({ ok: false, message: "url must point to " + TENANT_HOST }); return; }

      const map = uploads || {};
      map[b.item_id] = {
        url: String(b.url), name: String(b.name || "").slice(0, 300),
        date: dateFromName(b.name || ""),
      };
      await writeJsonAt(token, UPLOADS, map);
      res.status(200).json({ ok: true });
      return;
    }
    res.status(405).json({ ok: false, message: "GET or POST only" });
  } catch (e) {
    res.status(500).json({ ok: false, message: String(e.message || e) });
  }
};

module.exports.allowedItem = allowedItem;
