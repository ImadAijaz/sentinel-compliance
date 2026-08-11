// Public, minimal endpoint for the provider self-service portal: returns ONLY the
// requested provider's items (never the whole roster), plus the item's siblings for
// the upload dropdown. Used by provider.html and upload.html (no login needed).
const data = require("../data.json");
const { applyRosterDelta } = require("../lib/delta");

// No folderLink/fileLink here. The upload pages no longer need it (they pass the item id and
// /api/upload-start resolves the destination server-side), and returning it to an anonymous
// caller handed out the exact SharePoint path of every provider's credential folder.
function slim(i) {
  return { id: i.id, category: i.category, authority: i.authority, expires: i.expires, isFile: i.isFile, entity: i.entity, entityKey: i.entityKey };
}

module.exports = async (req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Cache-Control", "no-store");
  const u = new URL(req.url, "http://localhost");
  const e = u.searchParams.get("e");
  const id = u.searchParams.get("item");
  // Reachable without a login (QR upload from a phone), so keep it to the one scope that flow
  // needs. It previously also returned scope:"staff" records, which the signed-in API gates.
  const isSelfService = i => i && i.scope === "provider";
  // Apply the live roster delta so providers added via the dashboard show up here
  // (no full Python regenerate needed for the QR upload page to recognize them).
  const allItems = (await applyRosterDelta(data.items || [])).filter(isSelfService);
  let ekey = e;
  let current = null;
  if (id) { current = allItems.find(i => i.id === id) || null; if (current) ekey = current.entityKey; }
  if (!ekey) { res.status(400).json({ error: "missing e or item" }); return; }
  const items = allItems.filter(i => i.entityKey === ekey).map(slim);
  if (!items.length) { res.status(404).json({ error: "not found" }); return; }
  res.status(200).json({ entity: items[0] ? items[0].entity : "", items, current: current ? slim(current) : null });
};
