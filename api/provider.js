// Private provider self-service portal and photo delivery.
// Staff issue a signed, seven-day link. Opening it creates a private provider cookie, after
// which the provider can see and upload only their own credential checklist.
const data = require("../data.json");
const { applyRosterDelta } = require("../lib/delta");
const {
  getSession, getProviderSession, signProviderAccess, verifyProviderAccess,
  providerCookieHeader,
} = require("../lib/session");
const {
  accessToken, readJsonAt, drivePath, docsPathFromUrl, docsRoot, encPath,
} = require("../lib/graph");

const SUPP = drivePath("_Sentinel/supplemental_detected.json");
const PORTAL_SECONDS = 60 * 60 * 24 * 7;

function slim(i) {
  return {
    id: i.id, category: i.category, authority: i.authority, expires: i.expires,
    isFile: i.isFile, entity: i.entity, entityKey: i.entityKey, active: i.active,
  };
}
function canUseProviders(session) {
  return !!(session && (session.admin || (Array.isArray(session.tabs) && session.tabs.includes("provider"))));
}
function photoLike(i) {
  return !!(i && i.scope === "provider" && i.isFile && i.fileLink &&
    /(?:^|\b)(?:photo|headshot|portrait)(?:\b|$)|badge[ _-]*photo/i.test(
      [i.category, i.uploadName, i.fileLink].filter(Boolean).join(" ")
    ));
}
async function readBody(req) {
  if (req.body && typeof req.body === "object") return req.body;
  if (typeof req.body === "string") { try { return JSON.parse(req.body); } catch (e) { return {}; } }
  let raw = "";
  await new Promise(resolve => { req.on("data", c => raw += c); req.on("end", resolve); });
  try { return JSON.parse(raw || "{}"); } catch (e) { return {}; }
}
async function providerItems() {
  // Newly scanned files (including photos) should reach the portal without waiting for a deploy.
  // If Microsoft is briefly unavailable, keep serving the baked roster instead of blanking it.
  const [baseAll, live] = await Promise.all([
    applyRosterDelta(data.items || []),
    (async () => {
      try {
        const token = await accessToken();
        return Object.values((await readJsonAt(token, SUPP)) || {}).filter(i => i && i.scope === "provider");
      } catch (e) { return []; }
    })(),
  ]);
  const base = baseAll.filter(i => i && i.scope === "provider");
  const byId = new Map();
  base.concat(live).forEach(i => { if (i && i.id) byId.set(i.id, i); });
  return Array.from(byId.values());
}
function sendRedirect(res, location, cookie) {
  if (cookie) res.setHeader("Set-Cookie", cookie);
  res.writeHead(302, { Location: location, "Cache-Control": "no-store" });
  res.end();
}
async function sendPhoto(res, item) {
  const path = docsPathFromUrl(item && item.fileLink);
  if (!path) { res.status(404).end(); return; }
  const token = await accessToken();
  const headers = { Authorization: "Bearer " + token };
  let r = await fetch(docsRoot() + "/root:/" + encPath(path) + ":/thumbnails/0/large/content", { headers });
  // Photos are sometimes PDFs. Graph's thumbnail endpoint turns those into an image. For a real
  // image file, fall back to the file itself if no thumbnail is available.
  if (!r.ok) {
    r = await fetch(docsRoot() + "/root:/" + encPath(path) + ":/content", { headers });
    if (!r.ok || !/^image\//i.test(r.headers.get("content-type") || "")) { res.status(404).end(); return; }
  }
  res.setHeader("Content-Type", r.headers.get("content-type") || "image/jpeg");
  res.setHeader("Cache-Control", "private, max-age=300");
  res.setHeader("X-Content-Type-Options", "nosniff");
  res.status(200).send(Buffer.from(await r.arrayBuffer()));
}

module.exports = async (req, res) => {
  res.setHeader("Cache-Control", "no-store");
  res.setHeader("X-Content-Type-Options", "nosniff");
  const u = new URL(req.url, "http://localhost");

  // Redeem the signed URL once, then remove its secret from the browser address bar.
  const portalToken = u.searchParams.get("portal");
  if (portalToken) {
    const access = verifyProviderAccess(portalToken);
    if (!access) { sendRedirect(res, "/provider.html?expired=1"); return; }
    sendRedirect(res, "/provider.html", providerCookieHeader(portalToken, PORTAL_SECONDS));
    return;
  }

  const staff = getSession(req);
  const provider = getProviderSession(req);

  // Staff-only link creation. The link contains no Microsoft or SharePoint credentials.
  if (req.method === "POST" && u.searchParams.get("issue") === "1") {
    if (!canUseProviders(staff)) { res.status(403).json({ error: "provider access required" }); return; }
    const body = await readBody(req);
    const entityKey = String(body.entityKey || "").trim();
    const all = await providerItems();
    const first = all.find(i => i.entityKey === entityKey);
    if (!first) { res.status(404).json({ error: "provider not found" }); return; }
    const token = signProviderAccess(entityKey, PORTAL_SECONDS);
    res.status(200).json({
      ok: true, entity: first.entity, expiresInDays: 7,
      path: "/api/provider?portal=" + encodeURIComponent(token),
    });
    return;
  }

  if (!staff && !provider) { res.status(401).json({ error: "private link required" }); return; }
  if (staff && !canUseProviders(staff)) { res.status(403).json({ error: "provider access required" }); return; }

  try {
    const all = await providerItems();
    const requested = String(u.searchParams.get("e") || "").trim();
    const itemId = String(u.searchParams.get("item") || "").trim();
    let ekey = provider ? provider.entityKey : requested;
    let current = null;
    if (itemId) {
      current = all.find(i => i.id === itemId) || null;
      if (current && !ekey) ekey = current.entityKey;
    }
    // A provider cookie is permanently locked to its own entity, regardless of query strings.
    if (provider) ekey = provider.entityKey;
    if (!ekey) { res.status(400).json({ error: "provider not specified" }); return; }
    if (current && current.entityKey !== ekey) current = null;

    const mine = all.filter(i => i.entityKey === ekey);
    if (!mine.length) { res.status(404).json({ error: "provider not found" }); return; }
    const photo = mine.find(photoLike) || null;
    if (u.searchParams.get("photo") === "1") { await sendPhoto(res, photo); return; }

    res.status(200).json({
      entity: mine[0].entity || "", entityKey: ekey, hasPhoto: !!photo,
      items: mine.map(slim), current: current ? slim(current) : null,
    });
  } catch (e) {
    res.status(500).json({ error: String(e.message || e) });
  }
};

module.exports.photoLike = photoLike;
module.exports.providerItems = providerItems;
