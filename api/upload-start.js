// Opens a Microsoft Graph upload session in the item's OneDrive folder (app-only).
// The browser then PUTs the file bytes straight to the returned uploadUrl, so large
// phone photos never hit Vercel's request-size limit.
//
// SECURITY: only a signed-in Sentinel user may ask the app-only Graph identity to create an
// upload session. The destination is also derived SERVER-SIDE from the requested item's record;
// a caller-supplied path is never trusted.
const { accessToken, encPath, docsRoot, ensureFolderIn } = require("../lib/graph");
const { getSession } = require("../lib/session");
const data = require("../data.json");
const { applyRosterDelta } = require("../lib/delta");
const FAC = require("../lib/facility");
const { providerPhaseForCategory } = require("../lib/filerules");

// The 6 SOP phase subfolders a document may legitimately be filed into.
const PHASES = [
  "1. Application & Document Collection", "2. Primary Source Verification",
  "3. Background & Compliance Review", "4. Medical Staff Review",
  "5. Payer Enrollment & Facility Setup", "6. Approval & Ongoing Monitoring",
];

const SITE_DIR = {
  "Castle Hills ER": "Castle Hills", "Frisco ER": "Frisco",
  "Urgent Care Ennis": "Urgent Care Ennis", "Urgent Care Plano": "Urgent Care Plano",
};
function safeSeg(s, fallback) {
  const v = String(s || "").replace(/[\\/:*?"<>|]/g, " ").replace(/\s+/g, " ").trim();
  return v || fallback;
}
function staffFolder(item) {
  const toks = String(item.entity || "").trim().split(/\s+/).filter(Boolean);
  const last = toks.length ? toks[toks.length - 1] : "Staff";
  const first = toks.slice(0, -1).join(" ");
  const role = /front desk/i.test(item.role || "") ? "FD" : (item.role || "Staff");
  return safeSeg(last + (first ? ", " + first : "") + "_" + role, "Staff");
}
function targetFolderForItem(item) {
  const base = "Sama Farooqui/Sentinel";
  if (item.scope === "provider") {
    const phase = PHASES[providerPhaseForCategory(item.category)] || PHASES[0];
    return base + "/Provider/" + safeSeg(item.entity, item.entityKey || "Provider") + "/" + phase + "/" + safeSeg(item.category, "Misc");
  }
  if (item.scope === "staff") {
    const site = SITE_DIR[item.staffFacility || item.facility] || safeSeg(item.staffFacility || item.facility, "Unassigned");
    return base + "/Staff/" + site + "/" + staffFolder(item) + "/" + safeSeg(item.category, "Misc");
  }
  if (item.scope === "facility" || item.scope === "other") {
    const site = SITE_DIR[item.entity] || safeSeg(item.entity, "Unassigned");
    const classified = FAC.classifySection("", item.category);
    const section = classified ? classified.section : FAC.STATE_SECTIONS[0];
    return base + "/State Readiness/" + site + "/" + section + "/" + safeSeg(item.category, "Other");
  }
  return null;
}

module.exports = async (req, res) => {
  if (req.method === "OPTIONS") { res.status(204).end(); return; }

  const session = getSession(req);
  if (!session) { res.status(401).json({ ok: false, message: "sign-in required" }); return; }

  try {
    const url = new URL(req.url, "http://localhost");
    const itemId = url.searchParams.get("item") || "";
    const entityKey = url.searchParams.get("e") || "";
    const phase = url.searchParams.get("phase") || "";
    const name = (url.searchParams.get("name") || "upload.bin").replace(/[^A-Za-z0-9 ._-]/g, "_");
    if (!itemId && !entityKey) { res.status(400).json({ ok: false, message: "missing item" }); return; }

    // Resolve the destination from OUR data, not from the caller.
    const allItems = await applyRosterDelta(data.items || []);
    let target = null;
    if (itemId) target = allItems.find(i => i.id === itemId) || null;
    if (!target && entityKey) target = allItems.find(i => i.entityKey === entityKey) || null;
    if (!target) { res.status(404).json({ ok: false, message: "unknown item" }); return; }
    const tabs = Array.isArray(session.tabs) ? session.tabs : [];
    if (!session.admin && !tabs.includes(target.scope)) {
      res.status(403).json({ ok: false, message: "this account cannot upload to that scope" }); return;
    }

    const token = await accessToken();
    // Always file new uploads into the organized SharePoint tree using the tracked item's scope,
    // entity and category.  The former folderLink path put every staff upload at the facility
    // root and provider uploads at the provider root, where the scanner could not attribute them.
    // A caller cannot influence this path; all segments come from our own item record.
    let rootUrl = docsRoot();
    let folderPath = targetFolderForItem(target);
    if (!folderPath) { res.status(409).json({ ok: false, message: "unsupported item scope" }); return; }
    // Retain the explicit admin import phase only when it is one of the six known values.
    if (target.scope === "provider" && phase && PHASES.includes(phase)) {
      const seg = folderPath.split("/"); seg[4] = phase; folderPath = seg.join("/");
    }
    await ensureFolderIn(token, rootUrl, folderPath);

    const filePath = folderPath + "/Sentinel_Upload_" + name;
    const r = await fetch(rootUrl + "/root:/" + encPath(filePath) + ":/createUploadSession", {
      method: "POST",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename" } })
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error((j.error && j.error.message) || ("Graph HTTP " + r.status));
    res.status(200).json({ ok: true, uploadUrl: j.uploadUrl, path: filePath });
  } catch (e) {
    res.status(200).json({ ok: false, message: String(e.message || e) });
  }
};

module.exports.targetFolderForItem = targetFolderForItem;
