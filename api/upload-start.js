// Opens a Microsoft Graph upload session in the item's OneDrive folder (app-only).
// The browser then PUTs the file bytes straight to the returned uploadUrl, so large
// phone photos never hit Vercel's request-size limit.
//
// SECURITY: this endpoint is deliberately reachable WITHOUT a login, because the QR-code flow
// lets a provider upload their own document from a phone. It used to take the destination folder
// straight from the query string — and the app's Graph token is tenant-wide Files.ReadWrite.All,
// so anyone on the internet could open an upload session into ANY folder in the company's
// SharePoint/OneDrive. The destination is now derived SERVER-SIDE from the requested item's own
// record; a caller-supplied path is never trusted.
const { accessToken, drivePath, encPath, driveRoot, ensureFolder, docsRoot, docsPathFromUrl, ensureFolderIn } = require("../lib/graph");
const data = require("../data.json");
const { applyRosterDelta } = require("../lib/delta");

// The 6 SOP phase subfolders a document may legitimately be filed into.
const PHASES = [
  "1. Application & Document Collection", "2. Primary Source Verification",
  "3. Background & Compliance Review", "4. Medical Staff Review",
  "5. Payer Enrollment & Facility Setup", "6. Approval & Ongoing Monitoring",
];

module.exports = async (req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  if (req.method === "OPTIONS") { res.status(204).end(); return; }

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

    const folder = target.folderLink || target.fileLink || "";
    if (!folder) { res.status(409).json({ ok: false, message: "no folder on record for this item yet" }); return; }

    const token = await accessToken();
    // Absolute SharePoint URLs land in the docs library; legacy relative paths resolve to the
    // OneDrive ROOT. Both are values we stored ourselves, so neither is attacker-controlled.
    const docsPath = docsPathFromUrl(folder);
    let rootUrl, folderPath;
    if (docsPath != null) { rootUrl = docsRoot(); folderPath = docsPath; }
    else { rootUrl = driveRoot(); folderPath = drivePath(folder); }
    // Optional phase subfolder — only ever one of the six known names.
    if (phase && PHASES.includes(phase)) folderPath = folderPath + "/" + phase;
    if (docsPath != null) await ensureFolderIn(token, rootUrl, folderPath);
    else await ensureFolder(token, folderPath);

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
