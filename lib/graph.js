// ============================================================================
// Microsoft Graph helpers for Sentinel — APP-ONLY (client credentials).
// The app authenticates as itself (no user sign-in) and writes into the OneDrive
// where all the credential files live. Sama left the company, so there is no user
// to sign in; app-only access via the drive ID is the correct, durable approach.
// Secrets come ONLY from Vercel env vars: MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET.
// Requires the app to have the *Application* permission Files.ReadWrite.All (admin-consented).
// ============================================================================
const TENANT = process.env.MS_TENANT_ID;
const CLIENT_ID = process.env.MS_CLIENT_ID;
const CLIENT_SECRET = process.env.MS_CLIENT_SECRET;

// Sama's OneDrive (sfarooqui@wcgtx.com — where WCGTX Phyicians_04.08.2020 and all the
// real credential files live). Verified via app-only write. Not secret; env-overridable.
const DRIVE_ID = process.env.MS_DRIVE_ID ||
  "b!hICmGNzaFEiC8Z6vebrpNWzB937MR0tFsLlTxA2x3Z9-nxsW_blJTrLUhaL3IsBm";
// Files live under this single subfolder of that drive's root.
// Env-overridable so the app-state folder can be moved out of a former employee's personal
// OneDrive and into a durable team drive without another code change.
const ROOT = process.env.MS_DRIVE_ROOT || "WCGTX Phyicians_04.08.2020";
const GRAPH = "https://graph.microsoft.com/v1.0";

// DECOUPLED DOCS LOCATION: the reorganized Sentinel documents now live in the
// CorporateArchivesDirectory SharePoint library (a different drive). App-state JSON
// (allowed.json/uploads.json) stays on the OneDrive above so sign-in keeps working;
// only document upload targeting uses this docs drive. App-only Files.ReadWrite.All
// is tenant-wide, so the same token covers this drive too. Env-overridable.
const DOCS_DRIVE_ID = process.env.MS_DOCS_DRIVE_ID ||
  "b!21j9qMOtK0-ztrloaqcQnMnBmsH8JahImnJmCtRtYU-43O1iKIuQT4Zi09B1XG-y";
function docsRoot() { return GRAPH + "/drives/" + DOCS_DRIVE_ID; }

// Map an absolute SharePoint URL into the docs library to its path under the library
// root (everything after ".../Shared Documents/", URL-decoded). null if not such a URL.
function docsPathFromUrl(u) {
  const m = String(u || "").match(/\/Shared(?:%20| )Documents\/(.+)$/i);
  if (!m) return null;
  const raw = m[1].split(/[?#]/)[0].replace(/\/+$/, "");
  return raw.split("/").map(seg => { try { return decodeURIComponent(seg); } catch (e) { return seg; } }).join("/");
}

let tokenCache = null, tokenPending = null;
async function accessToken() {
  if (tokenCache && tokenCache.until > Date.now()) return tokenCache.value;
  if (tokenPending) return tokenPending;
  tokenPending = requestAccessToken();
  try { return await tokenPending; } finally { tokenPending = null; }
}
async function requestAccessToken() {
  if (!CLIENT_ID || !CLIENT_SECRET || !TENANT) throw new Error("MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET not set in Vercel env.");
  const r = await fetch("https://login.microsoftonline.com/" + TENANT + "/oauth2/v2.0/token", {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, grant_type: "client_credentials", scope: "https://graph.microsoft.com/.default" }).toString()
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(j.error_description || j.error || ("token HTTP " + r.status));
  tokenCache = { value: j.access_token, until: Date.now() + Math.max(0, (Number(j.expires_in) || 3600) - 120) * 1000 };
  return j.access_token;
}

function encPath(p) { return p.split("/").map(encodeURIComponent).join("/"); }
function driveRoot() { return GRAPH + "/drives/" + DRIVE_ID; }

// Map a dashboard fileLink/folderLink ("../..WCGTX Master Physician File/...") to a
// path under the drive root (prefixed with ROOT).
function drivePath(rel) {
  const clean = String(rel || "").replace(/^(\.\.\/)+/, "").replace(/^\/+|\/+$/g, "");
  return ROOT + (clean ? "/" + clean : "");
}

// Make sure every folder in a path exists (createUploadSession needs the parent to exist).
async function ensureFolderIn(token, rootUrl, folderPath) {
  const parts = String(folderPath).split("/").filter(Boolean);
  let parent = "";
  for (const name of parts) {
    const target = parent === "" ? rootUrl + "/root/children" : rootUrl + "/root:/" + encPath(parent) + ":/children";
    const res = await fetch(target, {
      method: "POST", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify({ name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" })
    });
    if (!res.ok && res.status !== 409) {
      const t = await res.text();
      if (!/nameAlreadyExists|already exists/i.test(t)) throw new Error("ensureFolder '" + name + "' HTTP " + res.status + " " + t.slice(0, 160));
    }
    parent = parent === "" ? name : parent + "/" + name;
  }
}
async function ensureFolder(token, folderPath) { return ensureFolderIn(token, driveRoot(), folderPath); }

// Read a JSON file from the drive (null if missing).
async function readJsonAt(token, path) {
  const r = await fetch(driveRoot() + "/root:/" + encPath(path) + ":/content", { headers: { Authorization: "Bearer " + token } });
  if (r.status === 404) return null;
  if (!r.ok) throw new Error("read " + path + " HTTP " + r.status);
  return await r.json().catch(() => null);
}
async function readDocsJsonAt(token, path) {
  const r = await fetch(docsRoot() + "/root:/" + encPath(path) + ":/content", { headers: { Authorization: "Bearer " + token } });
  if (r.status === 404) return null;
  if (!r.ok) throw new Error("read docs " + path + " HTTP " + r.status);
  return await r.json().catch(() => null);
}
// Write a JSON file to the drive (creates parent folders).
async function writeJsonAt(token, path, obj) {
  const parent = path.split("/").slice(0, -1).join("/");
  if (parent) await ensureFolder(token, parent);
  const r = await fetch(driveRoot() + "/root:/" + encPath(path) + ":/content", {
    method: "PUT", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" }, body: JSON.stringify(obj)
  });
  if (!r.ok) throw new Error("write " + path + " HTTP " + r.status);
}

// Extract an expiry date from a filename like "ACLS_09_18_2026.pdf" -> latest valid date as YYYY-MM-DD.
function dateFromName(name) {
  let best = null, m;
  const s = String(name || "");
  // ISO first: 2026-09-18 / 2026_09_18 / 2026.09.18
  const reIso = /(?<!\d)(\d{4})[_.\-](\d{1,2})[_.\-](\d{1,2})(?!\d)/g;
  while ((m = reIso.exec(s))) {
    let y = +m[1], mo = +m[2], d = +m[3];
    if (y >= 1900 && y <= 2199 && mo >= 1 && mo <= 12 && d >= 1 && d <= new Date(Date.UTC(y, mo, 0)).getUTCDate()) {
      const iso = y + "-" + String(mo).padStart(2, "0") + "-" + String(d).padStart(2, "0");
      if (!best || iso > best) best = iso;
    }
  }
  // M-D-Y / M_D_Y / M.D.Y
  const re = /(?<!\d)(\d{1,2})[_.\-](\d{1,2})[_.\-](\d{4}|\d{2})(?!\d)/g;
  while ((m = re.exec(s))) {
    let mo = +m[1], d = +m[2], y = +m[3]; if (y < 100) y += 2000;
    if (y >= 1900 && y <= 2199 && mo >= 1 && mo <= 12 && d >= 1 && d <= new Date(Date.UTC(y, mo, 0)).getUTCDate()) {
      const iso = y + "-" + String(mo).padStart(2, "0") + "-" + String(d).padStart(2, "0");
      if (!best || iso > best) best = iso;
    }
  }
  return best;
}

// Delete a provider's SharePoint folder under Sentinel/Provider/<name>/. The folder
// often doesn't exist yet (placeholder providers added via the dashboard before any
// files were uploaded), so a 404 is treated as "nothing to delete" — not an error.
async function deleteProviderFolder(token, providerName) {
  const name = String(providerName || "").trim();
  if (!name) return { ok: false, status: 0, error: "no name" };
  const path = "Sama Farooqui/Sentinel/Provider/" + name;
  const r = await fetch(docsRoot() + "/root:/" + encPath(path), {
    method: "DELETE",
    headers: { Authorization: "Bearer " + token }
  });
  if (r.status === 204 || r.status === 200) return { ok: true, status: r.status, name };
  if (r.status === 404) return { ok: true, status: 404, name, note: "folder did not exist" };
  let detail = ""; try { detail = (await r.text()).slice(0, 160); } catch (e) {}
  return { ok: false, status: r.status, name, error: detail };
}

module.exports = { TENANT, CLIENT_ID, CLIENT_SECRET, DRIVE_ID, ROOT, GRAPH, DOCS_DRIVE_ID, accessToken, encPath, driveRoot, docsRoot, docsPathFromUrl, drivePath, ensureFolder, ensureFolderIn, readJsonAt, readDocsJsonAt, writeJsonAt, dateFromName, deleteProviderFolder };
