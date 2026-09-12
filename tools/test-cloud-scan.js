// Simulate Microsoft responses; the real network and files are never touched.
const assert = require("node:assert/strict");
const G = require("../lib/graph");
const state = {};
const baseUrl = "https://wcgtx.sharepoint.com/sites/CorporateArchivesDirectory/Shared%20Documents/";
const folder = "Sama Farooqui/Sentinel/State Readiness/Castle Hills";
G.accessToken = async () => "test-token";
G.readJsonAt = async (token, path) => state[path] || null;
G.writeJsonAt = async (token, path, data) => { state[path] = structuredClone(data); };
require("../lib/session").getSession = () => ({ admin: true });
require("../lib/master-intake").lock = async () => async () => {};
require("../lib/delta").applyRosterDelta = async () => [
  { id: "cola", scope: "facility", entity: "Castle Hills ER", entityKey: "castle-hills-er", category: "COLA", folderLink: baseUrl + folder },
  { id: "qapi", scope: "other", entity: "Castle Hills ER", entityKey: "castle-hills-er", category: "Meeting — QAPI", folderLink: baseUrl + folder + "/10. Quality Improvement Program" },
  { id: "staff-bls", scope: "staff", entity: "Jane Doe", entityKey: "ennis-rn-doe-jane", facility: "Urgent Care Ennis", role: "RN", category: "BLS" },
  { id: "provider-acls", scope: "provider", entity: "Alex Smith", entityKey: "smith-alex", category: "ACLS Certification", folderLink: baseUrl + "Sama Farooqui/Sentinel/Provider/Alex Smith" },
];
function file(name) { return { id: name, name, file: {}, parentReference: { id: "parent" }, webUrl: baseUrl + folder + "/" + name }; }
let deltaCalls = 0;
global.fetch = async url => {
  if (url.includes("/root:/") && url.includes("?$select=id,folder")) return Response.json({ id: "sentinel-root", folder: {} });
  if (url.includes("token=latest")) return Response.json({ "@odata.deltaLink": G.docsRoot() + "/root/delta?token=done" });
  if (url.includes("/items/sentinel-root/children")) return Response.json({ value: [{ id: "readiness", name: "State Readiness", folder: {} }, { id: "staff", name: "Staff", folder: {} }, { id: "providers", name: "Provider", folder: {} }] });
  if (url.includes("/items/staff/children")) return Response.json({ value: [{ id: "ennis", name: "Urgent Care Ennis", folder: {} }] });
  if (url.includes("/items/ennis/children")) return Response.json({ value: [{ id: "jane", name: "Doe, Jane_RN", folder: {} }] });
  if (url.includes("/items/jane/children")) return Response.json({ value: [file("BLS_01_01_2029.pdf")] });
  if (url.includes("/items/providers/children")) return Response.json({ value: [{ id: "alex", name: "Alex Smith", folder: {} }] });
  if (url.includes("/items/alex/children")) return Response.json({ value: [file("ACLS_01_01_2029.pdf")] });
  if (url.includes("/items/readiness/children")) return Response.json({ value: [{ id: "ch", name: "Castle Hills", folder: {} }] });
  if (url.includes("/items/ch/children")) return Response.json({ value: [file("COLA_02_26_2028.pdf"), file("COLA_01_17_2026.pdf"), file("QAPI_Minutes_08_01_2026.pdf")] });
  if (url.includes("/items/parent?")) return Response.json({ id: "parent", name: "08. Laboratory Services", parentReference: { path: "/drive/root:/" + folder } });
  if (url.includes("/root/delta")) {
    deltaCalls++;
    return Response.json({ value: [file("COLA_02_26_2028.pdf"), file("COLA_01_17_2026.pdf"), file("QAPI_Minutes_08_01_2026.pdf")], "@odata.deltaLink": G.docsRoot() + "/root/delta?token=done" });
  }
  throw new Error("Unexpected test network call: " + url);
};
const handler = require("../api/scan");
let result;
const res = { setHeader() {}, status() { return this; }, json(v) { result = v; } };
(async () => {
  await handler({ method: "GET", url: "/api/scan?rescan=1" }, res);
  assert.equal(result.ok, true, JSON.stringify(result));
  assert.equal(result.moreToScan, false);
  assert.equal(result.inSentinel, 5);
  const detected = state[G.drivePath("_Sentinel/auto_detected.json")];
  assert.equal(detected.cola.date, "2028-02-26");
  assert.ok(detected.cola.url.endsWith("COLA_02_26_2028.pdf"));
  assert.equal(detected.qapi.date, "2026-09-01");
  assert.equal(detected["staff-bls"].date, "2029-01-01");
  assert.equal(detected["provider-acls"].date, "2029-01-01");
  assert.equal(deltaCalls, 0, "full scan stays inside Sentinel instead of walking the entire corporate drive");
  await handler({ method: "GET", url: "/api/scan" }, res);
  assert.equal(result.ok, true);
  assert.equal(deltaCalls, 1, "subsequent passes use the captured incremental cursor");
  console.log("PASS cloud scan resolves missing parent paths, retains newest COLA, matches Other across facility sections, and reports completion accurately.");
})().catch(e => { console.error(e); process.exitCode = 1; });
