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
];
function file(name) { return { id: name, name, file: {}, parentReference: { id: "parent" }, webUrl: baseUrl + folder + "/" + name }; }
let deltaCalls = 0;
global.fetch = async url => {
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
  assert.equal(result.inSentinel, 3);
  const detected = state[G.drivePath("_Sentinel/auto_detected.json")];
  assert.equal(detected.cola.date, "2028-02-26");
  assert.ok(detected.cola.url.endsWith("COLA_02_26_2028.pdf"));
  assert.equal(detected.qapi.date, "2026-09-01");
  assert.equal(deltaCalls, 1);
  console.log("PASS cloud scan resolves missing parent paths, retains newest COLA, matches Other across facility sections, and reports completion accurately.");
})().catch(e => { console.error(e); process.exitCode = 1; });
