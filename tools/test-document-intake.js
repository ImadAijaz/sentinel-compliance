// Offline safety checks. No Microsoft calls, no real document writes.
const assert = require("node:assert/strict");
const E = require("../lib/evidence");
const G = require("../lib/graph");
const F = require("../lib/master-filing");
const FAC = require("../lib/facility");
let checks = 0;
function check(name, fn) { fn(); checks++; console.log("PASS " + name); }
check("older attachment cannot regress a 2028 expiry or proof link", () => {
  const seed = { expires: "2028-02-26", fileLink: "new", isFile: true };
  const out = E.apply(seed, { date: "2026-01-17", url: "old" });
  assert.equal(out.expires, "2028-02-26"); assert.equal(out.fileLink, "new"); assert.deepEqual(seed, { expires: "2028-02-26", fileLink: "new", isFile: true });
});
check("newest proof/date stay paired regardless of arrival order", () => {
  const old = { url: "old", date: "2026-01-17" }, renewed = { url: "new", date: "2028-02-26" };
  assert.deepEqual(E.choose(old, renewed), renewed); assert.deepEqual(E.choose(renewed, old), renewed);
  assert.deepEqual(E.choose(renewed, { url: "undated" }), renewed);
});
check("undated proof and archived proof never hide a tracked expiry", () => {
  assert.equal(E.apply({ expires: "2020-01-01" }, { url: "proof" }).expires, "2020-01-01");
  assert.equal(E.apply({ expires: "2027-01-01" }, { url: "proof", date: "2029-01-01" }, true).expires, "2027-01-01");
});
check("shared parser handles underscore MDY and ISO dates, not impossible dates", () => {
  assert.equal(G.dateFromName("COLA_02_26_2028_CHER.pdf"), "2028-02-26");
  assert.equal(G.dateFromName("COLA_2028_02_26_CHER.pdf"), "2028-02-26");
  assert.equal(G.dateFromName("COLA_02_30_2028.pdf"), null);
  assert.equal(G.dateFromName("phone_12028_02_26.pdf"), null);
});
check("archives excluded; live z-prefixed vendor and health folders kept", () => {
  for (const p of ["z.Expired Docs", "Z_Expired_Documents_FriscoER", ".Expired Health Docs", "Licenses Expired", "zArchive", "ZZZ Archive", "z.Superseded"]) assert.equal(FAC.isArchivedSegment(p), true, p);
  for (const p of ["Z_Firetrol", "z.health", "z.Admin Office Staff"]) assert.equal(FAC.isArchivedSegment(p), false, p);
  assert.equal(F.excluded("..Frisco ER/HR ONLY"), "restricted HR folder");
});
check("COLA files into laboratory at both urgent cares", () => {
  for (const site of ["Urgent Care Ennis", "Urgent Care Plano"]) {
    const out = F.classify({ path: ".." + site + "/..State Requirements_" + site + "/COLA_02_26_2028.pdf" }, () => null);
    assert.equal(out.family, "facility"); assert.equal(out.base, "State Readiness/" + site + "/08. Laboratory Services");
  }
});
check("provider photo classified to its provider; not an unknown guessed person", () => {
  const resolve = F.providerResolver([{ scope: "provider", entity: "Christopher Couch" }], []);
  assert.equal(resolve("Couch, Chris MD"), "Christopher Couch");
  assert.equal(resolve("Okorie Ikechukwu MD"), null);
  const out = F.classify({ path: "..WCGTX Master Physician File/Couch, Chris MD/Misc/Photos/Photo.jpg" }, resolve);
  assert.equal(out.family, "provider"); assert.equal(out.entity, "Christopher Couch");
});
check("untracked site and miscellaneous site folders never become staff", () => {
  assert.ok(F.classify({ path: "..Dallas UC/Office/Paper.pdf" }, () => null).skip);
  assert.ok(F.classify({ path: "..Frisco ER/Vendors/Acme/Paper.pdf" }, () => null).skip);
  assert.equal(F.classify({ path: "..Frisco ER/Nursing Staff/Doe, Jane_RN/License/RN.pdf" }, () => null).family, "staff");
});
const file = (id, path, hash = "same") => ({ id, path, hashes: { quickXorHash: hash }, size: 100, tag: "v1" });
check("identical file skipped, changed same-name file preserved beside original", () => {
  const source = file("s", "..Frisco ER/..State Requirements/COLA.pdf");
  const dest = file("d", "State Readiness/Frisco/08. Laboratory Services/COLA.pdf");
  assert.equal(F.plan([source], [dest], []).operations.length, 0);
  const p = F.plan([{ ...source, hashes: { quickXorHash: "new" } }], [dest], []);
  assert.equal(p.operations.length, 1); assert.match(p.operations[0].target, /COLA_Sentinel_[a-f0-9]+\.pdf$/);
});
check("colliding documents retain subfolders", () => {
  const p = F.plan([file("a", "..Frisco ER/..State Requirements/A/CLIA.pdf"), file("b", "..Frisco ER/..State Requirements/B/CLIA.pdf")], [], []);
  assert.equal(new Set(p.operations.map(o => o.target)).size, 2);
  assert.ok(p.operations.every(o => /\/(A|B)\/CLIA.pdf$/.test(o.target)));
});
check("event record and recurring obligation are separate", () => {
  assert.equal(FAC.isNonExpiring("QAPI_Minutes_08_01_2026.pdf"), true);
  const rows = [{ scope: "other", category: "Meeting — QAPI" }];
  const r = FAC.matchRecurringObligations(rows, "QAPI", "QAPI_Minutes_08_01_2026.pdf");
  assert.equal(r[0].recordDate, "2026-08-01"); assert.equal(r[0].nextDue, "2026-09-01");
});
console.log(checks + " offline safety checks passed. No live documents changed.");
