const G = require("./graph");
const E = require("./evidence");

async function readEvidence(token) {
  const [detected, uploads, currentDates, oldDates] = await Promise.all([
    G.readJsonAt(token, G.drivePath("_Sentinel/auto_detected.json")),
    G.readJsonAt(token, G.drivePath("_Sentinel/uploads.json")),
    G.readDocsJsonAt(token, "Sama Farooqui/Sentinel/_State/doc_dates.json"),
    G.readJsonAt(token, G.drivePath("_Sentinel/doc_dates.json")),
  ]);
  const merged = {};
  for (const map of [detected, uploads]) {
    for (const id of Object.keys(map || {})) merged[id] = E.choose(merged[id], map[id]);
  }
  for (const cache of [oldDates, currentDates]) {
    for (const [id, d] of Object.entries((cache && cache.dates) || {})) {
      merged[id] = E.choose(merged[id], {
        url: d.link, name: d.name, date: d.expires, recordDate: d.recordDate,
        recurring: d.recurring, cadenceMonths: d.cadenceMonths, cadenceNote: d.cadenceNote,
      });
    }
  }
  return merged;
}
module.exports = { readEvidence };
