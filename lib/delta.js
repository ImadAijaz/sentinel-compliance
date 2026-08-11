// Apply roster_delta.json on top of the baked data.json items, so newly-added providers
// (added via the dashboard "+ Add provider" or by a new Provider/<Name>/ folder) appear
// immediately without a full Python regenerate. Shared by /api/data and /api/provider.
const { accessToken, drivePath, readJsonAt } = require("./graph");

function slug(l, f) { return (l + "-" + (f || "")).replace(/[^A-Za-z0-9]+/g, "-").replace(/^-+|-+$/g, "").toLowerCase(); }
function spUrl(name) {
  return "https://wcgtx.sharepoint.com/sites/CorporateArchivesDirectory/Shared%20Documents/Sama%20Farooqui/Sentinel/Provider/"
    + name.split("/").map(encodeURIComponent).join("/");
}

const CRED = [
  ["State Medical License","Texas Medical Board",90], ["Individual DEA Registration","DEA",60],
  ["ACLS Certification","AHA",60], ["ATLS Certification","ACS",60], ["PALS Certification","AHA",60],
  ["BLS Certification","AHA",60], ["Board Certification","Specialty Board",180],
  ["Driver's License","Texas DPS",30], ["Medical Diploma","Medical School",0],
  ["Influenza Vaccination","Employee Health",14], ["TB Screening","Employee Health",60],
  ["CME (20 hrs / 2 yrs)","TMB / CME",60], ["TSCA Documents","WCGTX",60],
  ["NPDB Query (2 yrs)","NPDB",60], ["OIG / SAM Exclusion Check","OIG",14],
  ["NPI Verification","NPPES",0], ["Initial Application","WCGTX",0],
  ["CV / Resume","WCGTX",0], ["Delineation of Privileges (DOP)","WCGTX",0],
  ["Peer References","WCGTX",0], ["Malpractice / COI Insurance","Carrier",60],
];

async function applyRosterDelta(items) {
  // COPY each item, don't just copy the array. `items` is the module-cached require("../data.json")
  // array, which lives for the whole lifetime of a warm serverless instance. The merge below sets
  // `active`, `expires` and `pending` on these objects — with a shallow slice() those writes stuck
  // to the shared cache, so a date later corrected or cleared in Excel kept showing on the
  // dashboard (nothing resets a field the new delta simply omits), and two instances served
  // different data. Only provider items are mutated, so only those need cloning.
  let out = items.map(i => (i && i.scope === "provider") ? Object.assign({}, i) : i);
  try {
    const token = await accessToken();
    const delta = await readJsonAt(token, drivePath("_Sentinel/roster_delta.json"));
    if (!delta) return out;

    // ---- Corruption guard --------------------------------------------------------------------
    // A delta produced by a mis-read roster (e.g. a stray leading column shifting the layout, so
    // the "last name" cell reads a status value like "THHS Email Sent") can list almost every
    // real provider as "removed" and inject a junk "new provider" per row. That wipes the real
    // board and floods it with blank stubs. Detect that shape and skip the provider-side merge
    // entirely, falling back to the clean baked data.json. Staff sync (below) is unaffected.
    const bakedKeys = new Set(out.filter(i => i.scope === "provider").map(i => i.entityKey));
    // A "junk" name is spreadsheet text that cannot be a person: a status/column label, or an
    // all-digits cell. Deliberately NOT "contains a digit" — a single odd row must never be able
    // to suppress every legitimately-added provider.
    const isJunkName = s => {
      const n = String(s || "").trim().toLowerCase();
      if (!n) return false;
      return /\b(thhs|email sent|requested|rqstd|verify|verifiy|yrly|exp date)\b/.test(n) || /^\d+$/.test(n.replace(/\s+/g, ""));
    };
    const newList = delta.newProviders || [];
    const junkCount = newList.filter(p => isJunkName(p.last) || isJunkName(p.first)).length;
    // Systematic misparse (not one bad row): most of the "new" names are unusable label text.
    const newBad = newList.length >= 2 && junkCount > newList.length * 0.3;
    const removedHit = (delta.removed || []).filter(k => bakedKeys.has(k)).length;
    const removesTooMany = bakedKeys.size > 0 && removedHit > bakedKeys.size * 0.3;
    const deltaCorrupt = newBad || removesTooMany;

    if (!deltaCorrupt) {
      const inactSet = new Set(delta.inactivated || []);
      const goneSet = new Set(delta.removed || []);
      out = out.filter(i => !(i.scope === "provider" && goneSet.has(i.entityKey)));
      out.forEach(i => { if (i.scope === "provider" && inactSet.has(i.entityKey)) i.active = false; });
      const existingKeys = new Set(out.filter(i => i.scope === "provider").map(i => i.entityKey));
      newList.forEach(p => {
        // Skip individual unusable rows (a stray label in the name column) even when the delta
        // as a whole looks healthy — one bad row shouldn't become a phantom provider card.
        if (isJunkName(p.last) || isJunkName(p.first)) return;
        const ekey = slug(p.last, p.first || "");
        if (existingKeys.has(ekey)) return;
        const entity = ((p.first || "") + " " + p.last).trim();
        const folder = spUrl(entity);
        CRED.forEach(([cat, auth, lead]) => {
          out.push({
            id: slug(ekey, cat), scope: "provider", entity, entityKey: ekey,
            category: cat, authority: auth, renewalLeadDays: lead,
            expires: null, pending: true, isFile: false, active: true,
            fileLink: folder, folderLink: folder, owner: entity,
            notes: "New provider — awaiting roster data + uploads",
            liveAdded: true,
          });
        });
      });
    }
    // Apply expiry dates read straight from the Excel Credentials cells (so dates a human
    // types directly into the sheet show on the dashboard). Keyed by entityKey|category.
    // Skipped when the delta is corrupt — those keys/dates came from a mis-read layout.
    if (!deltaCorrupt && Array.isArray(delta.dates) && delta.dates.length) {
      const dmap = {};
      delta.dates.forEach(d => { if (d && d.entityKey && d.category && d.expires) dmap[d.entityKey + "|" + d.category] = d.expires; });
      out.forEach(i => {
        if (i.scope !== "provider") return;
        const hit = dmap[i.entityKey + "|" + i.category];
        if (hit) { i.expires = hit; i.pending = false; i.fromExcel = true; }
      });
    }
    // Live staff: when a staff_delta exists (written by Sync from Excel / daily regen), drop the
    // baked scope:staff items and use the fresh ones — so staff stay in sync with the RN/FD sheets.
    try {
      const sd = await readJsonAt(token, drivePath("_Sentinel/staff_delta.json"));
      if (sd && Array.isArray(sd.items) && sd.items.length) {
        // Same corruption guard the provider side gets. This REPLACES all baked staff data, so a
        // misparsed RN/FD workbook (the same leading-column edit that broke the physician roster)
        // would otherwise wipe the Staff tab and fill it with label text. Only swap in the live
        // staff when the names look like names and the count is in the right ballpark.
        const bakedStaff = out.filter(i => i.scope === "staff").length;
        const names = [...new Set(sd.items.map(i => String(i.entity || "").trim()))].filter(Boolean);
        const junk = names.filter(isJunkName).length;
        const sane = names.length > 0 && junk <= names.length * 0.3 &&
                     (bakedStaff === 0 || sd.items.length >= bakedStaff * 0.5);
        if (sane) out = out.filter(i => i.scope !== "staff").concat(sd.items);
      }
    } catch (e2) { /* keep baked staff if the live staff delta can't be read */ }
  } catch (e) { /* delta is optional; failures don't break the response */ }
  return out;
}

module.exports = { applyRosterDelta };
