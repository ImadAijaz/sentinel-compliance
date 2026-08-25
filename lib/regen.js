// ONE shared implementation of "re-read the master Excel and refresh the cached deltas".
//
// This used to exist as two near-identical copies: one in api/data.js (?regen=1, the admin
// "Sync from Excel" button) and one in api/digest.js (the daily cron). When the header-detection
// fix for the shifted-roster bug landed, only the api/data.js copy was updated — so the cron kept
// the old "column A = last name" reader and silently re-corrupted the dashboard every night.
// Both callers now share this module. Do not reintroduce a second copy.
const xl = require("./excel");
const { drivePath, writeJsonAt } = require("./graph");
const data = require("../data.json");

function slug(l, f) { return (l + "-" + f).replace(/[^A-Za-z0-9]+/g, "-").replace(/^-+|-+$/g, "").toLowerCase(); }

// Read provider names out of a sheet using HEADER-detected name columns — never a fixed col A/B.
// The roster can carry a leading tracking column that shifts the whole layout right.
function namesFrom(sh) {
  const { lastIdx, firstIdx } = xl.detectNameCols((sh.values || [])[0] || []);
  return (sh.values || []).slice(1)
    .map(r => ({
      last: String(r[lastIdx] == null ? "" : r[lastIdx]).replace(/[\*,()]+/g, "").trim(),
      first: String(r[firstIdx] == null ? "" : r[firstIdx]).trim(),
    }))
    .filter(x => x.last);
}

// Rebuild roster_delta.json (+ staff_delta.json) from the live workbooks.
// Returns { ok:true, ...counts } or { ok:false, refused:true, reason } when the parse looks wrong.
async function regenerateRoster(token) {
  const [act, inact] = await Promise.all([
    xl.readSheet(token, xl.SHEET_ACTIVE),
    xl.readSheet(token, xl.SHEET_INACTIVE),
  ]);
  const liveActive = namesFrom(act);
  const liveInactive = namesFrom(inact);
  const liveActiveKeys = new Set(liveActive.map(p => slug(p.last, p.first)));
  const liveInactiveKeys = new Set(liveInactive.map(p => slug(p.last, p.first)));
  const seedKeys = new Set((data.items || []).filter(i => i.scope === "provider").map(i => i.entityKey));
  const newProviders = liveActive.filter(p => !seedKeys.has(slug(p.last, p.first)));
  // Read `active` off the baked data only — never off objects a previous merge may have mutated.
  // NOTE: data.json holds one ITEM PER CREDENTIAL (~40 rows per provider), so these must be
  // reduced to UNIQUE PROVIDER KEYS. Counting raw items made the safety check below compare
  // ~727 item-rows against 180 providers and refuse every legitimate sync.
  const bakedActive = (data.items || []).filter(i => i.scope === "provider" && i.active);
  const inactivated = [...new Set(bakedActive.filter(i => liveInactiveKeys.has(i.entityKey)).map(i => i.entityKey))];
  const removed = [...new Set(bakedActive.filter(i => !liveActiveKeys.has(i.entityKey) && !liveInactiveKeys.has(i.entityKey)).map(i => i.entityKey))];
  const dates = xl.expiryDatesFromValues(act.values);

  // WRITE-SIDE SAFETY: refuse to persist a delta built from a MISREAD roster.
  //
  // The failure this exists to catch is a column shift (a "THHS Email Sent" tracking column
  // inserted before the names), which turns every surname into label text — so the names look
  // like labels AND essentially every known provider stops matching.
  //
  // It must NOT fire on ordinary churn. The baked data.json is a point-in-time bake, the
  // workbook is edited daily, and providers legitimately come and go — so a moderate number of
  // dropped providers is expected and healthy. Only a near-total mismatch indicates a parse
  // failure, hence 60% rather than 30%.
  const isJunk = s => { const n = String(s || "").trim().toLowerCase(); return !!n && (/\b(thhs|email sent|requested|rqstd|verify|verifiy|yrly|exp date)\b/.test(n) || /^\d+$/.test(n.replace(/\s+/g, ""))); };
  const junkNames = liveActive.filter(p => isJunk(p.last) || isJunk(p.first)).length;
  const namesUnreadable = liveActive.length > 0 && junkNames > liveActive.length * 0.3;
  const nearTotalMismatch = seedKeys.size > 0 && removed.length > seedKeys.size * 0.6;
  if (liveActive.length < 5 || namesUnreadable || nearTotalMismatch) {
    const why = liveActive.length < 5
      ? "only " + liveActive.length + " provider rows could be read from the workbook"
      : (namesUnreadable
        ? junkNames + " of " + liveActive.length + " names read as column labels rather than people (the roster layout has probably changed)"
        : removed.length + " of " + seedKeys.size + " known providers would be dropped at once");
    return {
      ok: false, refused: true, providersRead: liveActive.length, removedCount: removed.length,
      reason: "the master roster did not read correctly — " + why +
        ". Nothing was overwritten, so the dashboard still shows the last good data. Check the roster's first two columns are still Last Name / First Name.",
    };
  }

  const delta = {
    generatedAt: new Date().toISOString(),
    newProviders, inactivated, removed, dates,   // already de-duplicated to provider keys above
  };
  await writeJsonAt(token, drivePath("_Sentinel/roster_delta.json"), delta);

  // Staff sync is best-effort — a slow or locked RN/FD workbook must not fail the roster refresh.
  let staff = null;
  try {
    const items = await xl.buildStaffItems(token);
    await writeJsonAt(token, drivePath("_Sentinel/staff_delta.json"), { generatedAt: new Date().toISOString(), items });
    staff = items.length;
  } catch (se) { staff = "staff sync skipped: " + String(se.message || se).slice(0, 100); }

  return {
    ok: true, providersRead: liveActive.length, newProviders: newProviders.length,
    inactivated: delta.inactivated.length, removed: delta.removed.length, dates: dates.length, staff,
  };
}

module.exports = { regenerateRoster, namesFrom, slug };
