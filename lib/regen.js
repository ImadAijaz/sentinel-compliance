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
  const bakedActive = (data.items || []).filter(i => i.scope === "provider" && i.active);
  const inactivated = bakedActive.filter(i => liveInactiveKeys.has(i.entityKey)).map(i => i.entityKey);
  const removed = bakedActive.filter(i => !liveActiveKeys.has(i.entityKey) && !liveInactiveKeys.has(i.entityKey)).map(i => i.entityKey);
  const dates = xl.expiryDatesFromValues(act.values);

  // WRITE-SIDE SAFETY: refuse to persist a delta that would wipe the board. If the roster read
  // produced almost nothing, or would mark most known providers as gone, the parse went wrong —
  // keep the previous good delta rather than overwriting it with garbage. (lib/delta.js has a
  // matching read-side guard; this stops the bad data from ever being stored in the first place.)
  const wouldWipe = seedKeys.size > 0 && removed.length > seedKeys.size * 0.3;
  if (liveActive.length < 5 || wouldWipe) {
    return {
      ok: false, refused: true, providersRead: liveActive.length,
      reason: "roster parse looks wrong (" + liveActive.length + " providers read; " + removed.length +
        " of " + seedKeys.size + " known providers would be dropped) — kept the previous data instead of overwriting it",
    };
  }

  const delta = {
    generatedAt: new Date().toISOString(),
    newProviders, inactivated: [...new Set(inactivated)], removed: [...new Set(removed)], dates,
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
