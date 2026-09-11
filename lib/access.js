// Access list with PER-USER tab permissions. Stored in OneDrive (_Sentinel/allowed.json).
// Shape: { users: [ { email, tabs:["provider","staff","facility","other"], admin:bool } ] }
// (Old shape { emails:[...] } is auto-migrated to all-tabs users.)
const { accessToken, readJsonAt, writeJsonAt, drivePath } = require("./graph");
const PATH = drivePath("_Sentinel/allowed.json");
const ALL_TABS = ["provider", "staff", "facility", "other"];
const OWNER = "iaijaz@wcgtx.com";
const DEFAULT = [{ email: OWNER, tabs: ALL_TABS.slice(), admin: true }];

function normTabs(t) {
  const arr = Array.isArray(t) ? t.filter(x => ALL_TABS.includes(x)) : [];
  return arr.length ? arr : ALL_TABS.slice();
}

async function getUsers() {
  try {
    const token = await accessToken();
    const j = await readJsonAt(token, PATH);
    if (j && Array.isArray(j.users) && j.users.length) {
      return j.users.map(u => ({ email: String(u.email || "").toLowerCase(), tabs: normTabs(u.tabs), admin: !!u.admin || String(u.email || "").toLowerCase() === OWNER }));
    }
    if (j && Array.isArray(j.emails) && j.emails.length) { // migrate old format
      return j.emails.map(e => ({ email: String(e).toLowerCase(), tabs: ALL_TABS.slice(), admin: String(e).toLowerCase() === OWNER }));
    }
    return DEFAULT.slice();
  } catch (e) { return DEFAULT.slice(); }
}
async function setUsers(users) {
  const token = await accessToken();
  await writeJsonAt(token, PATH, { users });
}
function findUser(email, users) {
  const e = String(email || "").toLowerCase();
  return users.find(u => u.email === e) || null;
}
module.exports = { getUsers, setUsers, findUser, ALL_TABS, OWNER, DEFAULT };
