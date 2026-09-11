// Document -> credential matching rules, in ONE place.
//
// Used by api/scan.js (the live Graph delta scan) and by tools/refresh-data.js (the local pass
// that reads the synced folder straight off disk). A document has to be matched to the same
// credential whichever route finds it — api/digest.js once kept a private copy of the roster
// regen, drifted from the shared one, and corrupted the board nightly. Not doing that again.
const FAC = require("./facility");

const FILE_RULES = {
  "ACLS Certification": "\\bacls\\b", "ATLS Certification": "\\batls\\b", "PALS Certification": "\\bpals\\b",
  "BLS Certification": "\\bbls\\b", "State Medical License": "tmb[ ]*cert|medical license|tmb certificate|tmb[ ]+\\d",
  "Medical License Verify (annual)": "tmb[ ]*ver|tmb veri",
  // "^dea " must NOT swallow "DEA Verification ..." — that catch-all was matching the annual
  // verification letters and filing them (and their dates) under the registration itself, which
  // both mislabels the document and writes the wrong expiry into the master Excel.
  "Individual DEA Registration": "dea[ ]*cert|dea certificate|dea[ ]*reg|^dea (?!ver)",
  "DEA Verify (annual)": "dea[ ]*ver", "Influenza Vaccination": "flu|influenza",
  "TB Screening": "\\btb\\b|ppd|tubercul|quantiferon|\\bcxr\\b|chest", "Driver's License": "txdl|driver|drivers? lic|\\bdl\\b",
  "NPDB Query (2 yrs)": "npdb", "OIG / SAM Exclusion Check": "oig|sam |exclusion", "NPI Verification": "nppes|\\bnpi\\b",
  "TSCA Documents": "tsca", "CME (20 hrs / 2 yrs)": "\\bcme\\b", "Delineation of Privileges (DOP)": "privilege|\\bdop\\b",
  "Peer References": "reference|peer", "Initial Application": "application|initial app",
  "CV / Resume": "\\bcv\\b|resume|curriculum", "Medical Diploma": "diploma|medical school|ecfmg",
  "Malpractice / COI Insurance": "malpractice|certificate of insurance|\\bcoi\\b|tail coverage|policy",
  // Board cert files are commonly named ONLY by the board acronym (e.g. "ABEM_2027.pdf"),
  // so we accept the common boards in addition to literal "board"/"recert".
  "Board Certification": "board|recert|\\b(abem|abfm|abim|abps|aobem|aobim|aboem|abog|abpn|abs|abucm|aagp|abo|abr)\\b",
  // Facility credentials. Without these, nothing matched ANY facility category, so every
  // facility document fell through to the generic supplemental path and whatever date was in
  // its filename became an expiry.
  ...FAC.FACILITY_FILE_RULES,
};

// Normalize a filename for rule matching. Underscores were handled but hyphens and dots were
// not, so "DEA-Cert-2027.pdf" and "Medical.License.2027.pdf" matched nothing while the
// underscore spellings matched — the same document filed two ways behaved differently.
// The blanket /ii/->/i/ also mangled real words ("Hawaii" -> "Hawai"), so it is limited to the
// one thing it was for: the roster's recurring "Verifiy" misspelling.
function normf(s) {
  return String(s).toLowerCase()
    .replace(/[_\-.]+/g, " ")
    .replace(/ii/g, "i")
    .replace(/\s+/g, " ")
    .trim();
}

// Which credential, among the items filed in one folder, a filename is evidence for.
function matchByRules(items, fileName) {
  const nf = normf(fileName);
  for (const it of items) {
    const rule = FILE_RULES[it.category];
    if (rule && new RegExp(rule, "i").test(nf)) return it;
  }
  return null;
}

module.exports = { FILE_RULES, normf, matchByRules };
