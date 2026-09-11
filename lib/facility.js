// Facility document filing rules — shared by api/data.js (facsync) and api/scan.js.
//
// WHY THIS EXISTS
// The facility documents the dashboard reads live under
//   Sama Farooqui/Sentinel/State Readiness/<Frisco|Castle Hills>/<NN. Section>/
// but the people who actually maintain them work in a DIFFERENT tree (the legacy
// "..State Requirements_<SITE>" folders, and increasingly in a personal OneDrive that the
// team library never sees). Sentinel therefore reads a stale snapshot: measured 2026-09-11,
// every facility file in the team library matched the legacy tree byte-for-byte, while the
// live COLA accreditation on the wall was two years newer than the one on the board.
//
// Two jobs here:
//   1. classifySection() — given a source path, decide which of the 14 State Readiness
//      sections a document belongs in, so an import lands somewhere a surveyor would look.
//   2. FACILITY_FILE_RULES — let the scanner match a facility file to the credential it is
//      evidence FOR. Without these, every facility document fell through to the generic
//      "supplemental" path, where any date in the filename becomes an expiry date. That is
//      how 57 sets of meeting minutes ended up on the board as expired credentials.

const STATE_SECTIONS = [
  "01. Licensing & Regulatory Compliance", "02. Personnel Files & Credentialing",
  "03. Medical Staff Services", "04. Patient Care & Clinical Documentation",
  "05. Medication Management", "06. Crash Cart & Emergency Equipment",
  "07. Infection Prevention & Control", "08. Laboratory Services",
  "09. Radiology Services", "10. Quality Improvement Program",
  "11. Environment of Care", "12. Emergency Preparedness",
  "13. Patient Rights & Compliance", "14. Daily Readiness Walkthrough",
];

// Archive folders. The old test only knew the provider tree's "z.Expired Docs" spelling, so
// the facility tree's "Z_Expired Documents_FriscoER" walked straight past it and its
// superseded 2025 facility licence showed on the board as an expired credential. Real folder
// names measured in this library: "z.Expired Docs", "Z_Expired Documents_FriscoER",
// "Expired Docs", "Expired Licenses", "Expired Certifications", "Expired Health Documents",
// "Licenses Expired", ".Expired Health Docs", "zArchive", "zArchiveRiggins", "Archive".
// Deliberately NOT matched: "Z_Firetrol" (a live vendor folder) and "Expired Meds Report"
// (a real report, and a file not a folder) — a bare "z" or "expired" prefix is not enough.
// Leading "." and "z"/"zz" prefixes are both used to sort these folders out of the way, and
// they combine (".Expired Health Docs", "z.Expired Docs", "Z_Expired Documents_FriscoER").
// "superseded" is in this list because the master sync retires documents into a "z.Superseded"
// folder. Without it the scanner treats that folder as live, re-detects every retired document
// as a current one, and puts back on the board exactly what was just taken off it.
const ARCHIVE_SEG = /^[.\s]*(z+[._ ]?)*(expired|archive|archived|inactive|old|superseded)\b|^z\.|^zz\.|^z+archive|\bexpired$/i;
function isArchivedSegment(seg) {
  return ARCHIVE_SEG.test(String(seg || "").trim());
}
// rel is a "/"-joined FOLDER path (no filename). True if any segment is an archive folder.
function isArchivedPath(rel) {
  return String(rel || "").split(/[\/\\]+/).filter(Boolean).some(isArchivedSegment);
}

// Ordered most-specific first: the first rule that matches wins. Each is tested against the
// full lowercased "folder/path/filename.ext", so a file inherits its folder's meaning
// ("Laboratory/CLIA_03_01_2026_CHER.pdf") and an explicitly named file can still override a
// vague folder.
const SECTION_RULES = [
  // 10. Quality Improvement — FIRST, and deliberately ahead of Laboratory. The QAPI meeting
  // packs are filed under a lab folder in several places; 50 of them were sitting in
  // "08. Laboratory Services" when this was written. They are quality documents.
  [9, /qapi|quality assurance|quality improvement|\bqa\/?pi\b|performance improvement|governing board|govern board|board m(ee)?t|front staff|staff m(ee)?t|_minutes|minutes_|\bminutes\b|\bagenda\b|addendum|conference call|incident report/i],
  // 09. Radiology
  [8, /x[\s_-]?ray|radiol|\bmrt\b|mammog|\bct\b[\s_-]?scan|ultrasound|sonograph|fluoroscop|radiation/i],
  // 08. Laboratory
  [7, /laborator|\blab\b|\bclia\b|\bcola\b|api[\s_-]?cop|sysmex|proficiency|\bqc\b|calibra|specimen|phlebotom|waived test/i],
  // 05. Medication Management
  [4, /pharmac|pharm\b|medication|\bmeds?\b|drug|controlled subs|\bdea\b.*(pharm|drug)|formulary|dispens/i],
  // 06. Crash Cart & Emergency Equipment
  [5, /crash cart|code cart|defibrillat|\baed\b|broselow|emergency equipment|airway cart/i],
  // 07. Infection Prevention & Control
  [6, /infection|sterili|autoclav|biohazard|sharps|hand hygiene|\bppe\b|exposure control|bloodborne|tubercul|immuniz|vaccin/i],
  // 12. Emergency Preparedness — fire/life-safety and disaster planning.
  [11, /\bfire\b|firetrol|alarm|sprinkler|extinguish|backflow|suppress|evacuat|disaster|emergency prep|emergency operation|drill|hazard vulnerab|severe weather|active shooter/i],
  // 11. Environment of Care — building systems, utilities, vendors, maintenance.
  [10, /generator|diversified|load test|\bhvac\b|utilit|biomed|preventive mainten|service (contract|agreement|mainten)|elevator|pest|medical waste|\bwaste\b|environment of care|\beoc\b|grounds|plumbing|electrical|work order|\bspbs\b|equipment/i],
  // 04. Patient Care & Clinical Documentation
  [3, /patient care|clinical (policy|protocol|documentation)|nursing (policy|protocol)|standing order|care plan|triage|chart audit|medical record/i],
  // 13. Patient Rights & Compliance
  [12, /patient right|grievance|complaint|\bhipaa\b|privacy notice|advance directive|consent form|\bemtala\b|non[\s_-]?discrimination|language (access|assist)/i],
  // 14. Daily Readiness Walkthrough
  [13, /daily readiness|walkthrough|walk through|rounding|daily check|readiness log|binders? (location|located)|binder loc/i],
  // 03. Medical Staff Services
  [2, /medical staff|bylaw|credential(ing)? committee|privileg|peer review|\bmec\b|call schedule|roster of provider|providers as of/i],
  // 02. Personnel Files & Credentialing
  [1, /personnel|employee file|\bhr\b|\bi-?9\b|offer letter|job description|orientation|competenc|in[\s_-]?service|staff as of|staff training/i],
  // 01. Licensing & Regulatory Compliance — the catch-all for licences, permits, certificates
  // and organizational/corporate paperwork. LAST of the positive rules on purpose.
  [0, /licen[cs]e|licensure|permit|accredit|certificat|registration|registry|nppes|\bnpi\b|\bein\b|occupancy|attestation|state requirement|organizational|corporat|articles of|tax|insurance|\bcoi\b|agreement|contract|\bnpdb\b|\boig\b|inspection report|regulatory|compliance|logo|safe[\s_-]?ready|administrative code|title 25|texas law|statute/i],
];

// Returns { idx, section, rule } or null when nothing matched confidently.
// `relFolder` is the path INSIDE the source folder (may be ""), `fileName` the leaf name.
function classifySection(relFolder, fileName) {
  const hay = (String(relFolder || "") + "/" + String(fileName || "")).toLowerCase();
  for (const [idx, re] of SECTION_RULES) {
    if (re.test(hay)) return { idx, section: STATE_SECTIONS[idx], rule: String(re) };
  }
  return null;
}

// Credential categories that real facility documents are evidence for. Same shape as the
// provider FILE_RULES in api/scan.js: category -> regex tested against a normalized filename.
// Category names must match the facility item categories in data.json exactly, or the scanner
// will not find the item to attach to.
const FACILITY_FILE_RULES = {
  "Facility License": "facility licen[cs]e|freestanding er|\\bhcfs\\b",
  "DEA for Facility License": "dea[\\s_-]*facility|facility[\\s_-]*dea",
  "DEA Power of Attorney": "power of attorney|\\bpoa\\b",
  "Certificate of Occupancy": "occupancy",
  "EIN for Facility": "\\bein\\b",
  "CLIA": "\\bclia\\b",
  "COLA": "\\bcola\\b",
  "Consultant Pharmacist License": "consult[\\s_-]*pharmacist",
  "Pharmacy License": "board of pharmacy|pharmacy (licen[cs]e|certificate)",
  "X-Ray Registration": "x[\\s_-]?ray registration|certificate of x[\\s_-]?ray",
  "Inspection — Generator Annual": "generator.*annual|annual.*generator|load test",
  "Inspection — Generator Semi-Annual": "generator.*semi|semi.*generator",
  "Fire Alarm Inspection": "(fire )?alarm (system )?(inspect|report)",
  "Sprinkler Inspection": "sprinkler",
  "Fire Extinguisher Inspection": "exting",
  "Backflow Test": "backflow",
  "Fire Inspection": "fire inspection|fire marshal",
  "X-Ray Inspection": "x[\\s_-]?ray[\\s_-]*inspect",
};

// Documents that record an EVENT rather than carrying an expiry. The date in
// "QAPI_Minutes_04_15_2026_FriscoER.pdf" is the date the meeting happened; minutes do not
// expire. The scanner used to take any date out of any filename and store it as `expires`, so
// the day after each meeting its minutes turned red on the board and started asking to be
// renewed. Measured 2026-09-11: 57 of the 65 "expired" facility items were documents of this
// kind — meeting minutes, inspection reports, service calls, "as of" staff snapshots, emails
// and incident reports. They still show, with their date, as a dated record.
const NON_EXPIRING = /\bminutes\b|\bagenda\b|addendum|\bmtg\b|\bmeeting\b|conference call|\bqapi\b|dashboard|\breport\b|work order|\bservice\b|maintenance|\binspection\b|\bas of\b|\bemail\b|incident|load test|backflow|acknowlg|acknowledg|\bbinders?\b|walkthrough|\blog\b|\bresults\b/i;
// A certificate/licence/registration always wins, even when the name also says "inspection" or
// "service" — "Sysmex Service Maintenance Certificate" is a certificate with a real end date.
const EXPIRING_OVERRIDE = /certificat|\blicen[cs]e\b|\blic\b|registration|accredit|\bpermit\b|\bexp\b/i;
function isNonExpiring(fileName) {
  // Their filenames separate words with "_" and ".", and "_" is a word character — so a plain
  // \bqapi\b never matched "QAPI_02_12_2026_January Data_CHER.pdf". Normalize first, the same
  // way api/scan.js normf() does for the provider rules.
  const n = String(fileName || "").replace(/\.[^.]+$/, "").replace(/[_\-.]+/g, " ");
  if (EXPIRING_OVERRIDE.test(n)) return false;
  return NON_EXPIRING.test(n);
}

module.exports = {
  STATE_SECTIONS, SECTION_RULES, classifySection,
  isArchivedPath, isArchivedSegment, FACILITY_FILE_RULES,
  isNonExpiring,
};
