// Pure filing plan. No file I/O and no changes to the source folders.
const FAC = require("./facility");
const { createHash } = require("crypto");
const SITES = {
  "..Frisco ER": "Frisco", "..Castle Hills ER": "Castle Hills",
  "..Urgent Care Ennis": "Urgent Care Ennis", "..Urgent Care Plano": "Urgent Care Plano",
};
const PROVIDERS = "..WCGTX Master Physician File";
const PHASES = ["1. Application & Document Collection", "2. Primary Source Verification", "3. Background & Compliance Review"];
const strip = s => String(s || "").replace(/\b(MD|DO|PA|PA-?C|NP|FNP-?C|APRN|DDS|PHD|RN)\b\.?/gi, "").replace(/[*,]/g, " ").replace(/\s+/g, " ").trim();
const key = s => strip(s).toLowerCase().replace(/[^a-z0-9]+/g, "");
function display(raw) {
  const parts = String(raw).replace(/_credential(ing)?$/i, "").split(",");
  return strip(parts.length > 1 ? parts.slice(1).join(" ") + " " + parts[0] : raw);
}
function excluded(path) {
  const segs = String(path).split("/").filter(Boolean);
  if (segs.some(s => /^(?:z[._ ]*)?hr(?:[ _]|$)/i.test(s))) return "restricted HR folder";
  if (segs.some(FAC.isArchivedSegment)) return "archive folder";
  if (segs.some(s => /^\.{0,2}(?:folder template|z\.folder template)$/i.test(s))) return "template folder";
  return null;
}
function providerResolver(items, destinationFiles) {
  const names = new Map();
  // Existing organized names take priority over roster credential suffix variants.
  items.filter(i => i.scope === "provider").forEach(i => { if (i.entity) names.set(key(i.entity), i.entity); });
  destinationFiles.forEach(f => { const p = f.path.split("/"); if (p[0] === "Provider" && p[1]) names.set(key(p[1]), p[1]); });
  return raw => {
    // Non-comma ambiguity is intentionally not guessed (e.g. Okorie Ikechukwu MD).
    const n = display(raw), k = key(n);
    if (names.has(k)) return names.get(k);
    if (!String(raw).includes(",")) return null;
    const t = n.split(" "), first = t[0], last = t[t.length - 1];
    for (const initialOnly of [false, true]) {
      const candidates = [...names.entries()].filter(([canonicalKey, name]) => {
        const nt = strip(name).split(" ");
        return key(nt[nt.length - 1]) === key(last) &&
          (initialOnly ? key(nt[0])[0] === key(first)[0] : key(nt[0]) === key(first));
      });
      if (candidates.length === 1) return candidates[0][1];
      if (candidates.length > 1) return null;
    }
    return null;
  };
}
const aliases = { "z.health": "Health", licenses: "License", cme: "CMEs", certificates: "Certifications", miscellaneous: "Misc", cois: "COI" };
function category(s) { return aliases[String(s).toLowerCase()] || s || "Misc"; }
function providerPhase(cat, name) {
  const f = String(name).replace(/[_ .-]+/g, " ");
  if (cat.toLowerCase() === "sanction checks" || /\boig\b|npdb|\bsam\b|exclusion/i.test(f)) return 2;
  if (/verif|primary source/i.test(f)) return 1;
  return 0;
}
function classify(file, resolve) {
  const p = file.path.split("/"), name = p.pop();
  const skip = excluded(p.join("/"));
  if (skip) return { skip };
  if (/^(~\$|\.ds_store$|thumbs\.db$|desktop\.ini$)/i.test(name)) return { skip: "temporary file" };
  if (!/\.(pdf|jpg|jpeg|png|webp|gif|tif|tiff|heic|heif|doc|docx|xls|xlsx|ppt|pptx|msg|eml|txt)$/i.test(name)) return { skip: "unsupported file type" };
  if (p[0] === PROVIDERS && p.length >= 2) {
    if (p[1].startsWith(".")) return { skip: "provider reference folder" };
    const entity = resolve(p[1]);
    if (!entity) return { skip: "unmatched provider", owner: p[1] };
    const cat = category(p[2]);
    return { family: "provider", entity, base: ["Provider", entity, PHASES[providerPhase(cat, name)], cat].join("/"), deeper: p.slice(3), name };
  }
  const site = SITES[p[0]];
  if (!site) return { skip: "untracked site or roster/reference material" };
  if (/^\.{0,2}state requirements/i.test(p[1] || "")) {
    const sect = FAC.classifySection(p.slice(2).join("/"), name);
    return { family: "facility", entity: site, base: "State Readiness/" + site + "/" + (sect ? sect.section : FAC.STATE_SECTIONS[0]), deeper: p.slice(2), name };
  }
  // Only documented staff-group branches, never every miscellaneous site subfolder.
  if (/nursing staff|front office|front desk|midshift|admin office staff/i.test(p[1] || "") && p.length >= 3) {
    return { family: "staff", entity: site, base: ["Staff", site, p[2], category(p[3])].join("/"), deeper: p.slice(4), name };
  }
  return { skip: "unclassified site folder", owner: p.slice(0, 2).join("/") };
}
function sameContent(a, b) {
  if (!a || !b || a.size !== b.size) return false;
  for (const k of ["sha256Hash", "sha1Hash", "quickXorHash"]) {
    if (a.hashes && b.hashes && a.hashes[k] && b.hashes[k]) return a.hashes[k] === b.hashes[k];
  }
  return false; // file size and modified time are not content identity
}
function revisionName(name, source) {
  const fingerprint = JSON.stringify(source.hashes || {}) + "|" + source.id + "|" + source.tag;
  const suffix = createHash("sha256").update(fingerprint).digest("hex").slice(0, 12);
  return name.replace(/(\.[^.]+)$/, "_Sentinel_" + suffix + "$1");
}
function plan(source, destination, items) {
  const resolve = providerResolver(items, destination);
  const have = new Map(destination.map(f => [f.path.toLowerCase(), f]));
  const candidates = [], skips = {}, unmatched = new Set();
  for (const f of source) {
    const c = classify(f, resolve);
    if (c.skip) { skips[c.skip] = (skips[c.skip] || 0) + 1; if (c.owner) unmatched.add(c.owner); continue; }
    candidates.push({ ...c, source: f, target: c.base + "/" + c.name });
  }
  const collisions = new Map();
  candidates.forEach(c => { const k = c.target.toLowerCase(); collisions.set(k, (collisions.get(k) || 0) + 1); });
  const operations = [], counts = { provider: 0, staff: 0, facility: 0 }, already = { provider: 0, staff: 0, facility: 0 };
  const coveredSites = {};
  for (const c of candidates) {
    counts[c.family]++;
    if (c.family === "facility") coveredSites[c.entity] = (coveredSites[c.entity] || 0) + 1;
    if (collisions.get(c.target.toLowerCase()) > 1 && c.deeper.length) c.target = [c.base, ...c.deeper, c.name].join("/");
    const hit = have.get(c.target.toLowerCase());
    if (sameContent(c.source, hit)) { already[c.family]++; continue; }
    // Preserve a different existing file. No replace/delete, and no rename-on-conflict churn.
    if (hit) c.target = c.target.slice(0, -c.name.length) + revisionName(c.name, c.source);
    if (sameContent(c.source, have.get(c.target.toLowerCase()))) { already[c.family]++; continue; }
    operations.push({ family: c.family, source: c.source, target: c.target, status: "pending" });
  }
  return { operations, counts, already, skips, unmatched: [...unmatched], coveredSites };
}
module.exports = { SITES, PROVIDERS, classify, excluded, plan, sameContent, providerResolver, providerPhase, category };
