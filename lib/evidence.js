// One forward-only evidence policy for the scanner, dashboard and private portal.
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.SentinelEvidence = factory();
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  function object(v) { return typeof v === "string" ? { url: v } : (v || {}); }
  function choose(previous, incoming) {
    const a = object(previous), b = object(incoming);
    if (!b.url) return a;
    if (!a.url) return b;
    // A date belongs to its proof file. Never borrow one file's date for another URL.
    if (a.date && (!b.date || b.date < a.date)) return a;
    if (a.recordDate && b.recordDate && b.recordDate < a.recordDate) return a;
    return b;
  }
  function apply(item, evidence, preserveEdit) {
    const out = Object.assign({}, item), v = object(evidence);
    if (!v.url) return out;
    const older = out.expires && (!v.date || v.date < out.expires);
    // Preserve a newer proof link as well as its date. An older certificate can still
    // establish that proof exists when no linked certificate was previously available.
    if (!older || !out.isFile || !out.fileLink) {
      out.fileLink = v.url; out.uploadName = v.name || out.uploadName;
    }
    out.isFile = true; out.uploaded = true;
    if (v.recordDate && (!out.recordDate || v.recordDate >= out.recordDate)) {
      out.recordDate = v.recordDate;
      out.recurringFromDocument = !!v.recurring;
    }
    if (v.date && !preserveEdit && (!out.expires || v.date >= out.expires)) {
      out.expires = v.date;
      out.permanent = false; out.pending = false;
      out.expiresAuto = true; out.expiresFromFilename = !v.recurring;
      out.recurringFromDocument = !!v.recurring;
    }
    return out;
  }
  return { choose, apply };
});
