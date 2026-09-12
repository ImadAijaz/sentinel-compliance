// Resumable master -> corporate-library intake. No source file is changed or deleted.
// Inventory and preview must finish before copying can be enabled by an administrator.
const G = require("./graph");
const F = require("./master-filing");
const { randomUUID } = require("crypto");
const BASE = "Sama Farooqui/Sentinel";
const STATE = BASE + "/_State/master_intake.json";
const LOCK = BASE + "/_State/master_intake_lock.json";
const urlFor = p => G.docsRoot() + "/root:/" + G.encPath(p);
const headers = token => ({ Authorization: "Bearer " + token });

async function get(token, url, missing) {
  const r = await fetch(url, { headers: headers(token), signal: AbortSignal.timeout(18000) });
  if (missing && r.status === 404) return null;
  if (!r.ok) {
    const e = new Error("Microsoft HTTP " + r.status + (r.status === 429 ? ": requests slowed; progress saved" : " while reading document metadata"));
    if (r.status === 429 || r.status === 503) e.retryAfter = Math.max(15, Math.min(120, Number(r.headers.get("retry-after")) || 30));
    throw e;
  }
  return r.json();
}
async function save(token, value) {
  const r = await fetch(urlFor(STATE) + ":/content", {
    method: "PUT", headers: { ...headers(token), "Content-Type": "application/json" },
    body: JSON.stringify(value), signal: AbortSignal.timeout(18000),
  });
  if (!r.ok) throw new Error("Could not save intake progress: HTTP " + r.status);
}
async function lock(token, namespace) {
  const lockPath = namespace === "scan" ? BASE + "/_State/document_scan_lock.json" : LOCK;
  const prev = await get(token, urlFor(lockPath), true);
  if (!prev) await G.ensureFolderIn(token, G.docsRoot(), BASE + "/_State");
  if (prev) {
    const value = await get(token, urlFor(lockPath) + ":/content");
    if (value.until > Date.now()) return null;
  }
  const id = randomUUID();
  const r = await fetch(urlFor(lockPath) + ":/content", {
    method: "PUT", headers: { ...headers(token), "Content-Type": "application/json", ...(prev ? { "If-Match": prev.eTag } : { "If-None-Match": "*" }) },
    body: JSON.stringify({ id, until: Date.now() + 90000 }),
  });
  if (r.status === 409 || r.status === 412) return null;
  if (!r.ok) throw new Error("Could not reserve the intake worker: HTTP " + r.status);
  const etag = (await r.json()).eTag;
  return async () => {
    await fetch(urlFor(lockPath) + ":/content", {
      method: "PUT", headers: { ...headers(token), "Content-Type": "application/json", "If-Match": etag },
      body: JSON.stringify({ id, until: 0 }),
    });
  };
}
async function inspect(token) {
  const [source, dest] = await Promise.all([
    get(token, G.driveRoot() + "/root:/" + G.encPath(G.ROOT) + ":/children?$select=id,name,folder&$top=200"),
    get(token, urlFor(BASE) + "?$select=id,name,folder"),
  ]);
  return {
    ok: true, source: G.ROOT, destination: BASE,
    folders: (source.value || []).filter(x => x.folder).map(x => x.name),
    destinationReachable: !!dest.folder,
    note: "Read-only access check. No documents or settings changed.",
  };
}
function report(state) {
  if (!state) return { ok: true, stage: "not-started" };
  const ops = state.plan ? state.plan.operations : [];
  const statuses = {};
  ops.forEach(o => { statuses[o.status] = (statuses[o.status] || 0) + 1; });
  return {
    ok: !state.error || !!state.retryAfter, stage: state.stage, error: state.error || null, retryAfter: state.retryAfter || null,
    startedAt: state.startedAt, updatedAt: state.updatedAt, completedAt: state.completedAt,
    sourceFiles: Object.keys(state.source.files).length,
    destinationFiles: Object.keys(state.destination.files).length,
    foldersRemaining: state.source.queue.length + state.destination.queue.length,
    counts: state.plan && state.plan.counts, already: state.plan && state.plan.already,
    coverage: state.plan && state.plan.coveredSites,
    skipped: state.plan && state.plan.skips,
    unmatched: state.plan && state.plan.unmatched,
    excludedFolders: state.excluded,
    copy: statuses,
    samples: ops.filter(o => o.status !== "done").slice(0, 12).map(o => ({ from: o.source.path, to: o.target, status: o.status, error: o.error })),
    more: ["inventory", "copying"].includes(state.stage),
  };
}
function keepFolder(side, rel) {
  const p = rel.split("/");
  if (side === "destination") return ["Provider", "Staff", "State Readiness"].includes(p[0]) && !F.excluded(rel);
  if (p[0] === F.PROVIDERS) return !F.excluded(rel);
  if (F.SITES[p[0]]) return !F.excluded(rel);
  // Inspect actual roster filenames to expose stale staff-workbook links; never copy them.
  return p[0] === "..WCGTX Master Rosters";
}
async function inventory(token, state, deadline) {
  for (const side of ["source", "destination"]) {
    const inv = state[side], root = side === "source" ? G.driveRoot() : G.docsRoot();
    while (inv.queue.length && Date.now() < deadline) {
      const batch = inv.queue.slice(0, 2);
      // All-or-nothing per batch: a failed page keeps its cursor queued, never claims completion.
      const pages = await Promise.all(batch.map(q => get(token, q.next || root + "/items/" + encodeURIComponent(q.id) + "/children?$select=id,name,file,folder,size,eTag,cTag&$top=500")));
      inv.queue.splice(0, batch.length);
      for (let i = 0; i < batch.length; i++) {
        const q = batch[i], j = pages[i];
        for (const item of j.value || []) {
          const rel = q.rel ? q.rel + "/" + item.name : item.name;
          if (item.folder) {
            if (keepFolder(side, rel)) inv.queue.push({ id: item.id, rel });
            else if (side === "source") state.excluded[rel] = F.excluded(rel) || "outside tracked sites";
          } else if (item.file) {
            inv.files[item.id] = { id: item.id, path: rel, size: item.size || 0, tag: item.cTag || item.eTag, hashes: item.file.hashes || {} };
          }
        }
        if (j["@odata.nextLink"]) inv.queue.push({ ...q, next: j["@odata.nextLink"] });
      }
      await new Promise(resolve => setTimeout(resolve, 200));
    }
  }
  if (!state.source.queue.length && !state.destination.queue.length) {
    const items = await require("./delta").applyRosterDelta(require("../data.json").items || []);
    state.plan = F.plan(Object.values(state.source.files), Object.values(state.destination.files), items);
    state.stage = "preview";
  }
}
async function folder(token, rel, cache) {
  if (cache[rel]) return cache[rel];
  let item = await get(token, urlFor(BASE + "/" + rel) + "?$select=id,folder", true);
  if (!item) {
    await G.ensureFolderIn(token, G.docsRoot(), BASE + "/" + rel);
    item = await get(token, urlFor(BASE + "/" + rel) + "?$select=id,folder");
  }
  if (!item.folder) throw new Error("Destination is not a folder");
  cache[rel] = item.id;
  return item.id;
}
async function copyOne(token, op, cache) {
  const targetUrl = urlFor(BASE + "/" + op.target);
  const exists = await get(token, targetUrl + "?$select=id,name,size,file", true);
  if (exists) {
    if (F.sameContent(op.source, { size: exists.size, hashes: exists.file && exists.file.hashes })) {
      op.status = "done"; delete op.error; return;
    }
    // Never replace an existing ambiguous file, even if a copy was interrupted.
    if (op.status !== "copying") throw new Error("Destination appeared or differs from the preview; file preserved for review");
  }
  if (op.status === "copying") {
    if (op.monitor) {
      // Copy monitors are server-issued URLs, kept on the server; never client-controlled.
      const u = new URL(op.monitor);
      if (u.protocol !== "https:" || !(u.hostname === "graph.microsoft.com" || u.hostname.endsWith(".sharepoint.com"))) throw new Error("Unexpected copy-monitor host");
      const m = await get(token, op.monitor);
      if (m.status === "failed" || m.status === "deleteFailed") throw new Error("Microsoft could not finish the copy");
      if (m.status === "completed" && exists && exists.size === op.source.size) { op.status = "done"; delete op.error; return; }
    }
    if (Date.now() - op.submittedAt > 10 * 60 * 1000) throw new Error("Copy not verified after 10 minutes");
    return;
  }
  const sourceNow = await get(token, G.driveRoot() + "/items/" + encodeURIComponent(op.source.id) + "?$select=id,eTag,cTag");
  if ((sourceNow.cTag || sourceNow.eTag) !== op.source.tag) throw new Error("Source changed after preview; refresh inventory before copying this file");
  const parts = op.target.split("/"), name = parts.pop();
  const parentId = await folder(token, parts.join("/"), cache);
  const r = await fetch(G.driveRoot() + "/items/" + encodeURIComponent(op.source.id) + "/copy", {
    method: "POST", headers: { ...headers(token), "Content-Type": "application/json" },
    body: JSON.stringify({ parentReference: { driveId: G.DOCS_DRIVE_ID, id: parentId }, name, "@microsoft.graph.conflictBehavior": "fail" }),
    signal: AbortSignal.timeout(18000),
  });
  if (r.status !== 202) throw new Error("Copy request HTTP " + r.status);
  op.status = "copying"; op.submittedAt = Date.now(); op.monitor = r.headers.get("location");
}
async function copies(token, state, deadline) {
  const cache = {}, todo = state.plan.operations.filter(o => o.status === "pending" || o.status === "copying");
  while (todo.length && Date.now() < deadline) {
    const attempts = await Promise.all(todo.splice(0, 2).map(async op => {
      try { await copyOne(token, op, cache); }
      catch (e) { if (e.retryAfter) return e; op.status = "error"; op.error = String(e.message || e); }
    }));
    const limited = attempts.find(Boolean);
    if (limited) throw limited;
  }
  if (!state.plan.operations.some(o => ["pending", "copying"].includes(o.status))) {
    state.stage = state.plan.operations.some(o => o.status === "error") ? "needs-review" : "complete";
    if (state.stage === "complete") state.completedAt = new Date().toISOString();
  }
}
async function run(token, action, term) {
  if (action === "inspect") return inspect(token);
  if (action === "status") return report(await G.readDocsJsonAt(token, STATE));
  if (action === "audit") {
    const state = await G.readDocsJsonAt(token, STATE);
    if (!state) return { ok: true, files: [], total: 0, note: "No saved inventory" };
    const q = String(term || "").slice(0, 100).toLowerCase();
    const operations = new Map(((state.plan && state.plan.operations) || []).map(o => [o.source.id, o]));
    const files = Object.values(state.source.files).filter(f => f.path.toLowerCase().includes(q));
    return { ok: true, total: files.length, files: files.slice(0, 100).map(f => {
      const op = operations.get(f.id);
      return { source: f.path, filenameDate: G.dateFromName(f.path.split("/").pop()), target: op && op.target, copyStatus: op && op.status, error: op && op.error };
    }) };
  }
  const release = await lock(token);
  if (!release) return { ok: true, busy: true, more: true, stage: "busy" };
  try {
    let state = await G.readDocsJsonAt(token, STATE);
    if (action === "start") {
      if (state && ["inventory", "copying"].includes(state.stage)) return report(state);
      const [source, destination] = await Promise.all([
        get(token, G.driveRoot() + "/root:/" + G.encPath(G.ROOT) + "?$select=id,folder"),
        get(token, urlFor(BASE) + "?$select=id,folder"),
      ]);
      if (!source.folder || !destination.folder) throw new Error("The configured source or destination is not a folder");
      state = {
        version: 1, stage: "inventory", startedAt: new Date().toISOString(), excluded: {},
        source: { files: {}, queue: [{ id: source.id, rel: "" }] },
        destination: { files: {}, queue: [{ id: destination.id, rel: "" }] },
      };
    }
    if (!state) throw new Error("Start the read-only inventory first");
    if (action === "apply") {
      if (state.stage !== "preview") throw new Error("A complete filing preview is required before copying");
      state.stage = "copying";
    }
    if (action === "retry" && state.stage === "needs-review") {
      state.plan.operations.forEach(o => { if (o.status === "error") { o.status = o.monitor ? "copying" : "pending"; delete o.error; } });
      state.stage = "copying";
    }
    state.error = null; state.retryAfter = null;
    try {
      const deadline = Date.now() + 25000;
      if (state.stage === "inventory") await inventory(token, state, deadline);
      if (state.stage === "copying") await copies(token, state, deadline);
    } catch (e) { state.error = String(e.message || e); state.retryAfter = e.retryAfter || null; }
    state.updatedAt = new Date().toISOString();
    await save(token, state);
    return report(state);
  } finally { await release(); }
}
module.exports = { run, report, keepFolder, lock };
