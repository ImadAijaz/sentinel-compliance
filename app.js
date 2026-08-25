/* ============================================================================
   SENTINEL — Compliance Command Center  (app logic)
   Pure vanilla JS. No build step, no server. Works by double-click (file://).
   ========================================================================== */
(function () {
  "use strict";

  // ---------- tiny DOM helpers ----------
  const $ = (s, r = document) => r.querySelector(s);
  const el = (t, c, h) => { const e = document.createElement(t); if (c) e.className = c; if (h != null) e.innerHTML = h; return e; };
  const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const hash = s => window.sha256(String(s));

  // ---------- constants ----------
  const OVERLAY_KEY = "sentinel_overlay_v1";
  const CFG_KEY = "sentinel_config_v1";
  const PREF_KEY = "sentinel_prefs_v1";
  const SESS_KEY = "sentinel_unlocked";
  const DAY = 86400000;

  const ICONS = {
    "license": '<path d="M4 4h16v16H4z"/><path d="M8 9h8M8 13h5"/>',
    "dea": '<path d="M12 2l8 4v6c0 5-3.5 8-8 10-4.5-2-8-5-8-10V6z"/>',
    "cert": '<circle cx="12" cy="9" r="5"/><path d="M9 13l-1 8 4-2 4 2-1-8"/>',
    "cme": '<path d="M3 7l9-4 9 4-9 4z"/><path d="M21 10v5M7 12v4c0 1.5 2.5 3 5 3s5-1.5 5-3v-4"/>',
    "vax": '<path d="M14 4l6 6M16 6l-9 9-3 5 5-3 9-9M9 11l4 4"/>',
    "check": '<path d="M20 6L9 17l-5-5"/>',
    "doc": '<path d="M6 2h9l5 5v15H6z"/><path d="M14 2v6h6"/>',
    "shield": '<path d="M12 2l8 4v6c0 5-3.5 8-8 10-4.5-2-8-5-8-10V6z"/>',
    "fire": '<path d="M12 2s5 5 5 10a5 5 0 01-10 0c0-2 1-3 1-3s3 1 4-7z"/>',
    "meeting": '<circle cx="9" cy="8" r="3"/><circle cx="17" cy="9" r="2.5"/><path d="M3 20c0-3 3-5 6-5s6 2 6 5M15 20c0-2 1-3.5 3-3.5"/>',
    "building": '<path d="M5 21V5l7-3 7 3v16M9 9h.01M15 9h.01M9 13h.01M15 13h.01M9 17h6"/>',
    "id": '<rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="11" r="2"/><path d="M14 10h4M14 14H7"/>'
  };
  function iconFor(it) {
    const c = (it.category || "").toLowerCase();
    if (c.includes("dea")) return "dea";
    if (c.includes("license")) return "license";
    if (c.includes("board")) return "shield";
    if (/acls|atls|pals|bls/.test(c)) return "cert";
    if (c.includes("cme")) return "cme";
    if (/flu|tb|vaccin/.test(c)) return "vax";
    if (/oig|npdb|verif|sam/.test(c)) return "check";
    if (c.includes("inspection") || c.includes("fire") || c.includes("generator")) return "fire";
    if (c.includes("meeting")) return "meeting";
    if (c.includes("clia") || c.includes("cola") || c.includes("pharmac") || c.includes("x-ray") || c.includes("occupancy") || c.includes("npi") || c.includes("ein")) return "building";
    if (c.includes("malpractice") || c.includes("coi")) return "shield";
    return "doc";
  }

  // ---------- state ----------
  let CFG = null;                  // resolved auth config (file or localStorage)
  let CURRENT_USER = null;         // {label, role, tabs[], readonly}
  let READONLY = false;            // true for viewer role or ?readonly=1 share link
  const USERS_KEY = "sentinel_users_v1";
  function allUsers() { let fromCfg = (CFG && CFG.users) || [], fromLs = []; try { fromLs = JSON.parse(localStorage.getItem(USERS_KEY) || "[]"); } catch (e) {} return fromCfg.concat(fromLs); }
  function resolveUser(id, pw) {
    if (hash(id) === CFG.loginIdHash && hash(pw) === CFG.loginPwHash) return { label: "Administrator", role: "admin", tabs: ["provider", "staff", "facility", "other"], readonly: false };
    const u = allUsers().find(u => u.idHash === hash(id) && u.pwHash === hash(pw));
    if (u) return { label: u.label || id, role: "staff", tabs: (u.tabs && u.tabs.length) ? u.tabs : ["provider", "staff", "facility", "other"], readonly: !!u.readonly };
    return null;
  }
  let DATA = [];                   // merged item list
  let OVERLAY = loadJSON(OVERLAY_KEY, { edits: {}, added: [], deleted: [], logs: {}, watch: [], audit: [], leads: {}, tasks: {}, snapshots: {}, verified: {} });
  ["watch", "audit"].forEach(k => { if (!OVERLAY[k]) OVERLAY[k] = []; });
  ["leads", "tasks", "snapshots", "logs", "edits", "verified"].forEach(k => { if (!OVERLAY[k]) OVERLAY[k] = {}; });
  let PREFS = loadJSON(PREF_KEY, { theme: "light" });
  let UNLOCKED = loadSession();
  const state = { tab: "provider", view: "list", search: "", status: "", category: "", facility: "all", sort: "name", showInactive: false, openGroups: {}, quickView: "", selectMode: false, selection: new Set() };
  function nowStamp() { return new Date().toISOString().slice(0, 16).replace("T", " "); }
  function logAudit(action, it, detail) { OVERLAY.audit.unshift({ action: action, id: it.id, entity: it.entity, category: it.category, detail: detail || "", at: nowStamp() }); if (OVERLAY.audit.length > 500) OVERLAY.audit.length = 500; saveOverlay(); }
  function isWatched(id) { return OVERLAY.watch.indexOf(id) >= 0; }
  function toggleWatch(id) { const i = OVERLAY.watch.indexOf(id); if (i >= 0) OVERLAY.watch.splice(i, 1); else OVERLAY.watch.push(id); saveOverlay(); }

  // ---------- Cloud — active when signed in via Microsoft (overlay stored in OneDrive, no Supabase) ----------
  const CLOUD = !!window.SENTINEL_AUTH;
  let _syncT = null;
  function cloudSyncOverlay() { if (!CLOUD) return; clearTimeout(_syncT); _syncT = setTimeout(() => { try { fetch("/api/state", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(OVERLAY) }).catch(() => {}); } catch (e) {} }, 800); }

  function loadJSON(k, d) { try { return Object.assign({}, d, JSON.parse(localStorage.getItem(k) || "{}")); } catch (e) { return d; } }
  function saveOverlay() { localStorage.setItem(OVERLAY_KEY, JSON.stringify(OVERLAY)); cloudSyncOverlay(); }
  function savePrefs() { localStorage.setItem(PREF_KEY, JSON.stringify(PREFS)); }
  // Tab unlocks are kept IN MEMORY ONLY — every page load / refresh re-locks all
  // tabs so the access code is required each time the app is opened.
  function loadSession() { return new Set(); }
  function saveSession() { /* intentionally not persisted */ }

  // ---------- date / status engine ----------
  function today() { const d = new Date(); d.setHours(0, 0, 0, 0); return d; }
  function parseD(s) {
    if (!s) return null;
    const p = String(s).split("-");
    // Only treat as ISO YYYY-MM-DD when the parts plausibly are year/month/day; otherwise fall
    // back to Date() parsing. Prevents "18-09-2026" being silently mis-parsed as year=18 day=2026.
    if (p.length === 3 && +p[0] >= 1900 && +p[1] >= 1 && +p[1] <= 12 && +p[2] >= 1 && +p[2] <= 31) {
      return new Date(+p[0], +p[1] - 1, +p[2]);
    }
    const d = new Date(s); return isNaN(d) ? null : d;
  }
  function fmtD(s) { const d = parseD(s); if (!d) return "—"; return d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" }); }
  function addDays(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }

  function computeStatus(it) {
    const exp = parseD(it.expires);
    if (!exp) {
      if (it.permanent) return { key: "permanent", label: "Permanent", days: null };
      if (it.pending) return { key: "pending", label: "Action needed", days: null };
      if (it.recurring) return { key: "recurring", label: "Recurring", days: null };
      return { key: "pending", label: "No date", days: null };
    }
    const days = Math.round((exp - today()) / DAY);
    let key, label;
    if (days < 0) { key = "expired"; label = "Expired"; }
    else if (days <= 30) { key = "critical"; label = "Critical"; }
    else if (days <= 90) { key = "due"; label = "Due soon"; }
    else { key = "good"; label = "Good"; }
    const lead = (OVERLAY.leads && OVERLAY.leads[it.category] != null) ? OVERLAY.leads[it.category] : (it.renewalLeadDays || 60);
    const startBy = addDays(exp, -lead);
    return { key, label, days, startBy, expDate: exp };
  }
  function countdownText(st) {
    if (st.days == null) return "";
    if (st.days < 0) return Math.abs(st.days) + "d ago";
    if (st.days === 0) return "today";
    return "in " + st.days + "d";
  }

  // Live-detected supplemental records from /api/uploads-map (new files in SharePoint that
  // aren't yet in the baked data.json — surfaced without a regenerate).
  let LIVE_SUPP = [];
  // Last known document map from /api/uploads-map, kept so it can be re-applied after any
  // rebuild of DATA (see attachUploads / buildData).
  let ATTACHMENTS = null;
  // Apply the known documents onto the current DATA items. Safe to call repeatedly.
  function attachUploads() {
    if (!ATTACHMENTS) return;
    DATA.forEach(it => {
      const v = ATTACHMENTS[it.id]; if (!v) return;
      const url = (typeof v === "string") ? v : v.url; if (!url) return;
      it.fileLink = url; it.isFile = true; it.uploaded = true;
      if (typeof v === "object" && v.name) it.uploadName = v.name;
      // A date parsed out of a FILENAME is a guess (a phone-scan name like "Scan_2026-01-05.pdf"
      // is the scan date, not the expiry), so flag it as unconfirmed rather than presenting it as
      // the real expiry. Never override a human's explicit edit.
      if (typeof v === "object" && v.date && !OVERLAY.edits[it.id]) {
        it.expires = v.date; it.permanent = false; it.pending = false;
        it.expiresAuto = true;
        it.expiresFromFilename = true;
      }
    });
  }

  // ---------- data assembly ----------
  function buildData() {
    const seed = (window.SENTINEL_SEED && window.SENTINEL_SEED.items) || [];
    const del = new Set(OVERLAY.deleted);
    const out = [];
    const seedIds = new Set();
    seed.forEach(it => {
      if (del.has(it.id)) return;
      seedIds.add(it.id);
      out.push(OVERLAY.edits[it.id] ? Object.assign({}, it, OVERLAY.edits[it.id]) : it);
    });
    (OVERLAY.added || []).forEach(it => { if (!del.has(it.id)) { seedIds.add(it.id); out.push(it); } });
    // Append live supplementals not already represented in the seed (avoid double-listing).
    LIVE_SUPP.forEach(it => { if (it && it.id && !seedIds.has(it.id) && !del.has(it.id)) out.push(it); });
    DATA = out;
    // Re-attach the known documents. Attachments are applied by MUTATING items, so any code that
    // replaces window.SENTINEL_SEED with a fresh /api/data payload (Sync from Excel, add/delete,
    // the on-open and 90s auto-sync, tab refocus) used to throw them away — every uploaded or
    // scanned document silently flipped back to "No proof" until a full page reload, and the
    // 45s poll saw no change so it never re-applied them. Re-applying here, at the single place
    // DATA is built, means every rebuild keeps them.
    attachUploads();
  }

  // ---------- scope helpers ----------
  function tabItems(tab) {
    let arr = DATA.filter(i => i.scope === tab);
    if (tab === "provider" || tab === "staff") arr = arr.filter(i => state.showInactive ? true : (i.active !== false));
    return arr;
  }
  // items only from tabs the user has unlocked (gates all global data tools)
  function unlockedItems() { return DATA.filter(i => UNLOCKED.has(i.scope) && i.active !== false); }
  const LOCK_MSG = '<div class="empty"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="5" y="11" width="14" height="10" rx="2"/><path d="M8 11V7a4 4 0 018 0v4"/></svg><h3>Locked</h3><p>Enter a tab’s access code first — this only shows data from tabs you’ve unlocked.</p></div>';

  function applyFilters(arr) {
    const q = state.search.toLowerCase();
    return arr.filter(it => {
      if (state.facility !== "all" && it.scope !== "provider") { var fv = it.scope === "staff" ? (it.facility || "") : it.entity; if (fv !== state.facility) return false; }
      if (state.category && it.category !== state.category) return false;
      if (state.status && computeStatus(it).key !== state.status) return false;
      const qv = state.quickView;
      if (qv === "missing" && it.isFile) return false;
      if (qv === "watch" && !isWatched(it.id)) return false;
      if ((qv === "expired" || qv === "critical" || qv === "due") && computeStatus(it).key !== qv) return false;
      if (q) {
        let fname = "";
        try { fname = decodeURIComponent(String(it.fileLink || "").split("/").pop() || "").replace(/[_%]/g, " "); } catch (e) { fname = String(it.fileLink || ""); }
        const blob = (it.entity + " " + it.category + " " + (it.authority || "") + " " + (it.number || "") + " " + (it.notes || "") + " " + fname).toLowerCase();
        // match every typed word (any order) so "driving license" finds "Driver's License"-type files
        if (!q.split(/\s+/).filter(Boolean).every(tok => blob.includes(tok))) return false;
      }
      return true;
    });
  }
  const STATUS_RANK = { expired: 0, critical: 1, due: 2, pending: 3, recurring: 4, good: 5, permanent: 6 };
  function sortItems(arr) {
    return arr.slice().sort((a, b) => {
      const sa = computeStatus(a), sb = computeStatus(b);
      if (STATUS_RANK[sa.key] !== STATUS_RANK[sb.key]) return STATUS_RANK[sa.key] - STATUS_RANK[sb.key];
      const da = sa.days == null ? 99999 : sa.days, db = sb.days == null ? 99999 : sb.days;
      return da - db;
    });
  }

  // ================= AUTH =================
  function resolveConfig() {
    const saved = JSON.parse(localStorage.getItem(CFG_KEY) || "null");
    if (saved && saved.configured) return saved;
    if (window.SENTINEL_CONFIG && window.SENTINEL_CONFIG.configured) return window.SENTINEL_CONFIG;
    return null;
  }

  function renderAuth() {
    CFG = resolveConfig();
    if (!CFG) return renderSetup();
    return renderLogin();
  }

  // Cloud: the user is already authenticated by Microsoft (window.SENTINEL_AUTH set by the
  // index.html gate), so skip the code login entirely and open the app with full access.
  function bootSentinel() {
    if (window._booted) return; window._booted = true;
    applyTheme(PREFS.theme);
    if (window.SENTINEL_AUTH) {
      CFG = resolveConfig() || { configured: true, loginIdHash: "", loginPwHash: "", tabHashes: {} };
      const who = window.SENTINEL_AUTH.name || window.SENTINEL_AUTH.email || "Staff";
      const tabs = (window.SENTINEL_AUTH.tabs && window.SENTINEL_AUTH.tabs.length) ? window.SENTINEL_AUTH.tabs : ["provider", "staff", "facility", "other"];
      CURRENT_USER = { label: who, role: window.SENTINEL_AUTH.admin ? "admin" : "staff", tabs: tabs, readonly: /[?&]readonly=1/.test(location.search) };
      READONLY = CURRENT_USER.readonly;
      UNLOCKED = new Set(tabs);  // only the tabs this Microsoft account is granted
      enterApp();
    } else {
      renderAuth();
    }
  }

  function authShell(inner) {
    $("#authCard").innerHTML =
      '<div class="logo-lockup"><svg class="mark" style="width:44px;height:44px"><use href="#logo-sentinel"/></svg>' +
      '<div><div class="name" style="color:#fff"><b style="color:#5eead4">Sentinel</b></div>' +
      '<div class="tag" style="color:#7fb8b0">Compliance Command Center</div></div></div>' + inner;
  }

  function renderLogin() {
    authShell(
      '<div class="auth-sub">Wellness &amp; Care Group of Texas</div>' +
      '<div class="field"><label>Login ID</label><input id="inId" autocomplete="off" name="sentinel-login-x" placeholder="Enter your login ID"></div>' +
      '<div class="field"><label>Password</label><input id="inPw" type="password" autocomplete="current-password" placeholder="Enter your password"></div>' +
      '<button class="btn-primary" id="doLogin">Unlock Sentinel</button>' +
      '<div class="auth-msg" id="authMsg"></div>'
    );
    const go = () => {
      const u = resolveUser($("#inId").value, $("#inPw").value);
      if (u) {
        CURRENT_USER = u;
        READONLY = u.readonly || /[?&]readonly=1/.test(location.search) || /readonly/.test(location.hash);
        enterApp();
      } else { $("#authMsg").textContent = "Incorrect login ID or password."; $("#inPw").value = ""; }
    };
    $("#doLogin").onclick = go;
    $("#inPw").onkeydown = e => { if (e.key === "Enter") go(); };
    $("#inId").focus();
  }

  function renderSetup() {
    authShell(
      '<div class="auth-sub">First-time setup — choose your access codes</div>' +
      '<div class="field"><label>Login ID</label><input id="sId" placeholder="e.g. compliance.admin"></div>' +
      '<div class="field"><label>Master Password</label><input id="sPw" type="password" placeholder="Strong password"></div>' +
      '<div class="field"><label>Provider tab code</label><input id="sProv" type="password" placeholder="Provider Compliance code"></div>' +
      '<div class="field"><label>Facility tab code</label><input id="sFac" type="password" placeholder="Facility Compliance code"></div>' +
      '<div class="field"><label>Other tab code</label><input id="sOth" type="password" placeholder="Other Compliance code"></div>' +
      '<button class="btn-primary" id="doSetup">Save &amp; enter</button>' +
      '<div class="auth-msg" id="authMsg"></div>' +
      '<div class="auth-foot">Only SHA-256 hashes are stored (in this browser). After saving you can download a <b>config.js</b> to ship the codes with the app.</div>'
    );
    $("#doSetup").onclick = () => {
      const v = id => $(id).value.trim();
      if (!v("#sId") || !v("#sPw") || !v("#sProv") || !v("#sFac") || !v("#sOth")) { $("#authMsg").textContent = "Please fill every field."; return; }
      const cfg = {
        configured: true,
        loginIdHash: hash(v("#sId")), loginPwHash: hash(v("#sPw")),
        tabHashes: { provider: hash(v("#sProv")), facility: hash(v("#sFac")), other: hash(v("#sOth")) }
      };
      localStorage.setItem(CFG_KEY, JSON.stringify(cfg));
      CFG = cfg;
      offerConfigDownload(cfg);
      enterApp();
    };
    $("#sId").focus();
  }

  function offerConfigDownload(cfg) {
    const body = "/* Sentinel access config — generated " + new Date().toISOString().slice(0, 10) +
      ". Only SHA-256 hashes, no plaintext. */\nwindow.SENTINEL_CONFIG = " + JSON.stringify(cfg, null, 2) + ";\n";
    const blob = new Blob([body], { type: "text/javascript" });
    const a = el("a");
    a.href = URL.createObjectURL(blob); a.download = "config.js";
    a.textContent = "download config.js";
    setTimeout(() => { try { a.click(); } catch (e) {} }, 400);
  }

  function enterApp() {
    $("#authWrap").classList.add("hidden");
    $("#app").classList.remove("hidden");
    if (READONLY) document.body.classList.add("readonly"); else document.body.classList.remove("readonly");
    if (CURRENT_USER && CURRENT_USER.tabs.indexOf(state.tab) < 0) state.tab = CURRENT_USER.tabs[0];
    buildData();
    initShell();
    if (CLOUD) document.body.classList.add("cloud");
    // Render IMMEDIATELY from the loaded roster — never block the screen on a network call.
    // Saved edits (Supabase) and document badges (Microsoft) load in the background and re-render.
    render(); handleDeepLink();
    if (CLOUD) {
      // Saved edits/notes/verify marks live in OneDrive now (no Supabase). Best-effort, background.
      fetch("/api/state").then(r => r.ok ? r.json() : null).then(d => {
        if (d && Object.keys(d).length) { OVERLAY = Object.assign({ edits: {}, added: [], deleted: [], logs: {}, watch: [], audit: [], leads: {}, tasks: {}, snapshots: {}, verified: {} }, d); if (!OVERLAY.verified) OVERLAY.verified = {}; buildData(); render(); }
      }).catch(() => {});
      pullCloudUploads(render);
    } else if (location.protocol.indexOf("http") === 0) {
      fetch("/api/uploads").then(r => r.json()).then(u => { applyUploads(u); render(); }).catch(() => {});
    }
  }

  // ---- Live sync of documents: QR uploads + files dropped straight into OneDrive ----
  // /api/scan detects new files in OneDrive; /api/uploads-map returns the merged list.
  let _uploadsJSON = "";
  function readUploadsMap() {
    return fetch("/api/uploads-map").then(r => r.json());
  }
  function pullCloudUploads(cb) {
    const done = (u) => { if (u && typeof u === "object" && u.ok !== false) { _uploadsJSON = JSON.stringify(u); applyUploads(u); } if (cb) cb(); startUploadSync(); startRosterSync(); };
    // kick a scan (catch new OneDrive files), then read the merged map
    fetch("/api/scan").catch(() => {}).then(() => readUploadsMap().then(done, () => done(null))).catch(() => done(null));
  }
  function startUploadSync() {
    if (window._upSync || !CLOUD) return;
    window._upSync = setInterval(() => {
      fetch("/api/scan").catch(() => {}).then(() => readUploadsMap()).then(u => {
        // A failed read (HTTP 200 with ok:false) must not be mistaken for "the documents changed".
        if (!u || u.ok === false) return;
        const j = JSON.stringify(u);
        if (j !== _uploadsJSON) { _uploadsJSON = j; applyUploads(u); render(); toast("Documents updated."); }
      }).catch(() => {});
    }, 45000);
  }
  // Near-real-time roster/Excel sync: every 90s, admins trigger a silent regen (re-reads the
  // Excel + RN/FD workbooks), then everyone re-fetches and only re-renders if something actually
  // changed. So edits in Excel or OneDrive surface on the dashboard within ~1–2 min, no Sync click.
  function rosterSig(items) {
    let a = 0;
    for (const i of items) {
      const s = (i.id || "") + (i.expires || "") + (i.active === false ? "0" : "1") + (i.pending ? "p" : "");
      for (let k = 0; k < s.length; k++) a = (a * 31 + s.charCodeAt(k)) >>> 0;
    }
    return items.length + "|" + a;
  }
  function startRosterSync() {
    if (window._rosterSync || !CLOUD) return;
    let last = null;
    const refresh = () => fetch("/api/data").then(r => r.json()).then(d => {
      if (!d || !d.items) return;
      const g = rosterSig(d.items);
      if (last === null) { last = g; return; }
      if (g !== last) { last = g; window.SENTINEL_SEED = d; buildData(); render(); toast("Updated from Excel / OneDrive."); }
    }).catch(() => {});
    // One sync pass: pull the live Excel into the cache, then re-read and re-render on change.
    // The server throttles the regen globally (60s), so calling it on every open is cheap and
    // safe even with many people signing in at once — and it's no longer admin-only, so staff
    // opening the dashboard also get fresh data.
    const syncNow = () => fetch("/api/data?regen=1", { method: "POST" }).catch(() => {}).then(refresh);
    // AUTO-SYNC ON OPEN: establish the baseline signature first (so refresh can tell what
    // actually changed), then immediately sync. Chaining — not firing both at once — is what
    // makes a change found on open render right away instead of waiting for the next tick.
    fetch("/api/data").then(r => r.json()).then(d => { if (d && d.items) last = rosterSig(d.items); }).catch(() => {}).then(syncNow);
    // Re-sync when the tab regains focus, so a dashboard left open overnight is current the
    // moment the user comes back to it, without waiting out the interval.
    if (!window._rosterFocus) {
      window._rosterFocus = () => { if (document.visibilityState === "visible") syncNow(); };
      document.addEventListener("visibilitychange", window._rosterFocus);
    }
    window._rosterSync = setInterval(syncNow, 90000);
  }

  function lockApp() {
    UNLOCKED = new Set(); saveSession(); CURRENT_USER = null; READONLY = false;
    document.body.classList.remove("readonly");
    $("#app").classList.add("hidden");
    $("#authWrap").classList.remove("hidden");
    renderAuth();
  }

  // tab access-code prompt
  function promptTabCode(tab, onOk) {
    const back = el("div", "auth-wrap"); back.style.background = "rgba(5,20,25,.7)";
    const labels = { provider: "Provider Compliance", staff: "Staff Compliance", facility: "Facility Compliance", other: "Other Compliance" };
    back.innerHTML = '<div class="auth-card" style="width:min(380px,100%)">' +
      '<div class="logo-lockup" style="justify-content:center"><svg class="mark"><use href="#logo-sentinel"/></svg></div>' +
      '<div class="auth-sub" style="margin-top:10px">Enter access code for<br><b style="color:#fff;font-size:15px">' + labels[tab] + '</b></div>' +
      '<div class="field"><input id="tabCode" type="password" placeholder="Access code" style="text-align:center"></div>' +
      '<button class="btn-primary" id="tabGo">Unlock tab</button>' +
      '<div class="auth-msg" id="tabMsg"></div>' +
      '<div class="auth-foot"><a id="tabCancel">Cancel</a></div></div>';
    document.body.appendChild(back);
    const inp = back.querySelector("#tabCode"); inp.focus();
    const go = () => {
      if (hash(inp.value) === CFG.tabHashes[tab]) { UNLOCKED.add(tab); saveSession(); document.body.removeChild(back); onOk(); }
      else { back.querySelector("#tabMsg").textContent = "Incorrect code."; inp.value = ""; }
    };
    back.querySelector("#tabGo").onclick = go;
    inp.onkeydown = e => { if (e.key === "Enter") go(); };
    back.querySelector("#tabCancel").onclick = () => document.body.removeChild(back);
  }

  // ================= SHELL =================
  function initShell() {
    applyTheme(PREFS.theme);
    $("#genStamp").textContent = "data " + ((window.SENTINEL_SEED && window.SENTINEL_SEED.generatedAt) || "");
    $("#btnTheme").onclick = () => { applyTheme(PREFS.theme === "dark" ? "light" : "dark"); savePrefs(); };
    $("#btnLock").onclick = () => { if (window.SENTINEL_AUTH) location.href = "/api/logout"; else lockApp(); };
    $("#btnExport").onclick = exportCSV;
    $("#btnPrint").onclick = doPrint;
    $("#homeLogo").onclick = goHome;
    $("#btnEmail").onclick = emailMe;
    $("#btnSearch").onclick = openPalette;
    $("#btnInsights").onclick = openInsights;
    $("#btnContacts").onclick = openContacts;
    $("#btnBell").onclick = openBell;
    buildFooterLinks();
    recordSnapshot();
    $("#drawerBack").onclick = closeDrawer;
    $("#securityLink").onclick = showSecurityNote;
    $("#modal").onclick = (e) => { if (e.target === $("#modal")) closeModal(); };
    document.addEventListener("keydown", e => {
      if (e.key === "Escape") { closeDrawer(); closeModal(); }
      if ((e.ctrlKey || e.metaKey) && (e.key === "k" || e.key === "K")) { e.preventDefault(); openPalette(); }
    });
  }
  function buildFooterLinks() {
    const defs = [
      { label: "Tasks", fn: openTasks },
      { label: "Lead times", fn: openSettings, edit: true },
      { label: "Full PDF report", fn: fullReport },
      { label: "Email providers", fn: emailProviders, edit: true, localOnly: true },
      { label: "Save to Excel", fn: exportXlsx, edit: true, localOnly: true },
      { label: "Change log", fn: openAudit },
      { label: "Document gaps", fn: openGapReport },
      { label: "Onboarding board", fn: openOnboardingBoard },
      { label: "Sync from Excel", fn: syncFromExcel, cloudOnly: true, admin: true },
      { label: "System check", fn: openSelfTest, cloudOnly: true, admin: true },
      { label: "Backup", fn: backup, admin: true },
      { label: "Restore", fn: restore, edit: true, admin: true },
      { label: "Manage access", fn: openAccessPanel, cloudOnly: true, admin: true },
      { label: "Staff logins", fn: openUsers, admin: true, localOnly: true }
    ];
    const show = defs.filter(d => {
      if (d.admin && (!CURRENT_USER || CURRENT_USER.role !== "admin")) return false;
      if (d.edit && READONLY) return false;
      if (d.localOnly && CLOUD) return false;
      if (d.cloudOnly && !CLOUD) return false;
      return true;
    });
    const c = $("#footLinks"); if (!c) return; c.innerHTML = "";
    show.forEach((d, i) => {
      const a = el("a", "", d.label); a.style.cursor = "pointer"; a.onclick = d.fn; c.appendChild(a);
      if (i < show.length - 1) c.appendChild(document.createTextNode(" · "));
    });
  }
  function applyTheme(t) { PREFS.theme = t; document.documentElement.setAttribute("data-theme", t); }
  function goHome() {
    closeDrawer(); closeModal();
    state.tab = "provider"; state.status = ""; state.quickView = ""; state.category = "";
    state.facility = "all"; state.search = ""; state.selectMode = false; state.selection.clear();
    state.view = "list"; state.openGroups = {}; state.drill = []; state.openPhases = {};
    const vl = $("#viewLabel"); if (vl) vl.textContent = "Timeline";
    render(); window.scrollTo({ top: 0, behavior: "smooth" }); toast("Reset to home.");
  }

  // ================= RENDER =================
  function render() { renderKPIs(); renderAlert(); renderTabs(); renderToolbar(); renderContent(); updateBell(); }

  function statsFor(arr) {
    const s = { total: arr.length, expired: 0, critical: 0, due: 0, good: 0, permanent: 0, recurring: 0, pending: 0 };
    arr.forEach(i => s[computeStatus(i).key]++);
    const datable = s.expired + s.critical + s.due + s.good + s.permanent;
    const score = datable ? Math.round(100 * (s.good + s.permanent + 0.7 * s.due + 0.35 * s.critical) / datable) : 100;
    return Object.assign(s, { score, datable });
  }

  function renderKPIs() {
    const g = $("#kpiGrid"); g.innerHTML = "";
    if (!UNLOCKED.has(state.tab)) return;   // hide summary until the tab passcode is entered
    const arr = tabItems(state.tab);
    const s = statsFor(arr);
    const health = el("div", "kpi kpi-health glass" + (!state.status && !state.quickView ? " active" : ""));
    health.style.cursor = "pointer";
    health.title = "Show all items (clear status filters)";
    health.innerHTML = '<div class="h-meta"><div class="k-label">Show all items</div>' +
      '<div class="h-title">' + (state.tab === "provider" ? "Provider" : state.tab === "staff" ? "Staff" : state.tab === "facility" ? "Facility" : "Operational") + ' readiness</div>' +
      '<div class="h-desc">' + s.total + ' tracked items · <b>' + s.expired + '</b> expired · <b>' + s.critical + '</b> critical</div></div>';
    health.onclick = () => { state.status = ""; state.quickView = ""; render(); };
    g.appendChild(health);

    const cards = [
      ["expired", "Expired", s.expired, "Past due — act now"],
      ["critical", "Critical", s.critical, "≤ 30 days"],
      ["due", "Due soon", s.due, "≤ 90 days"],
      ["good", "Good", s.good, "Current & valid"]
    ];
    cards.forEach(([k, label, val, sub]) => {
      const c = el("div", "kpi glass " + k + (state.status === k ? " active" : ""));
      c.innerHTML = '<div class="k-label">' + label + '</div><div class="k-val">' + val + '</div><div class="k-sub">' + sub + '</div>';
      c.onclick = () => { state.status = state.status === k ? "" : k; render(); };
      g.appendChild(c);
    });
    if (!kpiAnimated && s.total) {
      kpiAnimated = true;
      requestAnimationFrame(() => g.querySelectorAll(".k-val").forEach(e => countUp(e, parseInt(e.textContent, 10) || 0)));
    }
  }

  function renderAlert() {
    const b = $("#alertBanner");
    if (!UNLOCKED.has(state.tab)) { b.innerHTML = ""; return; }   // hide until tab passcode entered
    const arr = DATA.filter(i => i.active !== false);
    const s = statsFor(arr);
    if (s.expired + s.critical === 0) { b.innerHTML = ""; return; }
    b.className = "alert-banner glass";
    b.innerHTML = '<span class="pulse"></span><div class="ab-text"><b>' + (s.expired + s.critical) +
      ' item' + (s.expired + s.critical > 1 ? "s" : "") + ' need immediate attention</b> — ' +
      s.expired + ' expired and ' + s.critical + ' critical (≤30 days) across all facilities & providers.</div>';
  }

  function renderTabs() {
    let defs = [["provider", "Provider Compliance"], ["staff", "Staff Compliance"], ["facility", "Facility Compliance"], ["other", "Other Compliance"]];
    if (CURRENT_USER && CURRENT_USER.tabs) defs = defs.filter(d => CURRENT_USER.tabs.indexOf(d[0]) >= 0);
    const NAV_ICONS = {
      provider: '<path d="M4 21V7l8-4 8 4v14"/><path d="M9 21v-6h6v6"/>',
      staff: '<circle cx="9" cy="7" r="4"/><path d="M2 21v-2a4 4 0 014-4h6a4 4 0 014 4v2"/>',
      facility: '<rect x="4" y="3" width="16" height="18" rx="2"/><path d="M9 8h6M9 12h6M9 16h4"/>',
      other: '<path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><path d="M14 2v6h6"/>',
    };
    const SHORT = { provider: "Providers", staff: "Staff", facility: "Facilities", other: "Other" };
    const t = $("#tabs"); t.innerHTML = "";
    defs.forEach(([k, label]) => {
      const cnt = tabItems(k).length;
      const locked = !UNLOCKED.has(k);
      const b = el("button", "tab" + (state.tab === k ? " active" : "") + (locked ? " locked" : ""));
      b.innerHTML = (locked
          ? '<svg class="lock" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="5" y="11" width="14" height="10" rx="2"/><path d="M8 11V7a4 4 0 018 0v4"/></svg>'
          : '<svg class="nav-ic" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9">' + (NAV_ICONS[k] || "") + '</svg>') +
        (SHORT[k] || label) + ' <span class="count">' + cnt + '</span>';
      b.onclick = () => {
        if (!UNLOCKED.has(k)) { promptTabCode(k, () => { state.tab = k; state.status = ""; render(); }); return; }
        state.tab = k; state.status = ""; render();
      };
      t.appendChild(b);
    });
  }

  function renderToolbar() {
    const tb = $("#toolbar"); tb.innerHTML = "";
    const cats = [...new Set(tabItems(state.tab).map(i => i.category))].sort();
    const sorts = [["status", "Status (worst first)"], ["expiry", "Expiry (soonest)"], ["name", "Name (A–Z)"]];
    tb.innerHTML =
      '<div class="search big"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4-4"/></svg>' +
      '<input id="q" type="search" name="sf-' + Math.random().toString(36).slice(2) + '" autocomplete="off" autocorrect="off" spellcheck="false" readonly data-1p-ignore data-lpignore="true" placeholder="Search this tab — provider, document, license #, file name…" value="' + esc(state.search) + '">' +
      '<button id="qClear" class="q-clear" title="Clear search"' + (state.search ? '' : ' style="display:none"') + '>×</button></div>' +
      '<select class="ctrl" id="sortF" title="Sort by">' + sorts.map(([v, lab]) => '<option value="' + v + '"' + ((state.sort || "name") === v ? " selected" : "") + '>Sort: ' + lab + '</option>').join("") + '</select>' +
      '<select class="ctrl" id="catF"><option value="">All categories</option>' + cats.map(c => '<option' + (state.category === c ? " selected" : "") + '>' + esc(c) + '</option>').join("") + '</select>' +
      (state.tab !== "provider" ? (function () { var its = tabItems(state.tab); var fc = function (fac) { var s = {}; its.forEach(function (i) { var f = i.scope === "staff" ? i.facility : i.entity; if (fac === "all" || f === fac) s[i.entityKey] = 1; }); return Object.keys(s).length; }; return '<div class="chips" id="facChips">' + [["all", "All"], ["Castle Hills ER", "Castle Hills"], ["Frisco ER", "Frisco"]].map(function (o) { return '<button class="chip' + (state.facility === o[0] ? " on" : "") + '" data-f="' + esc(o[0]) + '">' + o[1] + ' <span class="c">' + fc(o[0]) + '</span></button>'; }).join("") + '</div>'; })() : '') +
      (state.tab === "provider" || state.tab === "staff" ? '<label class="toggle-pill"><input type="checkbox" id="inact"' + (state.showInactive ? " checked" : "") + '> Show inactive</label>' : '') +
      '<div class="spacer" style="flex:1"></div>' +
      '<div class="seg" id="viewSeg">' +
        ['list', 'timeline', 'calendar'].map(v => '<button data-v="' + v + '" class="' + (state.view === v ? "on" : "") + '">' + v.charAt(0).toUpperCase() + v.slice(1) + '</button>').join("") +
      '</div>';
    $("#q").addEventListener("focus", function () { this.removeAttribute("readonly"); }, { once: true });
    $("#q").oninput = e => { state.search = e.target.value; $("#qClear").style.display = e.target.value ? "" : "none"; renderContent(); };
    $("#qClear").onclick = () => { state.search = ""; $("#q").value = ""; $("#qClear").style.display = "none"; renderContent(); $("#q").focus(); };
    $("#sortF").onchange = e => { state.sort = e.target.value; renderContent(); };
    $("#catF").onchange = e => { state.category = e.target.value; renderContent(); };
    if ($("#facChips")) [...$("#facChips").querySelectorAll(".chip")].forEach(b => b.onclick = () => { state.facility = b.dataset.f; renderContent(); });
    if ($("#inact")) $("#inact").onchange = e => { state.showInactive = e.target.checked; render(); };
    [...$("#viewSeg").querySelectorAll("button")].forEach(b => b.onclick = () => {
      state.view = b.dataset.v;
      if (state.view === "calendar") { state.selectMode = false; state.selection.clear(); }
      render();
    });
  }

  function renderQuickViews() {
    const base = tabItems(state.tab);
    const cnt = k => base.filter(i => {
      if (k === "missing") return (i.isFile === false && !i.centralProof);
      if (k === "watch") return isWatched(i.id);
      return computeStatus(i).key === k;
    }).length;
    const wrap = el("div", "quickviews");
    const defs = [["", "All", base.length], ["expired", "Expired", cnt("expired")], ["critical", "Critical", cnt("critical")],
      ["due", "Due soon", cnt("due")], ["missing", "Missing proof", cnt("missing")], ["watch", "★ Watching", cnt("watch")]];
    defs.forEach(([k, label, n]) => {
      const b = el("button", "qv" + (state.quickView === k ? " on" : ""), esc(label) + (k && n != null ? ' <span class="n">' + n + '</span>' : ""));
      b.onclick = () => { state.quickView = state.quickView === k ? "" : k; render(); };
      wrap.appendChild(b);
    });
    if (state.view !== "calendar") {
      const sel = el("button", "qv" + (state.selectMode ? " on" : ""), state.selectMode ? "✓ Selecting" : "☑ Select");
      sel.style.marginLeft = "auto";
      sel.onclick = () => { state.selectMode = !state.selectMode; if (!state.selectMode) state.selection.clear(); render(); };
      wrap.appendChild(sel);
    }
    return wrap;
  }

  function renderBulkBar() {
    const bar = el("div", "bulkbar");
    const n = state.selection.size;
    bar.innerHTML = '<b>' + n + ' selected</b><span class="spacer"></span>' +
      '<button id="bAll">Select all shown</button><button id="bWatch">★ Watch</button><button id="bIcs">Calendar</button><button id="bCsv">Export CSV</button><button id="bPrint">Print</button><button id="bClear">Clear</button>';
    bar.querySelector("#bAll").onclick = () => { currentArr().forEach(i => state.selection.add(i.id)); renderContent(); };
    bar.querySelector("#bWatch").onclick = () => { selItems().forEach(i => { if (!isWatched(i.id)) toggleWatch(i.id); }); render(); toast(n + " added to watchlist."); };
    bar.querySelector("#bIcs").onclick = () => { selItems().forEach(exportICS); };
    bar.querySelector("#bCsv").onclick = () => exportCSV(selItems());
    bar.querySelector("#bPrint").onclick = () => { window.print(); };
    bar.querySelector("#bClear").onclick = () => { state.selection.clear(); render(); };
    return bar;
  }
  function selItems() { return DATA.filter(i => state.selection.has(i.id)); }

  function renderContent() {
    const c = $("#content");
    if (!UNLOCKED.has(state.tab)) { c.innerHTML = lockedView(); $("#unlockBtn").onclick = () => promptTabCode(state.tab, render); return; }
    c.innerHTML = "";
    c.appendChild(renderQuickViews());
    if (state.selectMode) c.appendChild(renderBulkBar());
    const arr = sortItems(applyFilters(tabItems(state.tab)));
    const body = el("div");
    if (state.view === "calendar") renderCalendar(body, arr);
    else if (!arr.length) body.innerHTML = emptyView();
    else if (state.view === "timeline") body.appendChild(renderTimeline(arr));
    else renderHierarchy(body, arr, state.tab);
    c.appendChild(body);
  }

  function lockedView() {
    return '<div class="empty glass" style="padding:60px"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="5" y="11" width="14" height="10" rx="2"/><path d="M8 11V7a4 4 0 018 0v4"/></svg>' +
      '<h3>This tab is locked</h3><p>Enter the access code to view ' + state.tab + ' compliance.</p>' +
      '<button class="btn-primary" id="unlockBtn" style="max-width:200px;margin:10px auto 0">Enter code</button></div>';
  }
  function emptyView() {
    return '<div class="empty glass"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4-4"/></svg>' +
      '<h3>No matching items</h3><p>Try clearing the search or status filter.</p></div>';
  }

  function wireGroupCheck(head, items) {
    const gc = head.querySelector(".grp-check"); if (!gc) return;
    const ids = items.map(i => i.id);
    const allSel = ids.length && ids.every(id => state.selection.has(id));
    const someSel = ids.some(id => state.selection.has(id));
    gc.checked = allSel; gc.indeterminate = someSel && !allSel;
    gc.onclick = (e) => { e.stopPropagation(); if (allSel) ids.forEach(id => state.selection.delete(id)); else ids.forEach(id => state.selection.add(id)); renderContent(); };
  }

  // --- Folder structure: 6 credentialing SOP phases (providers) and the
  //     14 State-Readiness sections (facilities). Mirrors the Sentinel\ subfolders.
  //     A category that maps to no bucket falls into the trailing "Other" group. ---
  const SOP_PHASES = [
    ["1. Application & Document Collection", ["CV / Resume", "Initial Application", "Medical Diploma", "Driver's License", "Malpractice / COI Insurance", "ACLS Certification", "ATLS Certification", "BLS Certification", "PALS Certification"]],
    ["2. Primary Source Verification", ["State Medical License", "Medical License Verify (annual)", "Individual DEA Registration", "DEA Verify (annual)", "Board Certification", "NPI Verification"]],
    ["3. Background & Compliance Review", ["NPDB Query (2 yrs)", "OIG / SAM Exclusion Check", "Peer References"]],
    ["4. Medical Staff Review", ["Delineation of Privileges (DOP)"]],
    ["5. Payer Enrollment & Facility Setup", ["TSCA Documents"]],
    ["6. Approval & Ongoing Monitoring", ["CME (20 hrs / 2 yrs)", "Influenza Vaccination", "TB Screening"]],
  ];
  const STATE_SECTIONS = [
    ["01. Licensing & Regulatory Compliance", ["Certificate of Occupancy", "Facility License", "DEA for Facility License", "DEA Power of Attorney", "NPI for Facility", "EIN for Facility"]],
    ["02. Personnel Files & Credentialing", []],
    ["03. Medical Staff Services", []],
    ["04. Patient Care & Clinical Documentation", []],
    ["05. Medication Management", ["Pharmacy License", "Consultant Pharmacist License", "Consultant Pharmacist Agreement"]],
    ["06. Crash Cart & Emergency Equipment", []],
    ["07. Infection Prevention & Control", []],
    ["08. Laboratory Services", ["CLIA", "COLA", "API COP"]],
    ["09. Radiology Services", ["Certificate of X-Ray Registration"]],
    ["10. Quality Improvement Program", []],
    ["11. Environment of Care", []],
    ["12. Emergency Preparedness", []],
    ["13. Patient Rights & Compliance", []],
    ["14. Daily Readiness Walkthrough", []],
  ];
  const PHASE_OF = {}; SOP_PHASES.forEach(([, cats], i) => cats.forEach(cat => PHASE_OF[cat] = i));
  const SECTION_OF = {}; STATE_SECTIONS.forEach(([, cats], i) => cats.forEach(cat => SECTION_OF[cat] = i));

  // --- UI/UX helpers: compliance ring, animated counters, section jump-nav ---
  let kpiAnimated = false;
  function miniRing(score) {
    const r = 15, c = 2 * Math.PI * r, off = c * (1 - score / 100);
    const col = score >= 85 ? "var(--st-good)" : score >= 60 ? "var(--st-due)" : "var(--st-critical)";
    return '<div class="ring-mini" title="Compliance score ' + score + '%">' +
      '<svg width="42" height="42" viewBox="0 0 40 40"><circle cx="20" cy="20" r="15" fill="none" stroke="var(--hair)" stroke-width="4"/>' +
      '<circle class="ring-fill" cx="20" cy="20" r="15" fill="none" stroke="' + col + '" stroke-width="4" stroke-linecap="round" transform="rotate(-90 20 20)" stroke-dasharray="' + c.toFixed(1) + '" stroke-dashoffset="' + off.toFixed(1) + '" style="--circ:' + c.toFixed(1) + '"/></svg>' +
      '<span class="ring-num" style="color:' + col + '">' + score + '</span></div>';
  }
  function phaseId(key) { return "ph-" + String(key).replace(/[^a-z0-9]+/gi, "-").toLowerCase(); }
  function jumpToPhase(key) {
    state.openPhases[key] = true; renderContent();
    requestAnimationFrame(() => { const e = document.getElementById(phaseId(key)); if (e) e.scrollIntoView({ behavior: "smooth", block: "center" }); });
  }
  function countUp(node, to) {
    const dur = 750, t0 = Date.now();
    node.textContent = "0";
    (function step() { const p = Math.min(1, (Date.now() - t0) / dur); node.textContent = Math.round(to * (1 - Math.pow(1 - p, 3))); if (p < 1) requestAnimationFrame(step); })();
  }

  // Render rows under COLLAPSIBLE sub-headings (dropdowns), in fixed order. Each section
  // shows its count + status pills even while collapsed, so nothing is hidden at a glance;
  // empty sections are greyed and non-clickable so the full taxonomy stays visible.
  function appendBucketed(body, items, order, ofMap, otherLabel, groupKey) {
    if (!state.openPhases) state.openPhases = {};
    const buckets = order.map(() => []);
    const extra = [];
    items.forEach(it => {
      // Honor an explicit phaseIdx (set by the gap-fill scanner) so supplemental files
      // bucket under their folder section even though their category is just the filename.
      let i = (typeof it.phaseIdx === "number") ? it.phaseIdx : ofMap[it.category];
      if (i == null || i >= order.length) extra.push(it); else buckets[i].push(it);
    });
    const sections = order.map(([label], i) => ({ label, list: buckets[i] }));
    if (extra.length) sections.push({ label: otherLabel || "Other", list: extra });
    const worstOf = list => list.map(i => computeStatus(i).key).sort((a, b) => STATUS_RANK[a] - STATUS_RANK[b])[0];

    // sticky jump-nav across the sections that actually hold documents
    const filled = sections.filter(s => s.list.length);
    if (filled.length > 1) {
      const nav = el("div", "sec-nav");
      filled.forEach(s => {
        const key = (groupKey || "") + "||" + s.label;
        const chip = el("button", "sec-chip wb-" + worstOf(s.list));
        chip.innerHTML = '<span class="dot"></span>' + esc(s.label.replace(/\.$/, "")) + '<span class="sec-chip-n">' + s.list.length + '</span>';
        chip.onclick = (e) => { e.stopPropagation(); jumpToPhase(key); };
        nav.appendChild(chip);
      });
      body.appendChild(nav);
    }

    sections.forEach(({ label, list }) => {
      const n = list.length;
      const key = (groupKey || "") + "||" + label;
      const open = n ? !!state.openPhases[key] : false;
      const gs = n ? statsFor(list) : null;
      const worst = n ? worstOf(list) : "";
      const pills = gs ? ["expired", "critical", "due"].filter(k => gs[k])
        .map(k => '<span class="pill s-' + k + '">' + gs[k] + " " + k + '</span>').join("") : "";
      const h = el("div", "phase-head" + (n ? " has-items wb-" + worst : " empty") + (open ? " open" : ""));
      if (n) h.id = phaseId(key);
      h.innerHTML =
        '<svg class="chev" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M9 6l6 6-6 6"/></svg>' +
        '<span class="phase-name">' + esc(label) + '</span>' +
        '<span class="phase-line"></span>' +
        '<span class="phase-pills">' + pills + '</span>' +
        '<span class="phase-count">' + n + '</span>';
      if (n) h.onclick = () => { state.openPhases[key] = !open; renderContent(); };
      body.appendChild(h);
      const pb = el("div", "phase-body" + (open ? " open" : ""));
      sortItems(list).forEach(it => pb.appendChild(itemRow(it)));
      body.appendChild(pb);
    });
  }

  // grouped (provider / other-by-facility)
  function renderGrouped(c, arr, key, isProvider) {
    const groups = {};
    arr.forEach(i => { (groups[i[key]] = groups[i[key]] || []).push(i); });
    Object.keys(groups).sort().forEach((name, gi) => {
      const items = groups[name];
      const gs = statsFor(items);
      const worst = items.map(i => computeStatus(i).key).sort((a, b) => STATUS_RANK[a] - STATUS_RANK[b])[0];
      const open = state.openGroups[name];
      const g = el("div", "group glass" + (open ? " open" : ""));
      const initials = name.split(/\s+/).map(w => w[0]).slice(0, 2).join("").toUpperCase();
      const pills = ["expired", "critical", "due"].filter(k => gs[k]).map(k => '<span class="pill s-' + k + '">' + gs[k] + ' ' + k + '</span>').join("");
      const inactiveTag = (isProvider && items[0].active === false) ? ' <span class="pill s-pending">inactive</span>' : "";
      const rosterNote = items[0].rosterNote ? ' <span class="pill s-pending">' + esc(items[0].rosterNote) + '</span>' : "";
      const isTest = items[0].entityKey === "aijaz-imad";
      const testTag = isTest ? ' <span class="pill s-good" style="background:#16a34a;color:#fff">TEST</span>' : "";
      const head = el("div", "group-head" + ((!isProvider || isTest) ? " green" : ""));
      head.innerHTML = (state.selectMode ? '<input type="checkbox" class="row-check grp-check" title="Select all items in this group">' : "") +
        '<div class="avatar" style="' + (isProvider ? "" : "background:linear-gradient(135deg,#6366f1,#4f46e5)") + '">' + (isProvider ? esc(initials) : '<svg style="width:20px;height:20px" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2">' + ICONS.building + '</svg>') + '</div>' +
        '<div><div class="g-name"><span style="opacity:.55">' + (gi + 1) + '.</span> ' + esc(name) + testTag + inactiveTag + rosterNote + '</div><div class="g-meta">' + items.length + ' tracked items · health ' + gs.score + '</div></div>' +
        '<div class="mini-stats">' + miniRing(gs.score) + pills + (isProvider ? '<button class="icon-btn profile-btn" title="Full credentialing profile in one table" style="padding:5px 10px">👤 Profile</button><button class="icon-btn pemail-btn" title="Email this provider (to their email)" style="padding:5px 10px">✉ Email provider</button><button class="icon-btn portal-btn" title="Provider self-service portal (QR / link)" style="padding:5px 10px">🔗 Portal</button><button class="icon-btn binder-btn" title="Print survey-ready binder" style="padding:5px 10px">🗂 Binder</button>' : '<button class="icon-btn gemail-btn" title="Email me this group\'s report" style="padding:5px 10px">✉ Email</button>') + '<span class="worst-dot bg-' + worst + '"></span></div>' +
        '<svg class="chev" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:20px;height:20px"><path d="M9 6l6 6-6 6"/></svg>';
      head.onclick = () => { state.openGroups[name] = !open; renderContent(); };
      g.appendChild(head);
      const prb = head.querySelector(".profile-btn"); if (prb) prb.onclick = (e) => { e.stopPropagation(); openProviderProfile(items[0].entityKey, name); };
      const bb = head.querySelector(".binder-btn"); if (bb) bb.onclick = (e) => { e.stopPropagation(); printBinder(name); };
      const pb = head.querySelector(".portal-btn"); if (pb) pb.onclick = (e) => { e.stopPropagation(); openProviderPortal(items[0].entityKey, name); };
      const peb = head.querySelector(".pemail-btn"); if (peb) peb.onclick = (e) => { e.stopPropagation(); openEmailTemplate(items[0]); };
      const geb = head.querySelector(".gemail-btn"); if (geb) geb.onclick = (e) => { e.stopPropagation(); emailGroupToSelf(name, items); };
      wireGroupCheck(head, items);
      const body = el("div", "group-body");
      if (isProvider) appendBucketed(body, items, SOP_PHASES, PHASE_OF, "Other documents", name);
      else sortItems(items).forEach(it => body.appendChild(itemRow(it)));
      g.appendChild(body);
      c.appendChild(g);
    });
  }

  function renderFacility(c, arr) {
    const facs = state.facility === "all" ? ["Castle Hills ER", "Frisco ER"] : [state.facility];
    const sel = el("div", "fac-selector");
    ["all", "Castle Hills ER", "Frisco ER"].forEach(f => {
      const chip = el("button", "fac-chip" + (state.facility === f ? " on" : ""), f === "all" ? "All facilities" : esc(f));
      chip.onclick = () => { state.facility = f; renderContent(); };
      sel.appendChild(chip);
    });
    c.appendChild(sel);
    facs.forEach((f, fi) => {
      const items = arr.filter(i => i.entity === f);
      if (!items.length) return;
      const gs = statsFor(items);
      const card = el("div", "group glass open");
      const head = el("div", "group-head green"); head.style.cursor = "default";
      head.innerHTML = (state.selectMode ? '<input type="checkbox" class="row-check grp-check" title="Select all in this facility">' : "") +
        '<div class="avatar"><svg style="width:20px;height:20px" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2">' + ICONS.building + '</svg></div>' +
        '<div><div class="g-name"><span style="opacity:.55">' + (fi + 1) + '.</span> ' + esc(f) + '</div><div class="g-meta">' + items.length + ' licenses & certificates · health ' + gs.score + '</div></div>' +
        '<div class="mini-stats">' + miniRing(gs.score) + ["expired", "critical", "due"].filter(k => gs[k]).map(k => '<span class="pill s-' + k + '">' + gs[k] + ' ' + k + '</span>').join("") + '<button class="icon-btn gemail-btn" title="Email me this facility\'s report" style="padding:5px 10px">✉ Email</button><button class="icon-btn binder-btn" title="Print survey-ready binder" style="padding:5px 10px">🗂 Binder</button></div>';
      const bb = head.querySelector(".binder-btn"); if (bb) bb.onclick = (e) => { e.stopPropagation(); printBinder(f); };
      const ge = head.querySelector(".gemail-btn"); if (ge) ge.onclick = (e) => { e.stopPropagation(); emailGroupToSelf(f, items); };
      wireGroupCheck(head, items);
      card.appendChild(head);
      const body = el("div", "group-body");
      appendBucketed(body, items, STATE_SECTIONS, SECTION_OF, "Other licenses", f);
      card.appendChild(body);
      c.appendChild(card);
    });
  }

  // ===== Flip hierarchy: square tiles, drill-down (entity -> section -> docs) =====
  function subgroupsFor(tab) {
    if (tab === "provider") return { order: SOP_PHASES, of: PHASE_OF, other: "Other documents" };
    if (tab === "facility") return { order: STATE_SECTIONS, of: SECTION_OF, other: "Other licenses" };
    return null;
  }
  function navigate(drill) { state.drill = drill; state._flip = true; renderContent(); }
  function worstKey(items) { return items.map(i => computeStatus(i).key).sort((a, b) => STATUS_RANK[a] - STATUS_RANK[b])[0] || "none"; }
  function earliestExp(items) { return Math.min.apply(null, items.map(i => i.expires ? parseD(i.expires) : Infinity).concat(Infinity)); }
  function sortDocs(list) {
    const by = state.sort || "name";
    if (by === "name") return list.slice().sort((a, b) => (a.category || "").localeCompare(b.category || ""));
    if (by === "expiry") return list.slice().sort((a, b) => (a.expires ? parseD(a.expires) : Infinity) - (b.expires ? parseD(b.expires) : Infinity));
    return sortItems(list);
  }
  function sortedEntityNames(groups) {
    const by = state.sort || "name", names = Object.keys(groups);
    if (by === "name") return names.sort((a, b) => a.localeCompare(b));
    if (by === "expiry") return names.sort((a, b) => earliestExp(groups[a]) - earliestExp(groups[b]));
    return names.sort((a, b) => (STATUS_RANK[worstKey(groups[a])] - STATUS_RANK[worstKey(groups[b])]) || a.localeCompare(b));
  }
  function initials(name) { return name.split(/\s+/).map(w => w[0]).slice(0, 2).join("").toUpperCase(); }

  // open a proof file straight in the Microsoft 365 (Office/Outlook) web viewer
  function fileViewerUrl(it) {
    let raw = it.fileLink || "";
    if (CLOUD && it.isFile && raw.indexOf("../") === 0 && !/^https?:/i.test(raw)) raw = sharePointLink(raw);
    if (!/^https?:/i.test(raw)) return null;
    return raw + (raw.indexOf("?") >= 0 ? "&" : "?") + "web=1";
  }
  function openFile(it) { const u = fileViewerUrl(it); if (u) window.open(u, "_blank", "noopener"); else openDrawer(it, false); }

  function entityTile(name, items, tab, serial) {
    const gs = statsFor(items);
    const isPeople = tab === "provider" || tab === "staff";
    const total = items.length;
    const onFile = items.filter(i => i.isFile || i.centralProof).length;
    const exp = gs.expired || 0, crit = gs.critical || 0, due = gs.due || 0;
    const docsPct = total ? Math.round(onFile / total * 100) : 0;
    const validPct = total ? Math.max(0, Math.round((total - exp - crit - due) / total * 100)) : 0;
    const urgent = exp + crit;
    const badge = exp ? ["action needed", "b-act"] : ((crit || due) ? ["in progress", "b-prog"] : ["ready", "b-ready"]);
    const barc = p => p < 40 ? "f-amber" : (p < 80 ? "f-pri" : "f-green");
    const icon = isPeople ? '<div class="c-ic">' + esc(initials(name)) + '</div>'
      : '<div class="c-ic"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="4" y="3" width="16" height="18" rx="2"/><path d="M9 8h6M9 12h6M9 16h4"/></svg></div>';
    const test = items[0] && items[0].entityKey === "aijaz-imad" ? ' <span class="badge b-ready" style="margin-left:6px">TEST</span>' : "";
    const t = el("div", "tile tile-entity");
    t.innerHTML =
      '<div class="c-top">' + icon + (serial ? '<span class="c-sn">#' + serial + '</span>' : '') + '<span class="badge ' + badge[1] + '">' + badge[0] + '</span></div>' +
      '<div class="c-name">' + esc(name) + test + '</div>' +
      '<div class="prow"><div class="prow-h"><span>Documents on file</span><b>' + docsPct + '%</b></div><div class="bar"><i class="' + barc(docsPct) + '" style="width:' + docsPct + '%"></i></div></div>' +
      '<div class="prow"><div class="prow-h"><span>Currently valid</span><b>' + validPct + '%</b></div><div class="bar"><i class="' + barc(validPct) + '" style="width:' + validPct + '%"></i></div></div>' +
      '<div class="c-foot"><span class="l">' + total + ' items · ' + onFile + ' on file</span>' + (urgent ? '<span class="w">⚠ ' + urgent + ' expiring</span>' : '') + '</div>';
    t.onclick = () => navigate([name]);
    return t;
  }
  function sectionTile(label, items, entity) {
    const n = items.length;
    const gs = n ? statsFor(items) : null;
    const num = (label.match(/^\d+/) || [""])[0];
    const nameOnly = label.replace(/^\d+\.\s*/, "");
    const onFile = items.filter(i => i.isFile || i.centralProof).length;
    const pct = n ? Math.round(onFile / n * 100) : 0;
    const barc = p => p < 40 ? "f-amber" : (p < 80 ? "f-pri" : "f-green");
    const badge = !n ? null : (gs.expired ? ["action needed", "b-act"] : ((gs.critical || gs.due) ? ["in progress", "b-prog"] : ["ready", "b-ready"]));
    const t = el("div", "tile tile-entity sec" + (n ? "" : " empty"));
    t.innerHTML =
      '<div class="c-top"><div class="c-ic">' + (num || "•") + '</div>' + (badge ? '<span class="badge ' + badge[1] + '">' + badge[0] + '</span>' : '') + '</div>' +
      '<div class="c-name">' + esc(nameOnly) + '</div>' +
      (n ? '<div class="prow"><div class="prow-h"><span>Documents on file</span><b>' + pct + '%</b></div><div class="bar"><i class="' + barc(pct) + '" style="width:' + pct + '%"></i></div></div>' : '') +
      '<div class="c-foot"><span class="l">' + (n ? onFile + ' of ' + n + ' on file' : 'No items yet') + '</span>' + (gs && gs.expired ? '<span class="w">⚠ ' + gs.expired + ' expired</span>' : '') + '</div>';
    if (n) {
      t.onclick = () => navigate([entity, label]);
    } else t.classList.add("disabled");
    return t;
  }
  function docTile(it, showEntity) {
    const s = computeStatus(it);
    const base = [it.authority, it.number ? "#" + it.number : ""].filter(Boolean).join(" · ");
    const sub = showEntity ? (it.entity + (base ? " · " + base : "")) : base;
    const bmap = { expired: ["expired", "b-act"], critical: ["critical", "b-act"], due: ["due soon", "b-prog"], good: ["valid", "b-ready"], permanent: ["on file", "b-ready"], recurring: ["recurring", "b-prog"], pending: ["pending", "b-prog"], none: ["—", "b-prog"] };
    const badge = bmap[s.key] || [s.label, "b-prog"];
    const proof = it.isFile ? '<span class="proof-badge has">📄 Proof</span>'
      : (it.centralProof ? '<span class="proof-badge has" title="Covered by a carrier-level policy from the COI roster">🛡 Via roster</span>'
      : (it.isFile === false ? '<span class="proof-badge no">⚠ No proof</span>' : ""));
    const mail = '<button class="tile-mail" title="Email me this item" aria-label="Email me this item">✉</button>';
    const t = el("div", "tile tile-entity doc" + (it.supplemental ? " supp" : ""));
    t.innerHTML =
      '<div class="c-top"><div class="c-ic"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' + ICONS[iconFor(it)] + '</svg></div><span class="badge ' + badge[1] + '">' + badge[0] + '</span></div>' +
      '<div class="c-name">' + esc(it.category) + '</div>' +
      (sub ? '<div class="doc-sub">' + esc(sub) + '</div>' : '') +
      '<div class="c-foot"><span class="l">' + (it.expires ? fmtD(it.expires) : (it.permanent ? "No expiry" : ((it.isFile === false && !it.centralProof) ? "Awaiting upload" : "—"))) + '</span><span class="doc-fr">' + proof + mail + '</span></div>';
    const mb = t.querySelector(".tile-mail"); if (mb) mb.onclick = (e) => { e.stopPropagation(); emailGroupToSelf(it.entity + " — " + it.category, [it]); };
    t.onclick = () => openDrawer(it, false);
    return t;
  }
  function entityHeader(name, items, tab) {
    const gs = statsFor(items);
    const isProv = tab === "provider";
    const wrap = el("div", "ent-head glass");
    wrap.innerHTML = '<div class="ent-ic">' + (isProv ? esc(initials(name)) : '<svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" style="width:20px;height:20px">' + ICONS.building + '</svg>') + '</div>' +
      '<div class="ent-info"><div class="ent-nm">' + esc(name) + '</div><div class="ent-meta">' + items.length + ' tracked items</div></div>' +
      '<div class="ent-actions">' + (isProv
        ? '<button class="icon-btn" data-a="profile" title="Full credentialing profile in one table">👤 Profile</button><button class="icon-btn" data-a="pemail">✉ Email provider</button><button class="icon-btn" data-a="portal">🔗 Portal</button><button class="icon-btn" data-a="binder">🗂 Binder</button>' +
          (isAdmin() ? '<button class="icon-btn" data-a="inact" title="Move to Inactive Providers (keeps the row, marks inactive)">📦 Mark inactive</button><button class="icon-btn danger" data-a="delprov" title="Permanently delete from the roster (recoverable from Recycle bin)">🗑 Delete</button>' : "")
        : '<button class="icon-btn" data-a="email">✉ Email</button><button class="icon-btn" data-a="binder">🗂 Binder</button>') +
      ((tab === "provider" || tab === "staff") ? '<button class="icon-btn" data-a="onboard">📋 Onboarding</button>' : '') + '</div>';
    const it0 = items[0] || {};
    const bind = (a, fn) => { const b = wrap.querySelector('[data-a="' + a + '"]'); if (b) b.onclick = fn; };
    bind("profile", () => openProviderProfile(it0.entityKey, name));
    bind("pemail", () => openEmailTemplate(it0));
    bind("portal", () => openProviderPortal(it0.entityKey, name));
    bind("binder", () => printBinder(name));
    bind("email", () => emailGroupToSelf(name, items));
    bind("inact", () => removeProviderFromRoster(name, it0.entityKey));    // moves to Inactive sheet
    bind("delprov", () => deleteProviderHard(name, it0.entityKey));         // hard delete + trash
    bind("onboard", () => openOnboarding(it0.entityKey, name));
    return wrap;
  }
  // ================= PROVIDER PROFILE =================
  // Renders WCGTX's own "Physician Pre-Qualification Profile" form (the HR + Credentialing
  // internal review sheet) for one provider, filled in from the tracked roster data.
  // Layout, section order, wording and the checkbox/decision blocks follow that document.
  //
  // Fields the roster genuinely does not hold (degree, phone, employer/group) are shown as
  // "not tracked" rather than guessed at — this is a screening record, so an invented value is
  // worse than an obvious blank.

  // Primary-source and background screening rows, in the document's order.
  //   verify = the category holding the primary-source CHECK (drives the status)
  //   doc    = the category holding the underlying certificate/document
  const PQ_SCREEN = [
    { check: "DEA registration", source: "DEA Diversion Control Division", verify: "DEA Verify (annual)", doc: "Individual DEA Registration", ok: "VERIFIED", detail: "Exp" },
    { check: "Texas medical license", source: "Texas Medical Board", verify: "Medical License Verify (annual)", doc: "State Medical License", ok: "VERIFIED / ACTIVE", detail: "Exp" },
    { check: "NPDB query", source: "National Practitioner Data Bank", verify: "NPDB Query (2 yrs)", ok: "CLEAR", detail: "Next query due" },
    { check: "OIG exclusion check", source: "OIG LEIE", verify: "OIG / SAM Exclusion Check", ok: "CLEAR", detail: "Next due" },
    { check: "NPI record", source: "NPI Registry / NPPES", verify: "NPI Verification", ok: "VERIFIED / CLEAR", detail: "NPI" },
    { check: "Professional peer references", source: "Required references received", doc: "Peer References", ok: "GOOD", detail: "On file" },
    { check: "Malpractice history", source: "Attestation / claims review", doc: "Malpractice / COI Insurance", ok: "NO CLAIMS", detail: "As of" },
  ];
  const PQ_CERTS = [["ACLS", "ACLS Certification"], ["ATLS", "ATLS Certification"], ["PALS", "PALS Certification"], ["BLS", "BLS Certification"]];
  // Board code -> the specialty checkbox on the form. Codes come from the roster's
  // "Board Certified" column (ABEM/ABFM/ABIM plus osteopathic equivalents and others).
  const PQ_BOARDS = [
    ["EMERGENCY MEDICINE", /\b(abem|aobem)\b/i],
    ["INTERNAL MEDICINE", /\b(abim|aobim)\b/i],
    ["FAMILY MEDICINE", /\b(abfm|aafp)\b/i],
  ];

  function pfFileName(it) {
    if (!it) return "";
    if (it.uploadName) return it.uploadName;
    try { return decodeURIComponent(String(it.fileLink || "").split("/").pop() || ""); } catch (e) { return String(it.fileLink || ""); }
  }
  // Pick the most useful item per category for this provider (prefer one with a file / a date).
  function pqIndex(entityKey) {
    const mine = DATA.filter(i => i.scope === "provider" && i.entityKey === entityKey);
    const by = {};
    const score = x => (x.isFile ? 2 : 0) + (x.expires ? 1 : 0) + (x.permanent ? 1 : 0);
    mine.forEach(i => { const p = by[i.category]; if (!p || score(i) > score(p)) by[i.category] = i; });
    return { mine, by };
  }
  // One screening row's outcome. "ok" means the check is done and not lapsed.
  function pqState(it) {
    if (!it) return { key: "none", label: "NOT DONE", ok: false };
    const st = computeStatus(it);
    if (st.key === "expired") return { key: "expired", label: "EXPIRED", ok: false, st };
    if (st.key === "pending") return { key: "none", label: "OUTSTANDING", ok: false, st };
    if (!it.expires && !it.permanent && !it.isFile && !it.docStatus) return { key: "none", label: "NOT DONE", ok: false, st };
    if (st.key === "critical" || st.key === "due") return { key: st.key, label: "DUE SOON", ok: true, st };
    return { key: "ok", label: "", ok: true, st };
  }
  function pqBox(on, label) {
    return '<span class="pq-box' + (on ? " on" : "") + '">' + (on ? "☑" : "☐") + " " + esc(label) + '</span>';
  }
  function pqStyles() {
    if ($("#pqStyles")) return;
    const s = document.createElement("style"); s.id = "pqStyles";
    s.textContent =
      '.pq{font-size:12.5px;color:var(--ink,#0f172a)}' +
      '.pq-band{background:var(--accent,#0f766e);color:#fff;border-radius:10px;padding:12px 15px;margin-bottom:4px}' +
      '.pq-band .t1{font-size:11px;font-weight:800;letter-spacing:.16em;opacity:.85}' +
      '.pq-band .t2{font-size:17px;font-weight:800;letter-spacing:.03em;margin-top:2px}' +
      '.pq-band .t3{font-size:10.5px;font-weight:700;letter-spacing:.1em;opacity:.85;margin-top:3px}' +
      '.pq-sub{font-size:11px;color:var(--ink-2,#64748b);margin:6px 0 14px}' +
      '.pq-h{font-size:10.5px;font-weight:800;letter-spacing:.11em;text-transform:uppercase;color:var(--ink-2,#475569);' +
      'border-bottom:2px solid var(--accent,#0f766e);padding-bottom:5px;margin:18px 0 9px}' +
      '.pq-id{display:grid;grid-template-columns:118px 1fr 1fr;gap:9px;align-items:start}' +
      '.pq-photo{border:1px dashed var(--hair,#cbd5e1);border-radius:9px;height:132px;display:flex;flex-direction:column;' +
      'align-items:center;justify-content:center;text-align:center;font-size:9.5px;font-weight:800;letter-spacing:.07em;' +
      'color:var(--ink-3,#94a3b8);padding:6px;gap:5px}' +
      '.pq-f{border:1px solid var(--hair,#e2e8f0);border-radius:8px;padding:7px 9px;background:var(--surface-solid,#fff)}' +
      '.pq-f .k{font-size:9.5px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:var(--ink-3,#94a3b8)}' +
      '.pq-f .v{font-size:13px;font-weight:650;margin-top:2px;word-break:break-word}' +
      '.pq-f .v.na{color:var(--ink-3,#94a3b8);font-weight:600;font-style:italic}' +
      '.pq-cols2{display:grid;grid-template-columns:1fr 1fr;gap:9px}' +
      '.pq-box{display:inline-block;border:1px solid var(--hair,#e2e8f0);border-radius:7px;padding:6px 10px;font-size:11.5px;' +
      'font-weight:800;letter-spacing:.03em;color:var(--ink-2,#64748b);background:var(--surface-solid,#fff);white-space:nowrap}' +
      '.pq-box.on{color:var(--st-good,#047857);border-color:var(--st-good,#047857);background:var(--st-good-bg,#ecfdf5)}' +
      '.pq-row{display:flex;gap:7px;flex-wrap:wrap}' +
      '.pq-t{width:100%;border-collapse:collapse}' +
      '.pq-t th{font-size:9.5px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:var(--ink-3,#94a3b8);' +
      'text-align:left;padding:6px 9px;border-bottom:1px solid var(--hair,#e2e8f0)}' +
      '.pq-t td{padding:8px 9px;border-bottom:1px solid var(--hair,#eef2f7);vertical-align:middle;font-size:12.5px}' +
      '.pq-t td.c{font-weight:650}' +
      '.pq-t td.s{color:var(--ink-2,#64748b);font-size:11.5px}' +
      '.pq-t tr.bad td{background:var(--st-expired-bg,rgba(220,38,38,.05))}' +
      '.pq-st{display:inline-block;padding:3px 8px;border-radius:999px;font-size:10px;font-weight:800;letter-spacing:.04em;' +
      'border-width:1px;border-style:solid;white-space:nowrap}' +
      '.pq-st.none{color:var(--ink-2,#64748b);background:var(--bg,#f1f5f9);border-color:var(--hair,#e2e8f0)}' +
      '.pq-dec{border:2px solid var(--accent,#0f766e);border-radius:10px;padding:12px 14px;margin-top:10px;background:var(--surface-solid,#fff)}' +
      '.pq-note{border:1px solid var(--hair,#e2e8f0);border-radius:8px;min-height:52px;padding:8px 10px;font-size:12px;' +
      'color:var(--ink-2,#64748b);background:var(--surface-solid,#fff);white-space:pre-wrap}' +
      '.pq-wrap{overflow-x:auto}';
    document.head.appendChild(s);
  }

  // Resolve everything the form needs for one provider.
  function pqProfile(entityKey) {
    const { mine, by } = pqIndex(entityKey);
    const name = (mine[0] && mine[0].entity) || entityKey;
    const findDoc = re => mine.find(i => i.isFile && (re.test(i.category || "") || re.test(pfFileName(i)))) || null;

    // Board certification: the roster's Board Certified column lands in authority/number.
    const bc = by["Board Certification"] || null;
    const bcode = String((bc && (bc.authority || bc.number)) || "").trim();
    const lapsed = /lapsed/i.test(bcode);
    const boards = PQ_BOARDS.map(([label, re]) => ({ label, on: !lapsed && re.test(bcode) }));
    const other = (!lapsed && bcode && !boards.some(b => b.on) && !/^(yes|no|\?|-+\s*::?|n\/?a)$/i.test(bcode)) ? bcode : "";
    const boardVerified = boards.some(b => b.on) || !!other;
    // Only the three mapped boards name a specialty. Any other code (ABS, ABOG, ABPN…) is shown
    // as the board itself rather than dressed up as a specialty we haven't actually established.
    const specialty = boards.find(b => b.on)
      ? boards.find(b => b.on).label.replace(/\b\w+/g, w => w[0] + w.slice(1).toLowerCase())
      : (other ? "Board: " + other : "");

    const screen = PQ_SCREEN.map(spec => {
      const vIt = spec.verify ? (by[spec.verify] || null) : null;
      const dIt = spec.doc ? (by[spec.doc] || null) : null;
      // Take the WORST of the parts this row declares. Reading only the verification made an
      // EXPIRED Texas medical license show as "VERIFIED / ACTIVE" whenever the annual
      // re-verification was still in date — the most dangerous way this form could be wrong.
      const parts = [];
      if (spec.verify) parts.push(pqState(vIt));
      if (spec.doc) parts.push(pqState(dIt));
      const RANK = { expired: 0, none: 1, critical: 2, due: 2, ok: 3 };
      parts.sort((a, b) => RANK[a.key] - RANK[b.key]);
      const state = parts[0] || pqState(null);
      const driver = vIt || dIt;
      let detail = "";
      if (spec.detail === "Exp") {
        // Two different dates live on this row: the certificate's own expiry (the document's
        // "Exp:" field) and the annual re-verification. Show both, or a row can read EXPIRED
        // beside a future expiry date and look self-contradictory.
        detail = dIt && dIt.expires ? "Exp: " + fmtD(dIt.expires) : (dIt && dIt.isFile ? "Certificate on file, no date" : "Exp: —");
        if (vIt && vIt.expires) {
          detail += computeStatus(vIt).key === "expired"
            ? " · re-verification overdue since " + fmtD(vIt.expires)
            : " · next verification " + fmtD(vIt.expires);
        } else if (!vIt) detail += " · not yet verified";
      }
      else if (spec.detail === "NPI") detail = (vIt && String(vIt.number || "").trim()) ? "NPI: " + vIt.number : (vIt && (vIt.isFile || vIt.docStatus) ? "Registry record on file" : "NPI: —");
      else if (spec.detail === "On file") detail = dIt && dIt.isFile ? pfFileName(dIt) : (dIt && dIt.docStatus ? dIt.docStatus : "—");
      else if (spec.detail === "As of") detail = dIt && dIt.expires ? "As of: " + fmtD(dIt.expires) : (dIt && dIt.isFile ? "Attestation on file" : "As of: —");
      else detail = driver && driver.expires ? spec.detail + ": " + fmtD(driver.expires) : spec.detail + ": —";
      return { spec, state, detail, doc: dIt, verify: vIt };
    });

    const certs = PQ_CERTS.map(([label, cat]) => {
      const it = by[cat] || null;
      const st = it ? computeStatus(it) : null;
      const current = !!(it && st && st.key !== "expired" && st.key !== "pending");
      return { label, it, st, current, exp: it && it.expires ? fmtD(it.expires) : (it && it.permanent ? "Permanent" : "—") };
    });

    const allScreenOk = screen.every(r => r.state.ok);
    // The form mandates board certification ("at least one must be verified") but does NOT say
    // all four life-support cards are required — which varies by role. So this line means "every
    // certification actually on record is current", not "holds all four". A card that is simply
    // absent is listed under outstanding items instead of being treated as a policy failure,
    // because inventing a stricter rule than the form states would misreport people.
    const onRecord = certs.filter(c => c.it);
    const certsCurrent = onRecord.length > 0 && onRecord.every(c => c.current);
    // Decision policy — deliberately conservative.
    //   GREEN               everything verified, current and clear
    //   HOLD / FOLLOW-UP    ANY gap, expiry or missing record. This is the default.
    //   DOES NOT MEET       only when the roster POSITIVELY records no board certification
    //                       (the column literally reads "No").
    // A missing or blank record is a data gap, not a finding about the physician, so it must
    // never auto-produce "does not meet criteria" — that is a statement about a real person's
    // professional standing. An overdue annual re-verification is likewise a follow-up, not a
    // disqualification: the underlying licence or DEA registration can still be valid.
    // The reviewer can always tick a different box; the app only proposes.
    const boardExplicitlyNone = /^no$/i.test(bcode);
    let decision = "HOLD / FOLLOW-UP";
    if (allScreenOk && boardVerified && certsCurrent) decision = "GREEN - ADVANCE";
    else if (boardExplicitlyNone) decision = "DOES NOT MEET CRITERIA";

    return {
      entityKey, name, mine, by, screen, certs, boards, other, boardVerified, bcode, lapsed, specialty,
      email: (mine.find(i => i.email) || {}).email || "",
      headshot: findDoc(/photo|headshot|badge/i),
      summary: { allScreenOk, boardVerified, certsCurrent },
      decision,
    };
  }

  function openProviderProfile(entityKey, name) {
    pqStyles();
    const p = pqProfile(entityKey);
    const fld = (k, v, na) => '<div class="pq-f"><div class="k">' + esc(k) + '</div><div class="v' + (na ? " na" : "") + '">' + esc(v) + '</div></div>';
    const NT = "not tracked";

    let h = '<div class="pq">';
    h += '<div class="pq-band"><div class="t1">WCGTX</div><div class="t2">PHYSICIAN PRE-QUALIFICATION PROFILE</div>' +
      '<div class="t3">HR + CREDENTIALING &nbsp;·&nbsp; INTERNAL REVIEW</div></div>' +
      '<div class="pq-sub">Texas physician staffing &nbsp;|&nbsp; Preliminary screening before full credentialing</div>';

    // Identity block
    h += '<div class="pq-id"><div class="pq-photo"><div>PROFESSIONAL<br>HEADSHOT</div>' +
      (p.headshot ? '<div style="color:var(--st-good,#047857)">☑ ON FILE</div>' : '<div>[ NOT ON FILE ]</div>') + '</div>' +
      '<div style="display:grid;gap:9px">' +
      fld("Physician name", p.name) +
      fld("Primary specialty", p.specialty || NT, !p.specialty) +
      fld("State", "TEXAS") +
      fld("Email", p.email || NT, !p.email) +
      '</div><div style="display:grid;gap:9px">' +
      fld("Credentials / degree", NT, true) +
      fld("Profile ID", p.entityKey) +
      fld("Phone", NT, true) +
      fld("Current employer / group", NT, true) +
      '</div></div>';

    // Board certification
    h += '<div class="pq-h">Required board certification — at least one must be verified</div>' +
      '<div class="pq-row">' + p.boards.map(b => pqBox(b.on, b.label)).join("") +
      pqBox(!!p.other, "OTHER: " + (p.other || "__________")) + '</div>' +
      (p.lapsed ? '<div class="pq-sub" style="margin:8px 0 0;color:var(--st-expired,#b91c1c)">Roster records this board certification as <b>' + esc(p.bcode) + '</b> — treated as not verified.</div>'
        : (!p.boardVerified ? '<div class="pq-sub" style="margin:8px 0 0;color:var(--st-expired,#b91c1c)">No board certification recorded on the roster' + (p.bcode ? ' (column reads "' + esc(p.bcode) + '")' : '') + '.</div>' : ''));

    // Screening table
    h += '<div class="pq-h">Primary-source and background screening</div><div class="pq-wrap">' +
      '<table class="pq-t"><thead><tr><th>Check</th><th>Verification source</th><th>Status</th><th>Detail</th></tr></thead><tbody>';
    p.screen.forEach(r => {
      h += '<tr' + (r.state.ok ? "" : ' class="bad"') + '><td class="c">' + esc(r.spec.check) + '</td>' +
        '<td class="s">' + esc(r.spec.source) + '</td>' +
        '<td>' + (r.state.ok
          ? '<span class="pq-st s-' + (r.state.key === "ok" ? "good" : r.state.key) + '">☑ ' + esc(r.state.label || r.spec.ok) + '</span>'
          : '<span class="pq-st ' + (r.state.key === "expired" ? "s-expired" : "none") + '">☐ ' + esc(r.state.label) + '</span>') + '</td>' +
        '<td class="s">' + esc(r.detail) + '</td></tr>';
    });
    h += '</tbody></table></div>';

    // Life-support certifications
    h += '<div class="pq-h">Life-support &amp; trauma certifications</div><div class="pq-row">' +
      p.certs.map(c => '<span class="pq-box' + (c.current ? " on" : "") + '">' + (c.current ? "☑" : "☐") + " " + esc(c.label) +
        ' &nbsp;' + (c.current ? "CURRENT" : (c.it ? "EXPIRED" : "MISSING")) + ' &nbsp;<span style="font-weight:600;opacity:.75">Exp: ' + esc(c.exp) + '</span></span>').join("") +
      '</div>';

    // Summary + decision
    const dOk = p.decision === "GREEN - ADVANCE", dBad = p.decision === "DOES NOT MEET CRITERIA";
    h += '<div class="pq-h">Credentialing summary &amp; pre-qualification decision</div>' +
      '<div style="display:grid;gap:7px">' +
      pqBox(p.summary.allScreenOk, "All required checks green / verified / clear") +
      pqBox(p.summary.boardVerified, "At least one required board certification verified") +
      pqBox(p.summary.certsCurrent, "Required certifications current or renewal plan documented") +
      '</div>' +
      '<div class="pq-dec"><div class="k" style="font-size:9.5px;font-weight:800;letter-spacing:.08em;color:var(--ink-3,#94a3b8);text-transform:uppercase">Pre-qualification decision</div>' +
      '<div class="pq-row" style="margin-top:7px">' +
      pqBox(dOk, "GREEN - ADVANCE") + pqBox(!dOk && !dBad, "HOLD / FOLLOW-UP") + pqBox(dBad, "DOES NOT MEET CRITERIA") +
      '</div><div class="pq-sub" style="margin:9px 0 0">Suggested from the tracked roster data only. A reviewer confirms and signs off below — this is not an automatic determination.</div></div>';

    // Reviewer + notes (left blank for the human, as on the form)
    h += '<div class="pq-h">Reviewer</div><div class="pq-cols2">' +
      fld("Reviewer name / title", "—", true) + fld("Review date", "—", true) + '</div>' +
      '<div class="pq-h">Notes / outstanding items</div><div class="pq-note">' +
      esc(pqOutstanding(p).join("\n") || "—") + '</div></div>';

    const head = '<button class="icon-btn" id="pfPrint" style="padding:5px 10px">\u{1F5A8} Print</button>';
    openModal("Pre-Qualification Profile — " + esc(name), h, head);
    const pb = $("#pfPrint"); if (pb) pb.onclick = () => printProfile(entityKey, name);
  }

  // Auto-fill the "outstanding items" box with what is actually missing, so the reviewer starts
  // from a real list instead of a blank line.
  function pqOutstanding(p) {
    const out = [];
    p.screen.forEach(r => { if (!r.state.ok) out.push(r.spec.check + " — " + r.state.label.toLowerCase() + " (" + r.spec.source + ")"); });
    if (!p.boardVerified) out.push("Board certification not verified" + (p.bcode ? " (roster reads \"" + p.bcode + "\")" : ""));
    p.certs.forEach(c => { if (!c.current) out.push(c.label + (c.it ? " expired " + c.exp : " not on file")); });
    if (!p.headshot) out.push("Headshot for badge not on file");
    if (!p.email) out.push("No email address on the roster");
    return out;
  }

  function printProfile(entityKey, name) {
    const p = pqProfile(entityKey);
    const box = (on, t) => (on ? "[X] " : "[ ] ") + t;
    const rows = p.screen.map(r =>
      '<tr><td>' + esc(r.spec.check) + '</td><td>' + esc(r.spec.source) + '</td><td>' +
      esc(r.state.ok ? box(true, r.state.label || r.spec.ok) : box(false, r.state.label)) + '</td><td>' + esc(r.detail) + '</td></tr>').join("");
    const pa = $("#printArea");
    pa.innerHTML = '<div class="binder">' +
      '<h1>' + esc(p.name) + '</h1>' +
      '<div class="bsub">WCGTX Physician Pre-Qualification Profile &middot; HR + Credentialing internal review &middot; Texas physician staffing &middot; printed ' + new Date().toLocaleDateString() + '</div>' +
      '<table class="btable"><tbody>' +
      '<tr><td><b>Physician name</b></td><td>' + esc(p.name) + '</td><td><b>Profile ID</b></td><td>' + esc(p.entityKey) + '</td></tr>' +
      '<tr><td><b>Primary specialty</b></td><td>' + esc(p.specialty || "not tracked") + '</td><td><b>State</b></td><td>TEXAS</td></tr>' +
      '<tr><td><b>Email</b></td><td>' + esc(p.email || "not tracked") + '</td><td><b>Headshot</b></td><td>' + (p.headshot ? "On file" : "Not on file") + '</td></tr>' +
      '<tr><td><b>Credentials / degree</b></td><td>not tracked</td><td><b>Phone</b></td><td>not tracked</td></tr>' +
      '<tr><td><b>Current employer / group</b></td><td colspan="3">not tracked</td></tr>' +
      '</tbody></table>' +
      '<h2 style="font-size:12pt;margin:14px 0 6px">Required board certification</h2>' +
      '<div>' + p.boards.map(b => box(b.on, b.label)).join(" &nbsp; ") + " &nbsp; " + box(!!p.other, "OTHER: " + (p.other || "__________")) + '</div>' +
      '<h2 style="font-size:12pt;margin:14px 0 6px">Primary-source and background screening</h2>' +
      '<table class="btable"><thead><tr><th>Check</th><th>Verification source</th><th>Status</th><th>Detail</th></tr></thead><tbody>' + rows + '</tbody></table>' +
      '<h2 style="font-size:12pt;margin:14px 0 6px">Life-support &amp; trauma certifications</h2>' +
      '<div>' + p.certs.map(c => box(c.current, c.label + " " + (c.current ? "CURRENT" : (c.it ? "EXPIRED" : "MISSING")) + " — Exp: " + c.exp)).join(" &nbsp; ") + '</div>' +
      '<h2 style="font-size:12pt;margin:14px 0 6px">Credentialing summary &amp; pre-qualification decision</h2>' +
      '<div>' + box(p.summary.allScreenOk, "All required checks green / verified / clear") + '<br>' +
      box(p.summary.boardVerified, "At least one required board certification verified") + '<br>' +
      box(p.summary.certsCurrent, "Required certifications current or renewal plan documented") + '</div>' +
      '<div style="margin-top:10px"><b>Pre-qualification decision:</b> ' +
      box(p.decision === "GREEN - ADVANCE", "GREEN - ADVANCE") + " &nbsp; " +
      box(p.decision === "HOLD / FOLLOW-UP", "HOLD / FOLLOW-UP") + " &nbsp; " +
      box(p.decision === "DOES NOT MEET CRITERIA", "DOES NOT MEET CRITERIA") +
      '<div style="font-size:9pt;margin-top:4px">Suggested from tracked roster data only; the reviewer below confirms.</div></div>' +
      '<h2 style="font-size:12pt;margin:14px 0 6px">Notes / outstanding items</h2>' +
      '<div style="white-space:pre-wrap;font-size:10pt">' + esc(pqOutstanding(p).join("\n") || "—") + '</div>' +
      '<table class="btable" style="margin-top:14px"><tbody><tr><td><b>Reviewer name / title</b></td><td>__________________________</td>' +
      '<td><b>Review date</b></td><td>______________</td></tr></tbody></table>' +
      '<div class="bfoot">Generated by Sentinel Compliance Command Center from the WCGTX master roster.</div></div>';
    document.body.classList.add("printing");
    window.print();
    setTimeout(() => document.body.classList.remove("printing"), 400);
  }
  // ---- Onboarding status board (per new-hire process checklist, saved to OneDrive) ----
  var ONBOARD_STEPS = [
    ["welcome", "Welcome letter sent"],
    ["docs_requested", "Documents requested (PandaDoc)"],
    ["docs_received", "Documents received"],
    ["background", "Background check"],
    ["references", "References checked"],
    ["roster", "Added to roster / Paycor"],
    ["orientation", "Orientation scheduled"],
  ];
  function onbFetch() { return fetch("/api/data?onboard=all&_t=" + Date.now(), { cache: "no-store" }).then(r => r.json()).catch(() => ({})); }
  function openOnboarding(entityKey, name) {
    if (!entityKey) { toast("No record for this entity."); return; }
    toast("Loading onboarding…");
    onbFetch().then(all => {
      var rec = (all && all[entityKey]) || { steps: {} };
      var draw = () => {
        var done = ONBOARD_STEPS.filter(s => rec.steps && rec.steps[s[0]] && rec.steps[s[0]].done).length;
        var pct = Math.round(done / ONBOARD_STEPS.length * 100);
        var rows = ONBOARD_STEPS.map(function (s) {
          var st = rec.steps && rec.steps[s[0]];
          var on = st && st.done;
          var meta = on && st.at ? '<div class="item-sub" style="font-size:11px">done ' + esc(new Date(st.at).toLocaleDateString()) + (st.by ? ' · ' + esc(st.by) : '') + '</div>' : '';
          return '<label style="display:flex;align-items:flex-start;gap:11px;padding:11px 12px;border:1px solid var(--hair);border-radius:10px;background:var(--surface-solid);cursor:' + (isAdmin() ? 'pointer' : 'default') + '">' +
            '<input type="checkbox" class="onb" data-s="' + s[0] + '"' + (on ? ' checked' : '') + (isAdmin() ? '' : ' disabled') + ' style="margin-top:2px">' +
            '<div style="flex:1"><div style="font-weight:600' + (on ? ';color:var(--st-good)' : '') + '">' + esc(s[1]) + '</div>' + meta + '</div></label>';
        }).join("");
        openModal("Onboarding — " + esc(name),
          '<div class="item-sub" style="margin-bottom:10px">' + done + ' of ' + ONBOARD_STEPS.length + ' steps complete</div>' +
          '<div class="ent-bar" style="height:8px;margin-bottom:14px"><i style="display:block;height:100%;border-radius:999px;background:var(--accent);width:' + pct + '%"></i></div>' +
          '<div style="display:flex;flex-direction:column;gap:8px">' + rows + '</div>' +
          (isAdmin() ? '' : '<div class="item-sub" style="margin-top:10px">View only — admins can update steps.</div>'));
        if (isAdmin()) [].forEach.call($("#modalInner").querySelectorAll(".onb"), function (cb) {
          cb.onchange = function () {
            var step = cb.dataset.s, done2 = cb.checked;
            cb.disabled = true;
            fetch("/api/data?onboard=set", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ entityKey: entityKey, entity: name, step: step, done: done2 }) })
              .then(r => r.json()).then(function (d) {
                if (!d.ok) { toast("Save failed: " + (d.error || "")); cb.checked = !done2; cb.disabled = false; return; }
                rec = d.record || rec; draw();
              // Revert the tick on a network/timeout failure too, not just on an {ok:false} body.
              // Leaving it checked showed the step as done while the server never recorded it —
              // the box looked saved until someone reopened it from another machine.
              }).catch(function (e) { toast("Save failed: " + (e.message || e)); cb.checked = !done2; cb.disabled = false; });
          };
        });
      };
      draw();
    }).catch(() => toast("Couldn't load onboarding."));
  }
  function openOnboardingBoard() {
    toast("Loading onboarding board…");
    onbFetch().then(all => {
      var recs = Object.keys(all || {}).map(function (k) {
        var r = all[k]; var done = ONBOARD_STEPS.filter(s => r.steps && r.steps[s[0]] && r.steps[s[0]].done).length;
        return { key: k, name: r.entity || k, done: done, total: ONBOARD_STEPS.length };
      }).filter(r => r.done < r.total).sort((a, b) => a.done - b.done);
      var body = recs.length
        ? '<div class="item-sub" style="margin-bottom:12px">People currently mid-onboarding (' + recs.length + '). Completed ones drop off automatically.</div>' +
          '<div style="display:flex;flex-direction:column;gap:8px">' +
          recs.map(function (r) {
            var pct = Math.round(r.done / r.total * 100);
            return '<div style="display:flex;align-items:center;gap:12px;padding:10px 12px;border:1px solid var(--hair);border-radius:10px;background:var(--surface-solid)">' +
              '<div style="flex:1"><div style="font-weight:700">' + esc(r.name) + '</div>' +
              '<div class="ent-bar" style="height:6px;margin-top:6px"><i style="display:block;height:100%;border-radius:999px;background:var(--accent);width:' + pct + '%"></i></div></div>' +
              '<div style="font-size:12px;font-weight:700;color:var(--ink-2)">' + r.done + '/' + r.total + '</div></div>';
          }).join("") + '</div>'
        : '<div class="empty" style="padding:30px"><h3>Nobody mid-onboarding</h3><p>Open a provider or staff member and click <b>📋 Onboarding</b> to start tracking their steps.</p></div>';
      openModal("Onboarding board", body);
    }).catch(() => toast("Couldn't load the board."));
  }
  function isAdmin() { return CLOUD && window.SENTINEL_AUTH && window.SENTINEL_AUTH.admin && !READONLY; }
  function splitName(full) {
    const cleaned = String(full || "").replace(/[\*,()]+/g, "").trim();
    const parts = cleaned.split(/\s+/);
    if (parts.length < 2) return { first: "", last: parts[0] || "" };
    return { first: parts.slice(0, -1).join(" "), last: parts[parts.length - 1] };
  }
  function addProviderToRoster() {
    if (!isAdmin()) { toast("Admins only."); return; }
    openModal("Add provider to roster",
      '<div class="item-sub" style="margin-bottom:10px">Writes a new row to the <b>WCGTX Credentials</b> sheet of the master Excel. Their tracked items will appear after the next dashboard refresh.</div>' +
      '<div class="dl">First name <span class="req">*</span></div><input id="pAddF" style="width:100%;padding:9px;border-radius:8px;border:1px solid var(--hair);background:var(--surface-solid);color:var(--ink);margin-bottom:8px">' +
      '<div class="dl">Last name <span class="req">*</span></div><input id="pAddL" style="width:100%;padding:9px;border-radius:8px;border:1px solid var(--hair);background:var(--surface-solid);color:var(--ink);margin-bottom:8px">' +
      '<div class="dl">Email (optional)</div><input id="pAddE" style="width:100%;padding:9px;border-radius:8px;border:1px solid var(--hair);background:var(--surface-solid);color:var(--ink);margin-bottom:12px">' +
      '<div class="drawer-actions" style="border:none;padding:0"><button class="save" id="pAddGo">Add to roster</button></div>');
    $("#pAddGo").onclick = () => {
      const f = $("#pAddF").value.trim(), l = $("#pAddL").value.trim(), em = $("#pAddE").value.trim();
      $("#pAddF").classList.toggle("err", !f); $("#pAddL").classList.toggle("err", !l);
      if (!f || !l) { toast("First and last name are required."); return; }
      // Disable while the request is in flight. The server round trip is slow (workbook
      // download + backup + append + upload, with lock retries, then 7 folder creates), so an
      // impatient second click sent a second add: both passed the duplicate check, both
      // rewrote the whole workbook, and last-write-wins silently dropped one of them.
      const btn = $("#pAddGo");
      btn.disabled = true; btn.textContent = "Adding…";
      toast("Writing to the master roster…");
      addProviderRequest(f, l, em, false);
    };
  }
  function isLockError(err) { return /ROSTER_LOCKED|resourceLocked|is locked|\b423\b/i.test(String(err || "")); }

  // ---- System check: verifies the Excel + OneDrive round trip really works. ----------------
  // "Check only" is read-only. "Full test" adds a clearly-marked test provider, confirms the
  // Excel row and the OneDrive folder + 6 subfolders were created, confirms it shows on the
  // dashboard, then deletes all of it and verifies the cleanup. Backs up the workbook first.
  function openSelfTest() {
    if (!isAdmin()) { toast("Admins only."); return; }
    const html =
      '<div style="font-size:13px;line-height:1.65;color:var(--ink-2,#475569)">' +
      'Verifies the parts that touch your real files: reading the master Excel, writing to it, and creating/deleting OneDrive folders.' +
      '<ul style="margin:10px 0 0 18px;padding:0">' +
      '<li><b>Check only</b> — reads and reports. Changes nothing.</li>' +
      '<li><b>Full test</b> — also adds a test provider named <b>Sentinel ZZSelftest</b>, verifies it appeared in Excel <i>and</i> OneDrive, then removes it and verifies it is gone. The workbook is backed up first.</li>' +
      '</ul>' +
      '<div style="margin-top:10px"><b>Close the master Excel file first</b> — if it is open, writing to it is blocked.</div></div>' +
      '<div style="display:flex;gap:8px;margin-top:14px;flex-wrap:wrap">' +
      '<button class="btn-primary" id="stDiag" style="flex:1;min-width:130px">Check only</button>' +
      '<button class="btn-primary" id="stFull" style="flex:1;min-width:130px;background:#b45309">Full test</button>' +
      '</div>' +
      // Trace one name through Excel -> baked data -> cache -> dashboard, so "I can't see Dr X"
      // gets a specific answer instead of guesswork.
      '<div style="margin-top:16px;padding-top:14px;border-top:1px solid var(--hair,#e2e8f0)">' +
      '<div style="font-weight:700;font-size:13px;margin-bottom:6px">Find a provider</div>' +
      '<div class="item-sub" style="margin-bottom:8px">Type a surname to see whether they are in the master Excel, and why they may not be showing.</div>' +
      '<div style="display:flex;gap:8px"><input id="stWho" placeholder="e.g. Crespo" style="flex:1;padding:8px 10px;border:1px solid var(--hair,#e2e8f0);border-radius:8px;font-size:13px">' +
      '<button class="btn-primary" id="stWhoGo" style="min-width:96px">Look up</button></div></div>' +
      '<div id="stOut" style="margin-top:14px"></div>';
    // openModal takes an HTML string and returns the live #modalInner node — query from that,
    // not from a detached element, or the buttons below would never be wired up.
    const body = openModal("System check", html);
    const out = body.querySelector("#stOut");
    const run = (mode) => {
      body.querySelector("#stDiag").disabled = true;
      body.querySelector("#stFull").disabled = true;
      out.innerHTML = '<div style="font-size:13px;color:var(--ink-2,#475569)">Running ' +
        (mode === "full" ? "the full round trip — this reads and writes several files, give it up to a minute…" : "checks…") + '</div>';
      fetch("/api/data?selftest=" + mode, { method: "POST" }).then(r => r.json()).then(d => {
        body.querySelector("#stDiag").disabled = false;
        body.querySelector("#stFull").disabled = false;
        if (d.error) { out.innerHTML = '<div style="color:#b91c1c;font-size:13px">' + esc(d.error) + '</div>'; return; }
        const rows = (d.steps || []).map(s =>
          '<div style="display:flex;gap:8px;padding:7px 0;border-bottom:1px solid var(--line,#e2e8f0);align-items:flex-start">' +
          '<div style="flex:0 0 18px;color:' + (s.pass ? "#059669" : "#b91c1c") + ';font-weight:700">' + (s.pass ? "✓" : "✗") + '</div>' +
          '<div style="flex:1"><div style="font-weight:600;font-size:13px">' + esc(s.name) + '</div>' +
          (s.detail ? '<div style="font-size:12px;color:var(--ink-2,#475569);margin-top:2px;word-break:break-word">' + esc(s.detail) + '</div>' : '') +
          '</div></div>').join("");
        const allPass = d.passed === d.total;
        out.innerHTML =
          '<div style="padding:10px 12px;border-radius:8px;margin-bottom:10px;background:' + (allPass ? "#ecfdf5" : "#fef2f2") + ';color:' + (allPass ? "#065f46" : "#991b1b") + ';font-weight:600;font-size:13px">' +
          (allPass ? "✓ All " + d.total + " checks passed" : "⚠ " + d.passed + " of " + d.total + " passed — see below") +
          '<div style="font-weight:400;font-size:12px;margin-top:3px">' + esc(d.mode || "") + (d.cleanup ? " · cleanup: " + esc(d.cleanup) : "") + (d.backup ? " · backup: " + esc(d.backup) : "") + '</div></div>' +
          (d.rosterShifted ? '<div style="padding:8px 12px;border-radius:8px;margin-bottom:10px;background:#fffbeb;color:#92400e;font-size:12px">Note: your roster has an extra column before the names. The app now finds the columns by their headings, so this is handled — but if someone deletes that column it will keep working too.</div>' : '') +
          rows + (d.hint ? '<div style="font-size:12px;color:#92400e;margin-top:10px">' + esc(d.hint) + '</div>' : '');
      }).catch(e => {
        body.querySelector("#stDiag").disabled = false;
        body.querySelector("#stFull").disabled = false;
        out.innerHTML = '<div style="color:#b91c1c;font-size:13px">Could not run: ' + esc(String(e.message || e)) + '</div>';
      });
    };
    const lookup = () => {
      const q = (body.querySelector("#stWho").value || "").trim();
      if (q.length < 2) { toast("Type at least 2 letters."); return; }
      out.innerHTML = '<div class="item-sub">Looking up “' + esc(q) + '” in the master Excel…</div>';
      fetch("/api/data?whois=" + encodeURIComponent(q)).then(r => r.json()).then(d => {
        if (d.error) { out.innerHTML = '<div style="color:#b91c1c;font-size:13px">' + esc(d.error) + '</div>'; return; }
        const list = (arr, label) => arr && arr.length
          ? '<div class="pf-sub"><b>' + label + ':</b> ' + arr.map(x => esc((x.first ? x.first + " " : "") + x.last + " (row " + x.row + ")")).join(", ") + '</div>'
          : '<div class="pf-sub pf-dim">' + label + ': none</div>';
        out.innerHTML =
          '<div style="padding:11px 13px;border-radius:9px;background:var(--bg,#f1f5f9);font-size:13px;font-weight:650;margin-bottom:10px">' + esc(d.verdict) + '</div>' +
          '<div style="font-size:12px;line-height:1.7">' +
          list(d.liveRoster && d.liveRoster.credentialsSheet, "In “WCGTX Credentials” (active)") +
          list(d.liveRoster && d.liveRoster.inactiveSheet, "In “Inactive Providers”") +
          '<div class="pf-sub"><b>Visible on the dashboard:</b> ' + ((d.visibleOnDashboard || []).length ? esc(d.visibleOnDashboard.join(", ")) : "no") + '</div>' +
          '<div class="pf-sub pf-dim">Workbook totals: ' + ((d.counts && d.counts.liveActive) || 0) + ' active, ' + ((d.counts && d.counts.liveInactive) || 0) + ' inactive · cache built ' + esc((d.inCachedDelta && d.inCachedDelta.deltaGeneratedAt) || "never") + '</div>' +
          '</div>';
      }).catch(e => { out.innerHTML = '<div style="color:#b91c1c;font-size:13px">Lookup failed: ' + esc(String(e.message || e)) + '</div>'; });
    };
    body.querySelector("#stWhoGo").onclick = lookup;
    body.querySelector("#stWho").onkeydown = (e) => { if (e.key === "Enter") lookup(); };
    body.querySelector("#stDiag").onclick = () => run("diagnose");
    body.querySelector("#stFull").onclick = () => {
      if (!confirm("Full test will add a test provider called \"Sentinel ZZSelftest\", verify it in Excel and OneDrive, then delete it.\n\nThe workbook is backed up first. Make sure the master Excel file is closed.\n\nRun it?")) return;
      run("full");
    };
  }
  // Pull the latest roster + expiry dates straight from the master Excel on demand.
  function syncFromExcel() {
    if (!isAdmin()) { toast("Admins only."); return; }
    toast("Syncing from the master Excel…");
    // force=1 bypasses the server's 60s regen throttle — a manual click must always do real
    // work, even if an automatic on-open sync just ran a few seconds ago.
    fetch("/api/data?regen=1&force=1", { method: "POST" }).then(r => r.json()).then(d => {
      if (d.error) { toast(isLockError(d.error) ? "⚠ The master Excel is open right now — close it and try again." : ("Sync failed: " + d.error)); return; }
      return fetch("/api/data").then(r => r.json()).then(j => {
        if (j && j.items) {
          window.SENTINEL_SEED = j; buildData(); render();
          var dd = d.delta || {};
          toast("✓ Synced from Excel — " + (dd.newProviders || 0) + " new provider(s), " + (dd.dates || 0) + " dates read.");
        }
      });
    }).catch(e => toast("Sync failed: " + (e.message || e)));
  }
  function addProviderRequest(f, l, em, force) {
    // Put the Add button back if we stop without closing the modal, so a failed attempt can be
    // retried (it is disabled on click to prevent a double-submit rewriting the workbook twice).
    const reEnable = () => { const b = $("#pAddGo"); if (b) { b.disabled = false; b.textContent = "Add to roster"; } };
    fetch("/api/data?roster=add", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ first: f, last: l, email: em || undefined, force: !!force }) })
        // A 504/500 with a non-JSON body would otherwise reject with "Unexpected token '<'",
        // which reads like a client bug and hides that the Excel write may well have committed.
        .then(r => r.json().catch(() => ({ ok: false, error: "The server did not respond properly (HTTP " + r.status + "). The provider MAY have been added — click 'Sync from Excel' and check before retrying." })))
        .then(d => {
          if (!d.ok) {
            reEnable();
            if (isLockError(d.error)) { toast("⚠ The master Excel is open right now — close WCGTX Physician Roster.xlsx in Excel (desktop + any browser tab), wait a few seconds, then try again."); return; }
            // 409: name already present in the active sheet. Offer to add anyway (true duplicate name).
            if (/already present/i.test(d.error || "")) {
              if (confirm('"' + f + " " + l + '" is already in the active roster.\n\nAdd a second row with the same name anyway? (Use this only for true duplicate-named providers — the dashboard will treat them as one entity.)')) {
                addProviderRequest(f, l, em, true);
                return;
              }
              toast("Add cancelled."); return;
            }
            toast("Add failed: " + (d.error || "unknown")); return;
          }
          const folderMsg = d.folder && d.folder.ok ? " folder ✓" : (d.folder ? " folder ✗ " + (d.folder.error || "") : "");
          closeModal(); toast("✓ Added to roster." + folderMsg + " Refreshing dashboard…");
          // Trigger an instant regen so the new provider appears now, not at next cron run.
          return fetch("/api/data?regen=1&force=1", { method: "POST" }).catch(()=>{}).then(()=>{
            // Re-fetch /api/data and rebuild
            return fetch("/api/data").then(r=>r.json()).then(d=>{ if(d&&d.items){window.SENTINEL_SEED=d; buildData(); render(); toast("✓ "+f+" "+l+" added — dashboard updated."); }});
          });
        })
        .catch(e => { reEnable(); toast("Add failed: " + (e.message || e)); });
  }
  function deleteProviderHard(name, entityKey) {
    if (!isAdmin()) { toast("Admins only."); return; }
    const { first, last } = splitName(name);
    if (!confirm('PERMANENTLY DELETE "' + name + '" from the roster?\n\nThis removes them from BOTH WCGTX Credentials and Inactive Providers sheets in the master Excel. The row is logged to the Recycle bin and can be restored within ~200 deletes.\n\nUse "📦 Mark inactive" if you just want to deactivate them.')) return;
    toast("Deleting from the master roster…");
    fetch("/api/data?roster=delete", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ first, last, entityKey: entityKey || "" }) })
      .then(r => r.json())
      .then(d => {
        if (!d.ok) {
          if (isLockError(d.error)) { toast("⚠ The master Excel is open right now — close WCGTX Physician Roster.xlsx in Excel (desktop + any browser tab), wait a few seconds, then try again."); return; }
          let msg = "Delete failed: " + (d.error || "unknown"); if (d.tried) msg += '\nLooked up: last="' + d.tried.last + '", first="' + d.tried.first + '", entityKey="' + d.tried.entityKey + '"'; toast(msg); return;
        }
        console.log("[delete provider]", d);
        const snapMsg = d.snapshot && d.snapshot.ok ? (" backup ✓") : (d.snapshot ? (" backup ✗ " + (d.snapshot.error || "")) : "");
        const folderMsg = d.folder && d.folder.ok ? (d.folder.note ? "" : " folder ✓") : (d.folder ? (" folder ✗ " + (d.folder.status || "") + " " + (d.folder.error || "")) : "");
        var vmsg = (d.trashVerified === undefined || d.trashVerified === null) ? "" : (", verified " + d.trashVerified);
        if (d.warning) toast("⚠ Deleted but " + d.warning + snapMsg + folderMsg);
        else toast("✓ Deleted (" + d.removedRows + " row, trash " + (d.trashEntries || 0) + vmsg + ")" + snapMsg + folderMsg);
        return fetch("/api/data?regen=1&force=1", { method: "POST" }).catch(()=>{}).then(()=>{
          return fetch("/api/data").then(r=>r.json()).then(d=>{ if(d&&d.items){window.SENTINEL_SEED=d; buildData(); navigate([]); render(); toast("✓ " + name + " permanently deleted. Recoverable from Recycle bin."); }});
        });
      })
      .catch(e => toast("Delete failed: " + (e.message || e)));
  }
  function openRecycleBin() {
    if (!isAdmin()) { toast("Admins only."); return; }
    toast("Loading recycle bin…");
    fetch("/api/data?roster=trash&_t=" + Date.now(), { cache: "no-store" }).then(r => r.json()).then(t => {
      console.log("[recycle bin] raw response:", t);
      const entries = (t && t.entries) || [];
      const errLine = (t && t.error) ? '<div class="empty" style="padding:30px"><h3>Couldn\'t load recycle bin</h3><p>Server said: ' + esc(t.error) + '</p></div>' : null;
      const body = entries.length
        ? '<div class="item-sub" style="margin-bottom:12px">Deleted providers. The most recent 200 are listed here. <b>Restore</b> puts each row back in the sheet it was deleted from — so someone deleted while inactive comes back inactive (tick “Show inactive” to see them). Note: their OneDrive folder is not recreated; recover that from the SharePoint recycle bin.</div>' +
          '<div style="display:flex;flex-direction:column;gap:8px">' +
          entries.map(e => '<div style="display:flex;align-items:center;gap:10px;padding:10px 12px;border:1px solid var(--hair);border-radius:10px;background:var(--surface-solid)">' +
            '<div style="flex:1"><div style="font-weight:700">' + esc(e.entity || "(unnamed)") + '</div>' +
            '<div class="item-sub" style="font-size:12px">deleted ' + esc(new Date(e.deletedAt).toLocaleString()) + ' by ' + esc(e.deletedBy || "—") + ' · ' + (e.rows ? e.rows.length : 0) + ' row(s)</div></div>' +
            '<button class="icon-btn" data-rid="' + esc(e.id) + '">↺ Restore</button></div>').join("") +
          '</div>'
        : (errLine || '<div class="empty" style="padding:30px"><h3>Recycle bin is empty</h3><p>No providers have been hard-deleted yet. Deleted providers show up here with a Restore button.</p></div>');
      openModal("Recycle bin (deleted providers)", body);
      [...$("#modalInner").querySelectorAll("[data-rid]")].forEach(b => b.onclick = () => {
        const id = b.dataset.rid;
        if (!confirm("Restore this provider to the active roster?")) return;
        b.disabled = true; b.textContent = "Restoring…";
        fetch("/api/data?roster=restore", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ id }) })
          .then(r => r.json()).then(d => {
            if (!d.ok) { toast("Restore failed: " + (d.error || "unknown")); b.disabled = false; b.textContent = "↺ Restore"; return; }
            closeModal(); toast("✓ Restored " + d.entity + ". Refreshing…");
            return fetch("/api/data?regen=1&force=1", { method: "POST" }).catch(()=>{}).then(()=>{
              return fetch("/api/data").then(r=>r.json()).then(d=>{ if(d&&d.items){window.SENTINEL_SEED=d; buildData(); render(); }});
            });
          }).catch(e => toast("Restore failed: " + (e.message || e)));
      });
    }).catch(() => toast("Couldn't load recycle bin."));
  }
  // ---- Facility folder management (admin): add / delete / restore folders in a facility's
  //      SharePoint State Readiness tree. Mirrors the provider Add + Recycle bin experience. ----
  var FAC_INPUT = 'width:100%;padding:9px;border-radius:8px;border:1px solid var(--hair);background:var(--surface-solid);color:var(--ink)';
  function openFacilityFolders(facility) {
    if (!isAdmin()) { toast("Admins only."); return; }
    function paint(folders) {
      var list = folders.length
        ? folders.map(function (n) {
            return '<div style="display:flex;align-items:center;gap:10px;padding:9px 12px;border:1px solid var(--hair);border-radius:10px;background:var(--surface-solid)">' +
              '<div style="flex:1;font-weight:600">📁 ' + esc(n) + '</div>' +
              '<button class="icon-btn danger" data-del="' + esc(n) + '">🗑 Delete</button></div>';
          }).join("")
        : '<div class="item-sub">No folders here yet.</div>';
      openModal("Add / manage folders — " + esc(facility),
        '<div class="item-sub" style="margin-bottom:10px">Folders live in SharePoint under <b>State Readiness › ' + esc(facility) + '</b>. A new folder shows on the dashboard within ~45s. Deleting moves it to the Recycle bin (its files are kept and it can be restored).</div>' +
        '<div style="display:flex;gap:8px;margin-bottom:14px"><input id="facNew" placeholder="New folder name" style="' + FAC_INPUT + ';flex:1"><button class="save" id="facAdd" style="white-space:nowrap">＋ Add folder</button></div>' +
        '<div style="display:flex;flex-direction:column;gap:8px">' + list + '</div>');
      $("#facAdd").onclick = function () {
        var name = ($("#facNew").value || "").trim();
        if (!name) { toast("Enter a folder name."); return; }
        $("#facAdd").disabled = true;
        fetch("/api/data?facility=add", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ facility: facility, name: name }) })
          .then(function (r) { return r.json(); }).then(function (d) {
            if (!d.ok) { toast("Add failed: " + (d.error || "unknown")); $("#facAdd").disabled = false; return; }
            toast("✓ Folder created — showing on the dashboard within ~45s."); fetch("/api/scan").catch(function () {}); load();
          }).catch(function (e) { toast("Add failed: " + (e.message || e)); $("#facAdd").disabled = false; });
      };
      [].forEach.call($("#modalInner").querySelectorAll("[data-del]"), function (btn) {
        btn.onclick = function () {
          var name = btn.dataset.del;
          if (!confirm('Delete folder "' + name + '"?\n\nIt moves to the Recycle bin and can be restored. Its files are kept.')) return;
          btn.disabled = true; btn.textContent = "Deleting…";
          fetch("/api/data?facility=delete", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ facility: facility, name: name }) })
            .then(function (r) { return r.json(); }).then(function (d) {
              if (!d.ok) { toast("Delete failed: " + (d.error || "unknown")); btn.disabled = false; btn.textContent = "🗑 Delete"; return; }
              toast("✓ Deleted (recoverable from Recycle bin)."); fetch("/api/scan").catch(function () {}); load();
            }).catch(function (e) { toast("Delete failed: " + (e.message || e)); btn.disabled = false; btn.textContent = "🗑 Delete"; });
        };
      });
    }
    function load() {
      toast("Loading folders…");
      fetch("/api/data?facility=list&_t=" + Date.now(), { method: "POST", cache: "no-store", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ facility: facility }) })
        .then(function (r) { return r.json(); }).then(function (d) { if (d.folders) paint(d.folders); else { toast("Couldn't load: " + (d.error || "")); paint([]); } })
        .catch(function () { toast("Couldn't load folders."); });
    }
    load();
  }
  function openFacilityRecycleBin(facility) {
    if (!isAdmin()) { toast("Admins only."); return; }
    toast("Loading recycle bin…");
    fetch("/api/data?facility=trash&_t=" + Date.now(), { cache: "no-store" }).then(function (r) { return r.json(); }).then(function (t) {
      var entries = ((t && t.entries) || []).filter(function (e) { return !facility || e.facility === facility; });
      var body = entries.length
        ? '<div class="item-sub" style="margin-bottom:12px">Deleted folders for <b>' + esc(facility) + '</b>. Restore puts a folder (and its files) back.</div>' +
          '<div style="display:flex;flex-direction:column;gap:8px">' +
          entries.map(function (e) {
            return '<div style="display:flex;align-items:center;gap:10px;padding:10px 12px;border:1px solid var(--hair);border-radius:10px;background:var(--surface-solid)">' +
              '<div style="flex:1"><div style="font-weight:700">📁 ' + esc(e.name) + '</div>' +
              '<div class="item-sub" style="font-size:12px">deleted ' + esc(new Date(e.deletedAt).toLocaleString()) + ' by ' + esc(e.deletedBy || "—") + '</div></div>' +
              '<button class="icon-btn" data-rid="' + esc(e.id) + '">↺ Restore</button></div>';
          }).join("") + '</div>'
        : '<div class="empty" style="padding:30px"><h3>Recycle bin is empty</h3><p>No deleted folders for ' + esc(facility) + '.</p></div>';
      openModal("Recycle bin — " + esc(facility), body);
      [].forEach.call($("#modalInner").querySelectorAll("[data-rid]"), function (b) {
        b.onclick = function () {
          var id = b.dataset.rid;
          if (!confirm("Restore this folder?")) return;
          b.disabled = true; b.textContent = "Restoring…";
          fetch("/api/data?facility=restore", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ id: id, facility: facility }) })
            .then(function (r) { return r.json(); }).then(function (d) {
              if (!d.ok) { toast("Restore failed: " + (d.error || "unknown")); b.disabled = false; b.textContent = "↺ Restore"; return; }
              closeModal(); toast("✓ Restored " + d.name + " — showing within ~45s."); fetch("/api/scan").catch(function () {});
            }).catch(function (e) { toast("Restore failed: " + (e.message || e)); b.disabled = false; b.textContent = "↺ Restore"; });
        };
      });
    }).catch(function () { toast("Couldn't load recycle bin."); });
  }
  function removeProviderFromRoster(name, entityKey) {
    if (!isAdmin()) { toast("Admins only."); return; }
    const { first, last } = splitName(name);
    if (!confirm('Move "' + name + '" to Inactive Providers in the master Excel?\n\nTheir records stay in the dashboard but the provider is marked inactive going forward. You can re-activate later by editing the Excel directly.')) return;
    toast("Updating the master roster…");
    // Pass entityKey so the server can match the row regardless of how the display name splits.
    fetch("/api/data?roster=remove", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ first, last, entityKey: entityKey || "" }) })
      .then(r => r.json())
      .then(d => {
        if (!d.ok) {
          if (isLockError(d.error)) { toast("⚠ The master Excel is open right now — close WCGTX Physician Roster.xlsx in Excel (desktop + any browser tab), wait a few seconds, then try again."); return; }
          let msg = "Remove failed: " + (d.error || "unknown");
          if (d.tried) msg += '\nLooked up: last="' + d.tried.last + '", first="' + d.tried.first + '"';
          if (d.sample && d.sample.length) msg += "\nSample rows from sheet: " + d.sample.slice(0, 3).map(r => r.last + ", " + r.first).join(" | ");
          console.warn("[delete provider]", d);
          toast(msg);
          return;
        }
        toast("✓ Moved to inactive. Refreshing dashboard…");
        return fetch("/api/data?regen=1&force=1", { method: "POST" }).catch(()=>{}).then(()=>{
          return fetch("/api/data").then(r=>r.json()).then(d=>{ if(d&&d.items){window.SENTINEL_SEED=d; buildData(); navigate([]); render(); toast("✓ "+name+" marked inactive."); }});
        });
      })
      .catch(e => toast("Remove failed: " + (e.message || e)));
  }
  function renderHierarchy(c, arr, tab) {
    if (!state.drill || state._drillTab !== tab) { state.drill = []; state._drillTab = tab; }
    const sg = subgroupsFor(tab);
    const tabLabel = tab === "provider" ? "Providers" : tab === "staff" ? "Staff" : tab === "facility" ? "Facilities" : "Operations";
    const groups = {};
    arr.forEach(i => (groups[i.entity] = groups[i.entity] || []).push(i));

    // flow-chart breadcrumb + back
    const trail = [tabLabel].concat(state.drill);
    const flow = el("div", "flow");
    if (state.drill.length) {
      const back = el("button", "flow-back", "‹ Back");
      back.onclick = () => navigate(state.drill.slice(0, -1));
      flow.appendChild(back);
    }
    trail.forEach((label, i) => {
      if (i) flow.insertAdjacentHTML("beforeend", '<span class="flow-arrow">→</span>');
      const node = el("button", "flow-node" + (i === trail.length - 1 ? " active" : ""), esc(label));
      if (i < trail.length - 1) node.onclick = () => navigate(state.drill.slice(0, i));
      flow.appendChild(node);
    });
    c.appendChild(flow);

    const grid = el("div", "hgrid hgrid-wide" + (state._flip ? " flip-in" : ""));
    state._flip = false;

    const searching = (state.search || "").trim().length > 0;
    if (searching) {
      // Search flattens to matching DOCUMENTS within the current scope (whole tab, or the
      // drilled-in entity) so you can find a file without knowing its folder.
      const scoped = state.drill.length ? arr.filter(i => i.entity === state.drill[0]) : arr;
      if (state.drill.length) c.appendChild(entityHeader(state.drill[0], groups[state.drill[0]] || [], tab));
      const sh = el("div", "srch-head");
      sh.innerHTML = '<b>' + scoped.length + '</b> result' + (scoped.length === 1 ? "" : "s") + ' for “' + esc(state.search) + '” in ' + esc(state.drill.length ? state.drill[0] : tabLabel);
      c.appendChild(sh);
      sortDocs(scoped).forEach(it => grid.appendChild(docTile(it, !state.drill.length)));
      if (!scoped.length) grid.innerHTML = '<div class="hempty">No matches for “' + esc(state.search) + '”. Try fewer words.</div>';
    } else if (state.drill.length === 0) {
      grid.classList.add("hgrid-wide");
      if (tab === "provider" && isAdmin()) {
        const add = el("div", "tile tile-add");
        add.innerHTML = '<div class="add-plus">＋</div><div class="tile-nm">Add provider</div><div class="tile-meta">Writes a new row to the master Excel roster</div>';
        add.onclick = addProviderToRoster;
        grid.appendChild(add);
        const bin = el("div", "tile tile-add tile-bin");
        bin.innerHTML = '<div class="add-plus">🗑</div><div class="tile-nm">Recycle bin</div><div class="tile-meta">Restore providers you\'ve deleted</div>';
        bin.onclick = openRecycleBin;
        grid.appendChild(bin);
      }
      sortedEntityNames(groups).forEach((name, i) => grid.appendChild(entityTile(name, groups[name], tab, i + 1)));
      if (!grid.children.length) grid.innerHTML = '<div class="hempty">No matching items.</div>';
    } else {
      const entity = state.drill[0];
      const items = groups[entity] || [];
      c.appendChild(entityHeader(entity, items, tab));
      if (state.drill.length === 1 && sg) {
        const buckets = sg.order.map(() => []); const extra = [];
        items.forEach(it => {
          // Honor phaseIdx (set by the gap-fill scanner) so supplemental files bucket under
          // their folder section instead of "Other".
          const idx = (typeof it.phaseIdx === "number") ? it.phaseIdx : sg.of[it.category];
          if (idx == null || idx >= sg.order.length) extra.push(it); else buckets[idx].push(it);
        });
        sg.order.forEach(([label], idx) => grid.appendChild(sectionTile(label, buckets[idx], entity)));
        if (extra.length) grid.appendChild(sectionTile(sg.other, extra, entity));
      } else {
        let docs = items;
        if (sg && state.drill.length === 2) {
          const secLabel = state.drill[1];
          docs = items.filter(it => {
            // Use phaseIdx (gap-fill supplementals) first, else fall back to category lookup.
            const idx = (typeof it.phaseIdx === "number") ? it.phaseIdx : sg.of[it.category];
            return (idx == null || idx >= sg.order.length ? sg.other : sg.order[idx][0]) === secLabel;
          });
        }
        sortDocs(docs).forEach(it => grid.appendChild(docTile(it)));
        if (!docs.length) grid.innerHTML = '<div class="hempty">No documents here yet.</div>';
      }
    }
    // Facility/Other drill: admin tiles to add folders + recycle bin (OneDrive-synced), scoped
    // to the drilled-in facility. Prepended so they lead the grid.
    if (isAdmin() && state.drill.length === 1 && (tab === "facility" || tab === "other") && !(state.search || "").trim()) {
      const fent = state.drill[0];
      const fbin = el("div", "tile tile-add tile-bin");
      fbin.innerHTML = '<div class="add-plus">🗑</div><div class="tile-nm">Recycle bin</div><div class="tile-meta">Restore deleted folders</div>';
      fbin.onclick = () => openFacilityRecycleBin(fent);
      grid.insertBefore(fbin, grid.firstChild);
      const fmf = el("div", "tile tile-add");
      fmf.innerHTML = '<div class="add-plus">📁</div><div class="tile-nm">Add / manage folders</div><div class="tile-meta">Create or delete folders in OneDrive</div>';
      fmf.onclick = () => openFacilityFolders(fent);
      grid.insertBefore(fmf, grid.firstChild);
    }
    c.appendChild(grid);

    if (!window.__hescBound) {
      window.__hescBound = true;
      document.addEventListener("keydown", (e) => {
        if (e.key !== "Escape") return;
        const dr = $("#drawer"); if (dr && dr.getAttribute("aria-hidden") === "false") return;
        const md = $("#modal"); if (md && getComputedStyle(md).display !== "none") return;
        if (state.drill && state.drill.length) { e.preventDefault(); navigate(state.drill.slice(0, -1)); }
      });
    }
  }

  function itemRow(it) {
    const st = computeStatus(it);
    const row = el("div", "item-row" + (state.selection.has(it.id) ? " selected" : ""));
    const sub = [it.authority, it.number ? "#" + it.number : ""].filter(Boolean).join(" · ");
    const proofBadge = it.isFile
      ? '<span class="proof-badge has" title="Proof document attached">📄 Proof</span>'
      : (it.centralProof ? '<span class="proof-badge has" title="Covered by a carrier-level policy from the COI roster">🛡 Via roster</span>'
      : (it.isFile === false ? '<span class="proof-badge no" title="No proof document attached — click to attach">⚠ No proof</span>' : ""));
    const ver = OVERLAY.verified[it.id];
    const verBadge = ver ? '<span class="proof-badge has" title="Reviewed by ' + esc(ver.by || "") + ' on ' + esc(ver.at || "") + '" style="background:#dcfce7;color:#166534;border-color:#bbf7d0">✅ Verified</span>' : "";
    row.innerHTML =
      (state.selectMode ? '<input type="checkbox" class="row-check"' + (state.selection.has(it.id) ? " checked" : "") + '>' : "") +
      '<div class="ico s-' + st.key + '"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' + ICONS[iconFor(it)] + '</svg></div>' +
      '<div class="main"><div class="item-cat">' + esc(it.category) + '</div><div class="item-sub">' + esc(sub || it.notes || "") + '</div></div>' +
      verBadge + proofBadge +
      '<div class="item-when">' + (it.expires ? fmtD(it.expires) : (it.permanent ? "No expiry" : "—")) + (st.startBy && st.key !== "good" && st.key !== "permanent" ? '<small>renew by ' + fmtD(st.startBy.toISOString().slice(0, 10)) + '</small>' : "") + '</div>' +
      '<div class="countdown s-' + st.key + '" style="background:none;border:none;padding:0">' + countdownText(st) + '</div>' +
      '<div class="status-badge s-' + st.key + '">' + st.label + '</div>' +
      (it.scope !== "provider" ? '<button class="star irowmail" title="Email me this item" style="font-size:15px">✉</button>' : "") +
      '<button class="star' + (isWatched(it.id) ? " on" : "") + '" title="Watch this item">' + (isWatched(it.id) ? "★" : "☆") + '</button>';
    const star = row.querySelector(".star:last-child");
    star.onclick = (e) => { e.stopPropagation(); toggleWatch(it.id); render(); };
    const irm = row.querySelector(".irowmail");
    if (irm) irm.onclick = (e) => { e.stopPropagation(); emailGroupToSelf(it.entity + " — " + it.category, [it]); };
    const cb = row.querySelector(".row-check");
    if (cb) cb.onclick = (e) => { e.stopPropagation(); if (state.selection.has(it.id)) state.selection.delete(it.id); else state.selection.add(it.id); renderContent(); };
    row.onclick = () => { if (state.selectMode) { if (state.selection.has(it.id)) state.selection.delete(it.id); else state.selection.add(it.id); renderContent(); } else openDrawer(it, false); };
    return row;
  }

  function renderTimeline(arr) {
    const wrap = el("div", "timeline");
    // Require a PARSEABLE date, not just a truthy one — a single unparseable expiry (an old
    // hand-typed value like "pending") used to throw inside the loop below and blank the whole view.
    const dated = arr.filter(i => i.expires && parseD(i.expires)).sort((a, b) => parseD(a.expires) - parseD(b.expires));
    if (!dated.length) { wrap.innerHTML = '<div class="empty">No dated items to plot.</div>'; return wrap; }
    let curMonth = "";
    dated.forEach(it => {
      const d = parseD(it.expires);
      const mk = d.toLocaleDateString("en-US", { month: "long", year: "numeric" });
      if (mk !== curMonth) { curMonth = mk; wrap.appendChild(el("div", "tl-month", mk)); }
      const st = computeStatus(it);
      const item = el("div", "tl-item");
      item.innerHTML = '<div class="tl-date s-' + st.key + '" style="background:none;border:none">' + d.getDate() + '<small>' + d.toLocaleDateString("en-US", { month: "short" }) + '</small></div>' +
        '<div class="ico s-' + st.key + '"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' + ICONS[iconFor(it)] + '</svg></div>' +
        '<div style="flex:1"><div class="item-cat">' + esc(it.category) + '</div><div class="item-sub">' + esc(it.entity) + '</div></div>' +
        '<div class="status-badge s-' + st.key + '">' + countdownText(st) + '</div>';
      item.onclick = () => openDrawer(it, false);
      wrap.appendChild(item);
    });
    return wrap;
  }

  // ================= DRAWER =================
  let drawerItem = null;
  function encodePath(p) { return p.split("/").map(s => s === ".." || s === "." ? s : encodeURIComponent(s)).join("/"); }
  // Turn a matched local-disk proof path ("../..WCGTX Master Physician File/.../x.pdf")
  // into its SharePoint web link, so it opens from any computer in the cloud.
  function sharePointLink(localRaw) {
    const HOST = "https://wcgtx-my.sharepoint.com", PERS = "/personal/sfarooqui_wcgtx_com";
    const DOC = PERS + "/Documents/WCGTX Phyicians_04.08.2020/";
    const rel = String(localRaw).replace(/^(\.\.\/)+/, "");
    const server = DOC + rel;
    const parent = server.slice(0, server.lastIndexOf("/"));
    return HOST + PERS + "/_layouts/15/onedrive.aspx?id=" + encodeURIComponent(server) + "&parent=" + encodeURIComponent(parent);
  }

  function openDrawer(it, isNew) {
    drawerItem = it;
    const d = $("#drawer"); d.classList.add("show"); $("#drawerBack").classList.add("show");
    if (isNew) return renderDrawerEdit({ scope: state.tab, entity: "", category: "", authority: "", number: "", expires: "", issued: "", renewalLeadDays: 60, notes: "", active: true, id: "new-" + Date.now() }, true);
    renderDrawerView(it);
  }
  function closeDrawer() { $("#drawer").classList.remove("show"); $("#drawerBack").classList.remove("show"); drawerItem = null; }

  // ===== Free in-browser OCR (Tesseract.js) — read an expiry date off a scanned document =====
  function ensureTesseract() {
    return new Promise((resolve, reject) => {
      if (window.Tesseract) return resolve();
      const sc = document.createElement("script");
      sc.src = "https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js";
      sc.onload = resolve; sc.onerror = () => reject(new Error("Could not load the OCR engine."));
      document.head.appendChild(sc);
    });
  }
  function extractDates(text) {
    const out = [];
    const push = (y, mo, d) => { if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) out.push(y + "-" + String(mo).padStart(2, "0") + "-" + String(d).padStart(2, "0")); };
    let m;
    const re = /\b(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})\b/g;
    while ((m = re.exec(text))) { let y = +m[3]; if (y < 100) y += 2000; push(y, +m[1], +m[2]); }
    const mon = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
    const re2 = /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+(\d{1,2}),?\s+(\d{4})\b/gi;
    while ((m = re2.exec(text))) push(+m[3], mon[m[1].toLowerCase().slice(0, 3)], +m[2]);
    return [...new Set(out)].sort();
  }
  function ocrStat(msg) { const e = $("#ocrStatus"); if (e) e.innerHTML = msg; }
  function ocrApply(it, dates, auto) {
    if (!dates || !dates.length) { ocrStat("OCR ran but found no date in this document. Open “Edit” to set it."); return; }
    const today = new Date().toISOString().slice(0, 10);
    const fut = dates.filter(d => d >= today);
    const pick = fut.length ? fut[fut.length - 1] : dates[dates.length - 1];
    if (!it.expires) {
      if (auto) { ocrStat("✓ Auto-read expiry " + fmtD(pick)); setOcrExpiry(it, pick, true); }
      else pickOcrDate(it, dates);
    } else {
      const same = pick === it.expires;
      ocrStat("📄 Document reads <b>" + fmtD(pick) + "</b> · current expiry " + fmtD(it.expires) +
        (same ? " ✓ match" : ' — <button id="ocrUse" style="border:none;background:none;color:var(--accent);font-weight:800;cursor:pointer;text-decoration:underline;padding:0">use document date</button>'));
      const u = $("#ocrUse"); if (u) u.onclick = () => setOcrExpiry(it, pick, false);
    }
  }
  async function ocrReadDate(it, auto) {
    if (!it || !it.fileLink || !/^https?:/i.test(it.fileLink)) { ocrStat("No document file to read."); return; }
    if (!OVERLAY.ocr) OVERLAY.ocr = {};
    if (auto && OVERLAY.ocr[it.id]) { ocrApply(it, OVERLAY.ocr[it.id].dates, true); return; }  // cached — read once
    try {
      ocrStat("🔎 Reading the document… (OCR, a few seconds)");
      const r = await fetch("/api/data?ocr=" + encodeURIComponent(it.fileLink));
      const j = await r.json().catch(() => ({}));
      if (!r.ok || j.error) { ocrStat("⚠ " + esc(j.error ? (j.error + (j.detail ? " — " + j.detail : "")) : ("OCR failed " + r.status))); return; }
      OVERLAY.ocr[it.id] = { dates: j.dates || [], at: new Date().toISOString().slice(0, 10) }; saveOverlay();
      ocrApply(it, j.dates || [], auto);
    } catch (e) { ocrStat("⚠ OCR error: " + esc(String(e.message || e).slice(0, 90))); }
  }
  function pickOcrDate(it, dates) {
    const html = '<div class="item-sub" style="margin-bottom:10px">OCR read these dates from <b>' + esc(it.category) + '</b>. Pick the expiry:</div>' +
      dates.map(d => '<button class="doc-link ocrd" data-d="' + d + '" style="display:block;width:100%;text-align:left;margin-bottom:6px">📅 ' + fmtD(d) + '</button>').join("") +
      '<div class="item-sub" style="margin-top:8px;opacity:.65">Not the right one? Close this and use Edit to type it.</div>';
    openModal("Read expiry (OCR)", html);
    [...$("#modalInner").querySelectorAll(".ocrd")].forEach(b => b.onclick = () => { setOcrExpiry(it, b.dataset.d); closeModal(); });
  }
  function setOcrExpiry(it, iso, quiet) {
    OVERLAY.edits[it.id] = Object.assign({}, it, OVERLAY.edits[it.id] || {}, { expires: iso, permanent: false, pending: false });
    OVERLAY.logs[it.id] = (OVERLAY.logs[it.id] || []);
    OVERLAY.logs[it.id].push({ text: "Expiry read via OCR — " + iso, date: new Date().toISOString().slice(0, 10) });
    logAudit("edit", it, "expiry read via OCR: " + iso);
    saveOverlay(); buildData(); render();
    if ($("#drawer").classList.contains("show")) renderDrawerView(DATA.find(x => x.id === it.id) || Object.assign({}, it, { expires: iso }));
    toast((quiet ? "📅 Auto-read expiry " : "Expiry set to ") + fmtD(iso) + (quiet ? " (OCR)" : " — read from the document."));
  }

  function docSection(it) {
    let raw = it.fileLink || "";
    // Cloud: a MATCHED proof still pointing at a local-disk path → build its SharePoint
    // web link so it opens from anywhere (instead of the old dead-end "stored locally" note).
    if (CLOUD && it.isFile && raw.indexOf("../") === 0 && !/^https?:/i.test(raw)) {
      raw = sharePointLink(raw);
    }
    const isUrl = /^https?:/i.test(raw);
    const openable = isUrl || (!CLOUD && it.isFile);
    if (!openable) {
      const action = READONLY ? '' : (CLOUD
        ? '<div class="item-sub" style="margin-top:4px">Use the <b>QR code</b> button below to scan &amp; upload one — it’ll appear here.</div>'
        : '<button class="doc-link ghost" id="dAttach">Attach document</button>');
      return '<div class="dfield"><div class="dl">Proof document</div>' +
        '<div class="item-sub" style="margin-bottom:8px">⚠️ No proof document attached yet for this item.</div>' + action + '</div>';
    }
    const fileHref = isUrl ? raw : encodePath(raw);
    // "Open in Outlook" = open in the Microsoft 365 web (Office/OneDrive) viewer, where the user is signed in.
    // The document URL comes from /api/uploads-map, which the login-less QR flow can write, so it
    // is untrusted input: a value like  https://host/x?a=" onmouseover="…  would otherwise break
    // out of this href attribute and run script in the admin's authenticated session. Escape it,
    // and only ever emit http(s) or a local file path — never javascript:/data:.
    const safeHref = /^\s*(https?:|\/|[A-Za-z]:|\\)/.test(fileHref) ? fileHref : "#";
    const viewerHref = safeHref === "#" ? "#" : safeHref + (safeHref.indexOf("?") >= 0 ? "&" : "?") + "web=1";
    // Inline preview only works for local PDFs; SharePoint blocks cross-origin framing,
    // so on cloud we skip the (always-blank) iframe and open straight in the M365 viewer.
    const inlinePdf = !CLOUD && /\.pdf(\?|#|$)/i.test(raw);
    const fname = it.uploadName || (isUrl ? "" : decodeURIComponent(raw.split("/").pop()));
    const preview = inlinePdf
      ? '<div class="pdf-frame"><div class="pdf-fallback"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6">' + ICONS.doc + '</svg>If this preview stays blank, click “Open file in Outlook”.</div>' +
        '<iframe src="' + esc(safeHref) + '#view=FitH&toolbar=1" title="Proof document preview" loading="lazy"></iframe></div>'
      : '';
    return '<div class="dfield"><div class="dl">Proof document</div>' +
      (fname ? '<div class="pdf-name">📄 ' + esc(fname) + '</div>' : '') + preview +
      '<div class="doc-btns"><a class="doc-link" href="' + esc(viewerHref) + '" target="_blank" rel="noopener"><svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="5" width="18" height="14" rx="2"/><path d="M3 7l9 6 9-6"/></svg>Open file in Outlook</a>' +
      ((!CLOUD && !READONLY) ? '<button class="doc-link ghost" id="dReadDate">📅 Read date</button>' : '') +
      '</div>' +
      ((CLOUD && it.isFile && !READONLY) ? '<div id="ocrStatus" class="ocr-status">🔎 Checking the document…</div>' : '') +
      '</div>';
  }
  function wireAttach(it) {
    const b = $("#dAttach"); if (!b) return;
    if (READONLY) { b.style.display = "none"; return; }
    b.onclick = () => openAttachPicker(it);
  }

  function renderDrawerView(it) {
    const st = computeStatus(it);
    const logs = (OVERLAY.logs[it.id] || []);
    const task = OVERLAY.tasks[it.id] || {};
    const ver = OVERLAY.verified[it.id];
    const statusOpt = s => '<option' + ((task.status || "none") === s ? " selected" : "") + '>' + s + '</option>';
    const f = (l, v) => '<div class="dfield"><div class="dl">' + l + '</div><div class="dv">' + v + '</div></div>';
    $("#drawer").innerHTML =
      '<div class="drawer-head"><div><div class="status-badge s-' + st.key + '" style="margin-bottom:8px">' + st.label + (st.days != null ? ' · ' + countdownText(st) : '') + '</div>' +
      '<div class="dh-cat">' + esc(it.category) + '</div><div class="dh-ent">' + esc(it.entity) + '</div></div>' +
      '<button class="drawer-close" id="dClose">×</button></div>' +
      '<div class="drawer-body">' +
      (it.authority ? f("Issuing authority", esc(it.authority)) : "") +
      (it.number ? f("ID / License number", esc(it.number)) : "") +
      f("Expires", it.expires ? fmtD(it.expires) : (it.permanent ? "No expiry (permanent / on file)" : "—")) +
      (st.startBy && st.key !== "permanent" ? f("Start renewal by", fmtD(st.startBy.toISOString().slice(0, 10)) + ' <span class="item-sub">(lead ' + (it.renewalLeadDays || 60) + ' days)</span>') : "") +
      (it.issued ? f("Issued / signed", fmtD(it.issued)) : "") +
      (it.docStatus ? f("Document status", esc(it.docStatus)) : "") +
      (it.notes ? f("Notes", esc(it.notes)) : "") +
      f("Scope", esc(it.scope) + (it.active === false ? " · inactive" : "")) +
      docSection(it) +
      '<div class="dfield"><div class="dl">Task / assignment</div>' +
      '<div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center">' +
      '<input id="tAssignee" placeholder="Assignee" value="' + esc(task.assignee || "") + '" style="flex:1;min-width:110px;padding:9px;border-radius:8px;border:1px solid var(--hair);background:var(--surface-solid);color:var(--ink)">' +
      '<input id="tDue" type="date" value="' + esc(task.due || "") + '" style="padding:9px;border-radius:8px;border:1px solid var(--hair);background:var(--surface-solid);color:var(--ink)">' +
      '<select id="tStatus" style="padding:9px;border-radius:8px;border:1px solid var(--hair);background:var(--surface-solid);color:var(--ink)">' + ["none", "To-do", "In progress", "Done"].map(statusOpt).join("") + '</select>' +
      '<button class="icon-btn" id="tSave">Save</button></div></div>' +
      '<div class="dfield"><div class="dl">Review</div>' +
      (ver ? '<div class="item-sub" style="margin-bottom:8px">✅ Verified by <b>' + esc(ver.by || "") + '</b> on ' + esc(ver.at || "") + '</div>' : '<div class="item-sub" style="margin-bottom:8px">Not yet reviewed.</div>') +
      (READONLY ? '' : '<div style="display:flex;gap:6px"><button class="icon-btn" id="vMark">' + (ver ? "Re-verify (today)" : "✓ Mark verified") + '</button>' + (ver ? '<button class="icon-btn" id="vClear">Clear</button>' : '') + '</div>') +
      '</div>' +
      '<div class="dfield"><div class="dl">Renewal log</div><div id="logList">' + (logs.length ? logs.map(L => '<div class="log-entry">' + esc(L.text) + ' <small>— ' + L.date + '</small></div>').join("") : '<div class="item-sub">No log entries yet.</div>') + '</div>' +
      '<div style="display:flex;gap:6px;margin-top:8px"><input id="logIn" placeholder="Add a renewal note…" style="flex:1;padding:9px 11px;border-radius:9px;border:1px solid var(--hair);background:var(--surface-solid);color:var(--ink)"><button class="icon-btn" id="logAdd">Add</button></div></div>' +
      '</div>' +
      '<div class="drawer-actions">' + (READONLY ? "" : '<button id="dRenew" class="save">✓ Mark renewed</button><button id="dEdit">Edit</button>') + '<button id="dWatch">' + (isWatched(it.id) ? "★ Watching" : "☆ Watch") + '</button><button id="dIcs">Calendar (.ics)</button><button id="dQR">QR code</button>' + (it.scope === "provider" ? '<button id="dEmail">✉️ Email provider</button>' : "") + (READONLY ? "" : '<button id="dDel" class="del">Delete</button>') + '</div>';
    $("#dClose").onclick = closeDrawer;
    if ($("#dEdit")) $("#dEdit").onclick = () => renderDrawerEdit(it, false);
    if ($("#dDel")) $("#dDel").onclick = () => { if (confirm("Delete this item?")) { deleteItem(it); } };
    $("#dIcs").onclick = () => exportICS(it);
    // Auto-OCR every document on open (cached): fill the date if missing, verify if present.
    if (CLOUD && it.isFile && !READONLY) setTimeout(() => ocrReadDate(it, true), 80);
    if ($("#dEmail")) $("#dEmail").onclick = () => openEmailTemplate(it);
    $("#dWatch").onclick = () => { toggleWatch(it.id); renderDrawerView(it); render(); };
    if ($("#dRenew")) $("#dRenew").onclick = () => markRenewed(it);
    if ($("#tSave")) $("#tSave").onclick = () => {
      const t = { assignee: $("#tAssignee").value.trim(), due: $("#tDue").value, status: $("#tStatus").value };
      if (!t.assignee && !t.due && (t.status === "none" || !t.status)) delete OVERLAY.tasks[it.id];
      else { if (t.status === "none") t.status = "To-do"; OVERLAY.tasks[it.id] = t; }
      saveOverlay(); render(); toast("Task saved.");
    };
    if ($("#vMark")) $("#vMark").onclick = () => { OVERLAY.verified[it.id] = { by: (CURRENT_USER && CURRENT_USER.label) || "Admin", at: new Date().toISOString().slice(0, 10) }; logAudit("verify", it, "marked reviewed/verified"); saveOverlay(); renderDrawerView(it); render(); toast("Marked verified."); };
    if ($("#vClear")) $("#vClear").onclick = () => { delete OVERLAY.verified[it.id]; saveOverlay(); renderDrawerView(it); render(); toast("Verification cleared."); };
    wireAttach(it);
    $("#logAdd").onclick = () => { const v = $("#logIn").value.trim(); if (!v) return; OVERLAY.logs[it.id] = (OVERLAY.logs[it.id] || []); OVERLAY.logs[it.id].push({ text: v, date: new Date().toISOString().slice(0, 10) }); saveOverlay(); renderDrawerView(it); };
    if (READONLY) {
      const t = $("#tSave"); if (t) t.closest(".dfield").style.display = "none";
      const la = $("#logAdd"); if (la) { la.style.display = "none"; if ($("#logIn")) $("#logIn").style.display = "none"; }
    }
    wireQR(it); wireReadDate(it);
  }

  function renderDrawerEdit(it, isNew) {
    const fld = (l, id, val, type) => '<div class="dfield"><div class="dl">' + l + '</div><input id="' + id + '" type="' + (type || "text") + '" value="' + esc(val == null ? "" : val) + '"></div>';
    $("#drawer").innerHTML =
      '<div class="drawer-head"><div><div class="dh-cat">' + (isNew ? "Add compliance item" : "Edit item") + '</div><div class="dh-ent">' + esc(it.entity || "") + '</div></div><button class="drawer-close" id="dClose">×</button></div>' +
      '<div class="drawer-body">' +
      '<div class="dfield"><div class="dl">Scope</div><select id="eScope"><option value="provider"' + (it.scope === "provider" ? " selected" : "") + '>Provider</option><option value="facility"' + (it.scope === "facility" ? " selected" : "") + '>Facility</option><option value="other"' + (it.scope === "other" ? " selected" : "") + '>Other</option></select></div>' +
      fld("Entity (provider or facility name)", "eEntity", it.entity) +
      fld("Category", "eCat", it.category) +
      fld("Issuing authority", "eAuth", it.authority) +
      fld("ID / License number", "eNum", it.number) +
      fld("Expiration date", "eExp", it.expires, "date") +
      fld("Renewal lead days", "eLead", it.renewalLeadDays || 60, "number") +
      fld("Source folder / file (relative path)", "eLink", it.fileLink) +
      '<div class="dfield"><div class="dl">Notes</div><textarea id="eNotes" rows="3">' + esc(it.notes || "") + '</textarea></div>' +
      '</div>' +
      '<div class="drawer-actions"><button class="save" id="eSave">Save</button><button id="eCancel">Cancel</button></div>';
    $("#dClose").onclick = closeDrawer;
    $("#eCancel").onclick = () => isNew ? closeDrawer() : renderDrawerView(it);
    $("#eSave").onclick = () => {
      const get = id => $(id).value.trim();
      const rec = Object.assign({}, it, {
        scope: $("#eScope").value, entity: get("#eEntity"), category: get("#eCat"), authority: get("#eAuth"),
        number: get("#eNum"), expires: get("#eExp") || null, renewalLeadDays: +get("#eLead") || 60,
        fileLink: get("#eLink"), notes: get("#eNotes"), _edited: true
      });
      if (!rec.entity || !rec.category) { toast("Entity and category are required."); return; }
      saveItem(rec, isNew); logAudit(isNew ? "add" : "edit", rec); closeDrawer(); render(); toast(isNew ? "Item added." : "Changes saved.");
    };
  }

  function saveItem(rec, isNew) {
    const seedIds = new Set(((window.SENTINEL_SEED && window.SENTINEL_SEED.items) || []).map(i => i.id));
    if (isNew || !seedIds.has(rec.id)) {
      const i = OVERLAY.added.findIndex(a => a.id === rec.id);
      if (i >= 0) OVERLAY.added[i] = rec; else OVERLAY.added.push(rec);
    } else { OVERLAY.edits[rec.id] = rec; }
    saveOverlay(); buildData();
  }
  function deleteItem(it) {
    logAudit("delete", it);
    const seedIds = new Set(((window.SENTINEL_SEED && window.SENTINEL_SEED.items) || []).map(i => i.id));
    if (seedIds.has(it.id)) { if (!OVERLAY.deleted.includes(it.id)) OVERLAY.deleted.push(it.id); }
    OVERLAY.added = OVERLAY.added.filter(a => a.id !== it.id);
    delete OVERLAY.edits[it.id];
    saveOverlay(); buildData(); closeDrawer(); render(); toast("Item deleted.");
  }

  // ================= EXPORT =================
  function currentArr() { return sortItems(applyFilters(tabItems(state.tab))); }
  function exportCSV(arr) {
    arr = (arr && arr.length) ? arr : currentArr();
    const rows = [["Scope", "Entity", "Category", "Authority", "Number", "Issued", "Expires", "Status", "Days", "StartRenewalBy", "Notes", "FileLink"]];
    arr.forEach(it => {
      const st = computeStatus(it);
      rows.push([it.scope, it.entity, it.category, it.authority || "", it.number || "", it.issued || "", it.expires || "", st.label, st.days == null ? "" : st.days, st.startBy ? st.startBy.toISOString().slice(0, 10) : "", (it.notes || "").replace(/\s+/g, " "), it.fileLink || ""]);
    });
    const csv = rows.map(r => r.map(c => '"' + String(c).replace(/"/g, '""') + '"').join(",")).join("\r\n");
    download("sentinel-" + state.tab + "-" + new Date().toISOString().slice(0, 10) + ".csv", csv, "text/csv");
    toast("CSV exported (" + (rows.length - 1) + " rows).");
  }
  function exportICS(it) {
    const st = computeStatus(it);
    const when = st.startBy || st.expDate || today();
    const dt = when.toISOString().slice(0, 10).replace(/-/g, "");
    const uid = it.id + "@sentinel";
    const ics = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Sentinel//WCGTX//EN", "BEGIN:VEVENT",
      "UID:" + uid, "DTSTART;VALUE=DATE:" + dt, "DTEND;VALUE=DATE:" + dt,
      "SUMMARY:Renew — " + it.category + " (" + it.entity + ")",
      "DESCRIPTION:" + (it.category + " expires " + (it.expires || "n/a") + ". " + (it.notes || "")).replace(/[\r\n,]/g, " "),
      "BEGIN:VALARM", "TRIGGER:-P7D", "ACTION:DISPLAY", "DESCRIPTION:Sentinel renewal reminder", "END:VALARM",
      "END:VEVENT", "END:VCALENDAR"].join("\r\n");
    download("renew-" + it.id + ".ics", ics, "text/calendar");
    toast("Calendar reminder downloaded.");
  }
  function download(name, content, type) {
    const a = el("a"); a.href = URL.createObjectURL(new Blob([content], { type: type })); a.download = name;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
  }
  function doPrint() { Object.keys(state.openGroups).forEach(k => state.openGroups[k] = true); renderContent(); setTimeout(() => window.print(), 200); }

  function emailMe() {
    const scopes = [["provider", "Provider Compliance"], ["facility", "Facility Compliance"], ["other", "Other Compliance"]].filter(s => !CURRENT_USER || CURRENT_USER.tabs.indexOf(s[0]) >= 0);
    const ip = "width:100%;padding:9px 11px;border-radius:8px;border:1px solid var(--hair);background:var(--surface-solid);color:var(--ink)";
    const inbox = (window.SENTINEL_AUTH && window.SENTINEL_AUTH.email) || "your inbox";
    const rows = scopes.map(s => '<label class="toggle-pill" style="display:flex;gap:10px;cursor:pointer;border:1px solid var(--hair);border-radius:10px;padding:10px 12px;margin-bottom:8px"><input type="checkbox" class="em-scope" value="' + s[0] + '" checked> <b>' + s[1] + '</b></label>').join("");
    const body = '<div class="item-sub" style="margin-bottom:14px">Email the report for your tab(s) to <b>' + esc(inbox) + '</b> (your signed-in account).</div>' +
      rows + '<div class="auth-msg" id="emMsg" style="color:var(--st-expired);min-height:16px"></div>' +
      '<div class="drawer-actions" style="border:none;padding:6px 0 0"><button class="save" id="emSend">Send email</button></div>';
    openModal("Email a report", body);
    $("#emSend").onclick = () => {
      const sel = [...$("#modalInner").querySelectorAll(".em-scope:checked")].map(c => c.value);
      if (!sel.length) { $("#emMsg").textContent = "Select at least one tab."; return; }
      closeModal(); toast("Sending email…");
      const endpoint = CLOUD ? "/api/digest" : "http://localhost:8765/api/send-digest";
      fetch(endpoint, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ scopes: sel }) })
        .then(r => r.json())
        .then(d => toast(d.ok ? ("✓ Emailed report to " + (d.to || inbox)) : ("Send failed: " + (d.message || "").slice(-120))))
        .catch(() => toast(CLOUD ? "Email function error — check Vercel env vars." : "Email service not running — open via Start-Sentinel.bat, then try again."));
    };
  }

  // ================= TEMPLATED PROVIDER EMAILS =================
  // Built from Cynthia's Word templates. Blanks auto-fill from the item; user can edit.
  // The server sends the preview to the signed-in staff member's own inbox.
  const SIG_CRED = "\nThank You,\nCynthia Wray\nCredentialing Specialist\ncredentialing@wcgtx.com\n21175 Tomball Parkway #504  Houston, TX 77070   Ph# 346-226-6811 ext 1001";
  const SIG_ICCT = "\nThank you,\nCynthia Wray\nCredentialing Specialist\nImmediate Care Centers of Texas\ncwray@wcgtx.com";
  function greetingNow() { const h = new Date().getHours(); return h < 12 ? "morning" : h < 17 ? "afternoon" : "evening"; }
  function lastNameOf(entity) { const p = String(entity || "").replace(/[*]/g, "").trim().split(/\s+/); return p[p.length - 1] || ""; }

  const EMAIL_TEMPLATES = {
    reminder: {
      label: "Reminder — credential expiring",
      fields: [["greeting", "Greeting (morning/afternoon)"], ["drname", "Dr. (last name)"], ["credential", "Credential"], ["date", "Expiry date"]],
      fill: it => ({ greeting: greetingNow(), drname: lastNameOf(it.entity), credential: it.category || "", date: it.expires ? fmtD(it.expires) : "" }),
      subj: v => "Reminder: your " + v.credential + " expires on " + v.date,
      body: v => "Good " + v.greeting + " Dr. " + v.drname + ".  This is a reminder that your " + v.credential + " expires on " + v.date + ".  When you renew, please send a copy of your new certificate so we can update your credentialing file.\n" + SIG_CRED
    },
    thankyou: {
      label: "Thank you — renewal received",
      fields: [["credential", "Credential renewed"]],
      fill: it => ({ credential: it.category || "" }),
      subj: v => "Received: your renewed " + v.credential,
      body: v => "Thank you for sending a copy of your renewed " + v.credential + ".  Your credentialing file has been updated.\n" + SIG_CRED
    },
    onboard_phys: {
      label: "Onboarding — Physician",
      fields: [["greeting", "Greeting (morning/afternoon)"], ["drname", "Dr. (last name)"]],
      fill: it => ({ greeting: greetingNow(), drname: lastNameOf(it.entity) }),
      subj: () => "Welcome to WCGTX — credentialing onboarding",
      body: v => "Good " + v.greeting + " Dr. " + v.drname + " and welcome to WCGTX.  You will be receiving an email from PandaDoc with your Initial Application & Peer Reference Contact Info Form to complete.  Please send current copies of the following documents to begin your credentialing file:\n\n-CV/Resume\n-Texas Medical License Certificate\n-DEA Certificate\n-Certifications (including ACLS, ATLS, PALS, BLS)\n-Board Certification Letter (if applicable)\n-Driver License\n-Social Security Card\n-Diploma and/or ECFMG\n-Residency\n-Recent Flu Documentation\n-Recent TB Skin Test\n-CME (past 2 years, at least 20 hours)\n-TSCA (signed within 90 days or a fillable one has been attached for your convenience)\n" + SIG_CRED
    },
    onboard_rn: {
      label: "Onboarding — RN",
      fields: [["name", "Name"], ["facility", "Facility (ER)"], ["shift", "Shift"]],
      fill: it => ({ name: String(it.entity || "").replace(/[*]/g, "").trim(), facility: "", shift: "" }),
      subj: v => "Welcome to " + (v.facility || "") + " ER",
      body: v => "Good afternoon " + v.name + " and welcome to " + v.facility + " ER as " + v.shift + " shift.  You will be receiving an email from Paycor to register, complete Initial Application, and upload these required documents:\n\n-Resume\n-RN License\n-Driver License\n-Social Security Card\n-ACLS certificate\n-PALS certificate\n-BLS certificate\n-copy of Diploma (education)\n-recent Flu documentation\n-complete TB Questionnaire\n\nAfter you complete your Paycor registration, we will run a background check.  You will also receive an offer letter and employee handbook through PandaDoc to review and sign.  Please let us know if you have any questions.\n" + SIG_ICCT
    },
    onboard_fd: {
      label: "Onboarding — Front Desk",
      fields: [["name", "Name"], ["facility", "Facility (ER)"], ["shift", "Shift"]],
      fill: it => ({ name: String(it.entity || "").replace(/[*]/g, "").trim(), facility: "", shift: "" }),
      subj: v => "Welcome to " + (v.facility || "") + " ER",
      body: v => "Good afternoon " + v.name + " and welcome to " + v.facility + " ER as " + v.shift + " shift.  You will be receiving an email from Paycor to register, complete Initial Application, and upload these required documents:\n\n-Resume\n-Driver License\n-Social Security Card\n-BLS certificate\n-copy of Diploma (education)\n-recent Flu documentation\n\nAfter you complete your Paycor registration, we will run a background check.  You will also receive an offer letter and employee handbook through PandaDoc to review and sign.  Please let us know if you have any questions.\n" + SIG_ICCT
    }
  };

  function tmplToHtml(text) {
    const lines = text.split("\n"); let out = "", ul = false;
    for (const ln of lines) {
      if (/^\s*-/.test(ln)) { if (!ul) { out += '<ul style="margin:6px 0;padding-left:20px">'; ul = true; } out += "<li>" + esc(ln.replace(/^\s*-\s*/, "")) + "</li>"; }
      else { if (ul) { out += "</ul>"; ul = false; } out += ln.trim() === "" ? '<div style="height:10px"></div>' : "<div>" + esc(ln) + "</div>"; }
    }
    if (ul) out += "</ul>";
    return '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#0f172a;line-height:1.55">' + out + '</div>';
  }

  // A scan-to-upload QR block appended to every templated email.
  function uploadQrBlock(it) {
    const origin = (CLOUD ? location.origin : "https://sentinel-compliance-kappa.vercel.app");
    const url = origin + "/upload.html?item=" + encodeURIComponent(it.id);
    const qr = "https://api.qrserver.com/v1/create-qr-code/?size=220x220&data=" + encodeURIComponent(url);
    return '<div style="margin-top:18px;padding-top:14px;border-top:1px solid #e6ebf1;text-align:center;font-family:Segoe UI,Arial,sans-serif">' +
      '<div style="font-size:13px;color:#0f172a;margin-bottom:8px">📲 Scan this code to upload your document right from your phone:</div>' +
      '<img src="' + qr + '" alt="Upload QR code" width="180" height="180" style="border:1px solid #e6ebf1;border-radius:10px;padding:6px;background:#fff">' +
      '<div style="font-size:12px;color:#64748b;margin-top:6px">or open: <a href="' + url + '">' + url + '</a></div></div>';
  }

  // Email a single group's (facility / other / provider) compliance report to the SIGNED-IN user.
  function emailGroupToSelf(title, items) {
    const inbox = (window.SENTINEL_AUTH && window.SENTINEL_AUTH.email) || "your inbox";
    const rows = sortItems(items).map(it => {
      const st = computeStatus(it);
      const c = st.key === "expired" ? "#dc2626" : st.key === "critical" ? "#ea580c" : st.key === "due" ? "#ca8a04" : "#0d9488";
      return '<tr><td style="padding:6px 10px;border-bottom:1px solid #eee">' + esc(it.category) + '</td>' +
        '<td style="padding:6px 10px;border-bottom:1px solid #eee">' + (it.expires ? fmtD(it.expires) : (it.permanent ? "No expiry" : "—")) + '</td>' +
        '<td style="padding:6px 10px;border-bottom:1px solid #eee;color:' + c + ';font-weight:700">' + st.label + '</td>' +
        '<td style="padding:6px 10px;border-bottom:1px solid #eee">' + (it.isFile ? "on file" : "—") + '</td></tr>';
    }).join("");
    const html = '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#0f172a"><h2 style="color:#0f766e">' + esc(title) + ' — compliance report</h2>' +
      '<table style="border-collapse:collapse;width:100%"><thead><tr style="text-align:left;color:#475569;font-size:12px;text-transform:uppercase"><th style="padding:6px 10px">Item</th><th style="padding:6px 10px">Expires</th><th style="padding:6px 10px">Status</th><th style="padding:6px 10px">Proof</th></tr></thead><tbody>' + rows + '</tbody></table></div>';
    toast("Sending report to " + inbox + "…");
    const endpoint = CLOUD ? "/api/send-template" : "http://localhost:8765/api/send-template";
    fetch(endpoint, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ subject: title + " — compliance report", html: html }) })
      .then(r => r.json())
      .then(d => toast(d.ok ? ("✓ Report emailed to " + (d.to || inbox)) : ("Send failed: " + (d.message || "").slice(-120))))
      .catch(() => toast(CLOUD ? "Email function error." : "Run via Start-Sentinel.bat to send locally."));
  }

  function openEmailTemplate(it) {
    let cur = "reminder";
    const ip = "width:100%;padding:9px;border-radius:8px;border:1px solid var(--hair);background:var(--surface-solid);color:var(--ink)";
    const opts = Object.keys(EMAIL_TEMPLATES).map(k => '<option value="' + k + '">' + EMAIL_TEMPLATES[k].label + '</option>').join("");
    openModal("Email provider (from template)",
      '<div class="item-sub" style="margin-bottom:10px">This goes to the provider\'s email below (auto-detected — edit if needed). A scan-to-upload QR is added automatically.</div>' +
      '<div class="dl">To — provider\'s email <span class="req">*</span></div><input id="etTo" value="' + esc(it.email || "") + '" placeholder="provider@email.com" style="' + ip + ';margin-bottom:10px">' +
      '<div class="dl">Template</div><select id="etSel" style="' + ip + ';margin-bottom:10px">' + opts + '</select>' +
      '<div id="etFields"></div>' +
      '<div class="dl" style="margin-top:8px">Preview</div><div id="etPrev" style="border:1px solid var(--hair);border-radius:10px;padding:12px;background:#fff;max-height:300px;overflow:auto"></div>' +
      '<div class="drawer-actions" style="border:none;padding:10px 0 0"><button class="save" id="etSend">Send to provider</button></div>');
    const vals = () => { const o = {}; [...$("#modalInner").querySelectorAll(".etf")].forEach(i => o[i.dataset.k] = i.value); return o; };
    const refreshPrev = () => { $("#etPrev").innerHTML = tmplToHtml(EMAIL_TEMPLATES[cur].body(vals())) + uploadQrBlock(it); };
    const drawFields = () => {
      const t = EMAIL_TEMPLATES[cur], v = t.fill(it);
      $("#etFields").innerHTML = t.fields.map(([k, lab]) => '<div style="margin-bottom:8px"><div class="dl">' + lab + ' <span class="req">*</span></div><input class="etf" data-k="' + k + '" value="' + esc(v[k] || "") + '" style="' + ip + '"></div>').join("");
      [...$("#modalInner").querySelectorAll(".etf")].forEach(i => i.oninput = refreshPrev);
      refreshPrev();
    };
    $("#etSel").onchange = e => { cur = e.target.value; drawFields(); };
    drawFields();
    $("#etSend").onclick = () => {
      const t = EMAIL_TEMPLATES[cur], v = vals();
      const to = ($("#etTo").value || "").trim();
      // Every field is required — flag blanks and block the send.
      const toEl = $("#etTo"); toEl.classList.toggle("err", !to);
      const missing = [];
      [...$("#modalInner").querySelectorAll(".etf")].forEach(i => { const blank = !i.value.trim(); i.classList.toggle("err", blank); if (blank) missing.push(i.dataset.k); });
      if (!to || missing.length) { toast("All fields are required — please fill the highlighted field(s)."); return; }
      const upUrl = (CLOUD ? location.origin : "https://sentinel-compliance-kappa.vercel.app") + "/upload.html?item=" + encodeURIComponent(it.id);
      const payload = { to: to, subject: t.subj(v), html: tmplToHtml(t.body(v)) + uploadQrBlock(it), text: t.body(v) + "\n\nUpload your document here: " + upUrl };
      closeModal(); toast("Sending email to " + to + "…");
      const endpoint = CLOUD ? "/api/send-template" : "http://localhost:8765/api/send-template";
      fetch(endpoint, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) })
        .then(r => r.json())
        .then(d => toast(d.ok ? ("✓ Email sent to " + d.to) : ("Send failed: " + (d.message || "").slice(-120))))
        .catch(() => toast(CLOUD ? "Email function error — check Vercel env vars." : "Run via Start-Sentinel.bat to send locally."));
    };
  }

  // ================= MISC =================
  let toastT;
  function toast(msg) { const t = $("#toast"); t.textContent = msg; t.classList.add("show"); clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove("show"), 2600); }
  function showSecurityNote() {
    drawerItem = null; $("#drawer").classList.add("show"); $("#drawerBack").classList.add("show");
    $("#drawer").innerHTML = '<div class="drawer-head"><div class="dh-cat">Security &amp; data notes</div><button class="drawer-close" id="dClose">×</button></div>' +
      '<div class="drawer-body" style="font-size:14px;line-height:1.6;color:var(--ink-2)">' +
      '<p><b>This login is a deterrent, not encryption.</b> Sentinel runs entirely in your browser from a static file. The access codes are stored only as SHA-256 hashes — there is no plaintext password in any file. But anyone who can open these files can read the hashes and, with effort, brute-force a weak code offline. Use long, unique codes.</p>' +
      '<p><b>For true security</b> (real authentication, audit logs, encryption at rest), host Sentinel behind Microsoft 365 / Entra sign-in: put this folder in a SharePoint document library or an Azure Static Web App and require org login. I can help set that up — your data file stays the same.</p>' +
      '<p><b>Your edits</b> (add/edit/delete + renewal log notes) are saved in this browser\'s localStorage, layered over the generated data. To make changes permanent for everyone, update the source spreadsheets and re-run <code>generate_data.py</code>.</p>' +
      '<p><b>Source of truth:</b> <code>data.json</code> / <code>data.js</code>, generated from <i>WCGTX Physician Roster.xlsx</i> and the CHER/Frisco compliance exports.</p></div>';
    $("#dClose").onclick = closeDrawer;
  }

  // ================= MODAL =================
  function openModal(title, bodyHtml, headExtra) {
    const m = $("#modal");
    $("#modalInner").className = "modal";
    $("#modalInner").innerHTML =
      // NOTE: `title` may legitimately contain markup (openListModal appends a count span), and
      // some callers pre-escape, so this is NOT escaped here — callers passing untrusted text
      // MUST esc() it themselves. See openInsights, which reads entity names back out of a
      // data- attribute (the browser decodes entities on read, undoing the original escaping).
      '<div class="modal-head"><h2>' + title + '</h2><span class="spacer"></span>' + (headExtra || "") +
      '<button class="drawer-close" id="mClose">×</button></div>' +
      '<div class="modal-body">' + bodyHtml + '</div>';
    m.classList.add("show");
    $("#mClose").onclick = closeModal;
    return $("#modalInner");
  }
  function closeModal() { $("#modal").classList.remove("show"); $("#modalInner").className = "modal"; }

  // ================= ANALYTICS / INSIGHTS =================
  function barRow(label, val, max, cls) {
    const pct = Math.round(100 * val / (max || 1));
    return '<div class="bar-row"><div class="bl">' + esc(label) + '</div><div class="bar-track"><div class="bar-fill" style="width:' + Math.max(pct, val ? 4 : 0) + '%' + (cls === "risk" ? ";background:linear-gradient(90deg,var(--st-critical),var(--st-expired))" : "") + '"></div></div><div class="bv">' + val + '</div></div>';
  }
  function donutSVG(parts) {
    const total = parts.reduce((a, p) => a + p[1], 0) || 1;
    const r = 52, cx = 70, cy = 70, c = 2 * Math.PI * r; let off = 0, segs = "";
    parts.forEach(p => { if (!p[1]) return; const len = c * p[1] / total;
      segs += '<circle cx="' + cx + '" cy="' + cy + '" r="' + r + '" fill="none" stroke="' + p[2] + '" stroke-width="16" stroke-dasharray="' + len + ' ' + (c - len) + '" stroke-dashoffset="' + (-off) + '" transform="rotate(-90 ' + cx + ' ' + cy + ')"/>'; off += len; });
    const legend = parts.filter(p => p[1]).map(p => '<span class="leg-click" data-status="' + (p[3] || "") + '" style="cursor:pointer"><i style="background:' + p[2] + '"></i>' + p[0] + ' (' + p[1] + ')</span>').join("");
    return '<div style="display:flex;gap:20px;align-items:center;flex-wrap:wrap"><svg width="140" height="140" viewBox="0 0 140 140"><circle cx="70" cy="70" r="52" fill="none" stroke="var(--hair)" stroke-width="16"/>' + segs + '<text x="70" y="77" text-anchor="middle" font-size="24" font-weight="800" fill="currentColor">' + total + '</text></svg><div class="legend" style="flex-direction:column;align-items:flex-start;gap:7px">' + legend + '</div></div>';
  }
  function heatColor(n) {
    if (n === 0) return ["var(--hair)", "var(--ink-3)"];
    if (n <= 2) return ["#99f6e4", "#134e4a"];
    if (n <= 5) return ["#2dd4bf", "#06302c"];
    if (n <= 9) return ["#f59e0b", "#3b2600"];
    return ["#ef4444", "#ffffff"];
  }
  function monthItems(arr, y, m) { return arr.filter(it => { const e = parseD(it.expires); return e && e.getFullYear() === y && e.getMonth() === m; }); }
  function openListModal(title, items) {
    items = sortItems(items);
    const body = items.length ? items.map((it, i) => { const st = computeStatus(it); return '<div class="pal-item" data-i="' + i + '"><div class="ico s-' + st.key + '" style="width:30px;height:30px"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' + ICONS[iconFor(it)] + '</svg></div><div style="flex:1"><div class="pal-cat">' + esc(it.category) + '</div><div class="pal-ent">' + esc(it.entity) + " · " + it.scope + (it.expires ? " · " + fmtD(it.expires) : "") + '</div></div><div class="status-badge s-' + st.key + '">' + st.label + '</div></div>'; }).join("") : '<div class="empty">Nothing here.</div>';
    openModal(title + ' <span style="color:var(--ink-3);font-weight:600;font-size:14px">(' + items.length + ')</span>', body);
    [...$("#modalInner").querySelectorAll(".pal-item")].forEach(r => r.onclick = () => { const it = items[+r.dataset.i]; closeModal(); const go = () => { state.tab = it.scope; render(); openDrawer(it, false); }; if (UNLOCKED.has(it.scope)) go(); else promptTabCode(it.scope, go); });
  }
  function monthBuckets(arr) {
    const start = today(); const b = [];
    for (let i = 0; i < 12; i++) { const d = new Date(start.getFullYear(), start.getMonth() + i, 1); b.push({ y: d.getFullYear(), m: d.getMonth(), label: d.toLocaleDateString("en-US", { month: "short", year: "2-digit" }), n: 0 }); }
    arr.forEach(it => { const e = parseD(it.expires); if (!e) return; for (const x of b) { if (e.getFullYear() === x.y && e.getMonth() === x.m) { x.n++; break; } } });
    return b;
  }
  function openInsights() {
    const all = unlockedItems();
    if (!all.length) { openModal("Analytics & insights", LOCK_MSG); return; }
    const s = statsFor(all);
    const donut = donutSVG([["Expired", s.expired, "var(--st-expired)", "expired"], ["Critical", s.critical, "var(--st-critical)", "critical"], ["Due soon", s.due, "var(--st-due)", "due"], ["Good", s.good, "var(--st-good)", "good"], ["Permanent", s.permanent, "var(--st-perm)", "permanent"], ["Recurring", s.recurring, "var(--st-rec)", "recurring"], ["Pending", s.pending, "var(--st-pend)", "pending"]]);
    const months = monthBuckets(all); const mmax = months.reduce((a, b) => Math.max(a, b.n), 1);
    const monthBar = months.map(m => { const pct = Math.round(100 * m.n / mmax); return '<div class="bar-row clk" data-y="' + m.y + '" data-m="' + m.m + '" style="cursor:pointer"><div class="bl">' + esc(m.label) + '</div><div class="bar-track"><div class="bar-fill" style="width:' + Math.max(pct, m.n ? 4 : 0) + '%"></div></div><div class="bv">' + m.n + '</div></div>'; }).join("");
    const riskMap = {}; all.forEach(it => { const k = computeStatus(it).key; if (k === "expired" || k === "critical") riskMap[it.entity] = (riskMap[it.entity] || 0) + 1; });
    const risk = Object.keys(riskMap).map(n => [n, riskMap[n]]).sort((a, b) => b[1] - a[1]).slice(0, 8);
    const topRisk = risk.length ? risk.map(r => { const pct = Math.round(100 * r[1] / risk[0][1]); return '<div class="bar-row clk-ent" data-ent="' + esc(r[0]) + '" style="cursor:pointer"><div class="bl">' + esc(r[0]) + '</div><div class="bar-track"><div class="bar-fill" style="width:' + Math.max(pct, 4) + '%;background:linear-gradient(90deg,var(--st-critical),var(--st-expired))"></div></div><div class="bv">' + r[1] + '</div></div>'; }).join("") : '<div class="item-sub">No expired or critical items 🎉</div>';
    const heat = months.map(x => { const cc = heatColor(x.n); return '<div class="heat-cell clk" data-y="' + x.y + '" data-m="' + x.m + '" style="background:' + cc[0] + ';color:' + cc[1] + '"><div class="hm" style="color:inherit;opacity:.85">' + x.label + '</div><div class="hn">' + x.n + '</div></div>'; }).join("");
    const heatLegend = '<div class="legend" style="margin-top:12px"><span><i style="background:var(--hair)"></i>0</span><span><i style="background:#99f6e4"></i>1–2</span><span><i style="background:#2dd4bf"></i>3–5</span><span><i style="background:#f59e0b"></i>6–9</span><span><i style="background:#ef4444"></i>10+</span></div>';
    const body = '<div class="item-sub" style="margin-bottom:14px">Tip: click any month, segment, or bar to see the items behind it.</div>' +
      '<div class="charts-grid">' +
      '<div class="chart-card"><h3>Status mix (active)</h3>' + donut + '</div>' +
      '<div class="chart-card"><h3>Expirations — next 12 months</h3>' + monthBar + '</div></div>' +
      '<div class="chart-card" style="margin-top:18px"><h3>Highest-risk entities (expired + critical)</h3>' + topRisk + '</div>' +
      '<div class="chart-card" style="margin-top:18px"><h3>Expiration heatmap — next 12 months</h3><div class="heat">' + heat + '</div>' + heatLegend + '</div>' +
      '<div class="chart-card" style="margin-top:18px"><h3>Compliance score trend (monthly)</h3>' + trendHTML() + '</div>';
    openModal("Analytics & insights", body);
    const root = $("#modalInner");
    root.querySelectorAll(".clk").forEach(e => e.onclick = () => { const y = +e.dataset.y, m = +e.dataset.m; openListModal("Expiring " + new Date(y, m, 1).toLocaleDateString("en-US", { month: "long", year: "numeric" }), monthItems(all, y, m)); });
    // esc() the title: dataset.ent is an entity name from the Excel roster, and reading it back
    // off the attribute returns the DECODED text — so the escaping done when the attribute was
    // written is already undone by the time it reaches openModal, which does not escape.
    root.querySelectorAll(".clk-ent").forEach(e => e.onclick = () => openListModal(esc(e.dataset.ent), all.filter(i => i.entity === e.dataset.ent)));
    root.querySelectorAll(".leg-click").forEach(e => e.onclick = () => { const k = e.dataset.status; openListModal(e.textContent.replace(/\s*\(\d+\)\s*$/, ""), all.filter(i => computeStatus(i).key === k)); });
  }

  // ================= CONTACTS =================
  function openContacts() {
    if (!UNLOCKED.has("facility")) { openModal("Facility contacts directory", LOCK_MSG); return; }
    const cs = (window.SENTINEL_SEED && window.SENTINEL_SEED.contacts) || [];
    if (!cs.length) { openModal("Facility contacts", '<div class="empty">No contacts in the data.</div>'); return; }
    const byF = {}; cs.forEach(c => { (byF[c.facility] = byF[c.facility] || []).push(c); });
    const section = f => '<div class="contact-col"><div class="contact-fac">' + esc(f) + '</div><div class="contact-grid">' +
      byF[f].map(c => '<div class="contact-card"><div class="cc-name">' + esc(c.name) + '</div><div class="cc-role">' + esc(c.role || "") + '</div>' +
        (c.phone ? '<a href="tel:' + esc(c.phone.split("/")[0].trim()) + '">📞 ' + esc(c.phone) + '</a>' : "") +
        (c.email ? '<a href="mailto:' + esc(c.email) + '">✉️ ' + esc(c.email) + '</a>' : "") + '</div>').join("") + '</div></div>';
    // Facilities side by side (Castle Hills left, Frisco right); Help & Support spans full width below.
    const facs = Object.keys(byF).filter(f => !/Help & Support/i.test(f));
    const support = Object.keys(byF).filter(f => /Help & Support/i.test(f));
    const html = '<div class="contact-cols">' + facs.map(section).join("") + '</div>' + support.map(section).join("");
    openModal("Facility contacts directory", html);
  }

  // ================= AUDIT =================
  function openAudit() {
    const a = OVERLAY.audit || [];
    const head = a.length ? '<button class="icon-btn" id="aExport">Export CSV</button>' : "";
    const body = a.length ? a.map(x => '<div class="audit-row"><span class="ac ' + x.action + '">' + x.action + '</span><div><b>' + esc(x.category || "") + '</b> — ' + esc(x.entity || "") + (x.detail ? ' · ' + esc(x.detail) : "") + '</div><span class="at">' + esc(x.at) + '</span></div>').join("")
      : '<div class="empty"><h3>No changes yet</h3><p>Adds, edits, deletes and document attachments you make in Sentinel are logged here with a timestamp.</p></div>';
    openModal("Change log / audit trail", body, head);
    const ex = $("#aExport"); if (ex) ex.onclick = () => { const rows = [["When", "Action", "Entity", "Category", "Detail"]].concat(a.map(x => [x.at, x.action, x.entity, x.category, x.detail])); download("sentinel-audit-" + new Date().toISOString().slice(0, 10) + ".csv", rows.map(r => r.map(c => '"' + String(c || "").replace(/"/g, '""') + '"').join(",")).join("\r\n"), "text/csv"); toast("Audit log exported."); };
  }

  // ================= GAP REPORT =================
  function openGapReport() {
    if (!unlockedItems().length) { openModal("Document gap report", LOCK_MSG); return; }
    const missing = unlockedItems().filter(i => (i.isFile === false && !i.centralProof));
    const byE = {}; missing.forEach(i => { (byE[i.entity] = byE[i.entity] || []).push(i); });
    const head = '<button class="icon-btn" id="gExport">Export CSV</button>';
    let body = '<div class="item-sub" style="margin-bottom:14px">' + missing.length + ' active items have <b>no proof document</b> attached, across ' + Object.keys(byE).length + ' providers/facilities. Use the panel’s “Attach document” to fix any.</div>';
    Object.keys(byE).sort().forEach(e => {
      body += '<div class="contact-fac">' + esc(e) + ' <span class="n">(' + byE[e].length + ')</span></div>' +
        byE[e].map(i => '<div class="audit-row"><span class="ac edit">gap</span><div>' + esc(i.category) + '</div><span class="at">' + (i.expires ? fmtD(i.expires) : "") + '</span></div>').join("");
    });
    openModal("Document gap report", body, head);
    const ex = $("#gExport"); if (ex) ex.onclick = () => { const rows = [["Entity", "Category", "Expires", "Folder"]].concat(missing.map(i => [i.entity, i.category, i.expires || "", i.folderLink || i.fileLink || ""])); download("sentinel-gaps-" + new Date().toISOString().slice(0, 10) + ".csv", rows.map(r => r.map(c => '"' + String(c || "").replace(/"/g, '""') + '"').join(",")).join("\r\n"), "text/csv"); toast("Gap report exported."); };
  }

  // ================= GLOBAL SEARCH / COMMAND PALETTE =================
  function openPalette() {
    const m = $("#modal");
    $("#modalInner").className = "modal palette";
    $("#modalInner").innerHTML = '<input id="palIn" type="search" name="ask-' + Math.random().toString(36).slice(2) + '" autocomplete="off" autocorrect="off" spellcheck="false" readonly data-1p-ignore data-lpignore="true" placeholder="Ask: what expires in March? · who has no DEA? · expired licenses · next 30 days"><div class="results" id="palRes"></div>';
    m.classList.add("show");
    const inp = $("#palIn"); setTimeout(() => { inp.removeAttribute("readonly"); inp.focus(); }, 60);
    let sel = 0, list = [];
    function draw() {
      if (!list.length) { $("#palRes").innerHTML = '<div class="pal-empty">' + (inp.value.trim() ? "No matches." : "Type to search across all tabs…") + '</div>'; return; }
      $("#palRes").innerHTML = list.map((it, i) => { const st = computeStatus(it); return '<div class="pal-item' + (i === sel ? " sel" : "") + '" data-i="' + i + '"><div class="ico s-' + st.key + '" style="width:30px;height:30px"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' + ICONS[iconFor(it)] + '</svg></div><div style="flex:1"><div class="pal-cat">' + esc(it.category) + '</div><div class="pal-ent">' + esc(it.entity) + " · " + it.scope + '</div></div><div class="status-badge s-' + st.key + '">' + st.label + '</div></div>'; }).join("");
      [...$("#palRes").querySelectorAll(".pal-item")].forEach(e => e.onclick = () => pick(+e.dataset.i));
    }
    function run(q) {
      q = q.trim();
      if (!q) { list = []; sel = 0; draw(); return; }
      const sm = smartQuery(q, unlockedItems());
      list = sm ? sortItems(sm).slice(0, 80)
        : unlockedItems().filter(it => (it.entity + " " + it.category + " " + (it.authority || "") + " " + (it.number || "") + " " + (it.notes || "")).toLowerCase().includes(q.toLowerCase())).slice(0, 50);
      sel = 0; draw();
    }
    function pick(i) { const it = list[i]; if (!it) return; closeModal(); const go = () => { state.tab = it.scope; state.quickView = ""; state.status = ""; render(); openDrawer(it, false); }; if (UNLOCKED.has(it.scope)) go(); else promptTabCode(it.scope, go); }
    inp.oninput = () => run(inp.value);
    inp.onkeydown = e => { if (e.key === "ArrowDown") { sel = Math.min(sel + 1, list.length - 1); draw(); e.preventDefault(); } else if (e.key === "ArrowUp") { sel = Math.max(sel - 1, 0); draw(); e.preventDefault(); } else if (e.key === "Enter") pick(sel); else if (e.key === "Escape") closeModal(); };
    run("");
  }

  // ================= PER-PROVIDER / FACILITY BINDER =================
  function printBinder(entity) {
    const items = sortItems(DATA.filter(i => i.entity === entity));
    const rows = items.map(it => { const st = computeStatus(it); return '<tr><td>' + esc(it.category) + '</td><td>' + esc(it.authority || "") + '</td><td>' + esc(it.number || "") + '</td><td>' + (it.expires ? fmtD(it.expires) : (it.permanent ? "Permanent" : "—")) + '</td><td>' + st.label + (st.days != null ? " (" + countdownText(st) + ")" : "") + '</td><td>' + (it.isFile ? "✓ on file" : "—") + '</td></tr>'; }).join("");
    const pa = $("#printArea");
    pa.innerHTML = '<div class="binder"><h1>' + esc(entity) + '</h1><div class="bsub">Compliance credential file · Sentinel / WCGTX · printed ' + new Date().toLocaleDateString() + '</div>' +
      '<table class="btable"><thead><tr><th>Item</th><th>Authority</th><th>Number</th><th>Expires</th><th>Status</th><th>Proof</th></tr></thead><tbody>' + rows + '</tbody></table>' +
      '<div class="bfoot">' + items.length + ' tracked items. Generated by Sentinel Compliance Command Center.</div></div>';
    document.body.classList.add("printing");
    window.print();
    setTimeout(() => document.body.classList.remove("printing"), 400);
  }

  // ================= MARK RENEWED =================
  function cycleDays(it) {
    const c = (it.category || "").toLowerCase();
    if (/\bdea\b/.test(c) && c.indexOf("verif") < 0) return 1095;     // individual DEA ~3 yrs
    if (/acls|bls|pals|atls|board|npdb|tsca|clia|cola/.test(c)) return 730;  // 2 yrs
    return 365;
  }
  function markRenewed(it) {
    const exp = parseD(it.expires) || today();
    const base = exp > today() ? exp : today();
    const def = addDays(base, cycleDays(it)).toISOString().slice(0, 10);
    const v = prompt('Mark "' + it.category + '" renewed.\nNew expiry date (YYYY-MM-DD):', def);
    if (!v) return;
    // Validate before saving. An unparseable value (someone typing "pending") was stored as the
    // expiry and later crashed the Timeline render and the .ics export mid-loop, blanking the
    // view for everyone sharing the overlay.
    const dv = v.trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dv) || !parseD(dv)) {
      toast("Please enter the date as YYYY-MM-DD (for example " + def + ").");
      return;
    }
    const rec = Object.assign({}, it, { expires: dv });
    saveItem(rec, false);
    OVERLAY.logs[it.id] = (OVERLAY.logs[it.id] || []);
    OVERLAY.logs[it.id].push({ text: "Renewed — new expiry " + v.trim(), date: new Date().toISOString().slice(0, 10) });
    saveOverlay(); logAudit("renew", rec, "new expiry " + v.trim());
    drawerItem = rec; renderDrawerView(rec); render(); toast("Marked renewed → " + v.trim());
  }

  // ================= TASKS =================
  function openTasks() {
    const ts = Object.keys(OVERLAY.tasks).map(id => ({ id: id, t: OVERLAY.tasks[id], it: DATA.find(x => x.id === id) })).filter(x => x.it);
    const order = { "To-do": 0, "In progress": 1, "Done": 2 };
    ts.sort((a, b) => (order[a.t.status] || 0) - (order[b.t.status] || 0));
    const body = ts.length ? ts.map((x, i) => '<div class="pal-item" data-i="' + i + '"><span class="ac ' + (x.t.status === "Done" ? "renew" : x.t.status === "In progress" ? "edit" : "delete") + '">' + esc(x.t.status || "To-do") + '</span><div style="flex:1"><div class="pal-cat">' + esc(x.it.category) + '</div><div class="pal-ent">' + esc(x.it.entity) + (x.t.assignee ? " · 👤 " + esc(x.t.assignee) : "") + (x.t.due ? " · due " + esc(x.t.due) : "") + '</div></div></div>').join("")
      : '<div class="empty"><h3>No tasks yet</h3><p>Open any item and use the “Task” fields to assign an owner, due date, and status.</p></div>';
    openModal("Tasks &amp; assignments", body);
    ts.length && [...$("#modalInner").querySelectorAll(".pal-item")].forEach(r => r.onclick = () => { const it = ts[+r.dataset.i].it; closeModal(); const go = () => { state.tab = it.scope; render(); openDrawer(it, false); }; if (UNLOCKED.has(it.scope)) go(); else promptTabCode(it.scope, go); });
  }

  // ================= CUSTOM LEAD TIMES (settings) =================
  function openSettings() {
    const cats = [...new Set(DATA.map(i => i.category))].sort();
    const rows = cats.map(c => {
      const sample = DATA.find(i => i.category === c) || {};
      const cur = (OVERLAY.leads[c] != null) ? OVERLAY.leads[c] : (sample.renewalLeadDays || 60);
      return '<div class="bar-row"><div class="bl" style="width:auto;flex:1;text-align:left">' + esc(c) + '</div><input type="number" class="lead-in" data-cat="' + esc(c) + '" value="' + cur + '" style="width:90px;padding:7px 9px;border-radius:8px;border:1px solid var(--hair);background:var(--surface-solid);color:var(--ink)"> <span class="item-sub">days</span></div>';
    }).join("");
    openModal("Renewal lead times", '<div class="item-sub" style="margin-bottom:14px">“Start renewal by” = expiry minus these days. Changes apply instantly and are saved in this browser.</div>' + rows + '<div class="drawer-actions" style="border:none;padding:16px 0 0"><button class="save" id="leadSave">Save</button><button id="leadReset">Reset to defaults</button></div>');
    $("#leadSave").onclick = () => { [...$("#modalInner").querySelectorAll(".lead-in")].forEach(inp => { OVERLAY.leads[inp.dataset.cat] = +inp.value || 0; }); saveOverlay(); closeModal(); render(); toast("Lead times saved."); };
    $("#leadReset").onclick = () => { OVERLAY.leads = {}; saveOverlay(); closeModal(); render(); toast("Lead times reset to defaults."); };
  }

  // ================= BACKUP / RESTORE =================
  function backup() { download("sentinel-backup-" + new Date().toISOString().slice(0, 10) + ".json", JSON.stringify(OVERLAY, null, 2), "application/json"); toast("Backup downloaded."); }
  function restore() {
    const inp = el("input"); inp.type = "file"; inp.accept = "application/json,.json";
    inp.onchange = () => { const f = inp.files[0]; if (!f) return; const rd = new FileReader(); rd.onload = () => { try { const o = JSON.parse(rd.result); OVERLAY = Object.assign({ edits: {}, added: [], deleted: [], logs: {}, watch: [], audit: [], leads: {}, tasks: {}, snapshots: {} }, o); saveOverlay(); buildData(); render(); toast("Backup restored."); } catch (e) { toast("Invalid backup file."); } }; rd.readAsText(f); };
    inp.click();
  }

  // ================= TREND SNAPSHOTS =================
  function recordSnapshot() {
    const all = DATA.filter(i => i.active !== false); const s = statsFor(all);
    const key = new Date().toISOString().slice(0, 7);
    OVERLAY.snapshots[key] = { score: s.score, expired: s.expired, critical: s.critical, due: s.due, total: s.total, date: new Date().toISOString().slice(0, 10) };
    saveOverlay();
  }
  function trendHTML() {
    const keys = Object.keys(OVERLAY.snapshots).sort();
    if (keys.length < 2) return '<div class="item-sub">Trend appears after Sentinel has data from at least two different months. (A snapshot is saved automatically each month you open it.)</div>';
    const max = 100;
    return keys.map(k => { const s = OVERLAY.snapshots[k]; return barRow(k, s.score, max); }).join("");
  }

  // ================= NOTIFICATION BELL =================
  function bellCount() { let n = 0; unlockedItems().forEach(i => { const k = computeStatus(i).key; if (k === "expired" || k === "critical") n++; }); return n; }
  function updateBell() { const b = $("#bellCount"); if (!b) return; const n = bellCount(); b.textContent = n; b.style.display = n ? "inline-flex" : "none"; }
  function openBell() {
    const all = unlockedItems();
    if (!all.length) { openModal("Alerts &amp; upcoming", LOCK_MSG); return; }
    const urgent = sortItems(all.filter(i => { const k = computeStatus(i).key; return k === "expired" || k === "critical"; }));
    const soon = sortItems(all.filter(i => computeStatus(i).key === "due"));
    function list(items) { return items.slice(0, 50).map((it, i) => { const st = computeStatus(it); return '<div class="pal-item" data-id="' + esc(it.id) + '"><div class="ico s-' + st.key + '" style="width:30px;height:30px"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' + ICONS[iconFor(it)] + '</svg></div><div style="flex:1"><div class="pal-cat">' + esc(it.category) + '</div><div class="pal-ent">' + esc(it.entity) + " · " + it.scope + '</div></div><div class="status-badge s-' + st.key + '">' + countdownText(st) + '</div></div>'; }).join(""); }
    const body = '<h3 style="margin:0 0 8px;font-size:13px;color:var(--st-expired)">Needs attention now (' + urgent.length + ')</h3>' + (urgent.length ? list(urgent) : '<div class="item-sub">None.</div>') +
      '<h3 style="margin:18px 0 8px;font-size:13px;color:var(--st-due)">Coming up — due soon (' + soon.length + ')</h3>' + (soon.length ? list(soon) : '<div class="item-sub">None.</div>');
    openModal("Alerts &amp; upcoming", body);
    [...$("#modalInner").querySelectorAll(".pal-item")].forEach(r => r.onclick = () => { const it = DATA.find(x => x.id === r.dataset.id); if (!it) return; closeModal(); const go = () => { state.tab = it.scope; render(); openDrawer(it, false); }; if (UNLOCKED.has(it.scope)) go(); else promptTabCode(it.scope, go); });
  }

  // ================= CALENDAR VIEW =================
  function renderCalendar(c, arr) {
    if (state.calY == null) { const t = today(); state.calY = t.getFullYear(); state.calM = t.getMonth(); }
    const y = state.calY, m = state.calM;
    const first = new Date(y, m, 1), startDow = first.getDay(), days = new Date(y, m + 1, 0).getDate();
    const byDay = {}; arr.forEach(it => { const e = parseD(it.expires); if (e && e.getFullYear() === y && e.getMonth() === m) (byDay[e.getDate()] = byDay[e.getDate()] || []).push(it); });
    const head = el("div"); head.style.cssText = "display:flex;align-items:center;gap:12px;margin-bottom:14px";
    head.innerHTML = '<button class="icon-btn" id="calPrev">‹ Prev</button><b style="font-size:16px;flex:1;text-align:center">' + first.toLocaleDateString("en-US", { month: "long", year: "numeric" }) + '</b><button class="icon-btn" id="calNext">Next ›</button><button class="icon-btn" id="calIcs">Export all upcoming (.ics)</button>';
    c.appendChild(head);
    const grid = el("div"); grid.style.cssText = "display:grid;grid-template-columns:repeat(7,1fr);gap:6px";
    ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].forEach(d => { const h = el("div", "", d); h.style.cssText = "text-align:center;font-size:11px;font-weight:700;color:var(--ink-3);padding:4px"; grid.appendChild(h); });
    for (let i = 0; i < startDow; i++) grid.appendChild(el("div"));
    for (let d = 1; d <= days; d++) {
      const items = byDay[d] || [];
      const worst = items.length ? items.map(i => computeStatus(i).key).sort((a, b) => STATUS_RANK[a] - STATUS_RANK[b])[0] : null;
      const cell = el("div", "heat-cell"); cell.style.cssText = "min-height:64px;text-align:left;cursor:" + (items.length ? "pointer" : "default");
      cell.innerHTML = '<div style="font-weight:700;font-size:12px">' + d + '</div>' + (items.length ? '<div class="status-badge s-' + worst + '" style="margin-top:4px">' + items.length + ' due</div>' : "");
      if (items.length) cell.onclick = () => openListModal(first.toLocaleDateString("en-US", { month: "short" }) + " " + d + ", " + y, items);
      grid.appendChild(cell);
    }
    c.appendChild(grid);
    head.querySelector("#calPrev").onclick = () => { state.calM--; if (state.calM < 0) { state.calM = 11; state.calY--; } renderContent(); };
    head.querySelector("#calNext").onclick = () => { state.calM++; if (state.calM > 11) { state.calM = 0; state.calY++; } renderContent(); };
    head.querySelector("#calIcs").onclick = () => exportAllICS(arr);
  }
  function exportAllICS(arr) {
    const items = arr.filter(i => i.expires && computeStatus(i).key !== "expired");
    // Skip items with no usable date instead of throwing partway through the export.
    const ev = items.map(it => { const st = computeStatus(it); const when = st.startBy || st.expDate; if (!when || isNaN(when)) return null; const dt = when.toISOString().slice(0, 10).replace(/-/g, ""); return ["BEGIN:VEVENT", "UID:" + it.id + "@sentinel", "DTSTART;VALUE=DATE:" + dt, "DTEND;VALUE=DATE:" + dt, "SUMMARY:Renew - " + (it.category + " (" + it.entity + ")").replace(/[\r\n,]/g, " "), "BEGIN:VALARM", "TRIGGER:-P7D", "ACTION:DISPLAY", "DESCRIPTION:Sentinel reminder", "END:VALARM", "END:VEVENT"].join("\r\n"); });
    const events = ev.filter(Boolean);
    download("sentinel-calendar.ics", ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Sentinel//WCGTX//EN"].concat(events).concat(["END:VCALENDAR"]).join("\r\n"), "text/calendar");
    toast(events.length + " renewals exported to calendar.");
  }

  // ================= FULL PDF REPORT =================
  function fullReport() {
    if (!unlockedItems().length) { toast("Unlock a tab first — the report only covers unlocked tabs."); return; }
    const scopes = [["provider", "Providers"], ["facility", "Facilities"], ["other", "Other / Operational"]].filter(x => UNLOCKED.has(x[0]));
    let html = '<div class="binder"><h1>WCGTX Compliance Report</h1><div class="bsub">Sentinel · generated ' + new Date().toLocaleString() + '</div>';
    const all = unlockedItems(); const s = statsFor(all);
    html += '<p style="font-size:14px"><b>' + s.total + '</b> tracked · <b style="color:#dc2626">' + s.expired + '</b> expired · <b style="color:#ea580c">' + s.critical + '</b> critical · <b style="color:#ca8a04">' + s.due + '</b> due soon · health <b>' + s.score + '</b></p>';
    scopes.forEach(([sc, label]) => {
      const items = sortItems(DATA.filter(i => i.scope === sc && i.active !== false));
      if (!items.length) return;
      html += '<h2 style="margin:18px 0 6px">' + label + ' (' + items.length + ')</h2><table class="btable"><thead><tr><th>Entity</th><th>Item</th><th>Expires</th><th>Status</th><th>Proof</th></tr></thead><tbody>' +
        items.map(it => { const st = computeStatus(it); return '<tr><td>' + esc(it.entity) + '</td><td>' + esc(it.category) + '</td><td>' + (it.expires ? fmtD(it.expires) : (it.permanent ? "Permanent" : "—")) + '</td><td>' + st.label + '</td><td>' + (it.isFile ? "on file" : "-") + '</td></tr>'; }).join("") + '</tbody></table>';
    });
    html += '</div>';
    $("#printArea").innerHTML = html;
    document.body.classList.add("printing"); window.print();
    setTimeout(() => document.body.classList.remove("printing"), 500);
  }

  // ================= ATTACH PICKER (auto-detect available docs) =================
  function openAttachPicker(it) {
    const files = ((window.SENTINEL_SEED && window.SENTINEL_SEED.entityFiles) || {})[it.entityKey] || [];
    const fileRows = files.length ? files.map((f, i) => '<div class="pal-item" data-i="' + i + '"><div class="ico s-good" style="width:30px;height:30px"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' + ICONS.doc + '</svg></div><div style="flex:1"><div class="pal-cat" style="font-size:13px">' + esc(f.name) + '</div></div></div>').join("")
      : '<div class="item-sub">No documents detected in this entity’s folder.</div>';
    openModal("Attach proof document — " + esc(it.entity), '<div class="item-sub" style="margin-bottom:12px">Pick the file that proves <b>' + esc(it.category) + '</b>, or paste a path below.</div>' + fileRows +
      '<div class="drawer-actions" style="border:none;padding:14px 0 0"><input id="attachPath" placeholder="…or paste a relative path to a PDF" style="flex:2;padding:10px;border-radius:9px;border:1px solid var(--hair);background:var(--surface-solid);color:var(--ink)"><button class="save" id="attachManual">Attach</button></div>');
    function apply(link) { const rec = Object.assign({}, it, { fileLink: link, isFile: /\.(pdf|jpe?g|png|webp|gif|tiff?|heif?c?|docx?)$/i.test(link) }); saveItem(rec, false); logAudit("edit", rec, "attached document"); drawerItem = rec; closeModal(); renderDrawerView(rec); render(); toast("Document attached."); }
    [...$("#modalInner").querySelectorAll(".pal-item")].forEach(r => r.onclick = () => apply(files[+r.dataset.i].link));
    $("#attachManual").onclick = () => { const v = $("#attachPath").value.trim(); if (v) apply(v); };
  }

  // ================= BACKEND (local service) ACTIONS =================
  function svc(path, okMsg, confirmMsg) {
    if (confirmMsg && !confirm(confirmMsg)) return;
    toast("Working…");
    fetch("http://localhost:8765" + path, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ items: DATA }) })
      .then(r => r.json()).then(d => toast(d.ok ? okMsg : ("Failed: " + (d.message || "").slice(-140))))
      .catch(() => toast("Local service not running — open via Start-Sentinel.bat."));
  }
  function exportXlsx() { svc("/api/export-xlsx", "✓ Excel written: sentinel-export.xlsx (in the dashboard folder)."); }
  function emailProviders() { svc("/api/email-providers", "✓ Reminder emails sent to providers.", "Email each provider their own expiring credentials now? Only providers with an email on file and items due within 90 days will be contacted."); }

  // ================= MANAGE ACCESS (Microsoft accounts + per-tab permissions) =================
  const TAB_DEFS = [["provider", "Provider"], ["staff", "Staff"], ["facility", "Facility"], ["other", "Other"]];
  function openAccessPanel() {
    openModal("Manage access", '<div class="item-sub">Loading…</div>');
    let me = "";
    const ip = "padding:9px;border-radius:8px;border:1px solid var(--hair);background:var(--surface-solid);color:var(--ink)";
    const tabBoxes = (id, tabs) => TAB_DEFS.map(([k, lab]) => '<label style="font-size:12px;display:inline-flex;align-items:center;gap:4px;margin-right:8px"><input type="checkbox" class="tb-' + id + '" data-t="' + k + '"' + (tabs.indexOf(k) >= 0 ? " checked" : "") + '>' + lab + '</label>').join("");
    const draw = (users) => {
      const rows = users.map((u, i) => {
        const isMe = u.email === (me || "").toLowerCase();
        return '<div style="border:1px solid var(--hair);border-radius:10px;padding:10px 12px;margin-bottom:8px">' +
          '<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px"><b style="flex:1">' + esc(u.email) + '</b>' +
          (u.admin ? '<span class="ac renew">admin</span>' : '') + (isMe ? '<span class="item-sub">(you)</span>' : '') +
          (u.admin ? '' : '<button class="icon-btn arem" data-e="' + esc(u.email) + '">Remove</button>') + '</div>' +
          '<div>' + tabBoxes(i, u.tabs) +
          '<label style="font-size:12px;display:inline-flex;align-items:center;gap:4px;margin-right:8px;border-left:1px solid var(--hair);padding-left:8px" title="Admins can add, edit and delete. Unchecked = view only."><input type="checkbox" class="adm-' + i + '"' + (u.admin ? " checked" : "") + (isMe ? " disabled" : "") + '> Admin (can edit)</label>' +
          '<button class="icon-btn asave" data-e="' + esc(u.email) + '" data-i="' + i + '" style="margin-left:6px">Save</button></div></div>';
      }).join("");
      openModal("Manage access — accounts &amp; tabs",
        '<div class="item-sub" style="margin-bottom:10px">Each Microsoft account below can sign in and see <b>only the ticked tabs</b>. Add a coworker’s <b>@wcgtx.com</b> email and choose their tabs.</div>' +
        rows +
        '<div style="border-top:1px solid var(--hair);margin-top:10px;padding-top:10px"><div class="dl">Add someone</div>' +
        '<div style="display:flex;gap:6px;margin:6px 0"><input id="aEmail" type="email" placeholder="name@wcgtx.com" style="flex:1;' + ip + '"></div>' +
        '<div id="newTabs">' + TAB_DEFS.map(([k, lab]) => '<label style="font-size:12px;display:inline-flex;align-items:center;gap:4px;margin-right:8px"><input type="checkbox" class="tb-new" data-t="' + k + '" checked>' + lab + '</label>').join("") +
        '<label style="font-size:12px;display:inline-flex;align-items:center;gap:4px;margin-right:8px;border-left:1px solid var(--hair);padding-left:8px" title="Admins can add, edit and delete. Leave unchecked for view-only access."><input type="checkbox" id="aAdmin"> Admin (can edit)</label></div>' +
        '<button class="btn-primary" id="aAdd" style="max-width:160px;margin-top:8px">Add account</button></div>' +
        '<div class="auth-msg" id="aMsg" style="min-height:16px;margin-top:6px"></div>');
      [...$("#modalInner").querySelectorAll(".arem")].forEach(b => b.onclick = () => post({ action: "remove", email: b.dataset.e }));
      [...$("#modalInner").querySelectorAll(".asave")].forEach(b => b.onclick = () => {
        const tabs = [...$("#modalInner").querySelectorAll(".tb-" + b.dataset.i + ":checked")].map(c => c.dataset.t);
        const admBox = $("#modalInner").querySelector(".adm-" + b.dataset.i);
        post({ action: "save", email: b.dataset.e, tabs: tabs, admin: admBox ? admBox.checked : false });
      });
      $("#aAdd").onclick = () => {
        const v = $("#aEmail").value.trim(); if (!v) { $("#aMsg").textContent = "Enter an email."; return; }
        const tabs = [...$("#modalInner").querySelectorAll(".tb-new:checked")].map(c => c.dataset.t);
        post({ action: "save", email: v, tabs: tabs, admin: $("#aAdmin") ? $("#aAdmin").checked : false });
      };
    };
    const post = (body) => {
      $("#aMsg") && ($("#aMsg").textContent = "Saving…");
      fetch("/api/allowed", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) })
        .then(r => r.json()).then(d => { if (d.users) { toast("Access updated."); draw(d.users); } else toast(d.error || "Failed"); })
        .catch(() => toast("Couldn’t reach the server."));
    };
    fetch("/api/allowed").then(r => r.json()).then(d => {
      if (d.error) { openModal("Manage access", '<div class="empty"><h3>Admins only</h3><p>' + esc(d.error) + '</p></div>'); return; }
      me = d.me || ""; draw(d.users || []);
    }).catch(() => openModal("Manage access", '<div class="empty"><h3>Unavailable</h3><p>This is only available on the signed-in cloud app.</p></div>'));
  }

  // ================= STAFF LOGINS / ROLES =================
  function lsUsers() { try { return JSON.parse(localStorage.getItem(USERS_KEY) || "[]"); } catch (e) { return []; } }
  function openUsers() {
    if (!CURRENT_USER || CURRENT_USER.role !== "admin") { openModal("Staff logins", LOCK_MSG); return; }
    const users = lsUsers();
    const list = users.length ? users.map((u, i) => '<div class="audit-row"><span class="ac ' + (u.readonly ? "edit" : "renew") + '">' + (u.readonly ? "view-only" : "staff") + '</span><div style="flex:1"><b>' + esc(u.label || "") + '</b> — tabs: ' + esc((u.tabs || ["all"]).join(", ")) + '</div><button class="icon-btn udel" data-i="' + i + '">Remove</button></div>').join("")
      : '<div class="item-sub">No extra logins yet. The main login (iaijaz) is the administrator.</div>';
    const ip = "padding:9px;border-radius:8px;border:1px solid var(--hair);background:var(--surface-solid);color:var(--ink)";
    const form = '<h3 style="margin:20px 0 10px;font-size:13px">Add a staff login</h3><div style="display:grid;gap:8px">' +
      '<input id="uLabel" placeholder="Name / label (e.g. Frisco DON, or Inspector)" style="' + ip + '">' +
      '<input id="uId" placeholder="Login ID" style="' + ip + '"><input id="uPw" type="password" placeholder="Password" style="' + ip + '">' +
      '<div style="display:flex;gap:14px;flex-wrap:wrap">' + ["provider", "facility", "other"].map(t => '<label class="toggle-pill"><input type="checkbox" class="uTab" value="' + t + '" checked> ' + t + '</label>').join("") + '</div>' +
      '<label class="toggle-pill"><input type="checkbox" id="uRO"> Read-only (can view, cannot edit) — use this for a share/inspector login</label>' +
      '<button class="save" id="uAdd" style="padding:10px;border-radius:9px;border:none;color:var(--accent-ink);background:var(--accent);font-weight:700">Add login</button></div>' +
      '<div class="item-sub" style="margin-top:10px">Stored in this browser (hashes only). To use these on other PCs, click “Download config.js” and replace the file.</div>' +
      '<button class="icon-btn" id="uDownload" style="margin-top:8px">Download config.js</button>';
    openModal("Staff logins & roles", list + form);
    [...$("#modalInner").querySelectorAll(".udel")].forEach(b => b.onclick = () => { const a = lsUsers(); a.splice(+b.dataset.i, 1); localStorage.setItem(USERS_KEY, JSON.stringify(a)); openUsers(); toast("Login removed."); });
    $("#uAdd").onclick = () => {
      const id = $("#uId").value.trim(), pw = $("#uPw").value, label = $("#uLabel").value.trim();
      if (!id || !pw) { toast("Login ID and password are required."); return; }
      const tabs = [...$("#modalInner").querySelectorAll(".uTab:checked")].map(c => c.value);
      const a = lsUsers(); a.push({ label: label || id, idHash: hash(id), pwHash: hash(pw), tabs: tabs, readonly: $("#uRO").checked });
      localStorage.setItem(USERS_KEY, JSON.stringify(a)); openUsers(); toast("Login added: " + (label || id));
    };
    $("#uDownload").onclick = () => {
      const cfg = { configured: true, loginIdHash: CFG.loginIdHash, loginPwHash: CFG.loginPwHash, tabHashes: CFG.tabHashes, users: lsUsers() };
      download("config.js", "/* Sentinel config with staff logins (hashes only). */\nwindow.SENTINEL_CONFIG = " + JSON.stringify(cfg, null, 2) + ";\n", "text/javascript");
      toast("config.js downloaded — replace the one in the folder.");
    };
  }

  // ================= QR CODES =================
  // Provider self-service portal: one QR/link showing the provider's whole checklist + per-item upload.
  function openProviderPortal(ekey, name) {
    const origin = (CLOUD ? location.origin : "https://sentinel-compliance-kappa.vercel.app");
    const url = origin + "/provider.html?e=" + encodeURIComponent(ekey);
    const qr = "https://api.qrserver.com/v1/create-qr-code/?size=240x240&data=" + encodeURIComponent(url);
    openModal("Provider portal — " + esc(name),
      '<div style="text-align:center"><img src="' + qr + '" alt="QR code" style="width:240px;height:240px;border-radius:12px;border:1px solid var(--hair);background:#fff;padding:8px">' +
      '<div class="item-sub" style="margin-top:12px;max-width:380px;margin-left:auto;margin-right:auto">Share this QR or link with <b>' + esc(name) + '</b>. They’ll see which documents are on file vs. needed, and can upload each one — no login required.</div>' +
      '<div class="item-sub" style="margin-top:10px;word-break:break-all;opacity:.85"><a href="' + url + '" target="_blank" rel="noopener">' + esc(url) + '</a></div>' +
      '<div style="margin-top:10px"><button class="icon-btn" id="ppCopy">Copy link</button></div></div>');
    const c = $("#ppCopy"); if (c) c.onclick = () => { try { navigator.clipboard.writeText(url); toast("Link copied."); } catch (e) { toast("Copy not available — select the link."); } };
  }
  function wireQR(it) { const b = $("#dQR"); if (b) b.onclick = () => openQR(it); }
  function openQR(it) {
    if (CLOUD) {
      const url = location.origin + "/upload.html?item=" + encodeURIComponent(it.id);
      const qr = "https://api.qrserver.com/v1/create-qr-code/?size=240x240&data=" + encodeURIComponent(url);
      openModal("Scan to upload a document",
        '<div style="text-align:center"><img src="' + qr + '" alt="QR code" style="width:240px;height:240px;border-radius:12px;border:1px solid var(--hair);background:#fff;padding:8px">' +
        '<div style="margin-top:12px;font-weight:700">' + esc(it.category) + '</div><div class="item-sub">' + esc(it.entity) + '</div>' +
        '<div class="item-sub" style="margin-top:12px;max-width:340px;margin:12px auto 0">Scan from <b>any</b> phone, anywhere — or use the link below if you can’t scan. Choose or photograph the document; it’s stored securely and attaches here.</div>' +
        '<div class="item-sub" style="margin-top:10px"><b>Or open this link:</b></div>' +
        '<div class="item-sub" style="margin-top:4px;word-break:break-all"><a href="' + url + '" target="_blank" rel="noopener">' + esc(url) + '</a></div></div>');
      return;
    }
    toast("Building QR…");
    fetch("http://localhost:8765/api/info").then(r => r.json()).then(info => {
      const base = info.base || "http://localhost:8765";
      const uploadURL = base + "/_Sentinel_Compliance/upload.html?item=" + encodeURIComponent(it.id);
      const qr = "http://localhost:8765/api/qr?data=" + encodeURIComponent(uploadURL);
      openModal("Scan to upload a document",
        '<div style="text-align:center"><img src="' + qr + '" alt="QR code" style="width:240px;height:240px;border-radius:12px;border:1px solid var(--hair);background:#fff;padding:8px">' +
        '<div style="margin-top:12px;font-weight:700">' + esc(it.category) + '</div><div class="item-sub">' + esc(it.entity) + '</div>' +
        '<div class="item-sub" style="margin-top:12px;max-width:340px;margin-left:auto;margin-right:auto">Scan with your phone, then choose or photograph the document. It saves straight into this item’s folder and attaches here.</div>' +
        '<div class="item-sub" style="margin-top:10px"><b>Or open this link:</b></div>' +
        '<div class="item-sub" style="margin-top:4px;word-break:break-all"><a href="' + uploadURL + '" target="_blank" rel="noopener">' + esc(uploadURL) + '</a></div>' +
        '<div class="item-sub" style="margin-top:8px">Phone must be on the same Wi-Fi as this PC. If it won’t connect, allow Python through the Windows firewall.</div></div>');
    }).catch(() => openModal("QR code", '<div class="empty"><h3>Start the service</h3><p>Open Sentinel via <b>Start-Sentinel.bat</b> so the QR can be generated and your phone can reach the upload page.</p></div>'));
  }
  function applyUploads(u) {
    if (!u) return;
    // /api/uploads-map answers HTTP 200 with { ok:false, message } when Graph errors. Treating
    // that error object as the attachment map made EVERY document on the board disappear (no
    // item id matches, supplemental empties) — and the sync still toasted "Documents updated."
    // A failed read must change nothing at all.
    if (u.ok === false) return;
    // /api/uploads-map now returns { attachments, supplemental }. Old shape (a flat map
    // keyed by item id) is still accepted for the local Python relay.
    ATTACHMENTS = (u && u.attachments) ? u.attachments : u;
    // Replace LIVE_SUPP with the freshly received supplementals (handles adds + deletes).
    LIVE_SUPP = (u && Array.isArray(u.supplemental)) ? u.supplemental : [];
    buildData();          // rebuilds DATA and re-applies ATTACHMENTS via attachUploads()
  }
  function handleDeepLink() {
    const m = /[#&]item=([^&]+)/.exec(location.hash); if (!m) return;
    const id = decodeURIComponent(m[1]); const it = DATA.find(x => x.id === id);
    try { history.replaceState(null, "", location.href.split("#")[0]); } catch (e) {}
    if (!it) return;
    if (CURRENT_USER && CURRENT_USER.tabs.indexOf(it.scope) < 0) return;
    const go = () => { state.tab = it.scope; render(); openDrawer(it, false); };
    if (UNLOCKED.has(it.scope)) go(); else promptTabCode(it.scope, go);
  }

  // ================= READ DATE FROM PDF =================
  function wireReadDate(it) {
    const b = $("#dReadDate"); if (!b) return;
    if (READONLY || CLOUD) { b.style.display = "none"; return; }
    b.onclick = () => {
      toast("Reading the document…");
      fetch("http://localhost:8765/api/read-dates", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ file: it.fileLink }) })
        .then(r => r.json())
        .then(d => { if (!d.ok || !d.dates || !d.dates.length) { toast("No dates found in that document."); return; } pickDate(it, d.dates); })
        .catch(() => toast("Service not running — open via Start-Sentinel.bat."));
    };
  }
  function pickDate(it, dates) {
    const body = '<div class="item-sub" style="margin-bottom:10px">Dates found in <b>' + esc(it.fileLink.split("/").pop()) + '</b>. Pick the one to set as the expiry date:</div>' +
      dates.map(d => '<div class="pal-item" data-d="' + esc(d) + '"><div class="pal-cat" style="flex:1">' + fmtD(d) + '</div><div class="item-sub">' + esc(d) + '</div></div>').join("");
    openModal("Set expiry from document", body);
    [...$("#modalInner").querySelectorAll(".pal-item")].forEach(r => r.onclick = () => {
      const rec = Object.assign({}, it, { expires: r.dataset.d }); saveItem(rec, false); logAudit("edit", rec, "expiry read from PDF");
      drawerItem = rec; closeModal(); renderDrawerView(rec); render(); toast("Expiry set to " + fmtD(r.dataset.d));
    });
  }

  // ================= ASK-A-QUESTION (smart query) =================
  const MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];
  function smartQuery(q, pool) {
    q = q.toLowerCase(); let used = false, arr = pool.slice();
    const has = w => q.indexOf(w) >= 0;
    if (has("provider")) { arr = arr.filter(i => i.scope === "provider"); used = true; }
    if (has("facilit")) { arr = arr.filter(i => i.scope === "facility"); used = true; }
    if (has("expired")) { arr = arr.filter(i => computeStatus(i).key === "expired"); used = true; }
    else if (has("critical")) { arr = arr.filter(i => computeStatus(i).key === "critical"); used = true; }
    else if (has("due") || has("expiring") || has("soon")) { arr = arr.filter(i => ["due", "critical", "expired"].indexOf(computeStatus(i).key) >= 0); used = true; }
    if (has("missing") && (has("proof") || has("document") || has("doc"))) { arr = arr.filter(i => (i.isFile === false && !i.centralProof)); used = true; }
    // month
    const mi = MONTHS.findIndex(m => has(m.slice(0, 4)) && (has(m) || true) && new RegExp(m.slice(0, 3)).test(q));
    const monthHit = MONTHS.findIndex(m => has(m));
    if (monthHit >= 0) { const now = today(); let y = now.getFullYear(); if (monthHit < now.getMonth()) y++; arr = arr.filter(i => { const e = parseD(i.expires); return e && e.getMonth() === monthHit && e.getFullYear() === y; }); used = true; }
    // next N days
    const nd = /next\s+(\d+)\s*day/.exec(q); if (nd) { const n = +nd[1]; arr = arr.filter(i => { const s = computeStatus(i); return s.days != null && s.days >= 0 && s.days <= n; }); used = true; }
    if (has("this month")) { const now = today(); arr = arr.filter(i => { const e = parseD(i.expires); return e && e.getMonth() === now.getMonth() && e.getFullYear() === now.getFullYear(); }); used = true; }
    // category keyword (dea, license, cme, bls, acls, clia, cola, board, oig, npdb, tsca, flu, tb, malpractice, pharmacy, x-ray)
    const cats = ["dea", "license", "cme", "bls", "acls", "pals", "atls", "board", "oig", "npdb", "tsca", "flu", "tb", "malpractice", "pharmacy", "x-ray", "clia", "cola", "occupancy", "npi"];
    const catHit = cats.find(c => has(c));
    if (catHit) {
      if (/\b(no|without|missing)\b/.test(q) && !has("proof")) {
        // entities lacking this category entirely
        const haveCat = new Set(pool.filter(i => i.category.toLowerCase().indexOf(catHit) >= 0).map(i => i.entityKey));
        const ents = [...new Set(pool.map(i => i.entityKey))].filter(k => !haveCat.has(k));
        arr = pool.filter(i => ents.indexOf(i.entityKey) >= 0);
        // collapse to one row per entity
        const seen = {}; arr = arr.filter(i => seen[i.entityKey] ? false : (seen[i.entityKey] = true));
        used = true;
      } else { arr = arr.filter(i => i.category.toLowerCase().indexOf(catHit) >= 0); used = true; }
    }
    return used ? arr : null;
  }

  // ---------- boot ----------
  if ("serviceWorker" in navigator && location.protocol.indexOf("http") === 0) {
    try { navigator.serviceWorker.register("service-worker.js").catch(() => {}); } catch (e) {}
  }
  document.addEventListener("DOMContentLoaded", bootSentinel);
  if (document.readyState !== "loading") bootSentinel();
})();
