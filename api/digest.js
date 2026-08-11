// Digest email.
//  POST {scopes:[...]} (signed-in "Email me") -> sends to the signed-in user's own inbox,
//       limited to the tabs they're allowed (and the scopes they picked).
//  GET (Vercel Cron, daily) -> emails EACH allowed person their own tabs, to their own inbox.
// Sent via the Gmail SMTP relay. No hardcoded recipient.
const nodemailer = require("nodemailer");
const data = require("../data.json");
const { getSession } = require("../lib/session");
const { getUsers } = require("../lib/access");

function color(s) { return s === "EXPIRED" ? "#dc2626" : s === "CRITICAL" ? "#ea580c" : s === "DUE SOON" ? "#ca8a04" : "#0d9488"; }

function buildHtml(scopes) {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const rows = [];
  for (const it of (data.items || [])) {
    if (it.active === false || !it.expires) continue;
    if (!scopes.includes(it.scope)) continue;
    const e = new Date(it.expires); if (isNaN(e)) continue;
    const days = Math.floor((e - today) / 86400000);
    if (days > 90) continue;
    rows.push({ cat: it.category, entity: it.entity, exp: e, days, status: days < 0 ? "EXPIRED" : days <= 30 ? "CRITICAL" : "DUE SOON" });
  }
  rows.sort((a, b) => a.days - b.days);
  const nE = rows.filter(r => r.status === "EXPIRED").length, nC = rows.filter(r => r.status === "CRITICAL").length, nD = rows.filter(r => r.status === "DUE SOON").length;
  const trs = rows.map(r => {
    const c = color(r.status), when = r.days < 0 ? Math.abs(r.days) + "d ago" : "in " + r.days + "d";
    return `<tr><td style='padding:7px 10px;border-bottom:1px solid #eee'><b>${r.cat}</b><br><span style='color:#777;font-size:12px'>${r.entity}</span></td><td style='padding:7px 10px;border-bottom:1px solid #eee'>${r.exp.toDateString()}</td><td style='padding:7px 10px;border-bottom:1px solid #eee;color:${c};font-weight:700'>${when}</td><td style='padding:7px 10px;border-bottom:1px solid #eee'><span style='background:${c};color:#fff;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:700'>${r.status}</span></td></tr>`;
  }).join("");
  const html = `<div style='font-family:Segoe UI,Arial,sans-serif;max-width:760px;margin:auto;color:#0f172a'>
    <div style='background:linear-gradient(135deg,#14b8a6,#0f766e);color:#fff;padding:22px 26px;border-radius:14px 14px 0 0'>
      <div style='font-size:13px;letter-spacing:2px;opacity:.85'>SENTINEL · COMPLIANCE DIGEST</div>
      <div style='font-size:22px;font-weight:800;margin-top:4px'>${today.toDateString()}</div></div>
    <div style='border:1px solid #e6ebf1;border-top:none;border-radius:0 0 14px 14px;padding:22px 26px'>
      <p style='font-size:15px'><b style='color:#dc2626'>${nE} expired</b> &middot; <b style='color:#ea580c'>${nC} critical</b> &middot; <b style='color:#ca8a04'>${nD} due soon</b> <span style='color:#94a3b8'>(${scopes.join(", ")})</span></p>
      ${rows.length ? `<table style='border-collapse:collapse;width:100%;font-size:14px'><thead><tr style='text-align:left;color:#475569;font-size:12px;text-transform:uppercase'><th style='padding:7px 10px'>Item</th><th style='padding:7px 10px'>Expires</th><th style='padding:7px 10px'>Countdown</th><th style='padding:7px 10px'>Status</th></tr></thead><tbody>${trs}</tbody></table>` : "<p style='color:#059669;font-weight:700'>All clear — nothing expiring within 90 days.</p>"}
      <p style='color:#94a3b8;font-size:12px;margin-top:22px'>Sentinel compliance digest.</p></div></div>`;
  return { html, subject: `Sentinel digest — ${nE} expired, ${nC} critical, ${nD} due` };
}

module.exports = async (req, res) => {
  if (req.method === "OPTIONS") { res.status(204).end(); return; }
  // The GET path (the scheduled refresh) had NO authentication, so any anonymous caller could
  // trigger privileged Graph/Excel reads and rewrite the cached roster on demand. GET is now
  // cron-only (CRON_SECRET bearer) or admin; POST still requires a session, checked below.
  if (req.method !== "POST") {
    const { getSession: gs, isCronRequest } = require("../lib/session");
    const sess = gs(req);
    if (!isCronRequest(req) && !(sess && sess.admin)) { res.status(401).json({ ok: false, message: "cron or admin only" }); return; }
  }
  const user = process.env.GMAIL_USER, pass = process.env.GMAIL_APP_PASSWORD;
  // Mail credentials are only needed for the POST ("Email me") path. The GET path no longer sends
  // any mail — it just refreshes the roster cache — so missing Gmail config must not abort it.
  if (req.method === "POST" && (!user || !pass)) { res.status(200).json({ ok: false, message: "Set GMAIL_USER and GMAIL_APP_PASSWORD in Vercel env vars." }); return; }
  const send = (to, scopes) => {
    if (!user || !pass) throw new Error("mail not configured");
    const tx = nodemailer.createTransport({ service: "gmail", auth: { user, pass } });
    const { html, subject } = buildHtml(scopes);
    return tx.sendMail({ from: user, to, subject, html });
  };

  try {
    if (req.method === "POST") {
      const s = getSession(req);
      if (!s) { res.status(401).json({ ok: false, message: "sign-in required" }); return; }
      const allowed = (s.tabs && s.tabs.length) ? s.tabs : ["provider", "facility", "other"];
      let raw = ""; await new Promise(r => { req.on("data", c => raw += c); req.on("end", r); });
      let scopes = allowed;
      try { const b = JSON.parse(raw || "{}"); if (Array.isArray(b.scopes) && b.scopes.length) scopes = b.scopes.filter(x => allowed.includes(x)); } catch (e) {}
      if (!scopes.length) scopes = allowed;
      await send(s.email, scopes);
      res.status(200).json({ ok: true, to: s.email, scopes });
      return;
    }
    // GET = scheduled cron: email each allowed person their own tabs
    // Also: regenerate the roster delta so the dashboard picks up Excel changes within ~24h.
    // Uses the SHARED regen in lib/regen.js — this used to be a private copy of that logic and it
    // kept the old fixed-column roster reader after the header-detection fix, silently corrupting
    // the dashboard nightly. One implementation only.
    let regenInfo = null;
    try {
      const { accessToken } = require("../lib/graph");
      const { regenerateRoster } = require("../lib/regen");
      regenInfo = await regenerateRoster(await accessToken());
    } catch (e) { regenInfo = { error: String(e.message || e).slice(0, 200) }; }
    // Automatic digest emails are DISABLED. The 15-day email cadence is handled by the user's
    // own scheduled routine, not the app — so the daily cron now ONLY refreshes roster/staff
    // data silently and never emails anyone (previously it emailed every allowed user daily,
    // and twice over because the cron runs on both the delta and kappa deployments).
    res.status(200).json({ ok: true, sent: 0, emailsDisabled: true, regen: regenInfo });
  } catch (e) { res.status(200).json({ ok: false, message: String(e.message || e) }); }
};
