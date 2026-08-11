// Microsoft redirects here after staff sign-in. We exchange the code, read the
// id_token (came straight from Microsoft over TLS, so trusted), verify the user is
// in OUR tenant, then set a signed session cookie and send them to the dashboard.
const { sign, cookieHeader } = require("../lib/session");
const { getUsers, findUser } = require("../lib/access");

function esc(s) { return String(s).replace(/[&<>]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c])); }
function page(title, body) {
  return `<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<body style="font-family:system-ui,Segoe UI,Arial;max-width:520px;margin:48px auto;padding:0 18px;color:#0f172a">
<h2 style="color:#0f766e">${title}</h2>${body}<p><a href="/api/auth-callback">Try sign-in again</a></p></body>`;
}

module.exports = async (req, res) => {
  const url = new URL(req.url, "http://localhost");
  const code = url.searchParams.get("code");
  const err = url.searchParams.get("error_description") || url.searchParams.get("error");
  const tenant = process.env.MS_TENANT_ID, cid = process.env.MS_CLIENT_ID, sec = process.env.MS_CLIENT_SECRET;
  // Derive the redirect from the request host so this works on ANY domain (no hardcoding).
  const redirect = "https://" + (req.headers.host || "sentinel-compliance-kappa.vercel.app") + "/api/auth-callback";

  // No code and no error => this is the START of sign-in: redirect to Microsoft.
  if (!code && !err) {
    if (!tenant || !cid) { res.status(500).send("Auth not configured."); return; }
    const p = new URLSearchParams({ client_id: cid, response_type: "code", redirect_uri: redirect, response_mode: "query", scope: "openid profile email User.Read", prompt: "select_account" });
    res.writeHead(302, { Location: "https://login.microsoftonline.com/" + tenant + "/oauth2/v2.0/authorize?" + p.toString() });
    res.end(); return;
  }
  // esc() the error — it comes from the query string and was being written into the page raw,
  // which is reflected XSS on the app's own origin (a crafted link could run script as the
  // signed-in user and call every authenticated endpoint as them).
  if (err || !code) { res.setHeader("Content-Type", "text/html"); res.status(200).send(page("Sign-in didn’t complete", "<p>" + esc(err || "No authorization code returned.") + "</p>")); return; }

  try {
    const r = await fetch("https://login.microsoftonline.com/" + tenant + "/oauth2/v2.0/token", {
      method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ client_id: cid, client_secret: sec, grant_type: "authorization_code", code, redirect_uri: redirect, scope: "openid profile email User.Read" }).toString()
    });
    const j = await r.json();
    if (!j.id_token) throw new Error(j.error_description || "no id_token");
    const claims = JSON.parse(Buffer.from(j.id_token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/"), "base64").toString());
    if (claims.tid !== tenant) { res.setHeader("Content-Type", "text/html"); res.status(403).send(page("Access denied", "<p>That account isn’t part of Wellness &amp; Care Group of Texas.</p>")); return; }
    const email = claims.preferred_username || claims.email || claims.upn || "staff";
    // Allowlist with per-user tab permissions.
    const u = findUser(email, await getUsers());
    if (!u) {
      res.setHeader("Content-Type", "text/html");
      res.status(403).send(page("Not authorized", "<p><b>" + esc(email) + "</b> isn’t on the access list for this dashboard.</p><p>Ask the administrator (Imad) to add you, then try again.</p>"));
      return;
    }
    const token = sign({ email, name: claims.name || email, tabs: u.tabs, admin: !!u.admin, exp: Date.now() + 1000 * 60 * 60 * 12 }); // 12 hours
    res.setHeader("Set-Cookie", cookieHeader(token, 60 * 60 * 12));
    res.writeHead(302, { Location: "/" });
    res.end();
  } catch (e) {
    res.setHeader("Content-Type", "text/html");
    res.status(200).send(page("Sign-in problem", "<p>" + esc(String(e.message || e)) + "</p>"));
  }
};
