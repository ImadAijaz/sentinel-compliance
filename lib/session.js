// Signed session cookie for Microsoft-authenticated staff.
// HMAC-signed with MS_CLIENT_SECRET (already a server-only secret) — no new env var.
// Cookie is httpOnly + Secure, so client JS / F12 cannot read or forge it.
const crypto = require("crypto");
// NO hardcoded fallback. There used to be a literal "dev-only" default, which meant any
// deployment missing the env var would happily verify a session token that anyone could mint
// offline — including one with admin:true. Missing config must fail CLOSED, never open.
const SECRET = process.env.SESSION_SECRET || process.env.MS_CLIENT_SECRET || "";
const COOKIE = "sentinel_session";

function b64url(buf) { return Buffer.from(buf).toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, ""); }
function unb64url(s) { return Buffer.from(s.replace(/-/g, "+").replace(/_/g, "/"), "base64").toString(); }

function sign(obj) {
  if (!SECRET) throw new Error("SESSION_SECRET (or MS_CLIENT_SECRET) is not set — refusing to issue a session.");
  const p = b64url(JSON.stringify(obj));
  const sig = b64url(crypto.createHmac("sha256", SECRET).update(p).digest());
  return p + "." + sig;
}
function verify(token) {
  if (!SECRET) return null;                      // no key configured -> nobody is authenticated
  if (!token || token.indexOf(".") < 0) return null;
  const parts = token.split(".");
  if (parts.length !== 2) return null;
  const [p, sig] = parts;
  const expect = b64url(crypto.createHmac("sha256", SECRET).update(p).digest());
  // Compare as BYTES. The old length check compared string length while timingSafeEqual
  // compares byte length, so a multi-byte cookie threw RangeError and 500'd the route.
  const a = Buffer.from(sig), b = Buffer.from(expect);
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return null;
  try {
    const o = JSON.parse(unb64url(p));
    // exp is REQUIRED. It used to be checked only "if present", so a forged or legacy token
    // without one never expired.
    if (!o || typeof o.exp !== "number" || Date.now() > o.exp) return null;
    return o;
  } catch (e) { return null; }
}
function parseCookies(req) {
  const h = req.headers.cookie || ""; const o = {};
  // decodeURIComponent throws URIError on a malformed value (e.g. "%zz") — a bad cookie must
  // read as "not signed in", never crash the route with a 500.
  h.split(";").forEach(c => {
    const i = c.indexOf("=");
    if (i <= 0) return;
    const raw = c.slice(i + 1).trim();
    let val; try { val = decodeURIComponent(raw); } catch (e) { val = raw; }
    o[c.slice(0, i).trim()] = val;
  });
  return o;
}
function getSession(req) { return verify(parseCookies(req)[COOKIE]); }
function cookieHeader(token, maxAgeSec) {
  return COOKIE + "=" + token + "; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=" + maxAgeSec;
}
function clearHeader() { return COOKIE + "=; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=0"; }

// Is this a genuine Vercel cron invocation?
// The old test was /vercel-cron/i.test(User-Agent) — a client-supplied header, so anyone could
// `curl -A vercel-cron` and be treated as the scheduler. Vercel sends
// `Authorization: Bearer $CRON_SECRET` instead; compare that in constant time.
// Fails CLOSED: with no CRON_SECRET set, nothing is ever accepted as the cron.
function isCronRequest(req) {
  const secret = process.env.CRON_SECRET || "";
  if (!secret) return false;
  const got = String((req.headers && req.headers.authorization) || "").replace(/^Bearer\s+/i, "");
  if (!got) return false;
  const a = Buffer.from(got), b = Buffer.from(secret);
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}

module.exports = { sign, verify, getSession, cookieHeader, clearHeader, isCronRequest, COOKIE };
