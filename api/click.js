/**
 * POST /api/click
 *
 * Fire-and-forget click beacon for Messenger outreach links. Logs the ref
 * code plus first-touch context to the same Apps Script backend that
 * records applications, so "clicked but never applied" is visible next to
 * submitted applications in the Sheet. Mirrors api/apply.js: same env vars,
 * same Apps Script URL and token never reach the browser.
 *
 * Vercel environment variables required (shared with api/apply.js):
 *   APPS_SCRIPT_URL    the Apps Script web app /exec URL
 *   APPS_SCRIPT_TOKEN  must match SHARED_TOKEN in Code.gs
 */

const TIMEOUT_MS = 8000;

// Facebook's link-preview crawler fetches the page itself (for the OG
// card) when a link is pasted into Messenger, and it doesn't execute the
// page's JS — so it never calls this endpoint in practice. This check is
// a defensive second layer in case anything ever forwards its request here.
const BOT_UA = /facebookexternalhit|Facebot/i;

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ ok: false, error: 'method_not_allowed' });
  }

  const ua = req.headers['user-agent'] || '';
  if (BOT_UA.test(ua)) {
    return res.status(204).end();
  }

  const url = process.env.APPS_SCRIPT_URL;
  const token = process.env.APPS_SCRIPT_TOKEN;
  if (!url || !token) {
    return res.status(500).json({ ok: false, error: 'not_configured' });
  }

  let body = req.body;
  if (typeof body === 'string') {
    try { body = JSON.parse(body); }
    catch (err) { return res.status(400).json({ ok: false, error: 'bad_json' }); }
  }
  if (!body || typeof body !== 'object') {
    return res.status(400).json({ ok: false, error: 'bad_payload' });
  }

  const click = {
    ref: String(body.ref || 'direct').slice(0, 100),
    fbclid: body.fbclid ? String(body.fbclid).slice(0, 200) : '',
    landingPath: body.landingPath ? String(body.landingPath).slice(0, 200) : '',
    ts: body.ts || new Date().toISOString(),
    ua: ua.slice(0, 300)
  };

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);

  try {
    const upstream = await fetch(url, {
      method: 'POST',
      // Apps Script does not answer CORS preflight; text/plain keeps this
      // a simple request. The body is still JSON, parsed as such in doPost.
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ token, click }),
      redirect: 'follow',
      signal: controller.signal
    });
    await upstream.text(); // drain the response; the caller doesn't need it
  } catch (err) {
    // A click beacon must never surface an error to the visitor or block
    // navigation to the form. Fail silently — this is best-effort logging.
  } finally {
    clearTimeout(timer);
  }

  return res.status(204).end();
};
