/**
 * GET /api/internal?key=<passphrase>
 *
 * Read-only proxy for the internal dashboard (site/internal). Checks the
 * passphrase against INTERNAL_KEY, then calls the Apps Script doGet with
 * the server's own copy of the key and returns its JSON snapshot of
 * Renters, Applications, and Vendor Leads. No writes, ever.
 *
 * Vercel environment variables required:
 *   APPS_SCRIPT_URL   the Apps Script web app /exec URL (same as api/apply.js)
 *   INTERNAL_KEY      the dashboard passphrase; must match INTERNAL_KEY in Code.gs
 */

const TIMEOUT_MS = 25000;

module.exports = async function handler(req, res) {
  if (req.method !== 'GET') {
    res.setHeader('Allow', 'GET');
    return res.status(405).json({ ok: false, error: 'method_not_allowed' });
  }

  const url = process.env.APPS_SCRIPT_URL;
  const key = process.env.INTERNAL_KEY;
  if (!url || !key) {
    return res.status(500).json({ ok: false, error: 'not_configured' });
  }

  const given = String((req.query && req.query.key) || '');
  // Reject on length first so a wrong-length guess never reaches the compare.
  if (given.length !== key.length || given !== key) {
    return res.status(401).json({ ok: false, error: 'unauthorized' });
  }

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);

  try {
    const target = url + (url.indexOf('?') > -1 ? '&' : '?') + 'key=' + encodeURIComponent(key);
    const upstream = await fetch(target, { method: 'GET', redirect: 'follow', signal: controller.signal });

    const text = await upstream.text();
    let parsed;
    try { parsed = JSON.parse(text); }
    catch (err) {
      return res.status(502).json({ ok: false, error: 'bad_upstream_response' });
    }

    if (!parsed.ok) return res.status(502).json(parsed);
    res.setHeader('Cache-Control', 'no-store');
    return res.status(200).json(parsed);
  } catch (err) {
    const aborted = err && err.name === 'AbortError';
    return res.status(504).json({ ok: false, error: aborted ? 'upstream_timeout' : 'upstream_unreachable' });
  } finally {
    clearTimeout(timer);
  }
};
