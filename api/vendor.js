/**
 * POST /api/vendor
 *
 * Supplier / service-partner intake. Same shape as api/apply.js: sits
 * between the public form and Apps Script so the Apps Script URL and its
 * shared token never appear in the page source. Forwards the record to
 * Apps Script, which appends it to the "Vendor Leads" tab and emails the
 * office. Not a renter application, so no documents and no attachments.
 *
 * Vercel environment variables required (shared with api/apply.js):
 *   APPS_SCRIPT_URL    the Apps Script web app /exec URL
 *   APPS_SCRIPT_TOKEN  must match SHARED_TOKEN in Code.gs
 */

const TIMEOUT_MS = 25000;

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ ok: false, error: 'method_not_allowed' });
  }

  const url = process.env.APPS_SCRIPT_URL;
  const token = process.env.APPS_SCRIPT_TOKEN;
  if (!url || !token) {
    return res.status(500).json({ ok: false, error: 'not_configured' });
  }

  let record = req.body;
  if (typeof record === 'string') {
    try { record = JSON.parse(record); }
    catch (err) { return res.status(400).json({ ok: false, error: 'bad_json' }); }
  }
  if (!record || !record.submissionId) {
    return res.status(400).json({ ok: false, error: 'bad_payload' });
  }

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);

  try {
    // Apps Script does not answer CORS preflight; text/plain keeps this a
    // simple request. The body is still JSON and is parsed as such in doPost.
    const upstream = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ token, vendor: record }),
      redirect: 'follow',
      signal: controller.signal
    });

    const text = await upstream.text();
    let parsed;
    try { parsed = JSON.parse(text); }
    catch (err) {
      return res.status(502).json({ ok: false, error: 'bad_upstream_response' });
    }

    if (!parsed.ok) return res.status(502).json(parsed);
    return res.status(200).json(parsed);
  } catch (err) {
    const aborted = err && err.name === 'AbortError';
    return res.status(504).json({ ok: false, error: aborted ? 'upstream_timeout' : 'upstream_unreachable' });
  } finally {
    clearTimeout(timer);
  }
};
