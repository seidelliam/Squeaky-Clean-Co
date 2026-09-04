# Squeaky Clean Co — form → Sheet + email

When someone submits the rental application, two things happen at once:
a row is appended to your Google Sheet, and an email lands in the office
inbox with the full record as JSON plus the applicant's documents attached.

```
browser ──► /api/apply (Vercel) ──► Apps Script ──┬──► append row to Sheet
                                                  └──► email + attachments
```

The Apps Script URL and token live in Vercel environment variables, so they
never appear in the page source.

---

## 1. The Sheet

1. Create a Google Sheet. Name the tab **Applications**.
2. Copy its id — the long string in the URL between `/d/` and `/edit`.

Add your header row in row 1. Column **order does not matter** and neither do
extra columns you add later: the script matches values to headers by name.

If you want the default set, you can run `writeDefaultHeaders()` from the
script editor once and it will write them for you.

## 2. The Apps Script

1. In the Sheet: **Extensions → Apps Script**.
2. Replace the contents of `Code.gs` with this repo's `apps-script/Code.gs`.
3. Fill in the CONFIG block at the top:
   - `SHEET_ID` — from step 1
   - `SHARED_TOKEN` — any long random string; you'll paste the same one into Vercel
4. Run `testEndToEnd()` once. Google will prompt for authorization — approve
   it. Check that a test row appeared and a test email arrived.
5. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Copy the `/exec` URL.

> Re-deploy after any edit to `Code.gs`, or the live endpoint keeps running
> the old code. Use **Manage deployments → edit → Version: New version** to
> keep the same URL.

## 3. Vercel

1. Import this repo at vercel.com. No build step — it's a static site with
   one serverless function.
2. **Settings → Environment Variables**:

   | Name | Value |
   |---|---|
   | `APPS_SCRIPT_URL` | the `/exec` URL from step 2 |
   | `APPS_SCRIPT_TOKEN` | the same string as `SHARED_TOKEN` |

3. Redeploy so the variables take effect.

## 4. Verify

Submit a real application through the live site. You should get a row, an
email, and "Application sent" on the success screen.

---

## How it fails

The form never loses data on a bad day. If the endpoint is missing, times
out, or Apps Script errors, `sendToServer()` resolves false and the page
falls back to its original behavior — the applicant saves the JSON file and
emails it in. They see instructions, not an error.

If the Sheet write fails but email still works, the email arrives flagged
with the error at the top so you can add the row by hand.

## Limits worth knowing

- **Vercel request body: 4.5 MB.** Images are compressed client-side to
  ~170 KB each before attaching, so a full application runs well under 1 MB.
- **Gmail attachment total: 25 MB**, far above what this sends.
- **Apps Script email quota: 100 recipients/day** on a consumer Gmail
  account, 1,500/day on Workspace.

## Changing the sheet format later

Add, remove, or reorder columns freely — matching is by header name.
If a header's text differs from the payload's key, add one line to the
`ALIASES` map in `Code.gs` instead of touching any other code.
