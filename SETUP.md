# Squeaky Clean Co — form → Sheet + email

When someone submits the rental application, two things happen at once: a row
is appended to the **Applications** tab of your live Google Sheet, and an email
lands in the office inbox with the full record as JSON plus the applicant's
documents attached.

```
browser ──► /api/apply (Vercel) ──► Apps Script ──┬──► Applications tab
                                                  └──► email + attachments
```

Approved applications are then promoted into **Renters** by hand, from a menu.

---

## Why applications don't write straight to Renters

Three reasons, all specific to how your workbook is built:

1. **Renters puts its headers on row 2.** Row 1 is the merged title banner.
2. **Eight Renters columns are formulas** — Vendor Name, Within Radius?, the
   three Pmts counts, Trust Score, Trust Rating, Lifetime Paid. A normal append
   writes across the whole row and would blank them.
3. **Those formulas are already filled down past the data.** An append lands
   *after* the last non-empty row, which is below where the formulas stop.

On top of that, an application has no Assigned Vendor ID and no Distance to
Vendor — you decide those. And your Dashboard counts **Active Renters**, so
unapproved leads sitting in Renters would inflate your numbers.

So the form only ever touches the Applications tab. When you approve one, the
menu action writes a Renters row using **only the input columns**, into the
first row whose Renter ID is blank, leaving every formula cell untouched.

---

## 1. Sheets setup

Your workbook already lives in Google Sheets, which is exactly what this needs
— no export, no conversion.

1. Open the workbook. Copy its id from the URL, between `/d/` and `/edit`.
2. **Extensions → Apps Script**. Delete whatever is in `Code.gs`, paste in this
   repo's [apps-script/Code.gs](apps-script/Code.gs).
3. Fill the CONFIG block at the top:
   - `SHEET_ID` — from step 1
   - `SHARED_TOKEN` — any long random string; you'll paste the same one into Vercel
4. Save, then reload the spreadsheet tab. A **Squeaky Clean** menu appears.
5. **Squeaky Clean → Set up Applications tab.** Approve the authorization prompt
   (Google will call the script unverified — it's your own script in your own
   account; click *Advanced → Go to project*).
6. **Squeaky Clean → Test intake.** Confirm a row lands in Applications and a
   test email arrives.
7. **Deploy → New deployment → Web app**, Execute as **Me**, Who has access
   **Anyone**. Copy the `/exec` URL.

Optional housekeeping: the Renters, Cities, Renter Payments and Vendor Payments
tabs still have their pink `EXAMPLE ROW — delete before use` rows. Clear those
before going live so the Dashboard totals are real.

## 2. Vercel

1. Import this repo. No build step — static site plus one function.
2. **Settings → Environment Variables:**

   | Name | Value |
   |---|---|
   | `APPS_SCRIPT_URL` | the `/exec` URL from step 7 |
   | `APPS_SCRIPT_TOKEN` | the same string as `SHARED_TOKEN` |

3. Redeploy so the variables take effect.

## 3. Day-to-day

A new application arrives → you get the email with the paystub, ID and outlet
photo attached, and a row appears in Applications with **Review Status: New**.

To approve: open the Applications tab, click any cell in that row, then
**Squeaky Clean → Promote selected application to Renters**. It creates the next
`R-###`, fills the input columns, sets Status to `Pending Delivery`, and stamps
the application row with the new Renter ID so it can't be promoted twice.

Three fields it deliberately leaves blank for you, because the form can't know
them: **Assigned Vendor ID**, **Distance to Vendor (mi)**, and it defaults
**My Cut** to $10 — change `DEFAULT_MY_CUT` in CONFIG if that's wrong.

---

## How it fails

The form never loses data. If the endpoint is missing, times out, or Apps Script
errors, the page falls back to its original behavior — the applicant saves the
JSON file and emails it in, with instructions rather than an error.

If the Sheet write fails but email still works, the email arrives flagged with
the error at the top so you can add the row by hand.

## Limits worth knowing

- **Vercel request body: 4.5 MB.** Images are compressed client-side to ~170 KB
  each before attaching, so a full application runs well under 1 MB.
- **Apps Script email: 100/day** on consumer Gmail, 1,500/day on Workspace.
- Re-deploy the Apps Script after **any** edit to `Code.gs`, or the live URL
  keeps running the old version. *Manage deployments → edit → New version*
  keeps the same URL so Vercel doesn't need updating.

## Changing columns later

Applications matches by header name, so you can reorder or add columns freely.
If you rename one, update `APPLICATION_HEADERS`. For Renters, the writable
columns are listed in `RENTER_WRITABLE` — anything not on that list is treated
as a formula and never written to.

## Messenger link attribution

Every page accepts `?ref=` in the URL and remembers it for the tab via
sessionStorage (no cookies — Messenger's in-app browser doesn't share them).
It rides along on submission as the **Ref** column in Applications, and a
click beacon logs to a new **Clicks** tab the moment the link loads, even if
the visitor never applies. Facebook's link-preview crawler is filtered out
server-side in `api/click.js` before anything is logged.

**Generate a link:**
```
node scripts/gen-ref.js fb https://YOUR-DOMAIN
```
Prints the ref code and the full URL to paste into Messenger. Set
`SITE_URL` in your shell so you can drop the second argument.

**Check results:** open the Sheet → **Squeaky Clean → Build ref report**.
Rebuilds a **Ref Report** tab: one row per ref code, click count, first
click time, and whether it converted — so links that got clicks but no
application stand out. File → Download → CSV from that tab to export it.

One-time setup after pulling this: re-run **Squeaky Clean → Set up
Applications tab** (safe — it only relabels row 1, your data rows are
untouched) so the live sheet gets the new **Ref** column, and re-deploy the
Apps Script (**Manage deployments → edit → New version**) so `doPost` picks
up the click-handling code. No new Vercel env vars — `api/click.js` reuses
`APPS_SCRIPT_URL` / `APPS_SCRIPT_TOKEN`.
