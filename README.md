# ChargeCap v2.1: Microsoft 365 + code library backup

## New in v2.1 (2026-10-31)
- **Code library backup.** The CPT/ICD-10 library is now also stored on a
  **Code Library** tab in `ChargeCap Log.xlsx` (columns Type, Code,
  Description, Category). Every add/edit/delete on the phone is logged and
  replayed onto that tab about 2 seconds later (or on the next launch if
  offline). After each sync the phone reloads its library from the tab, so
  an app update or reinstall can't wipe your codes.
- Codes typed straight onto the tab in Excel (Type = CPT or ICD-10) show up
  in the app on its next launch or after Settings → **Sync code library now**.
- First sync from a phone keeps any codes that exist only on that phone and
  adds them to the tab.
- If the same code is changed in Excel and on the phone between syncs, the
  phone's change wins.
- sw.js `chargecap-v26`.

---

# ChargeCap v2.0: Microsoft 365

v2.0 moves sync from Google Sheets to an **Excel workbook in the practice's
Microsoft 365 OneDrive** (`ChargeCap/ChargeCap Log.xlsx`), covered by the
practice's Microsoft 365 BAA. Capture, OCR, coding and the phone's own case
list are unchanged. For the v1.x history (Google), see the README in
`chargecap-v1.13`.

## What changed from v1.13
- **Sign-in:** MSAL.js 3.30.0 (bundled as `msal-browser.min.js`), Entra app
  "ChargeCap" (client `8296bb04-…`, tenant `3602a33d-…`), single-tenant,
  SPA redirect URI `https://stickerscanner.netlify.app/`, delegated
  User.Read + Files.ReadWrite (+ offline_access). Sign-in uses a full-page
  redirect, not a popup. Popups were what broke silent sign-in in the iOS
  home-screen app. Expect about one "tap to reconnect" a day, because
  Microsoft caps a browser app's refresh token at 24 hours. Before that it
  refreshes on its own.
- **Workbook API:** Microsoft Graph Excel REST. Tabs are unchanged:
  Primary & Co-Surgeon, Assistant, Non-Op Consults (data), All Cases,
  Monthly Tally (formulas).
- **Newest cases at the top.** A new case is inserted at its date position,
  above every row with the same or an older date. A case entered late goes
  in with its own date. All Cases and Monthly Tally are newest first too.
  Because inserts shift rows, each sync re-reads the tab and finds the case
  by date + MRN (or name). When two cases share both, the CPT list and the
  last row number decide.
- **Month breaks.** Conditional formatting on every data tab and All Cases:
  the first (newest) row of each month has a thick teal line above it,
  bold text, and a dark-teal date cell. Alternate months are lightly
  tinted. It's formatting, not extra rows, so it stays right through hand
  edits and sorting.
- **Typed cells.** Date, DOB and Billed Date are real Excel dates. Every
  other column is text, so MRNs keep leading zeros and "99222-57" stays as
  typed.
- **Upgrade from v1.13:** on first launch the app drops the old
  Google-sheet row pointers and the Google token/sheet id (cases are
  untouched). The first sync then re-links every case to its row in the
  workbook.
- `todayISO()` now uses the local date. It was UTC, so after 6pm the
  default date of service was tomorrow.
- sw.js `chargecap-v25`: caches only GETs of the app's own files and the
  CDN libraries. v1.x matched every URL by mistake: `"./"` became `""`,
  and every URL ends with `""`. With Graph, that could have served stale
  workbook reads.

## Files
- `app.js`, `index.html`, `sw.js`, `styles.css`, `data.js`, `manifest.json`,
  `icons/`: the app.
- `msal-browser.min.js`: Microsoft's sign-in library, served with the app.
- `chargecap-template.xlsx`: a blank workbook (tabs, headers, formats, month
  breaks, **no patient data**). The app uploads it only if
  `ChargeCap/ChargeCap Log.xlsx` doesn't exist yet in OneDrive.
- **Not deployed** (kept outside this folder, contains PHI):
  `ChargeCap Log.xlsx` has the 330 cases merged from Tribble (Apr 1–Sep 1)
  and the Google sheet (Sep 4 on).
- `../chargecap-v2.0-extras/weekly-biller-report.ts` is the Office Script
  for the Monday biller email.

## Setup, in this order
1. **Upload the seeded workbook first.** In OneDrive (office.com → OneDrive),
   create a folder named `ChargeCap` and upload `ChargeCap Log.xlsx` into it.
   The name must be exact. If the app connects before this file exists, it
   creates a blank workbook in that spot instead.
2. **Deploy** this folder to Netlify (drag the whole folder onto Deploys).
3. **Phone:** don't delete the home-screen app. Close it fully and reopen
   it once or twice until Settings → About says **ChargeCap v2.0**.
4. **Settings → Connect Microsoft 365** → sign in with the practice account
   → you land back in the app. The first sync re-links your cases to their
   rows.
5. **Settings → Missing from the workbook → Check.** Review anything it
   lists before tapping Add (see the notes below).
6. **Weekly biller email:** see the next section. Then **turn off the
   Google Apps Script trigger**. Otherwise it keeps emailing the frozen
   Google sheet's pending rows every Monday.

## Weekly biller email (Power Automate + Office Script)
1. Open `ChargeCap Log.xlsx` in Excel for the web → **Automate** tab →
   **New Script**. Replace the code with `weekly-biller-report.ts`, name it
   **ChargeCap weekly biller report**, then Save.
2. make.powerautomate.com → **Create → Scheduled cloud flow**. Name it
   "ChargeCap weekly biller report". Repeat every 1 week, Monday. Then edit
   the Recurrence step: Time zone (UTC-07:00) Mountain Time, At these hours
   6, At these minutes 0.
3. **+ New step → Excel Online (Business) → Run script.** Location OneDrive
   for Business, Document Library OneDrive, File
   `/ChargeCap/ChargeCap Log.xlsx`, Script **ChargeCap weekly biller
   report**, mode `collect`.
4. **Data Operation → Parse JSON.** Content = `result` from Run script.
   Schema:
   `{"type":"object","properties":{"count":{"type":"integer"},"csv":{"type":"string"},"keys":{"type":"string"}}}`
5. **Condition:** `count` is greater than 0. In the **If yes** branch:
   - **Office 365 Outlook → Send an email (V2).** To `rnaylor@rmapinc.com`.
     Subject `Weekly billing report for Dr. Paulk`. Body: a short note.
     Under Advanced, add yourself as BCC. Attachments Name
     `ChargeCap pending @{formatDateTime(convertFromUtc(utcNow(),'Mountain Standard Time'),'yyyy-MM-dd')}.csv`,
     Attachments Content = `csv` (from Parse JSON).
   - **Excel Online (Business) → Run script** again: same file and script,
     mode `mark`, keys = `keys` (from Parse JSON), billedDate =
     `@{formatDateTime(convertFromUtc(utcNow(),'Mountain Standard Time'),'M/d/yyyy')}`.
   Because this runs after Send email, rows are marked "done" only if the
   email went out.
6. Test with **Test → Manually** after temporarily changing To to your own
   address. Note: a test run marks the rows done, so pick a time when
   that's what you want, or flip them back to "pending" afterwards.

## Things to watch on the first real use
- **All Cases** and **Monthly Tally** use ordinary formulas only. They
  work in every Excel since 2010 and in Excel for the web. A first version
  used VSTACK, which older Excel doesn't have, and came up blank.
  All Cases merges the three newest-first data tabs row by row, using
  hidden helper columns P–W, and covers 5,000 cases. Monthly Tally covers
  240 months. Both are built into the workbook files. The app only rewrites
  them if one of those tabs is deleted.
- Every formula range into a data tab starts at **row 1**. The app inserts
  new cases at row 2, and Excel shifts, rather than grows, a range whose
  first row is the insert point. LibreOffice test: insert at A2:O2 → All
  Cases 330→331 with the new case on top, Oct tally 41→42.
- All Cases assumes each data tab stays newest-first (the app keeps it
  that way). A row typed out of order by hand still shows up, just slightly
  out of order.
- iOS home-screen app + Microsoft sign-in redirect: the login page opens
  inside the app, then comes back. If it gets stuck in Safari instead,
  that is the thing to report.
- Phone cases dated before 9/4 that aren't in Tribble (or whose MRN was
  typed differently) will be listed under "Missing from the workbook".
  Review them before adding.

## Local testing
`npx serve .`, then open the localhost URL. Microsoft sign-in only works
from the registered redirect URI, so a local copy can't sync. The mocked
test (`/tmp/cctest/test.js` in the build session) covers sync logic.
