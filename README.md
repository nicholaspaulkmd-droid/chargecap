# ChargeCap v1.0

Surgical charge capture PWA for Dr. Paulk. Built to spec against
`ChargeCap_Spec_and_Build_Guide.docx`. This replaces the earlier v0.6
Claude.ai chat prototype with a real, installable app.

## What's already done

- Full UI: capture flow (OCR sticker scan → patient info → role/modifier →
  CPT codes → ICD-10 codes → notes), case list with All/Pending/Billed
  filters and search, settings panel.
- On-device OCR via Tesseract.js — no server involved, extracts patient
  name / MRN / DOB with a confidence-based yellow highlight on fields
  worth double-checking.
- CPT and ICD-10 favorite libraries pulled directly from your OR billing
  sheets (`Billing Sheet - OR.xlsx` and `Billing Sheet - 11.04.22.xlsx`),
  organized by the categories from the spec (Bariatric, Revision, Hernia,
  Cholecystectomy, Appendix, Bowel, Upper GI, Endoscopy, HPB,
  Thyroid/Parathyroid, E&M, Component Separation), plus live ICD-10
  search against the free NLM ClinicalTables API.
- IndexedDB storage (via idb-keyval) — replaces the old localStorage
  prototype, persists across sessions and reinstalls.
- Red/green billing status toggle with timestamp, full case editing.
- Biller export (formatted text → clipboard) and CSV export (all /
  billed / pending).
- PWA manifest + service worker — installs to the iPhone home screen and
  the app shell works offline once it's been opened once.
- Google Drive sync, coded and ready, but **disabled until you supply a
  Client ID** (see Phase 2 below) — everything else works without it.
  Once connected, every saved case backs up automatically to a Google
  Sheet split into a Primary/Co-Surgeon tab and an Assistant tab, plus
  an auto-updating Monthly Tally of Bariatric/EGD/General Surgery case
  counts.

**Important — verify before relying on it for billing:** the CPT/ICD-10
codes in `data.js` were transcribed from your existing Excel billing
sheets, which had some evident typos in the source. Skim `data.js` (or
the code picker in the app) against current documentation before you
lean on it for actual claims. Codes can be edited directly in `data.js`.

## What's left — deployment (Phase 1, ~20 min)

You don't need to touch the app code for this part.

1. **Push to GitHub.** From this folder:
   ```
   git init
   git add .
   git commit -m "ChargeCap v1.0"
   ```
   Create a new repo called `chargecap` at github.com/new, then:
   ```
   git remote add origin https://github.com/YOUR-USERNAME/chargecap.git
   git branch -M main
   git push -u origin main
   ```
2. **Deploy to Netlify.**
   - netlify.com → Add new site → Import an existing project → GitHub →
     select `chargecap`.
   - Build command: leave blank. Publish directory: `/` (repo root).
   - Deploy. You'll get a URL like `chargecap-abc123.netlify.app`.
3. **Install on iPhone.** Open the Netlify URL in **Safari** (not
   Chrome — iOS only supports PWA install from Safari) → Share button →
   Add to Home Screen.

At this point the app is fully usable — capture, code, export CSV, copy
biller text — with data stored locally on the phone.

## What's left — Google Drive sync (Phase 2, optional, ~15 min)

1. console.cloud.google.com → New Project → name it ChargeCap.
2. APIs & Services → Library → enable **Google Drive API** and
   **Google Sheets API**.
3. APIs & Services → OAuth consent screen → External → fill in the app
   name (ChargeCap) and your email, then **Audience → Test users → Add
   users** and add your own Google account's email. Without this step
   sign-in fails with "Error 403: access_denied" even with a correct
   Client ID, because the app starts in "Testing" publish status and
   only whitelisted test users can sign in.
4. APIs & Services → Credentials → Create Credentials → OAuth client ID
   → Application type: **Web application**.
5. Under Authorized JavaScript origins, add your Netlify URL, e.g.
   `https://chargecap-abc123.netlify.app`. (No redirect URI needed —
   this uses Google Identity Services' token flow, which only checks
   the origin.)
6. Copy the generated Client ID (looks like
   `123456-abc.apps.googleusercontent.com`).
7. Open the app → Settings → paste it into "Google OAuth Client ID" →
   Save → Connect Google Drive → sign in and grant Sheets access.

From then on, **every saved case backs up automatically** — no need to
mark it billed first — to a spreadsheet named **ChargeCap Log** created
automatically in your Drive, with four tabs:

- **Primary & Co-Surgeon** — one row per case where your role is
  Primary or Co-surgeon.
- **Assistant** — one row per case where your role is Assistant.
- **All Cases** — a hidden helper tab that just unions the two above;
  no need to look at it directly.
- **Monthly Tally** — auto-counts, per month, how many Bariatric / EGD /
  General Surgery cases you logged (across both role tabs combined).
  Bariatric and EGD are determined by CPT code (see `BARIATRIC_CPT` /
  `EGD_CPT` near the top of `app.js` if those code lists ever need to
  change); a case with both a bariatric and an EGD code on it counts as
  Bariatric. Everything else counts as General Surgery. This tab is
  entirely spreadsheet formulas — it recalculates itself as new rows
  land, nothing in the app has to push counts to it.

Editing a case and re-saving it, or toggling its billed status, updates
that case's existing row in place rather than adding a duplicate — and
if you change a case's role after it's already synced, the old row is
cleared and a fresh one is added to the correct tab. If you're offline,
syncs queue and go out the next time you're online and signed in.

You can skip this phase entirely and just use CSV export / biller-text
copy as your sync method — nothing else in the app depends on it.

**Heads up on patient data:** this sends full case data — including
patient name and MRN — to a spreadsheet in your own Google Drive. Only
your Google account can see it (no other server is involved), but it's
worth being deliberate about who has access to that Drive account and
whether your practice's policies are fine with PHI living in Google
Sheets before you turn this on.

## Local testing before you deploy

Any static file server works, e.g. from this folder:
```
npx serve .
```
Then open the printed `localhost` URL in a browser. (The camera/OCR
flow needs a real device or a browser that lets you pick an image file
as the "camera" input — desktop Chrome will just open a file picker,
which is fine for testing the OCR parsing logic.)

## Updating the app later

Edit the files, `git add . && git commit -m "..." && git push` — Netlify
redeploys automatically in under a minute. If you change `app.js`,
`styles.css`, or `data.js`, bump `CACHE_VERSION` at the top of `sw.js`
so installed phones pick up the new version instead of serving a stale
cached copy.

## Architecture reference

| Layer | Technology |
|---|---|
| Frontend | Vanilla HTML/CSS/JS (no framework) |
| OCR | Tesseract.js 5.x, on-device |
| ICD-10 live search | NLM ClinicalTables API (free, no key) |
| Local storage | IndexedDB via idb-keyval |
| Sync | Google Sheets API v4 (one row per billed case) |
| Auth | Google Identity Services OAuth 2.0 (browser-only, no backend) |
| Hosting | Netlify free tier (HTTPS, auto-deploy from GitHub) |
| PWA | manifest.json + service worker, installs to iOS home screen |

## Security note (from the original spec)

OAuth tokens live in the browser's own storage. No patient data is sent
anywhere except your own Google Drive under your own Google account —
there is no backend server. HTTPS is enforced by Netlify. If you want to
avoid storing PHI in Drive entirely, an earlier version of this plan
called for stripping identifiers before sync since the EMR already holds
the chart — worth a final decision on which approach you want before
Phase 2, since switching later just means editing the `caseToRow()` /
`SHEET_HEADER` fields in `app.js`.
