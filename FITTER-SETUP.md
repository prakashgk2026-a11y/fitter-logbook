# Fitter logbook — Govt. ITI Nedumangad (setup guide)

This is the same logbook app as the COPA one, set up for the **Fitter** trade (CTS 3.0, NSQF 4) with its **22 learning outcomes**
(1–11 first year, 12–22 second year) taken from column 1 of the assessment-criteria table. It runs on its own free
Supabase database + GitHub Pages site, so the Fitter data is completely separate from the COPA logbook.

> The learning outcomes are the only syllabus content that was in the PDF. The **practical jobs, theory lessons and logbook tasks**
> (the "split-up") are typed into the Google Sheet below and imported — or added in the app.

## A. One-time setup (≈ 20 minutes, done once by whoever hosts it)

1. **Supabase** — https://supabase.com → *New project* (use the Nedumangad instructor's own account if the data should belong to them).
2. **SQL Editor** → paste all of `supabase/schema.sql` → Run.
3. **SQL Editor** → paste all of `supabase/seed_fitter-cts3.sql` → Run.
   Result: ITI *Govt. ITI Nedumangad*, trade *Fitter* (24 months) and the 22 learning outcomes. The last query lists them so you can check.
   (Run twice by mistake? Nothing is duplicated.)
4. **Authentication → Providers → Email**: turn **Confirm email** OFF (or set up SMTP). **Authentication → URL Configuration**: Site URL and Redirect URL = the GitHub Pages address of step 6.
5. **Project Settings → API**: note the *Project URL* and the *anon public* key (never the service_role key).
6. **GitHub** — new repository, upload this folder, then
   *Settings → Pages → Source = GitHub Actions*, and *Settings → Secrets and variables → Actions → Variables*:
   `VITE_SUPABASE_URL` and `VITE_SUPABASE_ANON_KEY`. Push → the site appears at `https://<name>.github.io/<repo>/`.
   The app name "Fitter Logbook" comes from the file `.env.production` (change it there, or add a `VITE_APP_NAME` variable).
7. Open the site → **Sign up** (the Fitter instructor, first). Then put that e-mail in `supabase/bootstrap.sql` and run it in the SQL Editor — this person is now the administrator. Reload the site: Home shows the 22 learning outcomes.
8. Other instructors: they sign up → administrator opens *Settings → Users & access → Approve…* and picks ITI Nedumangad, role and trade Fitter.

## B. Fill the syllabus split-up in Google Sheets (optional but recommended)

1. New blank Google Sheet (sheets.new) → *Extensions → Apps Script* → delete the sample code → paste the whole file
   `apps-script/fitter-cts3-sheet.gs` → Save → choose **setupLogbookSheet** → Run → allow access (it only touches that sheet).
2. The sheet now has the tabs **Read me · LearningOutcomes (22, filled) · Jobs · Lessons · Tasks** and a **📘 Logbook** menu.
   - *Jobs*: type the job number (e.g. `1.1`, `1.2` … `1.10`), title, optional notes and hours. The LO column fills itself.
   - *Lessons*: same for theory (number, title, hours, notes).
   - *Tasks*: logbook tasks — number `1.2.1` (= first task of job 1.2), name, Aim, Procedure, Result, Conclusion.
   - Numbers are text, so `1.10` stays `1.10`. A red number means it was typed twice.
3. Menu **📘 Logbook → Check my data before exporting** lists missing LOs, duplicates, tasks without a job, empty titles.
4. **File → Download → Microsoft Excel (.xlsx)**, save as `fitter-data.xlsx`.

## C. Import the sheet into the app

On a computer with Node.js 20+ (in this project folder):

```powershell
npm install
$env:SUPABASE_URL="https://YOUR-PROJECT.supabase.co"
$env:SUPABASE_ANON_KEY="your-anon-key"
$env:ADMIN_EMAIL="administrator@example.com"      # the account from A7
$env:ADMIN_PASSWORD="its-password"
# 1) look first — nothing is written:
node scripts/migrate-copa.mjs --file fitter-data.xlsx --preset presets/fitter-cts3.json --dry-run
# 2) then import for real:
node scripts/migrate-copa.mjs --file fitter-data.xlsx --preset presets/fitter-cts3.json
```
(On Mac/Linux use `export NAME=value` instead of `$env:NAME=`.)
The import adds jobs, lessons and tasks under the right learning outcomes and keeps the 22 LOs. It runs **once per trade**: if jobs/lessons already exist it stops
rather than duplicating (delete the trade in *Settings → Trades* and repeat A3 to start over). After that, keep editing inside the app (turn on **Edit mode**).
Photos and PDFs are added in the app.

## D. Day to day
Edit mode → add/edit jobs, lessons, tasks, photos · *Student diary → Add diary* (e.g. "Fitter 2026–27 Batch") · *Share to repository* to let other ITIs copy the Fitter syllabus ·
Settings → Trades → *Student read-only link* · Settings → Backup.

## E. Changing the Fitter learning outcomes
Edit `presets/fitter-cts3.json`, run `npm run make:fitter` (regenerates the SQL seed and the Sheet script) and `npm test`.
Single changes are easier in the app: Home → ✎ on a learning outcome.

## Notes on the PDF
- LO 10 and LO 22 have identical wording in the PDF ("Read and apply engineering drawing…"); both are kept as numbered there.
- Square-bracket notes in column 1 (e.g. "[Basic fitting operation – marking, Hacksawing …]") are stored as the learning outcome's *description*; the sentence before it is the *title*.
- Typo fixed: LO 5 "…gas (oxyacetylene) /LPG / oxy-hydrogen)" → "gas (oxyacetylene / LPG / oxy-hydrogen)".
