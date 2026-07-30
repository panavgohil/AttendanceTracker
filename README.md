# Stateline — What changed

## 1. Visual redesign
Both `index.html` and `login.html` got a full glassmorphism pass:
- Frosted glass cards (`backdrop-filter: blur`), soft drifting ambient gradient orbs in cyan/violet behind everything
- New type system: **Sora** for headings/brand, **Inter** for body text, **JetBrains Mono** for all numbers/stats (percentages, points, roll numbers)
- Circular animated ring gauges replacing the flat progress bars, color-coded green/amber/red by safety margin
- Dark/light theme toggle (🌙/☀️ in the header), preference saved locally
- Mobile bottom nav bar for quick jumps to Stats / Today / Schedule / Settings

## 2. New features
- **Overall attendance ring** — one glanceable number for the whole semester
- **Bunk calculator** — pick a subject and a target % (defaults to 75) and it tells you how many classes you can safely skip, or how many you must attend in a row to recover
- **Per-subject "safe to skip / attend next N"** hint under every subject's ring
- **Low-attendance banner** — auto-appears listing any subject currently under 75%
- **Activity heatmap** — 12-week GitHub-style calendar: green (no bunks that day), red (missed something), violet (only cancelled/rescheduled)
- **Streak counter** — consecutive days without a bunk
- **Weekly PDF report** — opt-in toggle + "send now" button, fully wired to the backend below

## 3. Schedule editing (no more full resets)
"Manage Subjects" now lists every subject per day with **✎ edit** and **✕ delete**, so you fix one entry instead of wiping the semester. The old full reset still exists as a small "danger zone" fallback.

## 4. Weekly PDF email reports — backend (`Code.gs`)
This is your original script with additions merged in — everything else is untouched:

1. **New `doPost` case** `"send-report"` → `handleSendReport()`, which builds a one-page PDF (subject, attended/total points, %) from the `Attendance` sheet and emails it via `MailApp`.
2. **`updateUserProfile()`** now also writes a TRUE/FALSE opt-in flag into **column 7 of `Users`** (previously written as an empty string `""` on registration — now defaults to `"FALSE"`). This is what the weekly toggle in Settings controls.
3. **`handleLogin()`** now also returns `email` and `reportOptIn`, so logging in on a new device shows the right toggle state immediately instead of defaulting to off.
4. **`getHistoryJSON()`** now includes a `date` field per entry. Your previous version only returned `subject/status/weight/semester` — the new dashboard's heatmap and streak counter need the date to group entries by day. Dates are normalized through `formatDateCell_()` since Sheets sometimes auto-converts the date string sent from the front end into an actual Date cell.
5. **`sendWeeklyReports()` + `createWeeklyTrigger()`** — a new scheduled job. Run `createWeeklyTrigger` **once** from the Apps Script editor (pick it in the function dropdown, click ▶ Run) to install a trigger that emails every opted-in student a PDF every Sunday at 6pm. Re-running it is safe; it clears any previous trigger with the same name first.

### To deploy
1. Open your Apps Script project (Extensions → Apps Script from the Sheet).
2. Replace the entire contents of your existing script file with `Code.gs`.
3. Run `createWeeklyTrigger` once and authorize the Gmail/Sheets permissions it asks for.
4. Optional: run `sendWeeklyReports` manually once to confirm a test email arrives before trusting the Sunday trigger — check **Executions** in the left sidebar for errors.
5. Re-deploy the web app (Deploy → Manage deployments → Edit → New version) so the updated `doPost`/`doGet` logic is live at your existing `SCRIPT_URL`.

No sheet restructuring needed — this only starts using the already-existing (previously blank) 7th column of `Users` for the opt-in flag.

## 5. Uploading to GitHub — keeping the Apps Script URL out of the repo
`index.html` and `login.html` no longer hardcode `SCRIPT_URL`/`API_URL`. They now load it from `config.js`, which is listed in `.gitignore` so it never gets committed:

- `config.example.js` — committed placeholder, shows the shape others need to fill in
- `config.js` — your real file with the actual URL, **stays local, never pushed**

To set up a fresh clone (including your own, after pulling this):
```
cp config.example.js config.js
# then edit config.js and paste your real Apps Script URL in
```
If `config.js` is missing, both pages show an alert telling you to create it instead of silently failing.

**Worth knowing:** this only keeps the URL out of your *git history* — it doesn't hide it from anyone using the live site, since it's a plain client-side page and the URL is always visible in the browser's Network tab / page source once loaded. That's inherent to a no-backend-of-your-own static app; true protection would mean adding request-level auth checks in `Code.gs` itself (e.g. requiring the correct password on `saveAttendance`/`saveUserSchedule`/`updateUserProfile`, which currently trust the roll number alone). Happy to add that if you change your mind later — for now this just keeps your repo clean.

## Files
- `index.html` — dashboard
- `login.html` — login/register/forgot-password
- `Code.gs` — full backend (replaces your existing Apps Script file)
- `config.example.js` — commit this
- `config.js` — don't commit this (gitignored) — your real Apps Script URL goes here
- `.gitignore` — excludes `config.js` and common OS/editor junk
