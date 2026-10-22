# Site Checklists – notes for Claude

Mobile site checklists for Headland Construction (owner: Mitsuko, mitsuko@headlandconstruction.com).
A Google Apps Script web app whose data lives in the owner's **"Site Checklists" Google Sheet**.
Field crews use it on their phones: sign in → choose a job → choose a checklist → tick items.

## Working rules (from the owner – always follow)
- **Do not change anything until the owner says "go"** (or "proceed with the changes"). Until then, collect requests, explain the plan, ask questions, and keep a numbered running list.
- When the owner says go: first `git fetch` and compare with `origin/main` so you never overwrite edits made elsewhere; then make the changes, run the backend tests, commit, push, and confirm the deploy succeeded (see below).
- Write for a non-technical reader: plain words, what changed, what (if anything) they need to do.
- **How to give the owner steps:** she is on **Windows** (PowerShell) – never give Mac/Finder steps. Keep steps short: a numbered command she types, one line on what happens, and a one-line path for clicks (A → B → C). No long background or explanations unless she asks.
- The owner is the only person who opens the Google Sheet. Anything other users need to do must be possible in the app.

## How publishing works
- Code: `src/Code.gs` (backend), `src/Index.html` (screens; an HtmlService *template* – `<?= jobParam ?>` and `<?= appUrl ?>` are filled in by `doGet`), `src/appsscript.json` (time zone America/Vancouver; web app `executeAs: USER_DEPLOYING`, `access: ANYONE_ANONYMOUS` – the app does its own name + PIN sign-in).
- Every push to `main` that touches `src/` runs `.github/workflows/deploy.yml`: `clasp push` to the Apps Script project, then `clasp deploy --deploymentId` to update the existing web-app deployment, so **the link never changes**.
- Repo settings (owner manages): secret `CLASPRC_JSON`, variables `SCRIPT_ID`, `DEPLOYMENT_ID` (the "Site Checklists" deployment, `AKfycbwWR3Fi…`).
- Check a deploy: `gh run list -R mitsuko-headland/site-checklists --limit 1`, then read the job's check-run annotations (`gh api repos/mitsuko-headland/site-checklists/check-runs/<job id>/annotations`) – the workflow writes "Uploaded …" / "Published … version N" notices or the exact Google error there. Raw logs are not reachable through the proxy.
- Web app link: `https://script.google.com/macros/s/AKfycbwWR3FihWpsZUhz7CMhESpjXXzfmLs1UrnoL66bp1GJt6Wjdk8MgN1Bj5_ddCDpfm4SfQ/exec` (per-job links add `?job=<JobID>`).
- Deploy error `invalid_grant` = the Google sign-in saved in `CLASPRC_JSON` has expired. Give the owner the steps in **Renewing the Google sign-in** below, word for word. When she says done, re-run with `gh api -X POST repos/mitsuko-headland/site-checklists/actions/workflows/deploy.yml/dispatches -f ref=main` (`gh workflow run` fails – GraphQL is blocked) and confirm the "Published" notice.
- If a change needs a new Google permission (e.g. sending email), the owner must re-authorise once in the Apps Script editor – tell them.

## Renewing the Google sign-in (send exactly this)
In PowerShell (Start → type PowerShell):
1. `clasp.cmd login`. A browser window opens. Sign in with your Google account, click Allow, then come back to PowerShell.
2. `Get-Content $HOME\.clasprc.json -Raw | Set-Clipboard`. Nothing appears, but the sign-in is now copied.

Then paste it into GitHub: site-checklists repository → Settings → Secrets and variables → Actions → pencil next to CLASPRC_JSON → Ctrl+V → Update secret. Tell me when it's done.

(Use `clasp.cmd`, not `clasp` – plain `clasp` is blocked by Windows' script setting. If it keeps expiring within days, the likely cause is a Google Workspace re-sign-in rule; offer to fix that.)

## Google Sheet (tabs and row-1 headings are looked up by name – never rename them)
| Tab | Columns |
|---|---|
| Users | Name, PIN, Active, Admin — Active/Admin count as yes if the cell starts with Y (blank Active blocks sign-in) |
| Items (master list) | ItemID, Checklist, Trade, Area, Critical, Item, Division, Order, Active |
| Jobs | JobID, Job Name, Address, Created By, Created At, Active |
| Job Items | JobID, ItemID, Type (Item / HIDE), Checklist, Trade, Area, Critical, Item, Added By, Added At, Removed |
| Check Log | JobID, ItemID, Checklist, Status (Done / N/A / Cleared), Name, Timestamp — append-only audit trail; latest row per item wins |
| Edit Log | Timestamp, Name, JobID, ItemID, Field, Old Value, New Value — created on first admin edit |
| Notes | JobID, ItemID, Checklist, Note, Name, Timestamp — created on first note; append-only, shown oldest first |
| Setup | Instructions for the owner; rewritten by `writeSetupTab_` when `SETUP_TEXT_VERSION` changes |

Checklists: 1 Project Startup & Site Protection · 2 Demolition · 3 Pre-Rough In · 4 Rough In & Pre-Drywall · 5 Drywall · 6 Finishing.
`upgrade_()` migrates older sheets once (Walk→Checklist, Hold→Critical, renumbering, startup items); keep it idempotent.

## Current behaviour
- Sign-in: name + PIN (5 wrong PINs → 15-minute lock; 30-day signed token stored on the phone).
- Everyone signed in: see all jobs, create jobs, tick/untick, add items to one job, copy a job's link, **Add note** on any task (running log per job, name + time).
- Trade headings: tap to fold/unfold – the screen stays exactly where it was – blank space (#foldPad) is added at the bottom when the page would otherwise get too short (a trade with every task done starts folded each time a checklist is opened); the **+** on a heading adds tasks to that trade on **this job only** (never the master list – that stays on "+ Add items").
- "+" scrolls the heading and its add box to the top of the screen (so the phone keyboard doesn't cover it). Tasks added to one job (either button) show at the **top** of their trade, newest batch first, typed order within a batch; trade order still follows the master list.
- Admins (Admin = Y): add items to the master list; **Edit** on every item (wording, trade, area, critical flag) – edits to master items apply to every job and are logged; **Remove from this job** is inside the edit form (two taps to confirm).
- Trade headings stay pinned under the app header while scrolling. Auto-refresh every 45 s, paused while an editor, note box or add form is open.
- No "Mark N/A" button (old N/A ticks still display and count as done).

## Tests
`node tests/backend.test.js` runs 42 checks against a mock of SpreadsheetApp/Utilities/etc. (sign-in, lockout, sheet upgrade, ticks, add/edit/remove, Edit Log, notes). Run it before every push that touches `src/Code.gs`, and add checks for new behaviour – the owner can't debug Apps Script. For `src/Index.html`, at least confirm the script parses: replace `<?= … ?>` scriptlets with a placeholder and run it through `new Function`.

## Open items / ideas not yet requested
- Admin features the owner may want later: retire a master item from the app, archive finished jobs, manage users/PINs from the app.
- Short link for the team (owner to create with TinyURL).
