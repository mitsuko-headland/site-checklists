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
| Job Items | JobID, ItemID, Type (Item / HIDE / EDIT / LIST – EDIT = master task changed for this job only, one row per job+task, updated in place; LIST = whole checklist hidden on this job, ItemID L<n>, Removed=Y when shown again), Checklist, Trade, Area, Critical, Item, Added By, Added At, Removed |
| Check Log | JobID, ItemID, Checklist, Status (Done / N/A / Cleared), Name, Timestamp — append-only audit trail; latest row per item wins |
| Edit Log | Timestamp, Name, JobID, ItemID, Field, Old Value, New Value — every edit and "Removed from job" made in the app; created on first use |
| Notes | JobID, ItemID, Checklist, Note, Name, Timestamp — created on first note; append-only, shown oldest first |
| Trades | Code, Trade — heading codes and sort order (lowest code first, ties keep tab order, no code = last). Seeded from `TRADES` in Code.gs on first run; after that the tab is the source. Codes stored as text |
| (readable columns) | **Job Name** (after JobID) and **Task** (after ItemID) on Job Items, Check Log, Notes and Edit Log – readable copies written at the time by `append_` (via `jobLabel_`/`taskLabel_`, which read sheets with `raw_`, never `rows_`). Backfilled once by `readableUpgrade_` (property `READABLE_V1`; older rows get current wording). The app only uses the ID columns |
| Setup | Instructions for the owner; rewritten by `writeSetupTab_` when `SETUP_TEXT_VERSION` changes |

Checklists: 1 Project Startup & Site Protection · 2 Demolition · 3 Pre-Rough In · 4 Rough In & Pre-Drywall · 5 Drywall · 6 Finishing.
`upgrade_()` migrates older sheets once (Walk→Checklist, Hold→Critical, renumbering, startup items); keep it idempotent.

## Current behaviour
- Sign-in: name + PIN (5 wrong PINs → 15-minute lock; 30-day signed token stored on the phone).
- Everyone signed in: see all jobs, create jobs, tick/untick, add items to one job, copy a job's link, **Add note** on any task (running log per job, name + time).
- Trade headings: tap to fold/unfold – the screen stays exactly where it was – blank space (#foldPad) is added at the bottom when the page would otherwise get too short (a trade with every task done starts folded each time a checklist is opened); the **+** on a heading opens the same form as "+ Add items" (tasks, trade – filled in with that heading but changeable – area, critical, Add to: This job only / All jobs (master list, admins only)); the default is always This job only.
- Filter row: **All / Still to do / Critical items / Completed** (Completed = Done or N/A; choosing it unfolds every trade). **Collapse all / Expand all** (#foldAll, class `fold-all`, not a `.chip`) sits at the far right as plain accent text with a small arrow; it folds or opens every trade shown, and its label flips when everything shown is folded. Filter handlers select `.chip[data-f]`.
- The "this job" tag on job-only tasks shows to **admins only**; the "Added by" line and "edited for this job" tag show to everyone.
- Job-only tasks show "Added by <name> · <Mon d>" (from Job Items Added By / Added At; `addedBy`/`addedOn` in getJob). Master-list tasks never show who added them; tick stamps (who ticked + date/time) show on every task.
- "+" scrolls the heading and its add box to the top of the screen (so the phone keyboard doesn't cover it). Tasks added to one job (either button) show at the **top** of their trade, newest batch first, typed order within a batch; trade order still follows the master list.
- **Everyone** signed in: **Edit** on every task (wording, trade, area, critical flag) and **Remove from this job** (a checkbox at the bottom of the edit form; ticking it turns Save into a red "Remove task" button, and only that tap removes – any other edits in the form are ignored) – always **this job only**. A master task edited on a job gets an EDIT row in Job Items and shows an "edited for this job" tag; other jobs keep the master wording. Every edit/removal is logged in Edit Log. Changing a master task on every job is done only on the sheet's Items tab (owner's rule).
- The open edit form is held in `S.editor` and re-attached after every redraw (renderList → renderList_; focus restored), so loads, ticks saving and auto-refresh never close it or lose what's typed; it closes only on Save / Remove task / Cancel, or when leaving the checklist/job. (Bug fixed 2026-10-23: a redraw used to wipe it.)
- Admins (Admin = Y): additionally add tasks to the master list from "+ Add items" → "All jobs (master list)".
- Admins can **Hide / Unhide a whole checklist on one job** (button on each checklist card, tap-again confirm within 4 s; `setListHidden`, logged in Edit Log as "Checklist visibility"). Non-admins don't see a hidden checklist on that job (and are sent back to the list if it's hidden while open); admins see it greyed out with a HIDDEN label and can still open it. Nothing is deleted. `getJob` returns `hiddenLists`.
- Trade headings stay pinned under the app header while scrolling. Auto-refresh every 45 s, paused while an editor, note box or add form is open.
- Trade headings show the code from the Trades tab (e.g. "1050 Site Supervisor") and trades sort by code within each checklist; the trade drop-downs list every Trades-tab trade plus any trade already used. A heading only appears when the checklist has items for that trade.
- 2026-10-23 one-time trade moves (`tradesUpgrade_`, script property `TRADES_V1`): Energy Consultant, Surveyor, Cleaning → Site Supervisor; Eavestrough → Metalwork; Siding, Masonry (Brick), Masonry (Stone), Deck Carpentry → Exterior Carpentry; Parging → Concrete; Closet Shelving → Finish Carpentry; master items S001–S004 → Mobilize & Startup Site Protection. Logged in Edit Log as "Setup (trade codes)"; counts + uncoded trades written to the Setup tab.
- 2026-10-23 Park clean-up, step 1 (read-only): `cleanupReport_(ss, 'Park', 'Mitsuko', 'Park Cleanup')` runs once from `upgrade_` (property `PARK_REPORT_V1`) and writes the **Park Cleanup** tab: her ticks/unticks (+ what each task goes back to), notes, tasks added (flags others' ticks/notes on them), removals, edits on Park, her '(all jobs)' master edits, and app-added master tasks (ItemID M…, date decoded from the id; who added isn't recorded). Step 2 `cleanupRun_` (property `PARK_CLEAN_V1`: started/done/error: …) deletes her Park rows from Check Log, Notes and Edit Log, deletes her added tasks (kept if anyone else ticked/noted them), her HIDE rows and EDIT rows (kept if others also edited), restores/reverts her removals/edits on other people's job-only tasks, copies every deleted row to **Park Cleanup – Removed rows** first, then rewrites the Park Cleanup tab with the result. Master list (sections 6–7) untouched pending the owner. Reads the sheet directly – never call rows_ from code that upgrade_ runs.
- 2026-10-23 one-time fix `restoreHidden_(ss, 'JMVICS4598Q', 't112', 'Mitsuko')` (property `RESTORE_T112_V1`): deletes her mistaken HIDE row so the task is back on that job; logged in Edit Log as "Restored to job".
- 2026-10-23 `stoneUpgrade_` (property `STONE_V1`): Trades tab "Stone Countertops" → "Stone" (12370); tasks under Stone Countertops moved to Stone, logged as "Setup (trade codes)".
- No "Mark N/A" button (old N/A ticks still display and count as done).
- Speed: the app starts with ONE backend call, `startup(token, jobId)` (who's signed in + job list + the open job), instead of whoAmI → getJobs → getJob. Each phone keeps a saved copy (localStorage keys `sc.c.*`: me, jobs, last 5 jobs' data) so the app paints instantly, then updates from the sheet; the copy is cleared on sign-out. Ticks show instantly and save in the background (one write at a time per task; a failed save undoes the tick and says so). `setCheck` and `addNote` only append a row, so they no longer take the script lock.

## Tests
`node tests/backend.test.js` runs 117 checks against a mock of SpreadsheetApp/Utilities/etc. (sign-in, lockout, sheet upgrade, ticks, add/edit/remove, Edit Log, notes, startup, trade codes, added-by, job-only edits, removal log, clean-up summary + clearing, hide checklist, restore, readable columns, Stone). The main suite starts with READABLE_V1/STONE_V1 set so column positions stay put; those two upgrades are tested at the end. Run it before every push that touches `src/Code.gs`, and add checks for new behaviour – the owner can't debug Apps Script. For `src/Index.html`, at least confirm the script parses: replace `<?= … ?>` scriptlets with a placeholder and run it through `new Function`.

## Open items / ideas not yet requested
- Admin features the owner may want later (master edits/retiring stay on the sheet by her rule): archive finished jobs, manage users/PINs from the app.
- Short link for the team (owner to create with TinyURL).
