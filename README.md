# Site Walks

Mobile site-walk checklists for Headland Construction. Google Apps Script web app backed by the "Site Walks" Google Sheet.

- `src/Code.gs` – backend (reads/writes the sheet, checks Headland Google sign-in)
- `src/Index.html` – the app screens
- `src/appsscript.json` – project settings (Vancouver time, web app open to headlandconstruction.com)
- `.github/workflows/deploy.yml` – publishes every change to the live web app automatically

## How updates work
Changes are committed to `main` (by Claude or anyone with access). GitHub Actions uploads `src/` to the Apps Script project and updates the existing deployment, so the web app link never changes. See the **Actions** tab for each run; a red ✗ means it did not publish.

## Settings this repository needs (Settings → Secrets and variables → Actions)
| Type | Name | Value |
|---|---|---|
| Secret | `CLASPRC_JSON` | Contents of the `.clasprc.json` file created by `clasp login` |
| Variable | `SCRIPT_ID` | Apps Script → Project Settings → Script ID |
| Variable | `DEPLOYMENT_ID` | Apps Script → Deploy → Manage deployments → Deployment ID |

Never commit `.clasprc.json` – it can change Apps Script projects in the owner's Google account.
