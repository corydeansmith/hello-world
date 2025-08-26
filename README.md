# Workflow Demo — Zero‑Setup Test

This proof‑of‑concept is a Google Apps Script web app that:
- Creates a Drive folder and a Google Doc per request
- Logs each request in a Google Sheet
- Shows a simple UI to submit and list recent requests

No manual IDs, no template, no prior setup.

## Test in 3 minutes
1. Open `script.google.com` → New project.
2. Create files and paste from this repo:
   - `Code.gs` → `gas/Code.gs`
   - `Index.html` → `gas/Index.html`
   - `appsscript.json` → `gas/appsscript.json` (View → Show manifest file)
3. Deploy → New deployment → Web app → Execute as: Me → Access: Anyone with the link → Deploy.
4. Open the URL. Authorize when prompted.
5. Submit a request.

What happens automatically on first run:
- A Spreadsheet named `Workflow Demo Tracker` is created (with sheet `MasterTracker`).
- A Drive folder named `Workflow Demo Requests` is created.
- A subfolder and Doc are created for your request. The Doc is prefilled with basic info.
- A row is appended to the tracker with links.

Where to find things later
- The created Spreadsheet and Folder are in your Drive root.
- The app stores their IDs in Script Properties, so subsequent runs reuse them.

Optional
- Change names in `DEFAULTS` (top of `Code.gs`).
- Add email notifications (requires Gmail scope) or more statuses.

Troubleshooting
- Authorization prompts: accept Drive/Sheets/Docs permissions.
- If recent requests is empty: try submitting once; the sheet is created on first run.
- To reset: in Apps Script → Project Settings → Script properties, delete `SPREADSHEET_ID` and/or `PARENT_FOLDER_ID`, then reload the web app.
