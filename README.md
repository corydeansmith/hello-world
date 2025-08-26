# Google Apps Script Workflow Demo (Proof of Concept)

This is a minimal web app that demonstrates automating a Google Drive/Docs/Sheets workflow:
- Create a Drive folder and a Google Doc from a template
- Replace placeholders in the Doc (e.g., `{{RequestID}}`, `{{Title}}`)
- Log the request into a master Google Sheet
- Provide a simple UI to create a request and view recent ones

## What you get
- `gas/Code.gs` — server-side Apps Script (V8) functions and endpoints
- `gas/Index.html` — client UI served by `doGet`
- `gas/appsscript.json` — manifest with required scopes

## Prerequisites
- A Google account with access to Drive, Docs, and Sheets
- One Google Sheet to act as the master tracker (record rows)
- One Google Doc to act as a template. Add placeholders you want replaced:
  - `{{RequestID}}` `{{Title}}` `{{Type}}` `{{CreatedAt}}`
- One Drive folder to act as the parent for created request folders

## Setup (about 10 minutes)
1. Create a Google Sheet and copy its ID (the part after `/d/` in the URL). Optionally add a sheet named `MasterTracker`.
2. Create a Google Doc to use as the template and copy its ID. Add placeholders like `{{RequestID}}` in the body.
3. Create a Drive folder to contain all created request folders; copy its ID.
4. In your browser, open a new Apps Script project (`script.google.com`) and choose Blank project.
5. Create these files in the project and paste the contents from this repository:
   - `Code.gs` (from `gas/Code.gs`)
   - `Index.html` (from `gas/Index.html`)
   - `appsscript.json` (from `gas/appsscript.json`) via File → Project properties → Scopes/Manifest editor (or from Editor: View → Show manifest file)
6. In `Code.gs`, update the configuration at the top:
   ```js
   const CONFIG = {
     spreadsheetId: 'YOUR_SHEET_ID',
     sheetName: 'MasterTracker',
     parentFolderId: 'YOUR_PARENT_FOLDER_ID',
     templateDocIdByType: { General: 'YOUR_TEMPLATE_DOC_ID' },
     statusValues: ['Submitted','In Review','Approved','Rejected','Delivered','Archived']
   };
   ```
7. Run `initializeProject` once to create headers if needed and authorize the script when prompted.

## Deploy the web app
1. Click Deploy → New deployment → Select type: Web app.
2. Description: `Workflow Demo`.
3. Execute as: `Me`.
4. Who has access: `Anyone with the link` (or your domain only, as desired).
5. Deploy and copy the web app URL. Open it to use the UI.

## Using the app
- Fill out the form and click Create request.
- The app will:
  - Generate a `RequestID`
  - Create a subfolder in your parent folder
  - Copy the template Doc into that folder and replace placeholders
  - Append a row to your master Sheet
- Links to the new folder/doc will be shown and added to the table of recent requests.

## Customization ideas
- Add more request `Type` values and map them to different templates in `templateDocIdByType`.
- Add email notifications (require adding the Gmail scope & `MailApp.sendEmail`).
- Add status change actions and buttons, backed by `setStatus(requestId, status)`.
- Add permissions sharing for requester/assignee using DriveApp permissions.

## Troubleshooting
- If you get a permissions error creating files, ensure you deployed the web app as `Me` and that your account has access to the parent folder/template.
- If the page loads but recent requests is empty, confirm you set the correct Sheet ID and Sheet name, then run `initializeProject` once.
- To see server logs: View → Executions in the Apps Script editor.

---

This proof-of-concept keeps your existing Google stack while removing manual steps. You can iterate from here to add approvals, reminders, or AppSheet front-ends.
