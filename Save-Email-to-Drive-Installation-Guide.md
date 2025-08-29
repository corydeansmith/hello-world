### Save Email to Drive — Installation Guide (Copy/Paste Version)

This guide installs the Gmail add-on on any Google account. It converts a Gmail message to PDF and saves it (plus any non-inline attachments) into a Drive folder named:

YYYY-MM-DD_HHMM - Sender, Subject

### What you’ll do
- Create a new Apps Script project
- Enable one service (Drive API)
- Paste two files (Code.gs and appsscript.json)
- Install the add-on for your account

### 1) Create the project
- Go to `https://script.new` while signed in to the target Google account
- Rename the project: Save Email to Drive
- Left sidebar: Services (+) > add “Drive API” (required)
- Project Settings (gear) > check “Show appsscript.json”

### 2) Paste these two files exactly
- In the editor, click `Code.gs`, delete everything, paste this code, and click Save

```javascript
var PARENT_FOLDER_ID = '';
var TIMEZONE = 'America/Regina'; // Set your timezone; or use Session.getScriptTimeZone()

function onGmailMessageOpen(e) {
  return buildMessageCard_(e);
}

function buildMessageCard_(e) {
  var messageId = (e && e.gmail && e.gmail.messageId) || '';
  var section = CardService.newCardSection()
    .addWidget(
      CardService.newTextParagraph().setText(
        'Save this message as PDF and attachments into a Drive folder named "YYYY-MM-DD_HHMM - Sender, Subject".'
      )
    )
    .addWidget(
      CardService.newTextButton()
        .setText('Save to Drive')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor('#1a73e8')
        .setOnClickAction(
          CardService.newAction()
            .setFunctionName('saveCurrentMessage_')
            .setParameters({ "messageId": messageId })
        )
    );

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Save Email to Drive'))
    .addSection(section)
    .build();
}

function saveCurrentMessage_(e) {
  try {
    var messageId = (e.parameters && e.parameters.messageId) || (e.gmail && e.gmail.messageId);
    if (!messageId) throw new Error('No messageId available.');

    var parent = PARENT_FOLDER_ID ? DriveApp.getFolderById(PARENT_FOLDER_ID) : DriveApp.getRootFolder();
    var message = GmailApp.getMessageById(messageId);
    var folder = saveMessageAndAttachments_(message, parent);

    var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Saved to Drive'))
      .addSection(
        CardService.newCardSection().addWidget(
          CardService.newKeyValue()
            .setTopLabel('Folder')
            .setContent(folder.getName())
            .setButton(
              CardService.newTextButton()
                .setText('Open Folder')
                .setOpenLink(CardService.newOpenLink().setUrl(folder.getUrl()))
            )
        )
      )
      .build();

    var nav = CardService.newNavigation().updateCard(card);
    return CardService.newActionResponseBuilder()
      .setNavigation(nav)
      .setNotification(CardService.newNotification().setText('Saved email and attachments to Drive.'))
      .build();
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Error: ' + err.message))
      .build();
  }
}

function saveMessageAndAttachments_(message, parentFolder) {
  var folderName = makeFolderName_(message);
  var folder = parentFolder.createFolder(folderName);

  var pdfBlob = messageHtmlToPdfBlob_(message);
  folder.createFile(pdfBlob).setName('Message.pdf');

  var attachments = message.getAttachments({
    includeInlineImages: false,
    includeAttachments: true
  });
  for (var i = 0; i < attachments.length; i++) {
    var att = attachments[i];
    try {
      folder.createFile(att).setName(att.getName());
    } catch (e) {}
  }

  return folder;
}

function makeFolderName_(message) {
  var date = message.getDate();
  var timestamp = Utilities.formatDate(date, TIMEZONE, 'yyyy-MM-dd_HHmm');
  var sender = sanitizeName_(extractDisplayName_(message.getFrom()));
  var subject = sanitizeName_(message.getSubject() || 'No subject');
  var base = timestamp + ' - ' + sender + ', ' + subject;
  return trimToLength_(base, 200);
}

function extractDisplayName_(fromHeader) {
  var nameMatch = fromHeader.match(/"?(.*?)"?\s*<.*?>/);
  if (nameMatch && nameMatch[1]) return nameMatch[1].trim();
  return fromHeader.replace(/[<>"]/g, '').trim();
}

function sanitizeName_(s) {
  return s.replace(/[\\/:*?"<>|#%{}@$'`+=]/g, ' ').replace(/\s+/g, ' ').trim();
}

function trimToLength_(s, maxLen) {
  return s.length <= maxLen ? s : s.slice(0, maxLen - 1) + '…';
}

function messageHtmlToPdfBlob_(message) {
  var htmlBody = message.getBody();

  // Try to place inline images in the body where referenced (cid:...)
  htmlBody = inlineInlineImages_(htmlBody, message);

  var subject = message.getSubject() || '';
  var from = message.getFrom() || '';
  var to = message.getTo() || '';
  var cc = message.getCc() || '';
  var bcc = message.getBcc ? (message.getBcc() || '') : '';
  var dateStr = Utilities.formatDate(message.getDate(), TIMEZONE, 'yyyy-MM-dd HH:mm');

  var header =
    '<div><strong>Subject:</strong> ' + escapeHtml_(subject) + '</div>' +
    '<div><strong>From:</strong> ' + escapeHtml_(from) + '</div>' +
    (to ? '<div><strong>To:</strong> ' + escapeHtml_(to) + '</div>' : '') +
    (cc ? '<div><strong>Cc:</strong> ' + escapeHtml_(cc) + '</div>' : '') +
    (bcc ? '<div><strong>Bcc:</strong> ' + escapeHtml_(bcc) + '</div>' : '') +
    '<div><strong>Date:</strong> ' + escapeHtml_(dateStr) + '</div>';

  // Always append a gallery of inline images so none are lost
  var inlineGallery = buildInlineImageGallery_(message);

  var html =
    '<meta charset="UTF-8">' +
    '<style>' +
    'body { font-family: Arial, Helvetica, sans-serif; font-size: 12pt; color: #222; }' +
    '.meta { margin-bottom: 16px; padding-bottom: 8px; border-bottom: 1px solid #ddd; }' +
    '.meta div { margin: 2px 0; overflow-wrap: anywhere; }' +
    '.content img { max-width: 100%; height: auto; }' +
    '.content table { border-collapse: collapse; }' +
    '.inline-gallery img { max-width: 100%; height: auto; }' +
    '</style>' +
    '<div class="meta">' + header + '</div>' +
    '<div class="content">' + htmlBody + inlineGallery + '</div>';

  var blob = Utilities.newBlob(html, 'text/html', 'email.html');
  var file = Drive.Files.insert(
    { title: 'tmp_email', mimeType: 'application/vnd.google-apps.document' },
    blob,
    { convert: true }
  );
  var pdfBlob = DriveApp.getFileById(file.id).getAs(MimeType.PDF).setName('Message.pdf');
  DriveApp.getFileById(file.id).setTrashed(true);
  return pdfBlob;
}

// Convert img src="cid:...": to data URLs from inline image blobs
function inlineInlineImages_(htmlBody, message) {
  try {
    var atts = message.getAttachments({ includeInlineImages: true, includeAttachments: false });
    if (!atts || atts.length === 0) return htmlBody;

    var cidToDataUrl = {};
    for (var i = 0; i < atts.length; i++) {
      var att = atts[i];
      var getCid = att.getContentId ? att.getContentId() : null;
      if (!getCid) continue;
      var cid = String(getCid).replace(/[<>]/g, '');
      var blob = att.copyBlob();
      var contentType = blob.getContentType() || 'image/png';
      var base64 = Utilities.base64Encode(blob.getBytes());
      cidToDataUrl[cid] = 'data:' + contentType + ';base64,' + base64;
    }

    htmlBody = htmlBody.replace(/src\s*=\s*"cid:([^"]+)"/gi, function(m, cid) {
      cid = String(cid).replace(/[<>]/g, '');
      return cidToDataUrl[cid] ? 'src="' + cidToDataUrl[cid] + '"' : m;
    });
    htmlBody = htmlBody.replace(/src\s*=\s*'cid:([^']+)'/gi, function(m, cid) {
      cid = String(cid).replace(/[<>]/g, '');
      return cidToDataUrl[cid] ? "src='" + cidToDataUrl[cid] + "'" : m;
    });

    return htmlBody;
  } catch (e) {
    return htmlBody;
  }
}

// Append all inline images at the end (in case some couldn't be placed inline)
function buildInlineImageGallery_(message) {
  try {
    var atts = message.getAttachments({ includeInlineImages: true, includeAttachments: false });
    if (!atts || atts.length === 0) return '';

    var parts = [];
    parts.push('<hr style="margin:16px 0;border:none;border-top:1px solid #ddd;">');
    parts.push('<div class="inline-gallery">');
    for (var i = 0; i < atts.length; i++) {
      var att = atts[i];
      var blob = att.copyBlob();
      var contentType = blob.getContentType() || 'image/png';
      var base64 = Utilities.base64Encode(blob.getBytes());
      var dataUrl = 'data:' + contentType + ';base64,' + base64;
      parts.push('<div style="margin:8px 0;"><img src="' + dataUrl + '"></div>');
    }
    parts.push('</div>');
    return parts.join('');
  } catch (e) {
    return '';
  }
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
```

- Click `appsscript.json`, replace everything with this JSON, and click Save

```json
{
  "timeZone": "Etc/UTC",
  "dependencies": {
    "enabledAdvancedServices": [
      {
        "userSymbol": "Drive",
        "serviceId": "drive",
        "version": "v2"
      }
    ]
  },
  "exceptionLogging": "STACKDRIVER",
  "oauthScopes": [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.addons.execute",
    "https://www.googleapis.com/auth/gmail.addons.current.message.readonly",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
  ],
  "addOns": {
    "common": {
      "name": "Save Email to Drive",
      "logoUrl": "https://www.gstatic.com/images/icons/material/system/1x/drive_black_48dp.png",
      "layoutProperties": {
        "primaryColor": "#1a73e8",
        "secondaryColor": "#188038"
      }
    },
    "gmail": {
      "contextualTriggers": [
        {
          "unconditional": {},
          "onTriggerFunction": "onGmailMessageOpen"
        }
      ]
    }
  }
}
```

Tip: If a blue link appears after adding “Drive API”, click it to open Google Cloud, click Enable for the Drive API once, then return.

### 3) Install the add-on
- Click Deploy > Test deployments
- In the dialog, open the “Add-on” tab (if present)
- Click Install > Continue > select your account > Allow
- If you don’t see Install, hard refresh (Ctrl+Shift+R) and try again

### 4) Use it in Gmail
- Open Gmail, then open any message
- In the right sidebar, click “Save Email to Drive”
- Click “Save to Drive”, then “Open Folder” in the confirmation

### Optional: choose a specific Drive folder
- In Drive, open your destination folder and copy the ID from the URL (after `/folders/`)
- In `Code.gs`, set:
  - `var PARENT_FOLDER_ID = 'PASTE_ID_HERE';`
- Save, then use the add-on again

### Optional: set your time zone
- Apps Script: Project Settings (gear) > Time zone > choose your local time (e.g., America/Regina)

### Troubleshooting
- Install button missing: ensure `appsscript.json` has the `addOns.gmail.contextualTriggers` block; Save and hard refresh
- Blank sidebar: make sure the two files are saved; try Uninstall then Install again under Test deployments
- “Found a service identifier used more than once: Drive”: remove any duplicate Drive entries in `appsscript.json`
- “Missing : after property ID”: ensure `.setParameters({ "messageId": messageId })` matches exactly
- Inline images not perfectly in place: they are included; fidelity depends on Google’s HTML→Docs converter

### Uninstall on this account
- Gmail: right sidebar > hover the add-on icon > 3‑dot menu > Uninstall
- Or in Apps Script: Deploy > Test deployments > Uninstall

### Updating later
- Open the project, replace `Code.gs` with the new version, Save
- Deploy > Test deployments > Uninstall (if shown) > Install again

You can reuse this guide for any Google account. If you want a version with the Gmail API for higher inline-image fidelity, ask me and I’ll provide a second guide with that variant.