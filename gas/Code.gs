const CONFIG = {
  spreadsheetId: 'PUT_SPREADSHEET_ID_HERE',
  sheetName: 'MasterTracker',
  parentFolderId: 'PUT_PARENT_FOLDER_ID_HERE',
  templateDocIdByType: {
    General: 'PUT_DOC_TEMPLATE_ID_HERE'
  },
  statusValues: ['Submitted', 'In Review', 'Approved', 'Rejected', 'Delivered', 'Archived']
};

function initializeProject() {
  const sheet = getOrCreateMasterSheet();
  ensureHeaderRow(sheet);
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Workflow Demo')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getOrCreateMasterSheet() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  let sheet = spreadsheet.getSheetByName(CONFIG.sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.sheetName);
  }
  return sheet;
}

function ensureHeaderRow(sheet) {
  const requiredHeaders = [
    'RequestID',
    'Status',
    'Type',
    'Title',
    'RequesterEmail',
    'Assignee',
    'DueDate',
    'FolderLink',
    'DocLink',
    'CreatedAt',
    'LastUpdatedAt',
    'Notes'
  ];

  const lastColumn = sheet.getLastColumn();
  const existing = lastColumn > 0
    ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    : [];

  if (existing.length === 0 || existing[0] === '') {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    sheet.setFrozenRows(1);
    return;
  }

  const existingSet = new Set(existing);
  const toAppend = requiredHeaders.filter(h => !existingSet.has(h));
  if (toAppend.length > 0) {
    sheet.getRange(1, existing.length + 1, 1, toAppend.length).setValues([toAppend]);
  }
  sheet.setFrozenRows(1);
}

function getHeaderIndexes(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => map[String(h)] = i + 1);
  return map;
}

function generateRequestId() {
  return Utilities.getUuid().slice(0, 8);
}

function nowIso() {
  return new Date().toISOString();
}

function safeString(value) {
  return value == null ? '' : String(value);
}

function getTemplateDocIdForType(type) {
  const key = type && CONFIG.templateDocIdByType[type] ? type : 'General';
  return CONFIG.templateDocIdByType[key];
}

function createOrCopyDocAndFolder(requestId, type, title) {
  const parentFolder = DriveApp.getFolderById(CONFIG.parentFolderId);
  const folderName = `${requestId} - ${type}${title ? ' - ' + title : ''}`;
  const folder = parentFolder.createFolder(folderName);

  const templateDocId = getTemplateDocIdForType(type);
  const templateFile = DriveApp.getFileById(templateDocId);
  const newDocFile = templateFile.makeCopy(`${requestId} - ${type} - Doc`, folder);

  const doc = DocumentApp.openById(newDocFile.getId());
  const body = doc.getBody();
  const replacements = {
    '{{RequestID}}': requestId,
    '{{Type}}': type,
    '{{Title}}': title || '',
    '{{CreatedAt}}': nowIso()
  };
  Object.keys(replacements).forEach(k => body.replaceText(k, safeString(replacements[k])));
  doc.saveAndClose();

  return { folder, doc, newDocFile };
}

function appendRow(sheet, valuesByHeader) {
  const headerIndex = getHeaderIndexes(sheet);
  const maxCol = sheet.getLastColumn();
  const row = new Array(maxCol).fill('');
  Object.keys(valuesByHeader).forEach(key => {
    const col = headerIndex[key];
    if (col) {
      row[col - 1] = valuesByHeader[key];
    }
  });
  sheet.appendRow(row);
}

function listRecentRequests(limit) {
  const sheet = getOrCreateMasterSheet();
  const headerIndex = getHeaderIndexes(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const startRow = Math.max(2, lastRow - (limit || 10) + 1);
  const range = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn());
  const values = range.getValues();
  const headers = Object.keys(headerIndex).sort((a, b) => headerIndex[a] - headerIndex[b]);
  const results = values.map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
  return results.reverse();
}

function setStatus(requestId, newStatus) {
  if (!CONFIG.statusValues.includes(newStatus)) {
    throw new Error('Invalid status: ' + newStatus);
  }
  const sheet = getOrCreateMasterSheet();
  const headerIndex = getHeaderIndexes(sheet);
  const idCol = headerIndex['RequestID'];
  const statusCol = headerIndex['Status'];
  const updatedCol = headerIndex['LastUpdatedAt'];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;
  const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues().map(r => r[0]);
  const idx = ids.findIndex(id => String(id) === String(requestId));
  if (idx === -1) return false;
  const rowIndex = 2 + idx;
  sheet.getRange(rowIndex, statusCol).setValue(newStatus);
  sheet.getRange(rowIndex, updatedCol).setValue(new Date());
  return true;
}

function createRequest(formData) {
  const sheet = getOrCreateMasterSheet();
  ensureHeaderRow(sheet);

  const type = safeString(formData && formData.type) || 'General';
  const title = safeString(formData && formData.title) || '';
  const requesterEmail = safeString(formData && formData.requesterEmail);
  const assignee = safeString(formData && formData.assignee);
  const dueDate = safeString(formData && formData.dueDate);

  const requestId = generateRequestId();
  const createdAt = nowIso();

  const { folder, doc } = createOrCopyDocAndFolder(requestId, type, title);

  const row = {
    RequestID: requestId,
    Status: 'Submitted',
    Type: type,
    Title: title,
    RequesterEmail: requesterEmail,
    Assignee: assignee,
    DueDate: dueDate,
    FolderLink: folder.getUrl(),
    DocLink: doc.getUrl(),
    CreatedAt: createdAt,
    LastUpdatedAt: createdAt,
    Notes: ''
  };
  appendRow(sheet, row);

  return {
    ok: true,
    requestId: requestId,
    status: 'Submitted',
    folderUrl: folder.getUrl(),
    docUrl: doc.getUrl()
  };
}