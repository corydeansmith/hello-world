const SCRIPT_PROP = PropertiesService.getScriptProperties();

const DEFAULTS = {
  trackerSpreadsheetTitle: 'Workflow Demo Tracker',
  sheetName: 'MasterTracker',
  parentFolderName: 'Workflow Demo Requests',
  statusValues: ['Submitted', 'In Review', 'Approved', 'Rejected', 'Delivered', 'Archived']
};

function doGet() {
  ensureBootstrap();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Workflow Demo')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function ensureBootstrap() {
  let spreadsheetId = SCRIPT_PROP.getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    const ss = SpreadsheetApp.create(DEFAULTS.trackerSpreadsheetTitle);
    spreadsheetId = ss.getId();
    SCRIPT_PROP.setProperty('SPREADSHEET_ID', spreadsheetId);
  }

  // Ensure sheet exists and headers are present
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let sheet = spreadsheet.getSheetByName(DEFAULTS.sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(DEFAULTS.sheetName);
  }
  ensureHeaderRow(sheet);

  let parentFolderId = SCRIPT_PROP.getProperty('PARENT_FOLDER_ID');
  if (!parentFolderId) {
    const existingFolders = DriveApp.getFoldersByName(DEFAULTS.parentFolderName);
    let folder = null;
    if (existingFolders.hasNext()) {
      folder = existingFolders.next();
    } else {
      folder = DriveApp.createFolder(DEFAULTS.parentFolderName);
    }
    SCRIPT_PROP.setProperty('PARENT_FOLDER_ID', folder.getId());
  }
}

function getSpreadsheet() {
  ensureBootstrap();
  const spreadsheetId = SCRIPT_PROP.getProperty('SPREADSHEET_ID');
  return SpreadsheetApp.openById(spreadsheetId);
}

function getMasterSheet() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(DEFAULTS.sheetName) || ss.insertSheet(DEFAULTS.sheetName);
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

function getParentFolder() {
  ensureBootstrap();
  const parentFolderId = SCRIPT_PROP.getProperty('PARENT_FOLDER_ID');
  return DriveApp.getFolderById(parentFolderId);
}

function createDocAndFolder(requestId, type, title) {
  const parentFolder = getParentFolder();
  const folderName = `${requestId} - ${type}${title ? ' - ' + title : ''}`;
  const folder = parentFolder.createFolder(folderName);

  const doc = DocumentApp.create(`${requestId} - ${type} - Doc`);
  const newDocFile = DriveApp.getFileById(doc.getId());
  folder.addFile(newDocFile);
  DriveApp.getRootFolder().removeFile(newDocFile);

  const body = doc.getBody();
  body.appendParagraph(`Request ${requestId}`);
  body.appendParagraph(`Type: ${type}`);
  body.appendParagraph(`Title: ${title || ''}`);
  body.appendParagraph(`CreatedAt: ${nowIso()}`);
  doc.saveAndClose();

  return { folder, doc };
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
  const sheet = getMasterSheet();
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
  if (!DEFAULTS.statusValues.includes(newStatus)) {
    throw new Error('Invalid status: ' + newStatus);
  }
  const sheet = getMasterSheet();
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
  const sheet = getMasterSheet();

  const type = safeString(formData && formData.type) || 'General';
  const title = safeString(formData && formData.title) || '';
  const requesterEmail = safeString(formData && formData.requesterEmail);
  const assignee = safeString(formData && formData.assignee);
  const dueDate = safeString(formData && formData.dueDate);

  const requestId = generateRequestId();
  const createdAt = nowIso();

  const { folder, doc } = createDocAndFolder(requestId, type, title);

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

function initializeProject() {
  ensureBootstrap();
}