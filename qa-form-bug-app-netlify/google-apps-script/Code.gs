/*******************************************************
 * QA Form Bug App - Google Apps Script bridge for Netlify
 *
 * Deploy as: Web app
 * Execute as: Me
 * Who has access: Anyone with the link
 *
 * Netlify calls this script server-side. This script stores:
 * - uploaded evidence in Google Drive
 * - issues in the Issues sheet
 * - reusable form presets in the Templates sheet
 *******************************************************/

const SHEET_ID = '1JlJtBq3GlsEG1Rc9cwTLxcXDAwpvoX2C2Ld8fqEr6u0';
const SHEET_TAB_NAME = 'Issues';
const TEMPLATE_TAB_NAME = 'Templates';
const DRIVE_ROOT_FOLDER_ID = '1VA4Awn12PKmMc1VIEdimK-qwYBieU0yC';

const MAKE_UPLOADED_FILES_ANYONE_WITH_LINK = false;

const BASE_HEADERS = [
  'Issue No',
  'ID',
  'Template ID',
  'Title',
  'Epic Name',
  'Epic ID',
  'Feature Name',
  'Feature ID',
  'Issue Date',
  'Reported To',
  'Reported By',
  'Status',
  'Severity',
  'Priority',
  'Attachment Links',
  'Created At',
  'Updated At',
  'Source',
  'Extra Fields JSON',
  'Field Meta JSON',
  'Attachments JSON'
];

const TEMPLATE_HEADERS = [
  'ID',
  'Name',
  'Epic Name',
  'Epic ID',
  'Feature Name',
  'Feature ID',
  'Fields JSON',
  'Created At',
  'Updated At'
];

function doPost(e) {
  try {
    const payload = JSON.parse((e.postData && e.postData.contents) || '{}');

    switch (payload.action) {
      case 'getTemplates': return jsonResponse({ ok: true, templates: getTemplates() });
      case 'saveTemplate': return jsonResponse({ ok: true, template: saveTemplate(payload.template) });
      case 'deleteTemplate': return jsonResponse(deleteTemplate(payload.id));
      case 'getIssues': return jsonResponse({ ok: true, issues: getIssues() });
      case 'createIssue': return jsonResponse({ ok: true, issue: createIssue(payload.payload) });
      case 'updateIssue': return jsonResponse({ ok: true, issue: updateIssue(payload.id, payload.patch || {}) });
      case 'deleteIssue': return jsonResponse(deleteIssue(payload.id));
      case 'importIssues': return jsonResponse(importIssues(payload.issues || []));
      case 'uploadFile': return jsonResponse(uploadFile(payload));
      case 'uploadRemoteFile': return jsonResponse(uploadRemoteFile(payload));
      case 'upsertIssue': return jsonResponse(upsertIssue(payload.issue));
      default: return jsonResponse({ ok: false, error: 'Unknown action.' });
    }
  } catch (error) {
    return jsonResponse({ ok: false, error: error.message || String(error) });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function cleanName(value, fallback) {
  const text = String(value || fallback || 'unknown').trim();
  return text.replace(/[\\/:*?"<>|#%{}~]/g, '-').slice(0, 150) || fallback || 'unknown';
}

function slugify(value, fallback) {
  const text = String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
  return text || fallback || 'item';
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function getOrCreateSheet(name, headers) {
  const spreadsheet = getSpreadsheet();
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  ensureHeaders(sheet, headers, []);
  return sheet;
}

function ensureHeaders(sheet, baseHeaders, dynamicFieldNames) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  let headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].filter(String);

  if (headers.length === 0) {
    headers = baseHeaders.slice();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  const requiredHeaders = baseHeaders.concat(dynamicFieldNames || []);
  const missing = requiredHeaders.filter((header) => headers.indexOf(header) === -1);

  if (missing.length) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    sheet.getRange(1, 1, 1, headers.length + missing.length).setFontWeight('bold');
    headers = headers.concat(missing);
  }

  return headers;
}

function getIssueSheet() {
  return getOrCreateSheet(SHEET_TAB_NAME, BASE_HEADERS);
}

function getTemplateSheet() {
  return getOrCreateSheet(TEMPLATE_TAB_NAME, TEMPLATE_HEADERS);
}

function getHeaders(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);
}

function rowToObject(headers, row) {
  const obj = {};
  headers.forEach((header, index) => {
    if (!header) return;
    obj[header] = row[index] == null ? '' : row[index];
  });
  return obj;
}

function parseJson(value, fallback) {
  try {
    return value ? JSON.parse(String(value)) : fallback;
  } catch (error) {
    return fallback;
  }
}

function findRowByHeaderValue(sheet, headerName, value) {
  const headers = getHeaders(sheet);
  const index = headers.indexOf(headerName);
  if (index < 0) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const values = sheet.getRange(2, index + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i += 1) {
    if (String(values[i][0]) === String(value)) return i + 2;
  }
  return 0;
}

function findIssueRow(sheet, issue) {
  if (!issue) return 0;
  if (issue.id) {
    const byId = findRowByHeaderValue(sheet, 'ID', issue.id);
    if (byId) return byId;
  }
  if (issue.issueNo) return findRowByHeaderValue(sheet, 'Issue No', issue.issueNo);
  return 0;
}

function nextIssueNumber() {
  const sheet = getIssueSheet();
  const headers = getHeaders(sheet);
  const issueNoIndex = headers.indexOf('Issue No');
  const lastRow = sheet.getLastRow();
  if (issueNoIndex < 0 || lastRow < 2) return 'ISSUE-0001';

  const values = sheet.getRange(2, issueNoIndex + 1, lastRow - 1, 1).getValues();
  let max = 0;
  values.forEach((row) => {
    const match = String(row[0] || '').match(/(\d+)/g);
    if (!match) return;
    const num = Number(match[match.length - 1]);
    if (Number.isFinite(num)) max = Math.max(max, num);
  });
  return `ISSUE-${String(max + 1).padStart(4, '0')}`;
}

function getOrCreateChildFolder(parentFolder, name) {
  const safeName = cleanName(name, 'folder');
  const existing = parentFolder.getFoldersByName(safeName);
  if (existing.hasNext()) return existing.next();
  return parentFolder.createFolder(safeName);
}

function getEvidenceFolder(payload) {
  const root = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
  const epicFolder = getOrCreateChildFolder(root, payload.epicId || 'epic');
  const featureFolder = getOrCreateChildFolder(epicFolder, payload.featureId || 'feature');
  return getOrCreateChildFolder(featureFolder, payload.issueNo || 'issue');
}

function guessFileNameFromUrl(url, fallback) {
  try {
    const withoutQuery = String(url || '').split('?')[0].split('#')[0];
    const rawName = withoutQuery.split('/').filter(Boolean).pop() || '';
    const name = decodeURIComponent(rawName);
    return cleanName(name || fallback || 'web-evidence', 'web-evidence');
  } catch (error) {
    return cleanName(fallback || 'web-evidence', 'web-evidence');
  }
}

function uploadFile(payload) {
  if (!payload.base64) throw new Error('base64 file content is required.');

  const issueFolder = getEvidenceFolder(payload);
  const bytes = Utilities.base64Decode(payload.base64);
  const fileName = cleanName(payload.fileName, 'attachment');
  const blob = Utilities.newBlob(bytes, payload.mimeType || 'application/octet-stream', fileName);
  const file = issueFolder.createFile(blob);

  if (MAKE_UPLOADED_FILES_ANYONE_WITH_LINK) {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }

  return {
    ok: true,
    fileId: file.getId(),
    name: file.getName(),
    url: file.getUrl(),
    folderPath: `${payload.epicId || 'epic'}/${payload.featureId || 'feature'}/${payload.issueNo || 'issue'}`
  };
}

function uploadRemoteFile(payload) {
  if (!payload.url) throw new Error('Remote URL is required.');

  const issueFolder = getEvidenceFolder(payload);
  const response = UrlFetchApp.fetch(payload.url, {
    followRedirects: true,
    muteHttpExceptions: true,
    headers: { 'User-Agent': 'QA-Form-Bug-App/1.0' }
  });

  const statusCode = response.getResponseCode();
  if (statusCode < 200 || statusCode >= 300) {
    throw new Error(`Could not fetch remote evidence URL. HTTP ${statusCode}`);
  }

  const blob = response.getBlob();
  const fileName = guessFileNameFromUrl(payload.url, `${payload.fieldName || 'web-evidence'}-${Date.now()}`);
  blob.setName(fileName);
  const file = issueFolder.createFile(blob);

  if (MAKE_UPLOADED_FILES_ANYONE_WITH_LINK) {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }

  return {
    ok: true,
    fileId: file.getId(),
    name: file.getName(),
    mimeType: blob.getContentType(),
    size: blob.getBytes().length,
    url: file.getUrl(),
    sourceUrl: payload.url,
    folderPath: `${payload.epicId || 'epic'}/${payload.featureId || 'feature'}/${payload.issueNo || 'issue'}`
  };
}

function getTemplates() {
  const sheet = getTemplateSheet();
  const headers = getHeaders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, headers.length).getValues().map((row) => {
    const obj = rowToObject(headers, row);
    return {
      id: String(obj['ID'] || ''),
      name: String(obj['Name'] || ''),
      epicName: String(obj['Epic Name'] || ''),
      epicId: String(obj['Epic ID'] || ''),
      featureName: String(obj['Feature Name'] || ''),
      featureId: String(obj['Feature ID'] || ''),
      fields: parseJson(obj['Fields JSON'], []),
      createdAt: String(obj['Created At'] || ''),
      updatedAt: String(obj['Updated At'] || '')
    };
  }).filter((item) => item.id && item.name);
}

function saveTemplate(template) {
  if (!template) throw new Error('Template payload is required.');
  const now = new Date().toISOString();
  const item = {
    id: String(template.id || Utilities.getUuid()),
    name: String(template.name || '').trim(),
    epicName: String(template.epicName || '').trim(),
    epicId: String(template.epicId || slugify(template.epicName, 'epic')).trim(),
    featureName: String(template.featureName || '').trim(),
    featureId: String(template.featureId || slugify(template.featureName, 'feature')).trim(),
    fields: Array.isArray(template.fields) ? template.fields : [],
    createdAt: template.createdAt || now,
    updatedAt: now
  };
  if (!item.name) throw new Error('Template name is required.');

  const sheet = getTemplateSheet();
  const headers = getHeaders(sheet);
  const row = headers.map((header) => {
    switch (header) {
      case 'ID': return item.id;
      case 'Name': return item.name;
      case 'Epic Name': return item.epicName;
      case 'Epic ID': return item.epicId;
      case 'Feature Name': return item.featureName;
      case 'Feature ID': return item.featureId;
      case 'Fields JSON': return JSON.stringify(item.fields);
      case 'Created At': return item.createdAt;
      case 'Updated At': return item.updatedAt;
      default: return '';
    }
  });

  const existingRow = findRowByHeaderValue(sheet, 'ID', item.id);
  if (existingRow) sheet.getRange(existingRow, 1, 1, headers.length).setValues([row]);
  else sheet.appendRow(row);
  return item;
}

function deleteTemplate(id) {
  const sheet = getTemplateSheet();
  const row = findRowByHeaderValue(sheet, 'ID', id);
  if (row) sheet.deleteRow(row);
  return { ok: true, deleted: row ? 1 : 0 };
}

function issueFromRow(headers, row) {
  const obj = rowToObject(headers, row);
  const fields = parseJson(obj['Extra Fields JSON'], {});
  const attachments = parseJson(obj['Attachments JSON'], []);
  const fieldMeta = parseJson(obj['Field Meta JSON'], []);

  const baseSet = new Set(BASE_HEADERS);
  headers.forEach((header) => {
    if (!header || baseSet.has(header)) return;
    if (obj[header] !== '' && obj[header] != null && fields[header] == null) fields[header] = obj[header];
  });

  return {
    id: String(obj['ID'] || obj['Issue No'] || Utilities.getUuid()),
    issueNo: String(obj['Issue No'] || ''),
    templateId: String(obj['Template ID'] || ''),
    title: String(obj['Title'] || ''),
    epicName: String(obj['Epic Name'] || ''),
    epicId: String(obj['Epic ID'] || ''),
    featureName: String(obj['Feature Name'] || ''),
    featureId: String(obj['Feature ID'] || ''),
    issueDate: String(obj['Issue Date'] || obj['Created At'] || ''),
    reportedTo: String(obj['Reported To'] || ''),
    reportedBy: String(obj['Reported By'] || ''),
    status: String(obj['Status'] || ''),
    severity: String(obj['Severity'] || ''),
    priority: String(obj['Priority'] || ''),
    fields,
    fieldMeta,
    attachments,
    source: String(obj['Source'] || ''),
    createdAt: String(obj['Created At'] || ''),
    updatedAt: String(obj['Updated At'] || '')
  };
}

function getIssues() {
  const sheet = getIssueSheet();
  const headers = getHeaders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, headers.length).getValues()
    .map((row) => issueFromRow(headers, row))
    .filter((issue) => issue.issueNo);
}

function upsertIssue(issue) {
  if (!issue || !issue.issueNo) throw new Error('issue.issueNo is required.');

  const fields = issue.fields || {};
  const dynamicFieldNames = Object.keys(fields);
  const sheet = getIssueSheet();
  const headers = ensureHeaders(sheet, BASE_HEADERS, dynamicFieldNames);
  const attachments = issue.attachments || [];
  const attachmentLinks = attachments.map((item) => item.url || '').filter(Boolean).join('\n');

  const row = headers.map((header) => {
    switch (header) {
      case 'Issue No': return issue.issueNo || '';
      case 'ID': return issue.id || '';
      case 'Template ID': return issue.templateId || '';
      case 'Title': return issue.title || '';
      case 'Epic Name': return issue.epicName || '';
      case 'Epic ID': return issue.epicId || '';
      case 'Feature Name': return issue.featureName || '';
      case 'Feature ID': return issue.featureId || '';
      case 'Issue Date': return issue.issueDate || '';
      case 'Reported To': return issue.reportedTo || '';
      case 'Reported By': return issue.reportedBy || '';
      case 'Status': return issue.status || '';
      case 'Severity': return issue.severity || '';
      case 'Priority': return issue.priority || '';
      case 'Attachment Links': return attachmentLinks;
      case 'Created At': return issue.createdAt || '';
      case 'Updated At': return issue.updatedAt || '';
      case 'Source': return issue.source || '';
      case 'Extra Fields JSON': return JSON.stringify(fields);
      case 'Field Meta JSON': return JSON.stringify(issue.fieldMeta || []);
      case 'Attachments JSON': return JSON.stringify(attachments);
      default: return fields[header] || '';
    }
  });

  const existingRow = findIssueRow(sheet, issue);
  if (existingRow) {
    sheet.getRange(existingRow, 1, 1, headers.length).setValues([row]);
    return { ok: true, mode: 'updated', row: existingRow };
  }

  sheet.appendRow(row);
  return { ok: true, mode: 'inserted', row: sheet.getLastRow() };
}

function createIssue(payload) {
  if (!payload) throw new Error('Issue payload is required.');
  const now = new Date().toISOString();
  const issueNo = nextIssueNumber();
  const epicId = slugify(payload.epicId || payload.epicName, 'epic');
  const featureId = slugify(payload.featureId || payload.featureName, 'feature');
  const attachments = [];

  (payload.files || []).forEach((filePayload) => {
    const saved = uploadFile({
      epicId,
      featureId,
      issueNo,
      fileName: filePayload.fileName,
      mimeType: filePayload.mimeType,
      base64: filePayload.base64
    });
    attachments.push({
      id: Utilities.getUuid(),
      fieldName: filePayload.fieldName || 'Attachment',
      storage: 'drive-apps-script',
      originalName: filePayload.fileName || saved.name,
      savedName: saved.name,
      mimeType: filePayload.mimeType || '',
      size: filePayload.size || 0,
      driveFileId: saved.fileId,
      url: saved.url,
      folderPath: saved.folderPath
    });
  });

  (payload.remoteAttachmentUrls || []).forEach((group) => {
    const fieldName = String(group.fieldName || 'Web evidence').trim();
    (group.urls || []).forEach((url) => {
      const saved = uploadRemoteFile({ epicId, featureId, issueNo, fieldName, url });
      attachments.push({
        id: Utilities.getUuid(),
        fieldName,
        storage: 'drive-apps-script-remote-url',
        originalName: saved.name || url,
        savedName: saved.name || url,
        mimeType: saved.mimeType || '',
        size: saved.size || 0,
        driveFileId: saved.fileId,
        url: saved.url || url,
        sourceUrl: url,
        folderPath: saved.folderPath
      });
    });
  });

  const issue = {
    id: Utilities.getUuid(),
    issueNo,
    templateId: payload.templateId || '',
    epicName: String(payload.epicName || '').trim(),
    epicId,
    featureName: String(payload.featureName || '').trim(),
    featureId,
    title: String(payload.title || '').trim() || issueNo,
    issueDate: String(payload.issueDate || now).trim(),
    reportedTo: String(payload.reportedTo || '').trim(),
    reportedBy: String(payload.reportedBy || '').trim(),
    status: String(payload.status || 'Open').trim(),
    severity: String(payload.severity || '').trim(),
    priority: String(payload.priority || '').trim(),
    fields: payload.fields || {},
    fieldMeta: payload.fieldMeta || [],
    attachments,
    source: 'netlify-form',
    createdAt: now,
    updatedAt: now
  };

  upsertIssue(issue);
  return issue;
}

function updateIssue(id, patch) {
  const issues = getIssues();
  const issue = issues.find((item) => String(item.id) === String(id) || String(item.issueNo) === String(id));
  if (!issue) throw new Error('Issue not found.');

  const updated = Object.assign({}, issue, patch || {}, {
    id: issue.id,
    issueNo: issue.issueNo,
    updatedAt: new Date().toISOString()
  });
  upsertIssue(updated);
  return updated;
}

function deleteIssue(id) {
  const sheet = getIssueSheet();
  const issues = getIssues();
  const issue = issues.find((item) => String(item.id) === String(id) || String(item.issueNo) === String(id));
  if (!issue) return { ok: true, deleted: 0 };
  const row = findIssueRow(sheet, issue);
  if (row) sheet.deleteRow(row);
  return { ok: true, deleted: row ? 1 : 0 };
}

function importIssues(issues) {
  let imported = 0;
  (issues || []).forEach((issue) => {
    if (!issue.issueNo) return;
    upsertIssue(issue);
    imported += 1;
  });
  return { ok: true, imported };
}
