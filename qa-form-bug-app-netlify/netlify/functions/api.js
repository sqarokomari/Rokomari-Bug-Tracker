const ExcelJS = require('exceljs');
const { randomUUID } = require('crypto');

const APPS_SCRIPT_URL = process.env.GOOGLE_APPS_SCRIPT_WEB_APP_URL || '';
const SHEET_VIEW_URL = process.env.GOOGLE_SHEET_VIEW_URL || '';

function json(statusCode, data) {
  return {
    statusCode,
    headers: {
      'Content-Type': 'application/json',
      'Cache-Control': 'no-store'
    },
    body: JSON.stringify(data)
  };
}

function error(statusCode, message) {
  return json(statusCode, { ok: false, error: message });
}

function normalizePath(event) {
  const raw = event.rawUrl || `https://local${event.path || ''}`;
  const url = new URL(raw);
  let pathname = url.pathname || '/';
  pathname = pathname.replace(/^\/\.netlify\/functions\/api/, '/api');
  if (!pathname.startsWith('/api')) pathname = `/api${pathname}`;
  return pathname.replace(/\/+$/, '') || '/api';
}

function parseBody(event) {
  if (!event.body) return {};
  const body = event.isBase64Encoded ? Buffer.from(event.body, 'base64').toString('utf8') : event.body;
  try {
    return JSON.parse(body || '{}');
  } catch (err) {
    throw new Error('Request body must be valid JSON.');
  }
}

async function postToAppsScript(payload) {
  if (!APPS_SCRIPT_URL) {
    throw new Error('GOOGLE_APPS_SCRIPT_WEB_APP_URL is not configured in Netlify environment variables.');
  }

  const response = await fetch(APPS_SCRIPT_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(payload)
  });

  const text = await response.text();
  let body = {};
  try {
    body = text ? JSON.parse(text) : {};
  } catch (err) {
    throw new Error(`Apps Script returned non-JSON: ${text.slice(0, 250)}`);
  }

  if (!response.ok || body.ok === false) {
    throw new Error(body.error || `Apps Script failed with HTTP ${response.status}.`);
  }

  return body;
}

function collectHeaders(issues) {
  const dynamicNames = [];
  for (const issue of issues) {
    for (const meta of issue.fieldMeta || []) {
      if (meta.type === 'file') continue;
      if (!dynamicNames.includes(meta.label)) dynamicNames.push(meta.label);
    }
    for (const key of Object.keys(issue.fields || {})) {
      if (!dynamicNames.includes(key)) dynamicNames.push(key);
    }
  }
  return dynamicNames;
}

async function exportIssues(event) {
  const params = event.queryStringParameters || {};
  const epicId = String(params.epicId || '').trim();
  const featureId = String(params.featureId || '').trim();

  if (!epicId || !featureId) {
    return error(400, 'Select both Epic and Feature/Task before downloading Excel.');
  }

  const result = await postToAppsScript({ action: 'getIssues' });
  const issues = (result.issues || []).filter((issue) => issue.epicId === epicId && issue.featureId === featureId);
  const dynamicNames = collectHeaders(issues);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'QA Form Bug App';
  workbook.created = new Date();
  const sheet = workbook.addWorksheet('Issues');

  const baseHeaders = [
    'Issue No', 'Title', 'Epic Name', 'Epic ID', 'Feature Name', 'Feature ID',
    'Status', 'Severity', 'Priority', 'Attachment Links', 'Created At', 'Updated At'
  ];
  const headers = [...baseHeaders, ...dynamicNames, 'Extra Fields JSON', 'Field Meta JSON', 'Attachments JSON'];
  sheet.addRow(headers);
  sheet.getRow(1).font = { bold: true };
  sheet.views = [{ state: 'frozen', ySplit: 1 }];

  for (const issue of issues) {
    const attachmentLinks = (issue.attachments || []).map((item) => item.url).filter(Boolean).join('\n');
    const row = [
      issue.issueNo, issue.title, issue.epicName, issue.epicId, issue.featureName, issue.featureId,
      issue.status, issue.severity, issue.priority, attachmentLinks, issue.createdAt, issue.updatedAt
    ];
    for (const name of dynamicNames) row.push(issue.fields?.[name] ?? '');
    row.push(JSON.stringify(issue.fields || {}));
    row.push(JSON.stringify(issue.fieldMeta || []));
    row.push(JSON.stringify(issue.attachments || []));
    sheet.addRow(row);
  }

  sheet.columns.forEach((column) => {
    let maxLength = 12;
    column.eachCell({ includeEmpty: true }, (cell) => {
      const length = String(cell.value || '').length;
      maxLength = Math.min(Math.max(maxLength, length + 2), 60);
    });
    column.width = maxLength;
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return {
    statusCode: 200,
    isBase64Encoded: true,
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="qa-issues-${epicId}-${featureId}.xlsx"`,
      'Cache-Control': 'no-store'
    },
    body: Buffer.from(buffer).toString('base64')
  };
}

function slugify(value, fallback = 'item') {
  const text = String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
  return text || fallback;
}

async function importIssues(event) {
  const body = parseBody(event);
  if (!body.base64) return error(400, 'Excel file is required.');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(Buffer.from(body.base64, 'base64'));
  const sheet = workbook.getWorksheet('Issues') || workbook.worksheets[0];
  if (!sheet) return error(400, 'No worksheet found.');

  const headerRow = sheet.getRow(1);
  const headers = [];
  headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    headers[colNumber] = String(cell.value || '').trim();
  });

  const existingResult = await postToAppsScript({ action: 'getIssues' });
  const existingNos = new Set((existingResult.issues || []).map((issue) => issue.issueNo));
  const imported = [];

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const values = {};
    headers.forEach((header, colNumber) => {
      if (!header) return;
      const cell = row.getCell(colNumber);
      values[header] = cell.text || cell.value || '';
    });

    let fields = {};
    let attachments = [];
    let fieldMeta = [];
    try { fields = values['Extra Fields JSON'] ? JSON.parse(values['Extra Fields JSON']) : {}; } catch { fields = {}; }
    try { attachments = values['Attachments JSON'] ? JSON.parse(values['Attachments JSON']) : []; } catch { attachments = []; }
    try { fieldMeta = values['Field Meta JSON'] ? JSON.parse(values['Field Meta JSON']) : []; } catch { fieldMeta = []; }

    const issueNo = values['Issue No'] || `IMPORTED-${String(imported.length + 1).padStart(4, '0')}`;
    if (existingNos.has(issueNo)) return;

    const baseHeaderSet = new Set([
      'Issue No', 'Title', 'Epic Name', 'Epic ID', 'Feature Name', 'Feature ID',
      'Status', 'Severity', 'Priority', 'Attachment Links', 'Created At', 'Updated At',
      'Extra Fields JSON', 'Field Meta JSON', 'Attachments JSON'
    ]);

    for (const [key, value] of Object.entries(values)) {
      if (!baseHeaderSet.has(key) && value !== '') fields[key] = value;
    }

    if (!fieldMeta.length) {
      fieldMeta = Object.keys(fields).map((name) => ({ label: name, type: 'text' }));
    }

    imported.push({
      id: randomUUID(),
      issueNo,
      title: values['Title'] || issueNo,
      epicName: values['Epic Name'] || '',
      epicId: values['Epic ID'] || slugify(values['Epic Name'] || 'epic'),
      featureName: values['Feature Name'] || '',
      featureId: values['Feature ID'] || slugify(values['Feature Name'] || 'feature'),
      status: values['Status'] || 'Imported',
      severity: values['Severity'] || '',
      priority: values['Priority'] || '',
      fields,
      fieldMeta,
      attachments,
      source: 'imported-excel',
      createdAt: values['Created At'] || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    });
  });

  const sync = await postToAppsScript({ action: 'importIssues', issues: imported });
  return json(200, { imported: sync.imported || imported.length, skippedDuplicates: Math.max(0, sheet.rowCount - 1 - imported.length) });
}

exports.handler = async (event) => {
  try {
    if (event.httpMethod === 'OPTIONS') return json(200, { ok: true });

    const path = normalizePath(event);
    const method = event.httpMethod;

    if (method === 'GET' && path === '/api/health') {
      return json(200, {
        ok: true,
        storageMode: 'apps_script_netlify',
        spreadsheetUrl: SHEET_VIEW_URL,
        appsScriptEnabled: Boolean(APPS_SCRIPT_URL)
      });
    }

    if (method === 'GET' && path === '/api/templates') {
      const result = await postToAppsScript({ action: 'getTemplates' });
      return json(200, result.templates || []);
    }

    if (method === 'POST' && path === '/api/templates') {
      const template = parseBody(event);
      const result = await postToAppsScript({ action: 'saveTemplate', template });
      return json(201, result.template);
    }

    const templateMatch = path.match(/^\/api\/templates\/([^/]+)$/);
    if (method === 'DELETE' && templateMatch) {
      const result = await postToAppsScript({ action: 'deleteTemplate', id: decodeURIComponent(templateMatch[1]) });
      return json(200, result);
    }

    if (method === 'GET' && path === '/api/issues') {
      const result = await postToAppsScript({ action: 'getIssues' });
      return json(200, result.issues || []);
    }

    if (method === 'POST' && path === '/api/issues') {
      const payload = parseBody(event);
      const result = await postToAppsScript({ action: 'createIssue', payload });
      return json(201, result.issue);
    }

    const issueMatch = path.match(/^\/api\/issues\/([^/]+)$/);
    if (method === 'PUT' && issueMatch) {
      const patch = parseBody(event);
      const result = await postToAppsScript({ action: 'updateIssue', id: decodeURIComponent(issueMatch[1]), patch });
      return json(200, result.issue);
    }

    if (method === 'DELETE' && issueMatch) {
      const result = await postToAppsScript({ action: 'deleteIssue', id: decodeURIComponent(issueMatch[1]) });
      return json(200, result);
    }

    if (method === 'GET' && path === '/api/export/issues.xlsx') {
      return exportIssues(event);
    }

    if (method === 'POST' && path === '/api/import/issues.xlsx') {
      return importIssues(event);
    }

    return error(404, `No route for ${method} ${path}.`);
  } catch (err) {
    console.error(err);
    return error(500, err.message || 'Internal server error.');
  }
};
