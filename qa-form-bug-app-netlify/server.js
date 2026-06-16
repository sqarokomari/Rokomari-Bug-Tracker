require('dotenv').config();

const path = require('path');
const fs = require('fs-extra');
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const ExcelJS = require('exceljs');
const { v4: uuidv4 } = require('uuid');

const {
  ensureDataFiles,
  readTemplates,
  writeTemplates,
  readIssues,
  writeIssues,
  nextIssueNumber
} = require('./lib/store');
const { slugify, safeFileName, toArray } = require('./lib/utils');
const { isDriveMode, uploadToDrive } = require('./lib/drive');
const { isAppsScriptMode, uploadFileToAppsScript, uploadRemoteUrlToAppsScript, syncIssueToAppsScript } = require('./lib/apps-script');

const app = express();
const PORT = Number(process.env.PORT || 3000);
const APP_BASE_URL = process.env.APP_BASE_URL || `http://localhost:${PORT}`;

const tmpDir = path.join(__dirname, 'tmp');
const uploadsDir = path.join(__dirname, 'uploads');
const importsDir = path.join(__dirname, 'imports');

fs.ensureDirSync(tmpDir);
fs.ensureDirSync(uploadsDir);
fs.ensureDirSync(importsDir);

const upload = multer({ dest: tmpDir, limits: { fileSize: 200 * 1024 * 1024 } });
const importUpload = multer({ dest: importsDir, limits: { fileSize: 50 * 1024 * 1024 } });

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use('/uploads', express.static(uploadsDir));
app.use(express.static(path.join(__dirname, 'public')));

function parseJsonField(value, fallback) {
  if (!value) return fallback;
  try {
    return JSON.parse(value);
  } catch (error) {
    return fallback;
  }
}

function localFileUrl(relativePath) {
  return `${APP_BASE_URL}/${relativePath.replace(/\\/g, '/')}`;
}


function isSafeRemoteEvidenceUrl(value) {
  try {
    const url = new URL(String(value || '').trim());
    return ['http:', 'https:'].includes(url.protocol);
  } catch {
    return false;
  }
}

async function saveRemoteAttachment(url, { fieldName, epicId, featureId, issueNo }) {
  if (!isSafeRemoteEvidenceUrl(url)) {
    throw new Error(`Invalid remote evidence URL: ${url}`);
  }

  if (isAppsScriptMode()) {
    const scriptResult = await uploadRemoteUrlToAppsScript({
      url,
      fieldName,
      epicId,
      featureId,
      issueNo
    });
    return {
      id: uuidv4(),
      fieldName,
      ...scriptResult
    };
  }

  return {
    id: uuidv4(),
    fieldName,
    storage: 'remote-url',
    originalName: url,
    savedName: url,
    mimeType: '',
    size: 0,
    url,
    sourceUrl: url,
    folderPath: `${epicId}/${featureId}/${issueNo}`
  };
}

async function saveAttachment(file, { epicId, featureId, issueNo }) {
  const ext = path.extname(file.originalname);
  const baseName = path.basename(file.originalname, ext);
  const finalName = safeFileName(`${epicId}_${featureId}_${issueNo}_${baseName}${ext}`);

  if (isAppsScriptMode()) {
    try {
      const scriptResult = await uploadFileToAppsScript({
        localPath: file.path,
        originalName: finalName,
        mimeType: file.mimetype,
        epicId,
        featureId,
        issueNo
      });
      await fs.remove(file.path);
      return {
        id: uuidv4(),
        originalName: file.originalname,
        savedName: finalName,
        mimeType: file.mimetype,
        size: file.size,
        ...scriptResult
      };
    } catch (error) {
      await fs.remove(file.path).catch(() => {});
      throw error;
    }
  }

  if (isDriveMode()) {
    try {
      const driveResult = await uploadToDrive({
        localPath: file.path,
        originalName: finalName,
        mimeType: file.mimetype,
        epicId,
        featureId,
        issueNo
      });
      await fs.remove(file.path);
      return {
        id: uuidv4(),
        originalName: file.originalname,
        savedName: finalName,
        mimeType: file.mimetype,
        size: file.size,
        ...driveResult
      };
    } catch (error) {
      await fs.remove(file.path).catch(() => {});
      throw error;
    }
  }

  const issueUploadDir = path.join(uploadsDir, epicId, featureId, issueNo);
  await fs.ensureDir(issueUploadDir);
  const destination = path.join(issueUploadDir, finalName);
  await fs.move(file.path, destination, { overwrite: true });

  const relative = path.relative(__dirname, destination).replace(/\\/g, '/');
  return {
    id: uuidv4(),
    storage: 'local',
    originalName: file.originalname,
    savedName: finalName,
    mimeType: file.mimetype,
    size: file.size,
    url: localFileUrl(relative),
    folderPath: `${epicId}/${featureId}/${issueNo}`
  };
}

function currentStorageMode() {
  if (isAppsScriptMode()) return 'apps_script';
  if (isDriveMode()) return 'drive';
  return 'local';
}

app.get('/api/health', async (_req, res) => {
  res.json({
    ok: true,
    storageMode: currentStorageMode(),
    spreadsheetUrl: process.env.GOOGLE_SHEET_VIEW_URL || '',
    appsScriptEnabled: isAppsScriptMode() && Boolean(process.env.GOOGLE_APPS_SCRIPT_WEB_APP_URL)
  });
});

app.get('/api/templates', async (_req, res, next) => {
  try {
    res.json(await readTemplates());
  } catch (error) {
    next(error);
  }
});

app.post('/api/templates', async (req, res, next) => {
  try {
    const templates = await readTemplates();
    const now = new Date().toISOString();
    const template = {
      id: req.body.id || uuidv4(),
      name: String(req.body.name || '').trim(),
      epicName: String(req.body.epicName || '').trim(),
      epicId: String(req.body.epicId || '').trim(),
      featureName: String(req.body.featureName || '').trim(),
      featureId: String(req.body.featureId || '').trim(),
      fields: Array.isArray(req.body.fields) ? req.body.fields : [],
      createdAt: req.body.createdAt || now,
      updatedAt: now
    };

    if (!template.name) return res.status(400).json({ error: 'Template name is required.' });

    const index = templates.findIndex((item) => item.id === template.id);
    if (index >= 0) templates[index] = template;
    else templates.push(template);

    await writeTemplates(templates);
    res.status(201).json(template);
  } catch (error) {
    next(error);
  }
});

app.delete('/api/templates/:id', async (req, res, next) => {
  try {
    const templates = await readTemplates();
    const filtered = templates.filter((item) => item.id !== req.params.id);
    await writeTemplates(filtered);
    res.json({ deleted: templates.length - filtered.length });
  } catch (error) {
    next(error);
  }
});

app.get('/api/issues', async (_req, res, next) => {
  try {
    const issues = await readIssues();
    res.json(issues);
  } catch (error) {
    next(error);
  }
});

app.post('/api/issues', upload.any(), async (req, res, next) => {
  try {
    const issues = await readIssues();
    const issueNo = nextIssueNumber(issues);
    const now = new Date().toISOString();

    const epicId = slugify(req.body.epicId || req.body.epicName, 'epic');
    const featureId = slugify(req.body.featureId || req.body.featureName, 'feature');
    const fields = parseJsonField(req.body.fields, {});
    const fieldMeta = parseJsonField(req.body.fieldMeta, []);

    const attachments = [];
    for (const file of req.files || []) {
      const saved = await saveAttachment(file, { epicId, featureId, issueNo });
      attachments.push({ fieldName: file.fieldname, ...saved });
    }

    const remoteAttachmentGroups = parseJsonField(req.body.remoteAttachmentUrls, []);
    for (const group of remoteAttachmentGroups) {
      const fieldName = String(group.fieldName || 'Web evidence').trim();
      const urls = Array.isArray(group.urls) ? group.urls : [];
      for (const url of urls) {
        const saved = await saveRemoteAttachment(url, { fieldName, epicId, featureId, issueNo });
        attachments.push(saved);
      }
    }

    const issue = {
      id: uuidv4(),
      issueNo,
      templateId: req.body.templateId || '',
      epicName: String(req.body.epicName || '').trim(),
      epicId,
      featureName: String(req.body.featureName || '').trim(),
      featureId,
      title: String(req.body.title || '').trim() || issueNo,
      status: String(req.body.status || 'Open').trim(),
      severity: String(req.body.severity || '').trim(),
      priority: String(req.body.priority || '').trim(),
      fields,
      fieldMeta,
      attachments,
      source: 'local-form',
      createdAt: now,
      updatedAt: now
    };

    issues.push(issue);
    await writeIssues(issues);

    await syncIssueToAppsScript(issue).catch((error) => {
      console.error('Google Sheet sync failed:', error.message);
      issue.sheetSyncError = error.message;
    });

    res.status(201).json(issue);
  } catch (error) {
    next(error);
  }
});

app.put('/api/issues/:id', async (req, res, next) => {
  try {
    const issues = await readIssues();
    const index = issues.findIndex((issue) => issue.id === req.params.id);
    if (index < 0) return res.status(404).json({ error: 'Issue not found.' });

    issues[index] = {
      ...issues[index],
      ...req.body,
      id: issues[index].id,
      issueNo: issues[index].issueNo,
      updatedAt: new Date().toISOString()
    };
    await writeIssues(issues);
    await syncIssueToAppsScript(issues[index]).catch((error) => {
      console.error('Google Sheet status sync failed:', error.message);
    });
    res.json(issues[index]);
  } catch (error) {
    next(error);
  }
});

app.delete('/api/issues/:id', async (req, res, next) => {
  try {
    const issues = await readIssues();
    const filtered = issues.filter((issue) => issue.id !== req.params.id);
    await writeIssues(filtered);
    res.json({ deleted: issues.length - filtered.length });
  } catch (error) {
    next(error);
  }
});

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

app.get('/api/export/issues.xlsx', async (req, res, next) => {
  try {
    const epicId = String(req.query.epicId || '').trim();
    const featureId = String(req.query.featureId || '').trim();

    if (!epicId || !featureId) {
      return res.status(400).json({ error: 'Select both Epic and Feature/Task before downloading Excel.' });
    }

    const allIssues = await readIssues();
    const issues = allIssues.filter((issue) => issue.epicId === epicId && issue.featureId === featureId);
    const dynamicNames = collectHeaders(issues);
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'QA Form Bug App';
    workbook.created = new Date();

    const sheet = workbook.addWorksheet('Issues');
    const baseHeaders = [
      'Issue No',
      'Title',
      'Epic Name',
      'Epic ID',
      'Feature Name',
      'Feature ID',
      'Status',
      'Severity',
      'Priority',
      'Attachment Links',
      'Created At',
      'Updated At'
    ];
    const headers = [...baseHeaders, ...dynamicNames, 'Extra Fields JSON', 'Attachments JSON'];
    sheet.addRow(headers);
    sheet.getRow(1).font = { bold: true };
    sheet.views = [{ state: 'frozen', ySplit: 1 }];

    for (const issue of issues) {
      const attachmentLinks = (issue.attachments || []).map((item) => item.url).join('\n');
      const row = [
        issue.issueNo,
        issue.title,
        issue.epicName,
        issue.epicId,
        issue.featureName,
        issue.featureId,
        issue.status,
        issue.severity,
        issue.priority,
        attachmentLinks,
        issue.createdAt,
        issue.updatedAt
      ];
      for (const name of dynamicNames) row.push(issue.fields?.[name] ?? '');
      row.push(JSON.stringify(issue.fields || {}));
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

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="qa-issues-${epicId}-${featureId}.xlsx"`);
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    next(error);
  }
});

app.post('/api/import/issues.xlsx', importUpload.single('report'), async (req, res, next) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Excel file is required.' });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(req.file.path);
    const sheet = workbook.getWorksheet('Issues') || workbook.worksheets[0];
    if (!sheet) return res.status(400).json({ error: 'No worksheet found.' });

    const headerRow = sheet.getRow(1);
    const headers = [];
    headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      headers[colNumber] = String(cell.value || '').trim();
    });

    const issues = await readIssues();
    const existingNos = new Set(issues.map((issue) => issue.issueNo));
    const imported = [];

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const values = {};
      headers.forEach((header, colNumber) => {
        if (!header) return;
        const cell = row.getCell(colNumber);
        if (cell.text) values[header] = cell.text;
        else values[header] = cell.value ?? '';
      });

      const fields = values['Extra Fields JSON'] ? (() => {
        try { return JSON.parse(values['Extra Fields JSON']); } catch { return {}; }
      })() : {};
      const attachments = values['Attachments JSON'] ? (() => {
        try { return JSON.parse(values['Attachments JSON']); } catch { return []; }
      })() : [];

      const issueNo = values['Issue No'] || nextIssueNumber([...issues, ...imported]);
      if (existingNos.has(issueNo)) return;

      const baseHeaderSet = new Set([
        'Issue No', 'Title', 'Epic Name', 'Epic ID', 'Feature Name', 'Feature ID',
        'Status', 'Severity', 'Priority', 'Attachment Links', 'Created At', 'Updated At',
        'Extra Fields JSON', 'Attachments JSON'
      ]);
      for (const [key, value] of Object.entries(values)) {
        if (!baseHeaderSet.has(key) && value !== '') fields[key] = value;
      }

      imported.push({
        id: uuidv4(),
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
        fieldMeta: Object.keys(fields).map((name) => ({ label: name, type: 'text' })),
        attachments,
        source: 'imported-excel',
        createdAt: values['Created At'] || new Date().toISOString(),
        updatedAt: new Date().toISOString()
      });
    });

    await writeIssues([...issues, ...imported]);
    await fs.remove(req.file.path);
    res.json({ imported: imported.length, skippedDuplicates: Math.max(0, sheet.rowCount - 1 - imported.length) });
  } catch (error) {
    if (req.file) await fs.remove(req.file.path).catch(() => {});
    next(error);
  }
});

app.get('*', (_req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.use((error, _req, res, _next) => {
  console.error(error);
  res.status(500).json({ error: error.message || 'Internal server error.' });
});

ensureDataFiles().then(() => {
  app.listen(PORT, () => {
    console.log(`QA Form Bug App running at ${APP_BASE_URL}`);
    console.log(`Storage mode: ${currentStorageMode()}`);
  });
});
