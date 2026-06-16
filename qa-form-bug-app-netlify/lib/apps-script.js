const fs = require('fs-extra');

function isAppsScriptMode() {
  return String(process.env.STORAGE_MODE || 'local').toLowerCase() === 'apps_script';
}

function getWebAppUrl() {
  const url = String(process.env.GOOGLE_APPS_SCRIPT_WEB_APP_URL || '').trim();
  if (!url) {
    throw new Error('GOOGLE_APPS_SCRIPT_WEB_APP_URL is required when STORAGE_MODE=apps_script.');
  }
  return url;
}

async function postToAppsScript(payload) {
  const response = await fetch(getWebAppUrl(), {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(payload)
  });

  const text = await response.text();
  let body = {};
  try {
    body = text ? JSON.parse(text) : {};
  } catch (error) {
    throw new Error(`Apps Script returned a non-JSON response: ${text.slice(0, 300)}`);
  }

  if (!response.ok || body.ok === false) {
    throw new Error(body.error || `Apps Script request failed with HTTP ${response.status}.`);
  }

  return body;
}

async function uploadFileToAppsScript({ localPath, originalName, mimeType, epicId, featureId, issueNo }) {
  const buffer = await fs.readFile(localPath);
  const result = await postToAppsScript({
    action: 'uploadFile',
    epicId,
    featureId,
    issueNo,
    fileName: originalName,
    mimeType: mimeType || 'application/octet-stream',
    base64: buffer.toString('base64')
  });

  return {
    storage: 'drive-apps-script',
    name: result.name || originalName,
    driveFileId: result.fileId || '',
    url: result.url || '',
    folderPath: `${epicId}/${featureId}/${issueNo}`
  };
}


async function uploadRemoteUrlToAppsScript({ url, fieldName, epicId, featureId, issueNo }) {
  const result = await postToAppsScript({
    action: 'uploadRemoteFile',
    epicId,
    featureId,
    issueNo,
    fieldName: fieldName || 'Web evidence',
    url
  });

  return {
    storage: 'drive-apps-script-remote-url',
    originalName: result.name || url,
    savedName: result.name || url,
    mimeType: result.mimeType || '',
    size: result.size || 0,
    driveFileId: result.fileId || '',
    url: result.url || url,
    sourceUrl: url,
    folderPath: `${epicId}/${featureId}/${issueNo}`
  };
}

async function syncIssueToAppsScript(issue) {
  if (!isAppsScriptMode()) return null;
  return postToAppsScript({
    action: 'upsertIssue',
    issue
  });
}

module.exports = {
  isAppsScriptMode,
  uploadFileToAppsScript,
  uploadRemoteUrlToAppsScript,
  syncIssueToAppsScript
};
