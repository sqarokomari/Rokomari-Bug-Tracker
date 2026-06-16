const fs = require('fs-extra');
const path = require('path');
const { google } = require('googleapis');
const { normalizeBool } = require('./utils');

let driveClient = null;

function isDriveMode() {
  return String(process.env.STORAGE_MODE || 'local').toLowerCase() === 'drive';
}

async function getDriveClient() {
  if (!isDriveMode()) return null;
  if (driveClient) return driveClient;

  const keyFile = process.env.GOOGLE_SERVICE_ACCOUNT_FILE;
  if (!keyFile) {
    throw new Error('GOOGLE_SERVICE_ACCOUNT_FILE is required when STORAGE_MODE=drive.');
  }

  const resolvedKeyFile = path.resolve(process.cwd(), keyFile);
  if (!(await fs.pathExists(resolvedKeyFile))) {
    throw new Error(`Google service account file not found: ${resolvedKeyFile}`);
  }

  const auth = new google.auth.GoogleAuth({
    keyFile: resolvedKeyFile,
    scopes: ['https://www.googleapis.com/auth/drive']
  });

  driveClient = google.drive({ version: 'v3', auth });
  return driveClient;
}

async function findOrCreateFolder(drive, name, parentId) {
  const escapedName = String(name).replace(/'/g, "\\'");
  const qParts = [
    `name='${escapedName}'`,
    "mimeType='application/vnd.google-apps.folder'",
    'trashed=false'
  ];
  if (parentId) qParts.push(`'${parentId}' in parents`);

  const found = await drive.files.list({
    q: qParts.join(' and '),
    fields: 'files(id, name, webViewLink)',
    pageSize: 1,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true
  });

  if (found.data.files && found.data.files.length > 0) {
    return found.data.files[0];
  }

  const created = await drive.files.create({
    requestBody: {
      name,
      mimeType: 'application/vnd.google-apps.folder',
      parents: parentId ? [parentId] : undefined
    },
    fields: 'id, name, webViewLink',
    supportsAllDrives: true
  });

  return created.data;
}

async function makeFilePublic(drive, fileId) {
  await drive.permissions.create({
    fileId,
    requestBody: {
      role: 'reader',
      type: 'anyone'
    },
    supportsAllDrives: true
  });
}

async function uploadToDrive({ localPath, originalName, mimeType, epicId, featureId, issueNo }) {
  const drive = await getDriveClient();
  const rootFolderId = process.env.GOOGLE_DRIVE_ROOT_FOLDER_ID;
  if (!rootFolderId) {
    throw new Error('GOOGLE_DRIVE_ROOT_FOLDER_ID is required when STORAGE_MODE=drive.');
  }

  const epicFolder = await findOrCreateFolder(drive, epicId, rootFolderId);
  const featureFolder = await findOrCreateFolder(drive, featureId, epicFolder.id);
  const issueFolder = await findOrCreateFolder(drive, issueNo, featureFolder.id);

  const uploaded = await drive.files.create({
    requestBody: {
      name: originalName,
      parents: [issueFolder.id]
    },
    media: {
      mimeType: mimeType || 'application/octet-stream',
      body: fs.createReadStream(localPath)
    },
    fields: 'id, name, webViewLink, webContentLink',
    supportsAllDrives: true
  });

  if (normalizeBool(process.env.GOOGLE_DRIVE_MAKE_PUBLIC)) {
    await makeFilePublic(drive, uploaded.data.id);
  }

  return {
    storage: 'drive',
    name: uploaded.data.name,
    driveFileId: uploaded.data.id,
    url: uploaded.data.webViewLink || `https://drive.google.com/file/d/${uploaded.data.id}/view`,
    folderPath: `${epicId}/${featureId}/${issueNo}`
  };
}

module.exports = { isDriveMode, uploadToDrive };
