const fs = require('fs-extra');
const path = require('path');

const dataDir = path.join(__dirname, '..', 'data');
const templatesPath = path.join(dataDir, 'templates.json');
const issuesPath = path.join(dataDir, 'issues.json');

async function ensureDataFiles() {
  await fs.ensureDir(dataDir);
  if (!(await fs.pathExists(templatesPath))) await fs.writeJson(templatesPath, [], { spaces: 2 });
  if (!(await fs.pathExists(issuesPath))) await fs.writeJson(issuesPath, [], { spaces: 2 });
}

async function readTemplates() {
  await ensureDataFiles();
  return fs.readJson(templatesPath);
}

async function writeTemplates(templates) {
  await ensureDataFiles();
  await fs.writeJson(templatesPath, templates, { spaces: 2 });
}

async function readIssues() {
  await ensureDataFiles();
  return fs.readJson(issuesPath);
}

async function writeIssues(issues) {
  await ensureDataFiles();
  await fs.writeJson(issuesPath, issues, { spaces: 2 });
}

function nextIssueNumber(issues) {
  const max = issues.reduce((highest, issue) => {
    const match = String(issue.issueNo || '').match(/ISSUE-(\d+)/i);
    if (!match) return highest;
    return Math.max(highest, Number(match[1]));
  }, 0);
  return `ISSUE-${String(max + 1).padStart(4, '0')}`;
}

module.exports = {
  ensureDataFiles,
  readTemplates,
  writeTemplates,
  readIssues,
  writeIssues,
  nextIssueNumber
};
