const state = {
  templates: [],
  issues: [],
  // fields = builder draft fields. These are not used by Submit Issue until the preset is saved/selected.
  fields: [],
  selectedTemplateId: '',
  activeTemplate: null,
  selectedIssue: null,
  spreadsheetUrl: '',
  editingFieldIndex: null,
  remoteAttachmentUrls: {}
};

const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => Array.from(document.querySelectorAll(selector));

function toast(message) {
  const el = $('#toast');
  el.textContent = message;
  el.classList.remove('hidden');
  setTimeout(() => el.classList.add('hidden'), 2600);
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function slugify(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function getFieldTypeLabel(type) {
  const labels = {
    text: 'Text box',
    textarea: 'Long text',
    dropdown: 'Dropdown',
    checkbox: 'Check box',
    file: 'Document / Image Upload',
    datetime: 'Date & Time'
  };
  return labels[type] || type || 'Text box';
}

function formatDateTimeValue(value) {
  if (!value) return '';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString();
}

function formatDateTimeInputValue(value) {
  if (!value) return '';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return '';
  const pad = (num) => String(num).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function setDefaultIssueDate() {
  const input = $('#issueDate');
  if (!input || input.value) return;
  input.value = formatDateTimeInputValue(new Date().toISOString());
}

function getSeverityTone(value) {
  const v = String(value || '').toLowerCase();
  if (v === 'critical') return 'danger';
  if (v === 'high') return 'warning';
  if (v === 'medium') return 'caution';
  if (v === 'low') return 'success';
  return 'neutral';
}

function getPriorityTone(value) {
  const v = String(value || '').toUpperCase();
  if (v === 'P1') return 'danger';
  if (v === 'P2') return 'warning';
  if (v === 'P3') return 'caution';
  if (v === 'P4') return 'success';
  return 'neutral';
}

function getStatusTone(value) {
  const v = String(value || '').toLowerCase();
  if (['open', 'reopened'].includes(v)) return 'danger';
  if (['in progress', 'ready for retest'].includes(v)) return 'warning';
  if (['fixed', 'closed'].includes(v)) return 'success';
  if (['rejected', 'duplicate'].includes(v)) return 'neutral';
  return 'info';
}

function renderBadge(text, tone) {
  return `<span class="tag tag-${tone}">${escapeHtml(text || '-')}</span>`;
}

function countBy(issues, key, orderedValues = []) {
  const map = new Map();
  for (const value of orderedValues) map.set(value, 0);
  for (const issue of issues) {
    const raw = issue[key] || 'Unspecified';
    map.set(raw, (map.get(raw) || 0) + 1);
  }
  return Array.from(map.entries()).filter(([, count]) => count > 0 || orderedValues.length > 0);
}

function renderHorizontalChart(containerId, title, items, type) {
  const el = $(containerId);
  if (!el) return;
  const total = items.reduce((sum, [, count]) => sum + count, 0);
  if (!total) {
    el.innerHTML = `<div class="chart-placeholder">No ${title.toLowerCase()} data for the selected issues.</div>`;
    return;
  }

  el.innerHTML = items.map(([label, count]) => {
    const percent = Math.round((count / total) * 100);
    const tone = type === 'severity' ? getSeverityTone(label) : getPriorityTone(label);
    return `
      <div class="chart-row">
        <div class="chart-row-head">
          <span class="chart-label">${escapeHtml(label)}</span>
          <span class="chart-value">${count} (${percent}%)</span>
        </div>
        <div class="chart-bar-track"><div class="chart-bar tone-${tone}" style="width:${percent}%"></div></div>
      </div>`;
  }).join('');
}

function renderStatusPie(issues) {
  const el = $('#statusPieChart');
  if (!el) return;
  const statusOrder = ['Open', 'In Progress', 'Ready for Retest', 'Fixed', 'Closed', 'Reopened', 'Rejected', 'Duplicate', 'Imported'];
  const items = countBy(issues, 'status', statusOrder);
  const total = items.reduce((sum, [, count]) => sum + count, 0);
  if (!total) {
    el.innerHTML = '<div class="chart-placeholder">No status data for the selected issues.</div>';
    return;
  }

  const colorMap = {
    danger: '#ef4444',
    warning: '#f59e0b',
    caution: '#eab308',
    success: '#22c55e',
    neutral: '#64748b',
    info: '#3b82f6'
  };

  let start = 0;
  const segments = [];
  const legend = [];
  for (const [label, count] of items) {
    const percent = (count / total) * 100;
    const end = start + percent;
    const tone = getStatusTone(label);
    const color = colorMap[tone] || colorMap.info;
    segments.push(`${color} ${start}% ${end}%`);
    legend.push(`<div class="pie-legend-item"><span class="dot tone-${tone}"></span><span>${escapeHtml(label)}</span><strong>${count}</strong></div>`);
    start = end;
  }

  el.innerHTML = `
    <div class="pie-chart-wrap">
      <div class="pie-chart" style="background: conic-gradient(${segments.join(', ')})"></div>
      <div class="pie-chart-center">${total}<small>issues</small></div>
    </div>
    <div class="pie-legend">${legend.join('')}</div>`;
}

function renderReportStats() {
  const issues = getFilteredIssues();
  if (!isReportScopeSelected()) {
    ['#severityChart', '#priorityChart', '#statusPieChart'].forEach((selector) => {
      const el = $(selector);
      if (el) el.innerHTML = '<div class="chart-placeholder">Select an Epic and Feature/Task to see statistics.</div>';
    });
    const summary = $('#issueCountSummary');
    if (summary) summary.textContent = 'Total issues: 0';
    return;
  }

  renderHorizontalChart('#severityChart', 'Severity', countBy(issues, 'severity', ['Critical', 'High', 'Medium', 'Low']), 'severity');
  renderHorizontalChart('#priorityChart', 'Priority', countBy(issues, 'priority', ['P1', 'P2', 'P3', 'P4']), 'priority');
  renderStatusPie(issues);

  const summary = $('#issueCountSummary');
  if (summary) summary.textContent = `Total issues: ${issues.length}`;
}

function applyTheme(theme) {
  const normalizedTheme = theme === 'dark' ? 'dark' : 'light';
  const isDark = normalizedTheme === 'dark';
  document.body.classList.toggle('dark-mode', isDark);
  document.documentElement.dataset.theme = normalizedTheme;

  const btn = $('#themeToggleBtn');
  if (btn) {
    btn.textContent = isDark ? '☀️ Light' : '🌙 Dark';
    btn.setAttribute('aria-label', isDark ? 'Switch to light mode' : 'Switch to dark mode');
  }

  localStorage.setItem('qaIssueFormTheme', normalizedTheme);
}

function initTheme() {
  const savedTheme = localStorage.getItem('qaIssueFormTheme');
  const systemPrefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
  applyTheme(savedTheme || (systemPrefersDark ? 'dark' : 'light'));

  const btn = $('#themeToggleBtn');
  if (btn) {
    btn.addEventListener('click', () => {
      const nextTheme = document.body.classList.contains('dark-mode') ? 'light' : 'dark';
      applyTheme(nextTheme);
    });
  }
}

async function api(path, options = {}) {
  const response = await fetch(path, options);
  if (!response.ok) {
    let message = response.statusText;
    try {
      const body = await response.json();
      message = body.error || message;
    } catch {}
    throw new Error(message);
  }
  return response.json();
}

function getPresetValues() {
  return {
    name: $('#templateName').value.trim(),
    epicName: $('#epicName').value.trim(),
    epicId: $('#epicId').value.trim() || slugify($('#epicName').value) || 'epic',
    featureName: $('#featureName').value.trim(),
    featureId: $('#featureId').value.trim() || slugify($('#featureName').value) || 'feature',
    fields: state.fields
  };
}

function cloneTemplate(template) {
  return template ? JSON.parse(JSON.stringify(template)) : null;
}

function getSubmitFields() {
  return Array.isArray(state.activeTemplate?.fields) ? state.activeTemplate.fields : [];
}

function setPresetValues(template, options = { activate: true }) {
  $('#templateName').value = template?.name || '';
  $('#epicName').value = template?.epicName || '';
  $('#epicId').value = template?.epicId || '';
  $('#featureName').value = template?.featureName || '';
  $('#featureId').value = template?.featureId || '';
  state.fields = Array.isArray(template?.fields) ? JSON.parse(JSON.stringify(template.fields)) : [];
  state.editingFieldIndex = null;
  if ($('#addFieldBtn')) $('#addFieldBtn').textContent = 'Add Field';
  $('#cancelEditFieldBtn')?.classList.add('hidden');

  if (options.activate) {
    state.activeTemplate = cloneTemplate(template);
  }

  renderFieldList();
  renderDynamicForm();
  updateFieldOptionsVisibility();
  updateActiveContext();
  renderPresetPreview();
  setDefaultIssueDate();
  updateSubmitPresetNotice();
}

function updateActiveContext() {
  const active = state.activeTemplate;
  const epic = active
    ? `${active.epicName || '-'} (${active.epicId || 'epic'})`
    : 'No saved preset selected';
  const feature = active
    ? `${active.featureName || '-'} (${active.featureId || 'feature'})`
    : 'No saved preset selected';
  $('#activeEpic').textContent = epic;
  $('#activeFeature').textContent = feature;
}

function updateSubmitPresetNotice() {
  const notice = $('#submitPresetNotice');
  if (!notice) return;

  if (!state.activeTemplate) {
    notice.textContent = 'Save or load a preset before submitting an issue. Unsaved custom fields do not appear in the Submit Issue form.';
    notice.classList.remove('hidden');
    return;
  }

  const builderSnapshot = JSON.stringify(getPresetValues());
  const activeSnapshot = JSON.stringify({
    name: state.activeTemplate.name,
    epicName: state.activeTemplate.epicName,
    epicId: state.activeTemplate.epicId,
    featureName: state.activeTemplate.featureName,
    featureId: state.activeTemplate.featureId,
    fields: state.activeTemplate.fields || []
  });

  if (builderSnapshot !== activeSnapshot) {
    notice.textContent = 'You have unsaved preset changes. Click Save Preset to apply them to the Submit Issue form.';
    notice.classList.remove('hidden');
  } else {
    notice.classList.add('hidden');
  }
}

function renderTemplateSelect() {
  const select = $('#templateSelect');
  select.innerHTML = '<option value="">-- No preset selected --</option>';
  for (const template of state.templates) {
    const option = document.createElement('option');
    option.value = template.id;
    option.textContent = `${template.name} — ${template.epicName || 'No epic'} / ${template.featureName || 'No feature'}`;
    select.appendChild(option);
  }
  select.value = state.selectedTemplateId;
}

function renderFieldList() {
  const el = $('#fieldList');
  if (state.fields.length === 0) {
    el.className = 'field-list empty';
    el.textContent = 'No custom fields yet.';
    renderPresetPreview();
    return;
  }

  el.className = 'field-list';
  el.innerHTML = state.fields.map((field, index) => `
    <div class="field-item ${state.editingFieldIndex === index ? 'editing' : ''}">
      <div>
        <strong>${escapeHtml(field.label)}</strong>
        <div class="field-meta">${escapeHtml(getFieldTypeLabel(field.type))}${field.options?.length ? ` · ${escapeHtml(field.options.join(', '))}` : ''}</div>
      </div>
      <button type="button" data-edit-field="${index}" class="ghost">Edit</button>
      <button type="button" data-move-up="${index}" class="ghost">↑</button>
      <button type="button" data-remove-field="${index}" class="danger ghost">Remove</button>
    </div>
  `).join('');

  $$('[data-edit-field]').forEach((button) => {
    button.addEventListener('click', () => startEditField(Number(button.dataset.editField)));
  });

  $$('[data-remove-field]').forEach((button) => {
    button.addEventListener('click', () => {
      const index = Number(button.dataset.removeField);
      state.fields.splice(index, 1);
      if (state.editingFieldIndex === index) resetFieldEditor();
      if (state.editingFieldIndex !== null && state.editingFieldIndex > index) state.editingFieldIndex -= 1;
      renderFieldList();
      updateSubmitPresetNotice();
    });
  });

  $$('[data-move-up]').forEach((button) => {
    button.addEventListener('click', () => {
      const index = Number(button.dataset.moveUp);
      if (index === 0) return;
      const [field] = state.fields.splice(index, 1);
      state.fields.splice(index - 1, 0, field);
      if (state.editingFieldIndex === index) state.editingFieldIndex = index - 1;
      else if (state.editingFieldIndex === index - 1) state.editingFieldIndex = index;
      renderFieldList();
      updateSubmitPresetNotice();
    });
  });

  renderPresetPreview();
}

function renderPresetPreview() {
  const wrapper = $('#presetPreview');
  if (!wrapper) return;

  const preset = getPresetValues();
  const fields = state.fields || [];

  if (!preset.name && !preset.epicName && !preset.featureName && fields.length === 0) {
    wrapper.innerHTML = '<p class="field-meta">Start typing preset details or add fields to see the draft preview here.</p>';
    return;
  }

  const fieldPreview = fields.length ? fields.map((field, index) => {
    const id = `preview_field_${index}`;
    const label = `<label for="${id}">${escapeHtml(field.label)}</label>`;

    if (field.type === 'textarea') {
      return `<div class="form-row">${label}<textarea id="${id}" disabled placeholder="${escapeHtml(field.label)}"></textarea></div>`;
    }

    if (field.type === 'dropdown') {
      const options = (field.options || []).map((opt) => `<option>${escapeHtml(opt)}</option>`).join('');
      return `<div class="form-row">${label}<select id="${id}" disabled><option>-- Select --</option>${options}</select></div>`;
    }

    if (field.type === 'checkbox') {
      const options = field.options?.length ? field.options : ['Yes'];
      const checks = options.map((opt) => `
        <label class="checkbox-item">
          <input type="checkbox" disabled />
          ${escapeHtml(opt)}
        </label>`).join('');
      return `<div class="form-row"><label>${escapeHtml(field.label)}</label><div class="checkbox-group">${checks}</div></div>`;
    }

    if (field.type === 'file') {
      return `<div class="form-row"><label>${escapeHtml(field.label)}</label><div class="drop-zone preview-drop-zone"><span>Drag & drop files here, or click to select</span></div></div>`;
    }

    if (field.type === 'datetime') {
      return `<div class="form-row">${label}<input id="${id}" type="datetime-local" disabled /></div>`;
    }

    return `<div class="form-row">${label}<input id="${id}" type="text" disabled placeholder="${escapeHtml(field.label)}" /></div>`;
  }).join('') : '<p class="field-meta">No custom fields added to this preset draft yet.</p>';

  wrapper.innerHTML = `
    <div class="preview-context">
      <div><strong>Preset:</strong> ${escapeHtml(preset.name || '-')}</div>
      <div><strong>Epic:</strong> ${escapeHtml(preset.epicName || '-')} <span class="field-meta">${escapeHtml(preset.epicId || '')}</span></div>
      <div><strong>Feature / Task:</strong> ${escapeHtml(preset.featureName || '-')} <span class="field-meta">${escapeHtml(preset.featureId || '')}</span></div>
    </div>
    <div class="preset-preview-fields">${fieldPreview}</div>
  `;
}

function updateFieldOptionsVisibility() {
  const type = $('#newFieldType')?.value || 'text';
  const row = $('#optionsRow');
  const input = $('#newFieldOptions');
  if (!row || !input) return;

  const needsOptions = type === 'dropdown' || type === 'checkbox';
  row.classList.toggle('hidden', !needsOptions);
  input.disabled = !needsOptions;
  if (!needsOptions) input.value = '';
}


function getRemoteUrlsForField(index) {
  if (!state.remoteAttachmentUrls[index]) state.remoteAttachmentUrls[index] = [];
  return state.remoteAttachmentUrls[index];
}

function normalizeUrl(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  try {
    const url = new URL(text);
    if (!['http:', 'https:', 'data:'].includes(url.protocol)) return '';
    return url.href;
  } catch {
    return '';
  }
}

function extractUrlsFromTransfer(dataTransfer) {
  const values = [];
  const uriList = dataTransfer.getData('text/uri-list');
  const text = dataTransfer.getData('text/plain');
  const html = dataTransfer.getData('text/html');

  if (uriList) values.push(...uriList.split('\n').filter((line) => !line.startsWith('#')));
  if (text) values.push(text);

  if (html) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    doc.querySelectorAll('img[src], a[href]').forEach((node) => {
      values.push(node.getAttribute('src') || node.getAttribute('href'));
    });
  }

  return Array.from(new Set(values.map(normalizeUrl).filter(Boolean)));
}

async function dataUrlToFile(dataUrl, fileName = 'pasted-image.png') {
  const response = await fetch(dataUrl);
  const blob = await response.blob();
  const extension = blob.type?.split('/')[1] || 'png';
  const safeName = fileName.includes('.') ? fileName : `${fileName}.${extension}`;
  return new File([blob], safeName, { type: blob.type || 'application/octet-stream' });
}

function addFilesToInput(input, files) {
  if (!input || !files.length) return;
  const dataTransfer = new DataTransfer();
  Array.from(input.files || []).forEach((file) => dataTransfer.items.add(file));
  files.forEach((file) => dataTransfer.items.add(file));
  input.files = dataTransfer.files;
}

function addRemoteUrlToField(index, url) {
  const urls = getRemoteUrlsForField(index);
  if (!urls.includes(url)) urls.push(url);
}

async function addExternalUrlsToField(input, urls) {
  const index = input.dataset.fieldIndex;
  const dataUrlFiles = [];
  for (const url of urls) {
    if (url.startsWith('data:')) {
      dataUrlFiles.push(await dataUrlToFile(url, `pasted-image-${Date.now()}.png`));
    } else {
      addRemoteUrlToField(index, url);
    }
  }
  if (dataUrlFiles.length) addFilesToInput(input, dataUrlFiles);
  renderSelectedFiles(input);
}

function removeRemoteUrl(index, urlIndex) {
  const urls = getRemoteUrlsForField(index);
  urls.splice(urlIndex, 1);
  const input = document.querySelector(`[data-field-index="${index}"][type="file"]`);
  if (input) renderSelectedFiles(input);
}

function bindUploadDropZones() {
  $$('.drop-zone').forEach((zone) => {
    const input = zone.querySelector('input[type="file"]');
    if (!input) return;

    zone.addEventListener('click', (event) => {
      if (event.target.closest('[data-remove-remote-url]')) return;
      zone.focus();
      input.click();
    });

    zone.addEventListener('keydown', (event) => {
      if (event.key !== 'Enter' && event.key !== ' ') return;
      event.preventDefault();
      input.click();
    });

    zone.addEventListener('dragover', (event) => {
      event.preventDefault();
      zone.classList.add('drag-over');
    });

    zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));

    zone.addEventListener('drop', async (event) => {
      event.preventDefault();
      zone.classList.remove('drag-over');

      const files = Array.from(event.dataTransfer.files || []);
      if (files.length) addFilesToInput(input, files);

      const urls = extractUrlsFromTransfer(event.dataTransfer);
      if (urls.length) await addExternalUrlsToField(input, urls);

      renderSelectedFiles(input);
      if (!files.length && !urls.length) toast('No supported file or image link was found in the drop.');
    });

    zone.addEventListener('paste', async (event) => {
      const files = Array.from(event.clipboardData?.files || []);
      const urls = extractUrlsFromTransfer(event.clipboardData);
      if (!files.length && !urls.length) return;
      event.preventDefault();

      if (files.length) addFilesToInput(input, files);
      if (urls.length) await addExternalUrlsToField(input, urls);
      renderSelectedFiles(input);
    });

    input.addEventListener('change', () => renderSelectedFiles(input));
  });
}

function renderDynamicForm() {
  const wrapper = $('#dynamicFormFields');
  const fields = getSubmitFields();

  if (!state.activeTemplate) {
    wrapper.innerHTML = '<p class="field-meta">No preset selected. Save or load a preset first.</p>';
    return;
  }

  if (fields.length === 0) {
    wrapper.innerHTML = '<p class="field-meta">This saved preset has no custom fields.</p>';
    return;
  }

  wrapper.innerHTML = fields.map((field, index) => {
    const id = `field_${index}`;
    const label = `<label for="${id}">${escapeHtml(field.label)}</label>`;

    if (field.type === 'textarea') {
      return `<div class="form-row" data-dynamic-field="${index}">${label}<textarea id="${id}" data-field-index="${index}" placeholder="Enter ${escapeHtml(field.label)}"></textarea></div>`;
    }

    if (field.type === 'dropdown') {
      const options = (field.options || []).map((opt) => `<option value="${escapeHtml(opt)}">${escapeHtml(opt)}</option>`).join('');
      return `<div class="form-row" data-dynamic-field="${index}">${label}<select id="${id}" data-field-index="${index}"><option value="">-- Select --</option>${options}</select></div>`;
    }

    if (field.type === 'checkbox') {
      const options = field.options?.length ? field.options : ['Yes'];
      const checks = options.map((opt, optIndex) => `
        <label class="checkbox-item">
          <input type="checkbox" data-field-index="${index}" value="${escapeHtml(opt)}" id="${id}_${optIndex}" />
          ${escapeHtml(opt)}
        </label>
      `).join('');
      return `<div class="form-row" data-dynamic-field="${index}"><label>${escapeHtml(field.label)}</label><div class="checkbox-group">${checks}</div></div>`;
    }

    if (field.type === 'file') {
      return `
        <div class="form-row" data-dynamic-field="${index}">
          <label>${escapeHtml(field.label)}</label>
          <div class="drop-zone" data-drop-index="${index}" tabindex="0" title="Drop local files, drag an image from a web page, or paste a copied image here.">
            <input id="${id}" data-field-index="${index}" type="file" multiple />
            <span>Drag/drop files, paste image, drag web image, or click to select</span>
          </div>
          <div class="file-list" id="file_list_${index}">No files selected.</div>
        </div>`;
    }

    if (field.type === 'datetime') {
      return `<div class="form-row" data-dynamic-field="${index}">${label}<input id="${id}" data-field-index="${index}" type="datetime-local" /></div>`;
    }

    return `<div class="form-row" data-dynamic-field="${index}">${label}<input id="${id}" data-field-index="${index}" type="text" placeholder="Enter ${escapeHtml(field.label)}" /></div>`;
  }).join('');

  bindUploadDropZones();
}

function renderSelectedFiles(input) {
  const index = input.dataset.fieldIndex;
  const target = $(`#file_list_${index}`);
  const files = Array.from(input.files || []);
  const remoteUrls = getRemoteUrlsForField(index);

  if (!files.length && !remoteUrls.length) {
    target.textContent = 'No files selected.';
    return;
  }

  const localHtml = files.map((file) => `
    <span class="file-chip">${escapeHtml(file.name)} (${Math.ceil(file.size / 1024)} KB)</span>
  `).join('');

  const remoteHtml = remoteUrls.map((url, urlIndex) => `
    <span class="file-chip remote">
      Web image/link ${urlIndex + 1}
      <button type="button" data-remove-remote-url="${index}:${urlIndex}" title="Remove this web link">×</button>
      <small>${escapeHtml(url)}</small>
    </span>
  `).join('');

  target.innerHTML = localHtml + remoteHtml;
  target.querySelectorAll('[data-remove-remote-url]').forEach((button) => {
    button.addEventListener('click', (event) => {
      event.stopPropagation();
      const [fieldIndex, remoteIndex] = button.dataset.removeRemoteUrl.split(':').map(Number);
      removeRemoteUrl(fieldIndex, remoteIndex);
    });
  });
}

function fileToBase64Payload(file, fieldName) {
  return new Promise((resolve, reject) => {
    const maxSize = 5 * 1024 * 1024;
    if (file.size > maxSize) {
      reject(new Error(`${file.name} is larger than 5 MB. Upload large videos to Drive manually and drag/paste the Drive link instead.`));
      return;
    }

    const reader = new FileReader();
    reader.onload = () => {
      const result = String(reader.result || '');
      const base64 = result.includes(',') ? result.split(',')[1] : result;
      resolve({
        fieldName,
        fileName: file.name,
        mimeType: file.type || 'application/octet-stream',
        size: file.size,
        base64
      });
    };
    reader.onerror = () => reject(new Error(`Could not read ${file.name}.`));
    reader.readAsDataURL(file);
  });
}

async function readDynamicValues(filePayloads = [], remoteAttachments = []) {
  const values = {};
  const fields = getSubmitFields();

  for (const [index, field] of fields.entries()) {
    if (field.type === 'file') {
      const input = document.querySelector(`[data-field-index="${index}"][type="file"]`);
      const files = Array.from(input?.files || []);
      const urls = getRemoteUrlsForField(index).slice();

      for (const file of files) {
        filePayloads.push(await fileToBase64Payload(file, field.label));
      }

      if (urls.length) {
        remoteAttachments.push({ fieldName: field.label, urls });
      }

      values[field.label] = [
        ...files.map((file) => file.name),
        ...urls
      ].join(', ');
      continue;
    }

    if (field.type === 'checkbox') {
      const checked = $$(`input[type="checkbox"][data-field-index="${index}"]:checked`).map((item) => item.value);
      values[field.label] = checked.join(', ');
      continue;
    }

    const input = document.querySelector(`[data-field-index="${index}"]`);
    values[field.label] = input ? input.value : '';
  }

  return values;
}

async function loadTemplates() {
  state.templates = await api('/api/templates');
  renderTemplateSelect();
}

async function saveTemplate() {
  const payload = getPresetValues();
  if (!payload.name) throw new Error('Preset name is required.');
  if (!payload.epicName) throw new Error('Epic name is required.');
  if (!payload.featureName) throw new Error('Feature name is required.');
  if (state.selectedTemplateId) payload.id = state.selectedTemplateId;

  const saved = await api('/api/templates', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });
  state.selectedTemplateId = saved.id;
  state.activeTemplate = cloneTemplate(saved);
  await loadTemplates();
  $('#templateSelect').value = saved.id;
  renderDynamicForm();
  updateActiveContext();
  updateSubmitPresetNotice();
  toast('Preset saved and applied to Submit Issue form.');
}

async function deleteTemplate() {
  if (!state.selectedTemplateId) return toast('Select a preset first.');
  if (!confirm('Delete this preset? Existing issues will remain.')) return;
  await api(`/api/templates/${state.selectedTemplateId}`, { method: 'DELETE' });
  state.selectedTemplateId = '';
  setPresetValues(null);
  await loadTemplates();
  toast('Preset deleted.');
}

function resetFieldEditor() {
  state.editingFieldIndex = null;
  $('#newFieldName').value = '';
  $('#newFieldOptions').value = '';
  $('#newFieldType').value = 'text';
  $('#addFieldBtn').textContent = 'Add Field';
  $('#cancelEditFieldBtn')?.classList.add('hidden');
  updateFieldOptionsVisibility();
  renderFieldList();
}

function startEditField(index) {
  const field = state.fields[index];
  if (!field) return;
  state.editingFieldIndex = index;
  $('#newFieldName').value = field.label || '';
  $('#newFieldType').value = field.type || 'text';
  $('#newFieldOptions').value = Array.isArray(field.options) ? field.options.join(', ') : '';
  $('#addFieldBtn').textContent = 'Update Field';
  $('#cancelEditFieldBtn')?.classList.remove('hidden');
  updateFieldOptionsVisibility();
  renderFieldList();
  $('#newFieldName').focus();
}

function addField() {
  const label = $('#newFieldName').value.trim();
  const type = $('#newFieldType').value;
  const rawOptions = $('#newFieldOptions').value
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean);
  const options = (type === 'dropdown' || type === 'checkbox') ? rawOptions : [];
  const editingIndex = state.editingFieldIndex;

  if (!label) return toast('Field name is required.');
  if (state.fields.some((field, index) => index !== editingIndex && field.label.toLowerCase() === label.toLowerCase())) {
    return toast('Field name already exists.');
  }

  if (editingIndex !== null && state.fields[editingIndex]) {
    state.fields[editingIndex] = {
      ...state.fields[editingIndex],
      label,
      type,
      options
    };
    resetFieldEditor();
    updateSubmitPresetNotice();
    toast('Field updated in preset draft. Click Save Preset to apply it to Submit Issue.');
    return;
  }

  state.fields.push({ id: crypto.randomUUID(), label, type, options });
  $('#newFieldName').value = '';
  $('#newFieldOptions').value = '';
  renderFieldList();
  updateSubmitPresetNotice();
  toast('Field added to preset draft. Click Save Preset to use it in Submit Issue.');
}

async function submitIssue(event) {
  event.preventDefault();
  const preset = state.activeTemplate;
  if (!preset) return toast('Save or select a preset first.');
  if (!preset.epicName) return toast('Epic name is required.');
  if (!preset.featureName) return toast('Feature name is required.');
  if (!$('#issueTitle').value.trim()) return toast('Issue title is required.');

  const submitButton = event.target.querySelector('button[type="submit"]');
  submitButton.disabled = true;
  submitButton.textContent = 'Saving...';

  try {
    const remoteAttachments = [];
    const files = [];
    const fields = await readDynamicValues(files, remoteAttachments);

    await api('/api/issues', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        templateId: state.selectedTemplateId || '',
        epicName: preset.epicName,
        epicId: preset.epicId,
        featureName: preset.featureName,
        featureId: preset.featureId,
        title: $('#issueTitle').value.trim(),
        status: $('#issueStatus').value,
        severity: $('#issueSeverity').value,
        priority: $('#issuePriority').value,
        issueDate: $('#issueDate').value,
        reportedTo: $('#reportedTo').value.trim(),
        reportedBy: $('#reportedBy').value.trim(),
        fields,
        fieldMeta: getSubmitFields(),
        files,
        remoteAttachmentUrls: remoteAttachments
      })
    });

    event.target.reset();
    state.remoteAttachmentUrls = {};
    renderDynamicForm();
    setDefaultIssueDate();
    await loadIssues();
    toast('Issue saved.');
    switchTab('report');
  } catch (error) {
    toast(error.message);
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = 'Save Issue';
  }
}

async function loadIssues() {
  state.issues = await api('/api/issues');
  renderReportFilters();
  renderIssuesTable();
  updateReportActions();
}

function getSelectedReportScope() {
  return {
    epicId: $('#reportEpicFilter')?.value || '',
    featureId: $('#reportFeatureFilter')?.value || ''
  };
}

function getUniqueOptions(items, idKey, nameKey) {
  const map = new Map();
  for (const item of items) {
    const id = item[idKey];
    if (!id || map.has(id)) continue;
    map.set(id, { id, name: item[nameKey] || id });
  }
  return Array.from(map.values()).sort((a, b) => a.name.localeCompare(b.name));
}

function renderReportFilters() {
  const epicSelect = $('#reportEpicFilter');
  const featureSelect = $('#reportFeatureFilter');
  if (!epicSelect || !featureSelect) return;

  const previousEpic = epicSelect.value;
  const previousFeature = featureSelect.value;
  const epics = getUniqueOptions(state.issues, 'epicId', 'epicName');

  epicSelect.innerHTML = '<option value="">Select Epic first</option>' + epics.map((epic) =>
    `<option value="${escapeHtml(epic.id)}">${escapeHtml(epic.name)} (${escapeHtml(epic.id)})</option>`
  ).join('');

  if (epics.some((epic) => epic.id === previousEpic)) epicSelect.value = previousEpic;

  populateFeatureFilter(previousFeature);
}

function populateFeatureFilter(preferredFeatureId = '') {
  const epicId = $('#reportEpicFilter')?.value || '';
  const featureSelect = $('#reportFeatureFilter');
  if (!featureSelect) return;

  if (!epicId) {
    featureSelect.innerHTML = '<option value="">Select Feature / Task</option>';
    featureSelect.disabled = true;
    featureSelect.value = '';
    return;
  }

  const features = getUniqueOptions(
    state.issues.filter((issue) => issue.epicId === epicId),
    'featureId',
    'featureName'
  );

  featureSelect.innerHTML = '<option value="">Select Feature / Task</option>' + features.map((feature) =>
    `<option value="${escapeHtml(feature.id)}">${escapeHtml(feature.name)} (${escapeHtml(feature.id)})</option>`
  ).join('');
  featureSelect.disabled = false;

  if (features.some((feature) => feature.id === preferredFeatureId)) {
    featureSelect.value = preferredFeatureId;
  }
}

function isReportScopeSelected() {
  const { epicId, featureId } = getSelectedReportScope();
  return Boolean(epicId && featureId);
}

function updateReportActions() {
  const { epicId, featureId } = getSelectedReportScope();
  const selected = Boolean(epicId && featureId);
  const downloadBtn = $('#downloadExcelBtn');
  const notice = $('#reportScopeNotice');

  if (downloadBtn) {
    if (selected) {
      downloadBtn.href = `/api/export/issues.xlsx?epicId=${encodeURIComponent(epicId)}&featureId=${encodeURIComponent(featureId)}`;
      downloadBtn.classList.remove('disabled');
      downloadBtn.setAttribute('aria-disabled', 'false');
    } else {
      downloadBtn.href = '#';
      downloadBtn.classList.add('disabled');
      downloadBtn.setAttribute('aria-disabled', 'true');
    }
  }

  if (notice) {
    notice.classList.toggle('hidden', selected);
  }
}

function getFilteredIssues() {
  const text = $('#searchBox')?.value?.toLowerCase() || '';
  const status = $('#statusFilter')?.value || '';
  const { epicId, featureId } = getSelectedReportScope();

  if (!epicId || !featureId) return [];

  return state.issues.filter((issue) => {
    if (issue.epicId !== epicId || issue.featureId !== featureId) return false;

    const haystack = [
      issue.issueNo,
      issue.title,
      issue.epicName,
      issue.epicId,
      issue.featureName,
      issue.featureId,
      issue.status,
      issue.severity,
      issue.priority,
      issue.issueDate,
      issue.reportedTo,
      issue.reportedBy,
      JSON.stringify(issue.fields || {})
    ].join(' ').toLowerCase();
    return (!text || haystack.includes(text)) && (!status || issue.status === status);
  });
}

function renderIssuesTable() {
  const table = $('#issuesTable');
  const head = table.querySelector('thead');
  const body = table.querySelector('tbody');
  head.innerHTML = `
    <tr>
      <th>Issue No</th>
      <th>Title</th>
      <th>Epic</th>
      <th>Feature</th>
      <th>Issue Date</th>
      <th>Reported To</th>
      <th>Reported By</th>
      <th>Status</th>
      <th>Severity</th>
      <th>Priority</th>
      <th>Attachments</th>
      <th>Created</th>
      <th>Action</th>
    </tr>`;

  const issues = getFilteredIssues();
  renderReportStats();

  if (!isReportScopeSelected()) {
    body.innerHTML = '<tr><td colspan="13">Select an Epic and Feature/Task to view issues.</td></tr>';
    return;
  }

  if (issues.length === 0) {
    body.innerHTML = '<tr><td colspan="13">No issues found for this Epic and Feature/Task.</td></tr>';
    return;
  }

  body.innerHTML = issues.map((issue) => `
    <tr>
      <td><strong>${escapeHtml(issue.issueNo)}</strong></td>
      <td>${escapeHtml(issue.title)}</td>
      <td>${escapeHtml(issue.epicName)}<br><span class="field-meta">${escapeHtml(issue.epicId)}</span></td>
      <td>${escapeHtml(issue.featureName)}<br><span class="field-meta">${escapeHtml(issue.featureId)}</span></td>
      <td>${escapeHtml(formatDateTimeValue(issue.issueDate || issue.createdAt) || '-')}</td>
      <td>${escapeHtml(issue.reportedTo || '-')}</td>
      <td>${escapeHtml(issue.reportedBy || '-')}</td>
      <td>${renderBadge(issue.status, getStatusTone(issue.status))}</td>
      <td>${renderBadge(issue.severity || '-', getSeverityTone(issue.severity))}</td>
      <td>${renderBadge(issue.priority || '-', getPriorityTone(issue.priority))}</td>
      <td>${(issue.attachments || []).length}</td>
      <td>${escapeHtml(new Date(issue.createdAt).toLocaleString())}</td>
      <td><button data-open-issue="${issue.id}">View</button></td>
    </tr>`).join('');

  $$('[data-open-issue]').forEach((button) => {
    button.addEventListener('click', () => openIssue(button.dataset.openIssue));
  });
}

function openIssue(id) {
  const issue = state.issues.find((item) => item.id === id);
  if (!issue) return;
  state.selectedIssue = issue;
  const fieldEntries = Object.entries(issue.fields || {});
  const fieldMetaByLabel = new Map((issue.fieldMeta || []).map((field) => [field.label, field]));
  const attachments = issue.attachments || [];

  $('#modalContent').innerHTML = `
    <h2>${escapeHtml(issue.issueNo)}: ${escapeHtml(issue.title)}</h2>
    <div class="modal-grid">
      <div class="detail-box"><strong>Epic</strong>${escapeHtml(issue.epicName)}<br>${escapeHtml(issue.epicId)}</div>
      <div class="detail-box"><strong>Feature</strong>${escapeHtml(issue.featureName)}<br>${escapeHtml(issue.featureId)}</div>
      <div class="detail-box"><strong>Issue Date</strong>${escapeHtml(formatDateTimeValue(issue.issueDate || issue.createdAt) || '-')}</div>
      <div class="detail-box"><strong>Reported To</strong>${escapeHtml(issue.reportedTo || '-')}</div>
      <div class="detail-box"><strong>Reported By</strong>${escapeHtml(issue.reportedBy || '-')}</div>
      <div class="detail-box"><strong>Source</strong>${escapeHtml(issue.source || '-')}</div>
      <div class="detail-box"><strong>Status</strong><select id="modalStatus">${['Open','In Progress','Fixed','Ready for Retest','Closed','Reopened','Rejected','Duplicate','Imported'].map((s) => `<option ${issue.status === s ? 'selected' : ''}>${s}</option>`).join('')}</select></div>
      <div class="detail-box"><strong>Severity</strong>${renderBadge(issue.severity || '-', getSeverityTone(issue.severity))}</div>
      <div class="detail-box"><strong>Priority</strong>${renderBadge(issue.priority || '-', getPriorityTone(issue.priority))}</div>
    </div>

    <h3>Fields</h3>
    ${fieldEntries.length ? fieldEntries.map(([key, value]) => {
      const meta = fieldMetaByLabel.get(key);
      const displayValue = meta?.type === 'datetime' ? formatDateTimeValue(value) : value;
      return `
        <div class="detail-box" style="margin-bottom:8px"><strong>${escapeHtml(key)}</strong>${escapeHtml(displayValue || '-')}</div>
      `;
    }).join('') : '<p>No custom fields saved.</p>'}

    <h3>Attachments</h3>
    ${attachments.length ? `<ul>${attachments.map((file) => `<li><a href="${escapeHtml(file.url)}" target="_blank" rel="noreferrer">${escapeHtml(file.originalName || file.savedName || file.url)}</a> <span class="field-meta">${escapeHtml(file.storage || '')}</span></li>`).join('')}</ul>` : '<p>No attachments.</p>'}

    <div class="actions" style="margin-top:16px">
      <button id="saveIssueStatusBtn" class="primary">Save Status</button>
      <button id="deleteIssueBtn" class="danger ghost">Delete Issue</button>
    </div>
  `;

  $('#saveIssueStatusBtn').addEventListener('click', saveModalStatus);
  $('#deleteIssueBtn').addEventListener('click', deleteSelectedIssue);
  $('#issueModal').classList.remove('hidden');
}

async function saveModalStatus() {
  if (!state.selectedIssue) return;
  const status = $('#modalStatus').value;
  await api(`/api/issues/${state.selectedIssue.id}`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ status })
  });
  $('#issueModal').classList.add('hidden');
  await loadIssues();
  toast('Issue status updated.');
}

async function deleteSelectedIssue() {
  if (!state.selectedIssue) return;
  if (!confirm(`Delete ${state.selectedIssue.issueNo}?`)) return;
  await api(`/api/issues/${state.selectedIssue.id}`, { method: 'DELETE' });
  $('#issueModal').classList.add('hidden');
  await loadIssues();
  toast('Issue deleted.');
}

async function importExcel(event) {
  event.preventDefault();
  const file = $('#importFile').files[0];
  if (!file) return toast('Select an Excel file first.');

  const payload = await fileToBase64Payload(file, 'Excel Report');
  const result = await api('/api/import/issues.xlsx', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });

  $('#importResult').classList.remove('hidden');
  $('#importResult').textContent = `Imported ${result.imported} issue(s). Skipped duplicates: ${result.skippedDuplicates}.`;
  await loadIssues();
  toast('Import complete.');
}

function switchTab(tabName) {
  $$('.tab').forEach((tab) => tab.classList.toggle('active', tab.dataset.tab === tabName));
  $$('.tab-panel').forEach((panel) => panel.classList.toggle('active', panel.id === tabName));
}

function bindEvents() {
  $$('.tab').forEach((tab) => tab.addEventListener('click', () => switchTab(tab.dataset.tab)));
  $('#templateSelect').addEventListener('change', (event) => {
    state.selectedTemplateId = event.target.value;
    const template = state.templates.find((item) => item.id === state.selectedTemplateId);
    if (template) setPresetValues(template);
  });
  ['templateName', 'epicName', 'epicId', 'featureName', 'featureId'].forEach((id) => {
    $(`#${id}`).addEventListener('input', () => {
      updateSubmitPresetNotice();
      renderPresetPreview();
    });
  });
  $('#newFieldType').addEventListener('change', updateFieldOptionsVisibility);
  $('#saveTemplateBtn').addEventListener('click', () => saveTemplate().catch((error) => toast(error.message)));
  $('#deleteTemplateBtn').addEventListener('click', () => deleteTemplate().catch((error) => toast(error.message)));
  $('#addFieldBtn').addEventListener('click', addField);
  $('#cancelEditFieldBtn')?.addEventListener('click', resetFieldEditor);
  $('#clearFieldsBtn').addEventListener('click', () => {
    if (!confirm('Clear all custom fields from this form?')) return;
    state.fields = [];
    state.editingFieldIndex = null;
    $('#addFieldBtn').textContent = 'Add Field';
    $('#cancelEditFieldBtn')?.classList.add('hidden');
    renderFieldList();
    updateSubmitPresetNotice();
  });
  $('#issueForm').addEventListener('submit', submitIssue);
  $('#issueForm').addEventListener('reset', () => {
    setTimeout(() => {
      state.remoteAttachmentUrls = {};
      renderDynamicForm();
      setDefaultIssueDate();
    }, 0);
  });
  $('#refreshIssuesBtn').addEventListener('click', () => loadIssues().then(() => toast('Report refreshed.')));
  $('#reportEpicFilter').addEventListener('change', () => {
    populateFeatureFilter();
    updateReportActions();
    renderIssuesTable();
  });
  $('#reportFeatureFilter').addEventListener('change', () => {
    updateReportActions();
    renderIssuesTable();
  });
  $('#searchBox').addEventListener('input', renderIssuesTable);
  $('#statusFilter').addEventListener('change', renderIssuesTable);
  $('#downloadExcelBtn').addEventListener('click', (event) => {
    if (!isReportScopeSelected()) {
      event.preventDefault();
      toast('Select an Epic and Feature/Task before downloading Excel.');
    }
  });
  $('#importForm').addEventListener('submit', (event) => importExcel(event).catch((error) => toast(error.message)));
  $('#closeModalBtn').addEventListener('click', () => $('#issueModal').classList.add('hidden'));
  $('#issueModal').addEventListener('click', (event) => {
    if (event.target.id === 'issueModal') $('#issueModal').classList.add('hidden');
  });
}

async function init() {
  initTheme();
  bindEvents();
  renderFieldList();
  renderDynamicForm();
  updateFieldOptionsVisibility();
  updateActiveContext();
  renderPresetPreview();
  setDefaultIssueDate();

  try {
    const health = await api('/api/health');
    $('#serverStatus').textContent = `Storage: ${health.storageMode}`;
    state.spreadsheetUrl = health.spreadsheetUrl || 'https://docs.google.com/spreadsheets/d/1JlJtBq3GlsEG1Rc9cwTLxcXDAwpvoX2C2Ld8fqEr6u0/edit?usp=sharing';
    const spreadsheetBtn = $('#viewSpreadsheetBtn');
    if (spreadsheetBtn) spreadsheetBtn.href = state.spreadsheetUrl || '#';
  } catch {
    $('#serverStatus').textContent = 'Server error';
  }

  await loadTemplates();
  await loadIssues();
}

init().catch((error) => toast(error.message));
