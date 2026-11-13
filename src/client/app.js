import {
  FIELD_SUGGESTIONS, BATCH_SIZE, parseCsv, tableFromRecords, worksheetRecords,
  suggestedMappings, buildExport, batchRows, toCsv, readProfile, lookupParent
} from './core.js';

const $ = id => document.getElementById(id);
const state = {
  workbook: null, csvText: null, table: null, mappings: [],
  filter: { column: '', value: '' }, batch: 1, fileName: '',
  sourceErrors: [], result: null, loading: false
};
state.parent = lookupParent();

function el(tag, content, className) {
  const node = document.createElement(tag);
  if (content !== undefined) node.textContent = content;
  if (className) node.className = className;
  return node;
}

function notify(message, success = false) {
  $('notice').textContent = message;
  $('notice').className = `notice${success ? ' success' : ''}`;
}
function report(error, context) {
  console.error(context, error);
  notify(`${context}: ${error.message}`);
}
function option(select, value, label = value) {
  const node = el('option', label);
  node.value = value;
  select.append(node);
}
function downloadFile(name, data, type) {
  const url = URL.createObjectURL(data instanceof Blob ? data : new Blob([data], { type }));
  const link = el('a');
  link.href = url;
  link.download = name;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}
function baseName() {
  return state.fileName.replace(/\.[^.]+$/, '').replace(/[^\p{L}\p{N}_-]+/gu, '-').slice(0, 80) || 'work-items';
}
async function restoreEmptyFormulaCaches(data, workbook) {
  if (!globalThis.JSZip) throw new Error('Workbook cache reader did not load. Reload the page after npm install.');
  const zip = await globalThis.JSZip.loadAsync(data);
  async function readXml(name) {
    const file = zip.file(name);
    if (!file) throw new Error(`Workbook is missing ${name}. Resave it as .xlsx in Excel.`);
    const document = new DOMParser().parseFromString(await file.async('string'), 'application/xml');
    if (document.getElementsByTagName('parsererror').length) throw new Error(`Invalid XML in ${name}.`);
    return document;
  }
  const relationships = await readXml('xl/_rels/workbook.xml.rels');
  const targets = new Map([...relationships.getElementsByTagNameNS('*', 'Relationship')]
    .map(node => [node.getAttribute('Id'), node.getAttribute('Target')]));
  const definition = await readXml('xl/workbook.xml');
  for (const sheetNode of definition.getElementsByTagNameNS('*', 'sheet')) {
    const sheet = workbook.getWorksheet(sheetNode.getAttribute('name'));
    if (!sheet) continue;
    const id = sheetNode.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    const target = targets.get(id);
    if (!target) throw new Error(`Missing worksheet relationship for ${sheet.name}.`);
    const filename = new URL(target, 'https://local.invalid/xl/').pathname.slice(1);
    const document = await readXml(filename);
    for (const node of document.getElementsByTagNameNS('*', 'c')) {
      if (node.getAttribute('t') !== 'str') continue;
      const cached = node.getElementsByTagNameNS('*', 'v')[0];
      if (!cached || cached.textContent !== '' || !node.getElementsByTagNameNS('*', 'f').length) continue;
      const cell = sheet.getCell(node.getAttribute('r'));
      // ExcelJS drops empty string caches; XML distinguishes them from absent caches.
      if (cell.value && typeof cell.value === 'object' && ('formula' in cell.value || 'sharedFormula' in cell.value)) {
        cell.value = { ...cell.value, result: '' };
      }
    }
  }
}
function populateFilters() {
  const select = $('filter-column');
  select.replaceChildren();
  option(select, '', 'All rows');
  for (const name of state.table?.headers || []) option(select, name);
  if (state.filter.column && !state.table?.headers.includes(state.filter.column)) {
    option(select, state.filter.column, `${state.filter.column} (missing)`);
  }
  select.value = state.filter.column;
  select.disabled = !state.table;
  $('filter-value').value = state.filter.value;
  $('filter-value').disabled = !state.table || !state.filter.column;
  $('filter-values').replaceChildren();
  if (state.table && state.filter.column) {
    const column = state.table.headers.indexOf(state.filter.column);
    const values = new Set(state.table.rows.map(row => row.values[column]));
    for (const value of [...values].slice(0, 200)) {
      if (value !== undefined) option($('filter-values'), value);
    }
  }
}

function populateParent() {
  $('parent-enabled').checked = state.parent.enabled;
  $('parent-settings').classList.toggle('hidden', !state.parent.enabled);
  $('parent-mode').value = state.parent.mode;
  $('epic-lookup-settings').classList.toggle('hidden', state.parent.mode !== 'lookup');
  $('epic-direct-settings').classList.toggle('hidden', state.parent.mode !== 'direct');
  const source = $('epic-source');
  source.replaceChildren();
  option(source, '', 'Choose your Epic column');
  for (const header of state.table?.headers || []) option(source, header);
  if (state.parent.epicSource && !state.table?.headers.includes(state.parent.epicSource)) {
    option(source, state.parent.epicSource, `${state.parent.epicSource} (missing)`);
  }
  source.value = state.parent.epicSource;
  renderEpicLookups();
  for (const key of ['id', 'title']) {
    const select = $(`parent-${key}-source`);
    select.replaceChildren();
    option(select, '', key === 'id' ? 'Use the fixed ID below' : 'Use the fixed title below');
    for (const header of state.table?.headers || []) option(select, header);
    const selected = state.parent[`${key}Source`];
    if (selected && !state.table?.headers.includes(selected)) option(select, selected, `${selected} (missing)`);
    select.value = selected;
    $(`parent-${key}-default`).value = state.parent[`${key}Default`];
  }
}

function updateEpicCounts() {
  const column = state.table?.headers.indexOf(state.parent.epicSource) ?? -1;
  state.parent.epics.forEach((epic, index) => {
    const cell = document.querySelector(`[data-epic-count="${index}"]`);
    if (cell) cell.textContent = column < 0 ? '—' : String(state.table.rows.filter(row =>
      row.values[column].trim().toLowerCase() === epic.sourceName.trim().toLowerCase()).length);
  });
}

function renderEpicLookups() {
  $('epic-lookups').replaceChildren();
  state.parent.epics.forEach((epic, index) => {
    const row = el('tr');
    for (const [key, label] of [['sourceName', 'Source Epic name'], ['id', 'ADO Epic ID'], ['title', 'ADO Epic title']]) {
      const cell = el('td');
      const input = el('input');
      input.type = 'text';
      if (key === 'id') input.inputMode = 'numeric';
      input.value = epic[key];
      input.setAttribute('aria-label', `${label} ${index + 1}`);
      input.addEventListener('input', () => {
        epic[key] = input.value;
        state.batch = 1;
        updateEpicCounts(); refreshExport();
      });
      cell.append(input); row.append(cell);
    }
    const count = el('td');
    count.dataset.epicCount = String(index);
    row.append(count);
    const actions = el('td');
    const remove = el('button', 'Remove');
    remove.setAttribute('aria-label', `Remove Epic mapping ${index + 1}`);
    remove.addEventListener('click', () => {
      state.parent.epics.splice(index, 1);
      state.batch = 1;
      renderEpicLookups(); refreshExport();
    });
    actions.append(remove); row.append(actions);
    $('epic-lookups').append(row);
  });
  updateEpicCounts();
}

function renderMappings() {
  $('mappings').replaceChildren();
  $('mapping-empty').classList.toggle('hidden', state.mappings.length > 0);
  state.mappings.forEach((mapping, index) => {
    const row = el('tr', undefined, mapping.enabled ? '' : 'disabled-row');
    const enabledCell = el('td');
    const enabled = el('input');
    enabled.type = 'checkbox';
    enabled.checked = mapping.enabled;
    enabled.setAttribute('aria-label', `Enable mapping ${index + 1}`);
    enabled.addEventListener('change', () => {
      mapping.enabled = enabled.checked;
      row.classList.toggle('disabled-row', !mapping.enabled);
      refreshExport();
    });
    enabledCell.append(enabled);
    row.append(enabledCell);
    for (const key of ['field', 'source', 'default', 'prefix', 'suffix', 'format']) {
      const cell = el('td');
      const control = el(['field', 'source', 'format'].includes(key) ? 'select' : 'input');
      control.setAttribute('aria-label', `${key} for mapping ${index + 1}`);
      if (key === 'field') {
        option(control, '', 'Choose ADO field');
        for (const field of FIELD_SUGGESTIONS) option(control, field);
        option(control, '__custom', 'Custom field…');
        control.value = !mapping.field || FIELD_SUGGESTIONS.includes(mapping.field) ? mapping.field : '__custom';
        const custom = el('input');
        custom.type = 'text';
        custom.placeholder = 'Exact custom CSV field name';
        custom.setAttribute('aria-label', `custom field for mapping ${index + 1}`);
        custom.classList.toggle('hidden', control.value !== '__custom');
        custom.value = control.value === '__custom' ? mapping.field : '';
        custom.addEventListener('input', () => { mapping.field = custom.value; refreshExport(); });
        control.addEventListener('change', () => {
          custom.classList.toggle('hidden', control.value !== '__custom');
          mapping.field = control.value === '__custom' ? custom.value : control.value;
          if (mapping.field === 'Title' || mapping.field === 'Work Item Type') {
            mapping.format = 'plain'; renderMappings();
          }
          refreshExport();
        });
        cell.append(control, custom);
        row.append(cell);
        continue;
      } else if (key === 'format') {
        for (const [value, label] of [['plain', 'Plain text'], ['paragraphs', 'Paragraphs (HTML)'], ['story', 'User story (HTML)'], ['markdown', 'Markdown → HTML']]) option(control, value, label);
        control.disabled = ['Title', 'Work Item Type'].includes(mapping.field);
      } else if (key === 'source') {
        option(control, '', 'Constant / default only');
        for (const header of state.table?.headers || []) option(control, header);
        if (mapping.source && !state.table?.headers.includes(mapping.source)) {
          option(control, mapping.source, `${mapping.source} (missing)`);
        }
      } else {
        control.type = 'text';
        if (key === 'default') control.placeholder = 'If blank';
      }
      control.value = key === 'format' ? mapping.format || 'plain' : mapping[key];
      control.addEventListener(['source', 'format'].includes(key) ? 'change' : 'input', () => {
        mapping[key] = control.value;
        refreshExport();
      });
      cell.append(control);
      row.append(cell);
    }
    const actions = el('td');
    const group = el('div', undefined, 'row-actions');
    for (const [label, action, disabled] of [
      ['↑', () => {
        [state.mappings[index - 1], state.mappings[index]] = [state.mappings[index], state.mappings[index - 1]];
      }, index === 0],
      ['↓', () => {
        [state.mappings[index + 1], state.mappings[index]] = [state.mappings[index], state.mappings[index + 1]];
      }, index === state.mappings.length - 1],
      ['×', () => { state.mappings.splice(index, 1); }, false]
    ]) {
      const button = el('button', label);
      button.disabled = disabled;
      button.setAttribute('aria-label', `${label === '×' ? 'Remove' : label === '↑' ? 'Move up' : 'Move down'} mapping ${index + 1}`);
      button.addEventListener('click', () => { action(); renderMappings(); refreshExport(); });
      group.append(button);
    }
    actions.append(group);
    row.append(actions);
    $('mappings').append(row);
  });
  $('add-mapping').disabled = !state.table;
  $('save-profile').disabled = !state.mappings.length;
}

function refreshExport() {
  const result = state.table ? buildExport(state.table, state.mappings, state.filter, state.parent) : null;
  if (result && state.sourceErrors.length) result.errors.unshift(...state.sourceErrors.slice(0, 20));
  if (result && state.sourceErrors.length > 20) result.errors.push(`${state.sourceErrors.length - 20} more source cell errors.`);
  state.result = result;
  const valid = result && !result.errors.length && !state.loading;
  $('status').textContent = state.loading ? 'Reading file…' : !result ? 'Waiting for source' : valid ? 'Ready to export' : 'Needs attention';
  $('status').className = `status ${!result ? 'idle' : valid ? 'ready' : 'blocked'}`;
  $('download').disabled = !valid;
  $('download-all').disabled = !valid || result.batches < 2;
  $('batch').disabled = !valid;
  $('batch').replaceChildren();
  if (!result?.batches) option($('batch'), '1', 'Batch 1');
  else for (let batch = 1; batch <= result.batches; batch++) {
    const entries = result.exportBatches?.[batch - 1];
    const count = entries ? entries.length : Math.min(BATCH_SIZE, result.rows.length - (batch - 1) * BATCH_SIZE);
    option($('batch'), String(batch), `Batch ${batch} · ${count.toLocaleString()} rows`);
  }
  state.batch = Math.min(Math.max(1, state.batch), result?.batches || 1);
  $('batch').value = String(state.batch);
  $('validation').replaceChildren();
  $('validation').classList.toggle('hidden', !result?.errors.length);
  if (result?.errors.length) {
    $('validation').append(el('strong', 'Resolve these before exporting'));
    const list = el('ul');
    for (const error of result.errors) list.append(el('li', error));
    $('validation').append(list);
  }
  const head = $('preview').querySelector('thead');
  const body = $('preview').querySelector('tbody');
  head.replaceChildren();
  body.replaceChildren();
  const visibleRows = result?.exportBatches?.[state.batch - 1] || result?.rows.slice((state.batch - 1) * BATCH_SIZE, state.batch * BATCH_SIZE) || [];
  $('preview-empty').classList.toggle('hidden', visibleRows.length > 0);
  if (visibleRows.length) {
    const tr = el('tr');
    tr.append(el('th', 'Source row'));
    result.headers.forEach(header => tr.append(el('th', header)));
    head.append(tr);
    for (const row of visibleRows.slice(0, 100)) {
      const tr = el('tr');
      if (row.isParent) tr.classList.add('parent-reference');
      tr.append(el('td', row.isParent ? 'Existing Epic' : String(row.sourceRow)));
      row.values.forEach((value, index) => {
        const cell = el('td');
        const mapping = state.mappings.find(mapping => mapping.enabled && mapping.field === result.headers[index]);
        if ($('preview-html').checked && mapping?.format && mapping.format !== 'plain' && !row.isParent) {
          // Non-plain values are generated by formatText from escaped source text only.
          cell.innerHTML = value;
          cell.classList.add('rich-text');
        } else cell.textContent = value;
        tr.append(cell);
      });
      body.append(tr);
    }
  }
  $('preview-label').textContent = result?.rows.length
    ? `${result.rows.length.toLocaleString()} new items · ${result.headers.length} CSV fields · preview ${Math.min(100, visibleRows.length)} of ${visibleRows.length} batch rows`
    : 'Your mapped output will appear here';
  $('batch-label').textContent = result?.batches ? `${result.batches} batch${result.batches === 1 ? '' : 'es'} · ${state.parent.enabled ? 'Includes Epic references' : 'UTF-8, comma-delimited'}` : 'Up to 1,000 rows per CSV';
}

function parseSource() {
  if (!state.workbook && state.csvText === null) return;
  try {
    const headerRow = Number($('header-row').value);
    let parsed;
    if (state.workbook) {
      parsed = worksheetRecords(state.workbook.getWorksheet(Number($('sheet').value)));
    } else {
      const delimiter = $('delimiter').value;
      parsed = { ...parseCsv(state.csvText, delimiter === 'tab' ? '\t' : delimiter), errors: [] };
    }
    state.table = tableFromRecords(parsed.records, headerRow);
    if (!state.parent.epicSource) state.parent.epicSource = state.table.headers.find(header => header.toLowerCase() === 'epic') || '';
    state.sourceErrors = parsed.errors.filter(error => error.row >= headerRow).map(error => error.message);
    if (!state.mappings.length) state.mappings = suggestedMappings(state.table.headers);
    state.batch = 1;
    $('source-rows').textContent = state.table.rows.length.toLocaleString();
    $('source-cols').textContent = String(state.table.headers.length);
    populateFilters();
    populateParent();
    renderMappings();
    refreshExport();
  } catch (error) {
    state.table = null;
    state.sourceErrors = [];
    $('source-rows').textContent = '—';
    $('source-cols').textContent = '—';
    populateFilters();
    populateParent();
    renderMappings();
    refreshExport();
    report(error, 'Cannot read this table');
  }
}

async function loadFile(file) {
  if (!file || state.loading) return;
  state.loading = true;
  state.table = null;
  state.workbook = null;
  state.csvText = null;
  state.sourceErrors = [];
  $('notice').classList.add('hidden');
  $('source-file').disabled = true;
  $('sheet').disabled = true;
  $('header-row').disabled = true;
  $('source-rows').textContent = '—';
  $('source-cols').textContent = '—';
  refreshExport();
  try {
    const extension = file.name.split('.').at(-1).toLowerCase();
    if (!['xlsx', 'csv'].includes(extension)) throw new Error('Choose an .xlsx or UTF-8 .csv file. Legacy .xls files must be saved as .xlsx first.');
    state.fileName = file.name;
    $('header-row').value = '1';
    const data = await file.arrayBuffer();
    $('sheet').replaceChildren();
    if (extension === 'xlsx') {
      if (!globalThis.ExcelJS) throw new Error('Excel reader did not load. Run npm install and reload this page.');
      const workbook = new globalThis.ExcelJS.Workbook();
      await workbook.xlsx.load(data);
      await restoreEmptyFormulaCaches(data, workbook);
      const visible = workbook.worksheets.filter(sheet => sheet.state === 'visible' || !sheet.state);
      if (!visible.length) throw new Error('The workbook has no visible worksheets.');
      state.workbook = workbook;
      visible.forEach(sheet => option($('sheet'), String(sheet.id), sheet.name));
      $('sheet').disabled = false;
      $('delimiter-label').classList.add('hidden');
    } else {
      try { state.csvText = new TextDecoder('utf-8', { fatal: true }).decode(data); }
      catch (error) { throw new Error(`CSV must be UTF-8. Resave as CSV UTF-8 in Excel. (${error.message})`); }
      option($('sheet'), 'csv', 'CSV data');
      $('delimiter-label').classList.remove('hidden');
    }
    $('file-label').textContent = file.name;
    $('header-row').disabled = false;
    state.loading = false;
    parseSource();
    if (state.table) notify(`Loaded ${file.name}. ${state.table.rows.length.toLocaleString()} nonempty source rows. ${state.mappings.length} mappings applied; review the preview before exporting.`, true);
  } catch (error) {
    state.table = null;
    state.workbook = null;
    state.csvText = null;
    $('file-label').textContent = 'Drop a spreadsheet here';
    $('sheet').replaceChildren();
    option($('sheet'), '', 'Load a file first');
    populateFilters();
    populateParent();
    renderMappings();
    report(error, 'File import failed');
  } finally {
    state.loading = false;
    $('source-file').disabled = false;
    $('source-file').value = '';
    refreshExport();
  }
}

$('source-file').addEventListener('change', event => loadFile(event.target.files[0]));
for (const eventName of ['dragenter', 'dragover']) $('dropzone').addEventListener(eventName, event => {
  event.preventDefault(); $('dropzone').classList.add('dragging');
});
for (const eventName of ['dragleave', 'drop']) $('dropzone').addEventListener(eventName, event => {
  event.preventDefault(); $('dropzone').classList.remove('dragging');
});
$('dropzone').addEventListener('drop', event => {
  if (event.dataTransfer.files.length !== 1) { notify('Drop exactly one source file at a time.'); return; }
  loadFile(event.dataTransfer.files[0]);
});
for (const id of ['sheet', 'header-row', 'delimiter']) $(id).addEventListener('change', () => {
  $('notice').classList.add('hidden');
  parseSource();
});
$('add-mapping').addEventListener('click', () => {
  state.mappings.push({ enabled: true, field: '', source: '', default: '', prefix: '', suffix: '', format: 'plain' });
  renderMappings(); refreshExport();
});
$('filter-column').addEventListener('change', () => {
  state.filter.column = $('filter-column').value;
  state.filter.value = '';
  state.batch = 1;
  populateFilters(); refreshExport();
});
$('filter-value').addEventListener('input', () => {
  state.filter.value = $('filter-value').value;
  state.batch = 1;
  refreshExport();
});
$('batch').addEventListener('change', () => {
  state.batch = Number($('batch').value);
  refreshExport();
});
$('preview-html').addEventListener('change', refreshExport);
$('parent-enabled').addEventListener('change', () => {
  state.parent.enabled = $('parent-enabled').checked;
  state.batch = 1;
  populateParent(); refreshExport();
});
$('parent-mode').addEventListener('change', () => {
  state.parent.mode = $('parent-mode').value;
  state.batch = 1;
  populateParent(); refreshExport();
});
$('epic-source').addEventListener('change', () => {
  state.parent.epicSource = $('epic-source').value;
  state.batch = 1;
  updateEpicCounts(); refreshExport();
});
$('add-epic').addEventListener('click', () => {
  state.parent.epics.push({ sourceName: '', id: '', title: '' });
  renderEpicLookups(); refreshExport();
});
for (const key of ['id', 'title']) {
  $(`parent-${key}-source`).addEventListener('change', () => {
    state.parent[`${key}Source`] = $(`parent-${key}-source`).value;
    state.batch = 1; refreshExport();
  });
  $(`parent-${key}-default`).addEventListener('input', () => {
    state.parent[`${key}Default`] = $(`parent-${key}-default`).value;
    state.batch = 1; refreshExport();
  });
}
$('download').addEventListener('click', () => {
  try {
    const result = state.result;
    const rows = batchRows(result, state.batch);
    downloadFile(`${baseName()}-ado-batch-${state.batch}.csv`, toCsv(result.headers, rows), 'text/csv;charset=utf-8');
    const items = rows.filter(row => !row.isParent).length;
    notify(`Downloaded batch ${state.batch}: ${items} new items${rows.length > items ? ` and ${rows.length - items} existing Epic references` : ''}. Review in ADO and import only once.`, true);
  } catch (error) { report(error, 'CSV export failed'); }
});
$('download-all').addEventListener('click', async () => {
  $('download-all').disabled = true;
  try {
    const result = state.result;
    const name = baseName();
    if (!globalThis.JSZip) throw new Error('ZIP library did not load. Run npm install and reload the page.');
    const zip = new globalThis.JSZip();
    for (let batch = 1; batch <= result.batches; batch++) {
      zip.file(`${name}-ado-batch-${String(batch).padStart(3, '0')}.csv`, toCsv(result.headers, batchRows(result, batch)));
    }
    const data = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
    downloadFile(`${name}-ado-batches.zip`, data, 'application/zip');
    notify(`Downloaded ${result.batches} CSV batches in one ZIP. Extract and import each CSV once.`, true);
  } catch (error) { report(error, 'ZIP export failed'); }
  finally { refreshExport(); }
});
$('save-profile').addEventListener('click', () => {
  const profile = { version: 3, mappings: state.mappings.map(mapping => ({ ...mapping, format: mapping.format || 'plain' })), filter: state.filter, parent: state.parent };
  downloadFile('ado-mapping-profile.json', JSON.stringify(profile, null, 2), 'application/json');
  notify('Profile downloaded, including field formats, filters, and your Epic name-to-ID table. Source rows are not included. Store securely if settings are sensitive.', true);
});
$('load-profile').addEventListener('click', () => $('profile-file').click());
$('profile-file').addEventListener('change', async event => {
  const file = event.target.files[0];
  if (!file) return;
  try {
    const profile = readProfile(JSON.parse(await file.text()));
    state.mappings = profile.mappings;
    state.filter = profile.filter;
    state.parent = profile.parent
      ? { ...lookupParent(), ...profile.parent, mode: profile.parent.mode || 'direct' }
      : lookupParent();
    state.batch = 1;
    populateFilters(); populateParent(); renderMappings(); refreshExport();
    notify(`Loaded ${file.name}. Review any missing source columns or field validation errors before exporting.`, true);
  } catch (error) { report(error, 'Profile was not loaded; previous mappings are unchanged'); }
  finally { $('profile-file').value = ''; }
});
FIELD_SUGGESTIONS.forEach(field => option($('ado-fields'), field));
renderMappings();
populateParent();
refreshExport();
