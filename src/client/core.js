export const BATCH_SIZE = 1000;
export const FIELD_SUGGESTIONS = [
  'Work Item Type', 'Title', 'Description', 'Acceptance Criteria',
  'Iteration Path', 'Area Path', 'Tags', 'Priority', 'Story Points',
  'Effort', 'Original Estimate', 'Remaining Work', 'Completed Work',
  'Activity', 'Value Area', 'Business Value', 'Risk', 'Size',
  'Severity', 'Repro Steps', 'System Info', 'Found In', 'Integrated In',
  'Resolved Reason', 'Blocked', 'Due Date', 'Start Date', 'Finish Date',
  'Target Date', 'Stack Rank', 'Backlog Priority'
];
export const FORMATS = ['plain', 'paragraphs', 'story', 'markdown'];
const MAPPING_FORMATS = [...FORMATS, 'iteration-root'];
export const emptyParent = () => ({ enabled: false, idSource: '', idDefault: '', titleSource: '', titleDefault: '' });
export const DEFAULT_EPICS = [
  ['1572169', 'Order History & Reordering'],
  ['1572170', 'Login & User Session'],
  ['1572171', 'Pricing, Promotions & Entitlements'],
  ['1572172', 'My Account'],
  ['1572173', 'Product Discovery'],
  ['1572174', 'Guest Experience'],
  ['1572175', 'Home Page'],
  ['1572176', 'Global Navigation & Styling'],
  ['1572177', 'Registration'],
  ['1620336', 'Quick Order'],
  ['1620342', 'Cart & Related Validations'],
  ['1620343', 'Checkout & Related Validations'],
  ['1620344', 'Lists'],
  ['1620345', 'Self- Service & Content'],
  ['1620346', 'Inventory Availability'],
  ['1620347', 'Order Management'],
  ['1620348', 'Order Fulfillment'],
  ['1620349', 'Invoicing & Returns'],
  ['1620350', 'Notifications & Exceptions'],
  ['1620351', 'Account Management'],
  ['1620352', 'Merchandising']
];
export const lookupParent = () => ({
  ...emptyParent(), mode: 'lookup', epicSource: '',
  epics: DEFAULT_EPICS.map(([id, title]) => ({ sourceName: title, id, title }))
});
const epicKey = value => value.trim().toLowerCase();
const validEpicId = id => /^[1-9]\d*$/.test(id) && Number.isSafeInteger(Number(id));
const forbidden = new Set([
  'id', 'state', 'assigned to', 'created by', 'created date', 'changed by', 'changed date',
  'parent', 'title 1', 'title 2'
]);

function escapeHtml(value) {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;').replaceAll("'", '&#39;');
}

function inlineMarkdown(value) {
  const tokens = [];
  const escaped = escapeHtml(value).replace(/`([^`]+)`/g, (_, code) => {
    tokens.push(`<code>${code}</code>`);
    return `\u0000${tokens.length - 1}\u0000`;
  });
  return escaped.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
    .replace(/__([^_]+)__/g, '<strong>$1</strong>')
    .replace(/(^|[\s(])\*([^*\n]+)\*(?=$|[\s.,!?)])/g, '$1<em>$2</em>')
    .replace(/\u0000(\d+)\u0000/g, (_, index) => tokens[Number(index)]);
}

export function formatText(value, mode = 'plain') {
  if (!FORMATS.includes(mode)) throw new Error(`Unknown text format "${mode}".`);
  if (mode === 'plain' || value === '') return value;
  const normalized = value.replace(/\r\n?/g, '\n').replaceAll('\u0000', '');
  if (mode === 'story') {
    const story = normalized.trim().match(/^As a\s+([\s\S]+?)\s+I want(?: to)?\s+([\s\S]+?)\s+so that\s+([\s\S]+)$/i);
    if (story) return ['As a', 'I want to', 'So that'].map((label, index) =>
      `<p><strong>${label}</strong> ${escapeHtml(story[index + 1].trim()).replaceAll('\n', '<br>')}</p>`).join('');
  }
  if (mode !== 'markdown') return normalized.split(/\n\s*\n/)
    .map(paragraph => `<p>${escapeHtml(paragraph).replaceAll('\n', '<br>')}</p>`).join('');
  const output = [];
  let paragraph = [];
  let list = null;
  function flushParagraph() {
    if (paragraph.length) output.push(`<p>${paragraph.map(inlineMarkdown).join('<br>')}</p>`);
    paragraph = [];
  }
  function closeList() {
    if (list) output.push(`</${list}>`);
    list = null;
  }
  for (const line of normalized.split('\n')) {
    const heading = line.match(/^(#{1,3})\s+(.+)$/);
    const bullet = line.match(/^\s*(?:[-*+]\s+|(\d+)[.)]\s+)(.+)$/);
    if (!line.trim()) { flushParagraph(); closeList(); }
    else if (heading) {
      flushParagraph(); closeList();
      const tag = `h${heading[1].length + 1}`;
      output.push(`<${tag}>${inlineMarkdown(heading[2])}</${tag}>`);
    } else if (bullet) {
      flushParagraph();
      const tag = bullet[1] ? 'ol' : 'ul';
      if (list !== tag) {
        closeList();
        output.push(`<${tag}${bullet[1] ? ` start="${Number(bullet[1])}"` : ''}>`);
        list = tag;
      }
      output.push(`<li>${inlineMarkdown(bullet[2])}</li>`);
    } else if (list && /^\s+\S/.test(line)) {
      output[output.length - 1] = output[output.length - 1].replace(/<\/li>$/, `<br>${inlineMarkdown(line.trim())}</li>`);
    } else { closeList(); paragraph.push(line); }
  }
  flushParagraph(); closeList();
  return output.join('');
}

export function text(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return String(value);
}

export function excelCellValue(value, row, column, errors) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value.toISOString();
  if (typeof value !== 'object') return text(value);
  if ('formula' in value || 'sharedFormula' in value) {
    if (value.result === undefined || value.result === null) {
      errors.push({ row, message: `Source row ${row}, column ${column}: formula has no saved result. Recalculate and save in Excel, then reload.` });
      return '';
    }
    return excelCellValue(value.result, row, column, errors);
  }
  if (value.error) {
    errors.push({ row, message: `Source row ${row}, column ${column}: Excel error ${value.error}. Fix it in the source file.` });
    return '';
  }
  if (Array.isArray(value.richText)) return value.richText.map(run => run.text).join('');
  if ('text' in value) return text(value.text);
  errors.push({ row, message: `Source row ${row}, column ${column}: unsupported Excel cell value. Replace it with a plain value.` });
  return '';
}

export function worksheetRecords(sheet) {
  const records = [];
  const errors = [];
  sheet.eachRow({ includeEmpty: false }, row => {
    const values = [];
    row.eachCell({ includeEmpty: false }, (cell, column) => {
      let value = cell.value;
      if (value && typeof value === 'object' && ('formula' in value || 'sharedFormula' in value) && cell.result !== undefined) {
        value = { ...value, result: cell.result };
      }
      values[column - 1] = excelCellValue(value, row.number, column, errors);
    });
    records[row.number - 1] = values;
  });
  return { records: Array.from({ length: records.length }, (_, index) => records[index] || []), errors };
}

export function parseCsv(input, delimiter = 'auto') {
  const source = input.replace(/^\uFEFF/, '');
  if (delimiter === 'auto') {
    const counts = new Map([[',', 0], [';', 0], ['\t', 0]]);
    let quoted = false;
    for (let index = 0; index < source.length; index++) {
      const char = source[index];
      if (char === '"') {
        if (quoted && source[index + 1] === '"') index++;
        else quoted = !quoted;
      }
      if (!quoted && (char === '\r' || char === '\n')) break;
      if (!quoted && counts.has(char)) counts.set(char, counts.get(char) + 1);
    }
    delimiter = [...counts].sort((a, b) => b[1] - a[1])[0][0];
  }
  if (![',', ';', '\t'].includes(delimiter)) throw new Error('Unsupported CSV delimiter.');
  const records = [];
  let row = [];
  let field = '';
  let quoted = false;
  let closedQuote = false;
  for (let index = 0; index < source.length; index++) {
    const char = source[index];
    if (quoted) {
      if (char === '"') {
        if (source[index + 1] === '"') { field += '"'; index++; }
        else { quoted = false; closedQuote = true; }
      } else field += char;
    } else if (char === delimiter) {
      row.push(field); field = ''; closedQuote = false;
    } else if (char === '\n' || char === '\r') {
      if (char === '\r' && source[index + 1] === '\n') index++;
      row.push(field); records.push(row); row = []; field = ''; closedQuote = false;
    } else if (char === '"') {
      if (field || closedQuote) throw new Error('Malformed CSV: a quote appears inside an unquoted field.');
      quoted = true;
    } else {
      if (closedQuote) throw new Error('Malformed CSV: unexpected text after a closing quote.');
      field += char;
    }
  }
  if (quoted) throw new Error('Malformed CSV: an opening quote has no closing quote.');
  if (field || row.length || closedQuote) { row.push(field); records.push(row); }
  return { records, delimiter };
}

export function tableFromRecords(records, headerRow = 1) {
  if (!Number.isInteger(headerRow) || headerRow < 1 || headerRow > records.length) {
    throw new Error('Choose a header row that exists in this sheet.');
  }
  const header = records[headerRow - 1];
  const data = records.slice(headerRow);
  let width = header.length;
  for (const row of data) width = Math.max(width, row.length);
  while (width > 0 && text(header[width - 1]).trim() === '' &&
    data.every(row => text(row[width - 1]) === '')) width--;
  const headers = Array.from({ length: width }, (_, index) => text(header[index]).trim());
  if (!headers.length) throw new Error('No columns found in the selected header row.');
  if (headers.some(name => !name)) throw new Error('Every populated source column needs a header. Fix blank headers or choose another header row.');
  if (new Set(headers.map(name => name.toLowerCase())).size !== headers.length) {
    throw new Error('Source headers must be unique (ignoring case and surrounding spaces). Rename duplicate columns.');
  }
  const rows = data.map((row, index) => ({
    sourceRow: headerRow + index + 1,
    values: headers.map((_, column) => text(row[column]))
  })).filter(row => row.values.some(value => value !== ''));
  return { headers, rows };
}

export function suggestedMappings(headers) {
  const match = names => names.map(name => headers.find(header => header.toLowerCase() === name.toLowerCase())).find(Boolean) || '';
  const item = (field, source = '', defaultValue = '', enabled = true) => ({
    enabled, field, source, default: defaultValue, prefix: '', suffix: '',
    format: field === 'Description' ? 'story' : field === 'Acceptance Criteria' ? 'markdown' : 'plain'
  });
  const result = [
    item('Work Item Type', '', 'User Story'),
    item('Title', match(['Work Summary (Title)', 'Title', 'Work Summary'])),
  ];
  for (const [field, candidates] of [
    ['Description', ['User Story (Description)', 'Description']],
    ['Acceptance Criteria', ['Acceptance Criteria']]
  ]) {
    const source = match(candidates);
    if (source) result.push(item(field, source));
  }
  const sprint = match(['Sprint', 'Iteration Path']);
  if (sprint) result.push(item('Iteration Path', sprint, '', false));
  return result;
}

export function mappedValue(row, headers, mapping) {
  const value = mapping.source ? row.values[headers.indexOf(mapping.source)] : '';
  if (mapping.format === 'iteration-root') {
    const root = mapping.default.trim().replace(/\\+$/, '');
    const sprint = value.trim().split('\\').pop().trim();
    return sprint ? `${root}\\${sprint}` : root;
  }
  const base = value === '' ? mapping.default : value;
  return formatText(base === '' ? '' : `${mapping.prefix}${base}${mapping.suffix}`, mapping.format || 'plain');
}

export function buildExport(table, mappings, filter = { column: '', value: '' }, parent = emptyParent()) {
  const errors = [];
  const epicLookup = new Map();
  const enabled = mappings.filter(mapping => mapping.enabled);
  const seen = new Set();
  for (const [index, mapping] of mappings.entries()) {
    if (!mapping.enabled) continue;
    const field = mapping.field.trim();
    const key = field.toLowerCase();
    const label = `Mapping ${index + 1}`;
    if (!field) errors.push(`${label}: enter an ADO field name.`);
    else if (field !== mapping.field) errors.push(`${label}: remove spaces around "${mapping.field}".`);
    if (seen.has(key)) errors.push(`${label}: duplicate ADO field "${field}".`);
    seen.add(key);
    if (forbidden.has(key)) errors.push(`${label}: "${field}" is not allowed in this new-item creation workflow.`);
    if (!MAPPING_FORMATS.includes(mapping.format || 'plain')) errors.push(`${label}: select a valid text format.`);
    if (mapping.format === 'iteration-root') {
      if (field !== 'Iteration Path') errors.push(`${label}: Iteration root mode requires the "Iteration Path" field.`);
      if (!mapping.default.trim().replace(/\\+$/, '')) errors.push(`${label}: enter an Iteration Path root in Default / constant.`);
      if (mapping.prefix || mapping.suffix) errors.push(`${label}: remove the prefix and suffix when using Iteration root mode.`);
    }
    if (field === 'Work Item Type' || field === 'Title') {
      if (mapping.format && mapping.format !== 'plain') errors.push(`${label}: "${field}" must use plain text.`);
    }
    if (mapping.source && !table.headers.includes(mapping.source)) errors.push(`${label}: source column "${mapping.source}" is missing.`);
    if (!mapping.source && mapping.default === '') errors.push(`${label}: select a source column or enter a constant/default.`);
  }
  for (const field of ['Work Item Type', 'Title']) {
    if (!enabled.some(mapping => mapping.field === field)) errors.push(`Enable exactly one "${field}" mapping.`);
  }
  if (filter.column && !table.headers.includes(filter.column)) errors.push(`Filter column "${filter.column}" is missing.`);
  if (parent.enabled) {
    if (parent.mode === 'lookup') {
      if (!parent.epicSource || !table.headers.includes(parent.epicSource)) errors.push('Epic linking: choose the source column containing Epic names.');
      if (!Array.isArray(parent.epics) || !parent.epics.length) errors.push('Epic linking: add at least one name-to-ID mapping.');
      const titles = new Map();
      for (const [index, epic] of (parent.epics || []).entries()) {
        const key = epicKey(epic.sourceName);
        const id = epic.id.trim();
        const title = epic.title.trim();
        if (!key) errors.push(`Epic mapping ${index + 1}: enter the Epic name used in your source.`);
        if (epicLookup.has(key)) errors.push(`Epic mapping ${index + 1}: duplicate source Epic name "${epic.sourceName}".`);
        if (!validEpicId(id)) errors.push(`Epic mapping ${index + 1}: ADO Epic ID must be a positive integer.`);
        if (!title || title.length > 255) errors.push(`Epic mapping ${index + 1}: current ADO title must have 1–255 characters.`);
        if (titles.has(id) && titles.get(id) !== title) errors.push(`Epic mapping ${index + 1}: conflicting ADO titles for Epic ID ${id}.`);
        titles.set(id, title);
        epicLookup.set(key, { id, title });
      }
    } else {
      for (const [key, label] of [['id', 'Epic ID'], ['title', 'Epic title']]) {
        if (parent[`${key}Source`] && !table.headers.includes(parent[`${key}Source`])) errors.push(`Parent ${label}: source column is missing.`);
        if (!parent[`${key}Source`] && !parent[`${key}Default`]) errors.push(`Parent ${label}: choose a source column or enter a constant.`);
      }
    }
  }
  if (errors.length) return { errors, headers: enabled.map(mapping => mapping.field), rows: [], batches: 0 };
  const filterIndex = table.headers.indexOf(filter.column);
  const selected = table.rows.filter(row => !filter.column ||
    row.values[filterIndex].toLowerCase() === text(filter.value).toLowerCase());
  const titleIndex = enabled.findIndex(mapping => mapping.field === 'Title');
  const typeIndex = enabled.findIndex(mapping => mapping.field === 'Work Item Type');
  let invalidCount = 0;
  const parentTitles = new Map();
  const rows = selected.map(row => {
    const values = enabled.map(mapping => mappedValue(row, table.headers, mapping));
    for (const [index, field] of [[titleIndex, 'Title'], [typeIndex, 'Work Item Type']]) {
      if (!values[index].trim()) {
        invalidCount++;
        if (errors.length < 20) errors.push(`Source row ${row.sourceRow}: "${field}" is blank.`);
      }
    }
    if (values[titleIndex].length > 255) {
      invalidCount++;
      if (errors.length < 20) errors.push(`Source row ${row.sourceRow}: Title exceeds 255 characters.`);
    }
    let parentId = '';
    let parentTitle = '';
    if (parent.enabled) {
      if (parent.mode === 'lookup') {
        const sourceName = row.values[table.headers.indexOf(parent.epicSource)];
        const epic = epicLookup.get(epicKey(sourceName));
        if (!epic) {
          if (errors.length < 20) errors.push(`Source row ${row.sourceRow}: ${sourceName.trim() ? `Epic "${sourceName}" has no mapping` : 'Epic name is blank'}. Add or correct it in the Epic lookup table.`);
          invalidCount++;
          return { sourceRow: row.sourceRow, values };
        }
        parentId = epic.id;
        parentTitle = epic.title;
      } else {
        const readParent = key => {
          const value = parent[`${key}Source`] ? row.values[table.headers.indexOf(parent[`${key}Source`])] : '';
          return (value === '' ? parent[`${key}Default`] : value).trim();
        };
        parentId = readParent('id');
        parentTitle = readParent('title');
      }
      if (!validEpicId(parentId)) {
        if (errors.length < 20) errors.push(`Source row ${row.sourceRow}: parent Epic ID must be a positive integer ADO ID.`);
        invalidCount++;
      }
      if (!parentTitle || parentTitle.length > 255) {
        if (errors.length < 20) errors.push(`Source row ${row.sourceRow}: parent Epic title must have 1–255 characters.`);
        invalidCount++;
      }
      if (parentTitles.has(parentId) && parentTitles.get(parentId) !== parentTitle) {
        if (errors.length < 20) errors.push(`Source row ${row.sourceRow}: conflicting titles for parent Epic ID ${parentId}.`);
        invalidCount++;
      }
      parentTitles.set(parentId, parentTitle);
      if (values[typeIndex] !== 'User Story') {
        if (errors.length < 20) errors.push(`Source row ${row.sourceRow}: Epic parent mode currently supports User Story children only.`);
        invalidCount++;
      }
    }
    return { sourceRow: row.sourceRow, values, ...(parent.enabled ? { parentId, parentTitle } : {}) };
  });
  if (invalidCount > 20) errors.push(`${invalidCount - 20} more required-value errors. Fix source data before exporting.`);
  if (!selected.length) errors.push('No source rows match the current filter.');
  const headers = enabled.map(mapping => mapping.field);
  if (!parent.enabled || errors.length) return { errors, headers, rows, batches: Math.ceil(rows.length / BATCH_SIZE) };
  const extra = headers.map((_, index) => index).filter(index => index !== titleIndex && index !== typeIndex);
  const hierarchyHeaders = ['ID', 'Work Item Type', 'Title 1', 'Title 2', ...extra.map(index => headers[index])];
  const groups = new Map();
  for (const row of rows) {
    if (!groups.has(row.parentId)) groups.set(row.parentId, []);
    groups.get(row.parentId).push(row);
  }
  const exportBatches = [];
  let current = [];
  for (const children of groups.values()) {
    let anchored = false;
    for (const child of children) {
      const required = anchored ? 1 : 2;
      if (current.length + required > BATCH_SIZE) {
        exportBatches.push(current); current = []; anchored = false;
      }
      if (!anchored) {
        current.push({ sourceRow: null, isParent: true, values: [child.parentId, 'Epic', child.parentTitle, '', ...extra.map(() => '')] });
        anchored = true;
      }
      current.push({ sourceRow: child.sourceRow, values: ['', child.values[typeIndex], '', child.values[titleIndex], ...extra.map(index => child.values[index])] });
    }
  }
  if (current.length) exportBatches.push(current);
  return { errors, headers: hierarchyHeaders, rows, exportBatches, batches: exportBatches.length };
}

export function batchRows(result, batch) {
  if (result.errors.length) throw new Error('Resolve all validation errors before exporting.');
  if (!Number.isInteger(batch) || batch < 1 || batch > result.batches) throw new Error('Invalid export batch.');
  return result.exportBatches ? result.exportBatches[batch - 1] : result.rows.slice((batch - 1) * BATCH_SIZE, batch * BATCH_SIZE);
}

export function toCsv(headers, rows) {
  const quote = value => `"${text(value).replaceAll('"', '""')}"`;
  return '\uFEFF' + [headers, ...rows.map(row => row.values)]
    .map(row => row.map(quote).join(',')).join('\r\n') + '\r\n';
}

export function readProfile(raw) {
  if (!raw || ![1, 2, 3].includes(raw.version) || !Array.isArray(raw.mappings) || !raw.mappings.length) {
    throw new Error('Not a valid version 1, 2 or 3 ADO Mapper profile.');
  }
  const mappings = raw.mappings.map((mapping, index) => {
    if (!mapping || typeof mapping.enabled !== 'boolean' ||
      ['field', 'source', 'default', 'prefix', 'suffix'].some(key => typeof mapping[key] !== 'string')) {
      throw new Error(`Invalid mapping ${index + 1} in profile.`);
    }
    if (raw.version >= 2 && !MAPPING_FORMATS.includes(mapping.format)) throw new Error(`Invalid text format in mapping ${index + 1}.`);
    return { enabled: mapping.enabled, ...Object.fromEntries(['field', 'source', 'default', 'prefix', 'suffix'].map(key => [key, mapping[key]])),
      ...(raw.version >= 2 ? { format: mapping.format } : {}) };
  });
  if (!raw.filter || typeof raw.filter.column !== 'string' || typeof raw.filter.value !== 'string') {
    throw new Error('Profile must include a valid filter configuration.');
  }
  if (raw.version >= 2 && (!raw.parent || typeof raw.parent.enabled !== 'boolean' ||
    ['idSource', 'idDefault', 'titleSource', 'titleDefault'].some(key => typeof raw.parent[key] !== 'string'))) {
    throw new Error('Invalid parent Epic settings in profile.');
  }
  if (raw.version === 3) {
    if (!['lookup', 'direct'].includes(raw.parent.mode) || typeof raw.parent.epicSource !== 'string' ||
      !Array.isArray(raw.parent.epics) || raw.parent.epics.some(epic =>
        !epic || ['sourceName', 'id', 'title'].some(key => typeof epic[key] !== 'string'))) {
      throw new Error('Invalid Epic lookup table in profile.');
    }
  }
  return { mappings, filter: { column: raw.filter.column, value: raw.filter.value },
    ...(raw.version >= 2 ? { parent: {
      ...emptyParent(), ...Object.fromEntries(Object.keys(emptyParent()).map(key => [key, raw.parent[key]])),
      ...(raw.version === 3 ? {
        mode: raw.parent.mode, epicSource: raw.parent.epicSource,
        epics: raw.parent.epics.map(epic => ({ sourceName: epic.sourceName, id: epic.id, title: epic.title }))
      } : {})
    } } : {}) };
}
