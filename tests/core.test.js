import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import ExcelJS from 'exceljs';
import { fileURLToPath } from 'node:url';
import {
  parseCsv, tableFromRecords, suggestedMappings, buildExport,
  batchRows, toCsv, readProfile, text, excelCellValue, worksheetRecords,
  formatText, emptyParent, FIELD_SUGGESTIONS, lookupParent, DEFAULT_EPICS
} from '../src/client/core.js';

const table = tableFromRecords([
  ['Title', 'Description', 'Sprint', 'Points', 'Flag'],
  ['One', 'Comma, "quote"\nand café', 'Sprint 1', 0, false],
  ['Two', '', 'Sprint 2', '', true]
]);
const mappings = () => [
  { enabled: true, field: 'Work Item Type', source: '', default: 'User Story', prefix: '', suffix: '' },
  { enabled: true, field: 'Title', source: 'Title', default: '', prefix: '', suffix: '' }
];
const mapping = (field, source, defaultValue = '', prefix = '', suffix = '') =>
  ({ enabled: true, field, source, default: defaultValue, prefix, suffix });

test('CSV handles BOM, delimiters, quotes, multiline, Unicode and empty fields', () => {
  const source = '\uFEFFTitle;Description;Empty\r\nOne;"hello; ""café""\nnext";\r\n';
  const parsed = parseCsv(source);
  assert.equal(parsed.delimiter, ';');
  assert.deepEqual(parsed.records, [['Title', 'Description', 'Empty'], ['One', 'hello; "café"\nnext', '']]);
  assert.deepEqual(parseCsv('A\tB\nx\ty').records, [['A', 'B'], ['x', 'y']]);
  assert.deepEqual(parseCsv('A,B\n"",""').records, [['A', 'B'], ['', '']]);
  for (const invalid of ['A\n"unterminated', 'A\n"quoted"x', 'A\nun"quoted']) {
    assert.throws(() => parseCsv(invalid), /Malformed CSV/);
  }
});

test('header selection, duplicate and blank headers, empty rows, zero/false', () => {
  assert.equal(table.rows[0].values[3], '0');
  assert.equal(table.rows[0].values[4], 'FALSE');
  assert.equal(text(false), 'FALSE');
  assert.equal(tableFromRecords([['A'], [''], [0], [false], []]).rows.length, 2);
  assert.deepEqual(tableFromRecords([['intro'], ['A', 'B'], ['x', 'y']], 2).headers, ['A', 'B']);
  assert.throws(() => tableFromRecords([['A', ' a ']]), /unique/);
  assert.throws(() => tableFromRecords([['A', ''], ['x', 'y']]), /header/);
  assert.throws(() => tableFromRecords([['A']], 2), /header row/);
  assert.deepEqual(tableFromRecords([['A', '', ''], ['x']]).headers, ['A']);
});

test('defaults, paths, ordering, zero and false preservation', () => {
  const result = buildExport(table, [...mappings(),
    mapping('Iteration Path', 'Sprint', '', 'Project\\'),
    mapping('Story Points', 'Points', '9'),
    mapping('Custom Flag', 'Flag', 'fallback', '[', ']'),
    mapping('Description', 'Description')
  ]);
  assert.deepEqual(result.errors, []);
  assert.deepEqual(result.rows[0].values, ['User Story', 'One', 'Project\\Sprint 1', '0', '[FALSE]', 'Comma, "quote"\nand café']);
  assert.equal(result.rows[1].values[3], '9');
  assert.equal(result.rows[1].values[5], '');
});

test('iteration root appends only the final source segment and uses root for blank sprints', () => {
  const source = tableFromRecords(parseCsv(
    'Title,Sprint\nOne,Sprint 1\nTwo,Old Project\\Release\\Sprint 2\nThree,\nFour,  \nFive,0\nSix,Path\\'
  ).records);
  const iteration = { ...mapping('Iteration Path', 'Sprint', 'Project\\Release\\'), format: 'iteration-root' };
  const result = buildExport(source, [...mappings(), iteration]);
  assert.deepEqual(result.errors, []);
  assert.deepEqual(result.rows.map(row => row.values[2]), [
    'Project\\Release\\Sprint 1', 'Project\\Release\\Sprint 2',
    'Project\\Release', 'Project\\Release', 'Project\\Release\\0', 'Project\\Release'
  ]);
  assert.equal(buildExport(source, [...mappings(), { ...iteration, source: '' }]).rows[0].values[2], 'Project\\Release');
  for (const invalid of [
    { ...iteration, default: '' },
    { ...iteration, default: '\\' },
    { ...iteration, prefix: 'Project\\' },
    { ...iteration, suffix: 'x' },
    { ...iteration, field: 'Area Path' }
  ]) {
    assert.notDeepEqual(buildExport(source, [...mappings(), invalid]).errors, []);
  }
});

test('SF import profile maps Sprint to its configured iteration root', () => {
  const profile = readProfile(JSON.parse(readFileSync(new URL('../sf-first-import-profile.json', import.meta.url), 'utf8')));
  const iteration = profile.mappings.find(item => item.field === 'Iteration Path');
  const source = tableFromRecords([
    ['Title', 'Sprint'], ['One', 'Another\\Sprint 3'], ['Two', '']
  ]);
  const result = buildExport(source, [...mappings(), iteration]);
  assert.deepEqual(result.errors, []);
  assert.deepEqual(result.rows.map(row => row.values[2]), [
    'eB2B Global Platform Mars Snacking\\01_Mars Snacking MVP Release\\Sprint 3',
    'eB2B Global Platform Mars Snacking\\01_Mars Snacking MVP Release'
  ]);
});

test('creation validation blocks forbidden, missing, duplicate and invalid fields', () => {
  for (const field of ['ID', 'State', 'Assigned To', 'Created By', 'Changed Date']) {
    assert.match(buildExport(table, [...mappings(), mapping(field, '', 'value')]).errors.join(), /not allowed/);
  }
  assert.match(buildExport(table, [...mappings(), mapping('title', 'Title')]).errors.join(), /duplicate/);
  assert.match(buildExport(table, [mappings()[0]]).errors.join(), /Title/);
  assert.match(buildExport(table, [...mappings(), mapping('Custom', 'Missing', 'fallback')]).errors.join(), /missing/);
  assert.match(buildExport(table, [...mappings(), mapping('Custom', '')]).errors.join(), /constant/);
  const empty = tableFromRecords([['Title'], ['   ']]);
  assert.match(buildExport(empty, mappings()).errors.join(), /Source row 2.*blank/);
  assert.match(buildExport(tableFromRecords([['Title'], ['x'.repeat(256)]]), mappings()).errors.join(), /255/);
});

test('filter is exact, case-insensitive, independent of mapping and matches blanks', () => {
  assert.equal(buildExport(table, mappings(), { column: 'Sprint', value: 'sprint 1' }).rows.length, 1);
  assert.equal(buildExport(table, mappings(), { column: 'Description', value: '' }).rows[0].values[1], 'Two');
  assert.match(buildExport(table, mappings(), { column: 'Sprint', value: 'Sprint 1 ' }).errors.join(), /No source rows/);
  assert.match(buildExport(table, mappings(), { column: 'Missing', value: '' }).errors.join(), /Filter column/);
});

test('1000-item boundary, last batch, no overlap, invalid batch and blocked export', () => {
  const large = tableFromRecords([['Title'], ...Array.from({ length: 2001 }, (_, index) => [`Item ${index + 1}`])]);
  const result = buildExport(large, mappings());
  assert.equal(result.batches, 3);
  assert.equal(batchRows(result, 1).length, 1000);
  assert.equal(batchRows(result, 2)[0].values[1], 'Item 1001');
  assert.equal(batchRows(result, 3).length, 1);
  assert.equal(batchRows(result, 3)[0].values[1], 'Item 2001');
  for (const batch of [0, 4, 1.5]) assert.throws(() => batchRows(result, batch), /Invalid/);
  assert.throws(() => batchRows(buildExport(table, []), 1), /Resolve/);
});

test('ADO CSV roundtrip preserves exact output and omits source-row markers', () => {
  const result = buildExport(table, [...mappings(), mapping('Description', 'Description')]);
  const csv = toCsv(result.headers, batchRows(result, 1));
  assert.ok(csv.startsWith('\uFEFF"Work Item Type","Title","Description"\r\n'));
  assert.deepEqual(parseCsv(csv).records, [result.headers, ...result.rows.map(row => row.values)]);
  assert.ok(!csv.includes('Source row'));
});

test('mapping profiles validate shape and retain only settings', () => {
  const profile = { version: 1, mappings: mappings(), filter: { column: '', value: '' } };
  assert.deepEqual(readProfile(profile), { mappings: mappings(), filter: { column: '', value: '' } });
  assert.throws(() => readProfile({ ...profile, version: 4 }), /version 1, 2 or 3/);
  assert.throws(() => readProfile({ ...profile, mappings: [{ enabled: true }] }), /Invalid mapping/);
  assert.throws(() => readProfile({ ...profile, filter: null }), /filter/);
});

test('Excel cached formulas preserve empty, zero and false; missing caches/errors are explicit', () => {
  const errors = [];
  assert.equal(excelCellValue({ formula: '1-1', result: 0 }, 2, 1, errors), '0');
  assert.equal(excelCellValue({ sharedFormula: 'A2', result: false }, 3, 1, errors), 'FALSE');
  assert.equal(excelCellValue({ formula: '""', result: '' }, 4, 1, errors), '');
  assert.deepEqual(errors, []);
  excelCellValue({ formula: '1+1' }, 5, 1, errors);
  excelCellValue({ formula: '1/0', result: { error: '#DIV/0!' } }, 6, 1, errors);
  assert.match(errors[0].message, /row 5.*no saved result/);
  assert.match(errors[1].message, /row 6.*#DIV\/0!/);
  assert.equal(excelCellValue({ richText: [{ text: 'Hello ' }, { text: 'world' }] }, 2, 1, []), 'Hello world');
  assert.equal(excelCellValue({ text: 'Link', hyperlink: 'https://example.com' }, 2, 1, []), 'Link');
  assert.equal(excelCellValue(new Date('2026-01-01T00:00:00Z'), 2, 1, []), '2026-01-01T00:00:00.000Z');
  const book = new ExcelJS.Workbook();
  const sheet = book.addWorksheet('Formula values');
  sheet.addRow(['Title', 'Description', 'Points', 'Flag']);
  sheet.addRow(['Story', { formula: '""', result: '' }, { formula: '1-1', result: 0 }, { formula: '1=2', result: false }]);
  const parsed = worksheetRecords(sheet);
  assert.deepEqual(parsed.errors, []);
  assert.deepEqual(parsed.records[1], ['Story', '', '0', 'FALSE']);
});

test('original XLSX loads cached user-story descriptions and multiline criteria', async () => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(fileURLToPath(new URL('../assets/examples/Example.xlsx', import.meta.url)));
  const sheet = workbook.worksheets[0];
  const { records, errors } = worksheetRecords(sheet);
  assert.deepEqual(errors, []);
  const sample = tableFromRecords(records);
  const result = buildExport(sample, suggestedMappings(sample.headers));
  assert.deepEqual(result.errors, []);
  assert.equal(result.rows.length, 2);
  assert.deepEqual(result.headers, ['Work Item Type', 'Title', 'Description', 'Acceptance Criteria']);
  assert.match(result.rows[0].values[2], /^<p><strong>As a<\/strong> System Admin<\/p>/);
  assert.equal(result.rows[0].values[3], '');
  assert.match(result.rows[1].values[3], /<ol start="1"><li>CI\/CD/);
  assert.match(result.rows[1].values[3], /<li>Main Salesforce/);
  assert.deepEqual(parseCsv(toCsv(result.headers, result.rows)).records[2], result.rows[1].values);
});

test('rich text modes generate safe HTML, story sections, lists and paragraph breaks', () => {
  assert.equal(formatText('unchanged\r\ntext'), 'unchanged\r\ntext');
  assert.equal(formatText('', 'markdown'), '');
  assert.equal(formatText('First\r\nsecond\r\n\r\nThird & last', 'paragraphs'),
    '<p>First<br>second</p><p>Third &amp; last</p>');
  assert.equal(formatText('As a developer I want to deploy\nsafely so that releases are reliable.', 'story'),
    '<p><strong>As a</strong> developer</p><p><strong>I want to</strong> deploy<br>safely</p><p><strong>So that</strong> releases are reliable.</p>');
  assert.equal(formatText('Not a story\nstill readable', 'story'), '<p>Not a story<br>still readable</p>');
  const html = formatText('# Criteria\n\n1. **Pipeline** works\n2. *Tests* pass\n\n- `code`\n- <img src=x onerror=alert(1)>', 'markdown');
  assert.match(html, /<h2>Criteria<\/h2>/);
  assert.match(html, /<ol start="1"><li><strong>Pipeline<\/strong> works<\/li><li><em>Tests<\/em> pass<\/li><\/ol>/);
  assert.match(html, /<ul><li><code>code<\/code><\/li>/);
  assert.ok(!html.includes('<img'));
  assert.match(html, /&lt;img/);
  assert.equal(formatText('`**not bold**`', 'markdown'), '<p><code>**not bold**</code></p>');
  assert.throws(() => formatText('text', 'invalid'), /Unknown text format/);
  const config = mappings();
  config[1].format = 'markdown';
  assert.match(buildExport(table, config).errors.join(), /Title.*plain text/);
});

const epicTable = (rows = [
  ['One', '165', 'Platform', 'First\nline'],
  ['Two', '200', 'Commerce', 'Second'],
  ['Three', '165', 'Platform', 'Third']
]) => tableFromRecords([['Title', 'Epic ID', 'Epic', 'Description'], ...rows]);
const parentConfig = () => ({ ...emptyParent(), enabled: true, idSource: 'Epic ID', titleSource: 'Epic' });

test('existing Epic hierarchy uses reference IDs and indented titles, not child IDs', () => {
  const result = buildExport(epicTable(), [...mappings(), { ...mapping('Description', 'Description'), format: 'paragraphs' }],
    { column: '', value: '' }, parentConfig());
  assert.deepEqual(result.errors, []);
  assert.deepEqual(result.headers, ['ID', 'Work Item Type', 'Title 1', 'Title 2', 'Description']);
  const batch = batchRows(result, 1);
  assert.deepEqual(batch.map(row => row.values.slice(0, 4)), [
    ['165', 'Epic', 'Platform', ''], ['', 'User Story', '', 'One'],
    ['', 'User Story', '', 'Three'], ['200', 'Epic', 'Commerce', ''],
    ['', 'User Story', '', 'Two']
  ]);
  assert.equal(batch[0].values[4], '');
  assert.equal(batch[1].values[4], '<p>First<br>line</p>');
  assert.equal(result.rows.length, 3);
  assert.deepEqual(parseCsv(toCsv(result.headers, batch)).records, [result.headers, ...batch.map(row => row.values)]);
});

test('parent source mapping, constants, defaults, validation and filtering', () => {
  const constant = { ...emptyParent(), enabled: true, idDefault: '165', titleDefault: 'Platform' };
  assert.equal(batchRows(buildExport(table, mappings(), undefined, constant), 1)[0].values[0], '165');
  const parent = parentConfig();
  parent.idDefault = '165'; parent.titleDefault = 'Platform';
  assert.deepEqual(buildExport(epicTable([['One', '', '', '']]), mappings(), undefined, parent).errors, []);
  for (const id of ['0', '-1', 'abc', '165.0', '1e2', '', '9007199254740992']) {
    assert.match(buildExport(epicTable([['One', id, 'Platform', '']]), mappings(), undefined, parentConfig()).errors.join(), /parent Epic ID/);
  }
  assert.match(buildExport(epicTable([['One', '165', '', '']]), mappings(), undefined, parentConfig()).errors.join(), /parent Epic title/);
  assert.match(buildExport(epicTable([['One', '165', 'A', ''], ['Two', '165', 'B', '']]), mappings(), undefined, parentConfig()).errors.join(), /conflicting titles/);
  assert.match(buildExport(table, mappings(), undefined, parentConfig()).errors.join(), /source column is missing/);
  const taskMappings = mappings(); taskMappings[0].default = 'Task';
  assert.match(buildExport(epicTable(), taskMappings, undefined, parentConfig()).errors.join(), /User Story children only/);
  const filtered = buildExport(epicTable(), mappings(), { column: 'Epic ID', value: '200' }, parentConfig());
  assert.equal(filtered.rows.length, 1);
  assert.equal(batchRows(filtered, 1).length, 2);
  assert.equal(batchRows(filtered, 1)[0].values[0], '200');
  assert.match(buildExport(table, [...mappings(), mapping('Parent', '', '165')]).errors.join(), /not allowed/);
});

test('hierarchy batches include repeat parent anchors within the exact 1000-row limit', () => {
  for (const count of [998, 999, 1000, 2001]) {
    const source = epicTable(Array.from({ length: count }, (_, index) => [`Item ${index + 1}`, '165', 'Platform', '']));
    const result = buildExport(source, mappings(), undefined, parentConfig());
    assert.equal(result.batches, Math.ceil(count / 999));
    const exportedChildren = [];
    for (let batch = 1; batch <= result.batches; batch++) {
      const rows = batchRows(result, batch);
      assert.ok(rows.length <= 1000);
      assert.deepEqual(rows[0].values.slice(0, 4), ['165', 'Epic', 'Platform', '']);
      exportedChildren.push(...rows.filter(row => !row.isParent).map(row => row.sourceRow));
    }
    assert.equal(exportedChildren.length, count);
    assert.equal(new Set(exportedChildren).size, count);
  }
  const distinct = epicTable(Array.from({ length: 501 }, (_, index) => [`Item ${index}`, String(index + 1), `Epic ${index}`, '']));
  const result = buildExport(distinct, mappings(), undefined, parentConfig());
  assert.equal(batchRows(result, 1).length, 1000);
  assert.equal(batchRows(result, 2).length, 2);
});

test('version 2 profiles retain HTML formatting and parent config; version 1 stays plain', () => {
  const config = [...mappings(), { ...mapping('Description', 'Description'), format: 'story' }].map(row => ({ ...row, format: row.format || 'plain' }));
  const profile = { version: 2, mappings: config, filter: { column: '', value: '' }, parent: parentConfig() };
  assert.deepEqual(readProfile(profile), { mappings: config, filter: profile.filter, parent: parentConfig() });
  assert.throws(() => readProfile({ ...profile, parent: { enabled: true } }), /parent Epic/);
  assert.throws(() => readProfile({ ...profile, mappings: [{ ...config[0], format: 'html' }] }), /text format/);
  const legacy = readProfile({ version: 1, mappings: mappings(), filter: profile.filter });
  assert.equal(legacy.parent, undefined);
  assert.equal(legacy.mappings[0].format, undefined);
  assert.equal(new Set(FIELD_SUGGESTIONS).size, FIELD_SUGGESTIONS.length);
  assert.ok(FIELD_SUGGESTIONS.includes('Acceptance Criteria'));
  assert.ok(!FIELD_SUGGESTIONS.includes('Parent'));
});

test('preloaded Epic names resolve screenshot IDs without source ID columns', () => {
  assert.equal(DEFAULT_EPICS.length, 21);
  const source = tableFromRecords([['Title', 'Epic'],
    ...DEFAULT_EPICS.map(([, title], index) => [`Story ${index}`, `  ${title.toUpperCase()}  `])]);
  const parent = { ...lookupParent(), enabled: true, epicSource: 'Epic' };
  const result = buildExport(source, mappings(), undefined, parent);
  assert.deepEqual(result.errors, []);
  const rows = batchRows(result, 1);
  assert.deepEqual(rows.filter(row => row.isParent).map(row => row.values.slice(0, 3)),
    DEFAULT_EPICS.map(([id, title]) => [id, 'Epic', title]));
  assert.equal(rows.filter(row => !row.isParent).length, 21);
  assert.equal(lookupParent().epics[0].id, '1572169');
  parent.epics[0].sourceName = 'Changed';
  assert.equal(lookupParent().epics[0].sourceName, 'Order History & Reordering');
});

test('manual Epic aliases resolve correct titles; missing or ambiguous names block export', () => {
  const parent = { ...lookupParent(), enabled: true, epicSource: 'Epic',
    epics: [{ sourceName: 'Products', id: '123', title: 'Product Discovery' },
      { sourceName: 'Checkout', id: '456', title: 'Checkout & Related Validations' }] };
  const source = tableFromRecords([['Title', 'Epic'], ['One', 'Products'], ['Two', 'Checkout']]);
  const result = buildExport(source, mappings(), undefined, parent);
  assert.deepEqual(result.errors, []);
  assert.deepEqual(batchRows(result, 1).filter(row => row.isParent).map(row => row.values.slice(0, 3)),
    [['123', 'Epic', 'Product Discovery'], ['456', 'Epic', 'Checkout & Related Validations']]);
  const bad = tableFromRecords([['Title', 'Epic'], ['Unknown', 'Missing'], ['Blank', '']]);
  assert.match(buildExport(bad, mappings(), undefined, parent).errors.join(), /row 2.*"Missing".*no mapping/);
  assert.match(buildExport(bad, mappings(), undefined, parent).errors.join(), /row 3.*blank/);
  assert.equal(buildExport(bad, mappings(), { column: 'Title', value: 'not selected' }, parent).errors.length, 1);
  assert.match(buildExport(source, mappings(), undefined, { ...parent, epicSource: 'Missing' }).errors.join(), /choose the source column/);
  const duplicate = { ...parent, epics: [...parent.epics, { sourceName: ' products ', id: '999', title: 'Other' }] };
  assert.match(buildExport(source, mappings(), undefined, duplicate).errors.join(), /duplicate source Epic name/);
  const conflicting = { ...parent, epics: [...parent.epics, { sourceName: 'Alias', id: '123', title: 'Wrong title' }] };
  assert.match(buildExport(source, mappings(), undefined, conflicting).errors.join(), /conflicting ADO titles/);
  for (const epic of [
    { sourceName: '', id: '123', title: 'Title' },
    { sourceName: 'Products', id: 'abc', title: 'Title' },
    { sourceName: 'Products', id: '123', title: '' }
  ]) {
    const invalid = buildExport(source, mappings(), undefined, { ...parent, epics: [epic] });
    assert.ok(invalid.errors.length);
    assert.throws(() => batchRows(invalid, 1), /Resolve/);
  }
});

test('lookup parents preserve hierarchy batching and filtering', () => {
  const parent = { ...lookupParent(), enabled: true, epicSource: 'Epic' };
  const source = tableFromRecords([['Title', 'Epic'],
    ...Array.from({ length: 1000 }, (_, index) => [`Story ${index}`, 'Product Discovery']),
    ['Last', 'Checkout & Related Validations']]);
  const result = buildExport(source, mappings(), undefined, parent);
  assert.deepEqual(result.errors, []);
  assert.equal(result.batches, 2);
  assert.equal(batchRows(result, 1).length, 1000);
  assert.equal(batchRows(result, 2).length, 4);
  assert.equal(batchRows(result, 2)[0].values[0], '1572173');
  assert.equal(batchRows(result, 2)[2].values[0], '1620343');
  const filtered = buildExport(source, mappings(), { column: 'Epic', value: 'Checkout & Related Validations' }, parent);
  assert.equal(filtered.rows.length, 1);
  assert.equal(batchRows(filtered, 1)[0].values[0], '1620343');
});

test('version 3 profiles save editable Epic lookup table and reject malformed settings', () => {
  const parent = { ...lookupParent(), enabled: true, epicSource: 'Epic' };
  const profile = { version: 3, mappings: mappings().map(row => ({ ...row, format: 'plain' })),
    filter: { column: '', value: '' }, parent };
  assert.deepEqual(readProfile(JSON.parse(JSON.stringify(profile))).parent, parent);
  for (const invalid of [{ ...parent, mode: 'guess' }, { ...parent, epics: null },
    { ...parent, epics: [{ sourceName: 'Products', id: 123, title: 'Products' }] },
    { ...parent, epicSource: null }]) {
    assert.throws(() => readProfile({ ...profile, parent: invalid }), /lookup table/);
  }
});
