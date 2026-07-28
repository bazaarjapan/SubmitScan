const test = require('node:test');
const assert = require('node:assert/strict');

const {
  CONFIG,
  normalizeBarcode,
  buildBarcodeRows,
  buildBarcodeFormula,
  resolvePrintRowNumbers,
  buildPrintRows,
  buildUniqueSheetName,
  findDuplicateIds,
  evaluateSubmissionScans,
  findResultColumn,
  findResultColumns,
  collectExistingSubmissionFlags,
  mergeSubmissionFlags,
  buildCheckSummary,
} = require('../コード.js');

test('normalizeBarcode normalizes values without discarding display text', () => {
  assert.equal(normalizeBarcode(' 001 '), '001');
  assert.equal(normalizeBarcode(123), '123');
  assert.equal(normalizeBarcode(null), '');
});

test('buildBarcodeRows creates encoded formulas and preserves empty rows', () => {
  const rows = buildBarcodeRows([
    ['01', 'A', '002'],
    ['', '', ''],
  ], 2);

  assert.equal(rows[0].id, '01A002');
  assert.match(rows[0].formula, /ENCODEURL\(D2\)/);
  assert.match(rows[0].formula, /xres=3&height=90&width=306/);
  assert.match(rows[0].formula, /,4,90,306\)$/);
  assert.equal(rows[1].id, '');
  assert.equal(rows[1].formula, '');
});

test('barcode layout uses high-resolution image and matching cell sizes', () => {
  assert.equal(CONFIG.barcodeXResolution, 3);
  assert.equal(CONFIG.barcodeImageWidth, 306);
  assert.equal(CONFIG.barcodeImageHeight, 90);
  assert.equal(CONFIG.barcodeColumnWidth, 320);
  assert.equal(CONFIG.barcodeRowHeight, 100);
});

test('buildBarcodeFormula supports a print-sheet cell reference', () => {
  const formula = buildBarcodeFormula('B2');
  assert.match(formula, /ENCODEURL\(B2\)/);
  assert.match(formula, /xres=3&height=90&width=306/);
  assert.match(formula, /,4,90,306\)$/);
});

test('resolvePrintRowNumbers uses selected roster rows within the data range', () => {
  assert.deepEqual(
    resolvePrintRowNumbers(CONFIG.rosterSheetName, 3, 3, 10),
    [3, 4, 5]
  );
  assert.deepEqual(
    resolvePrintRowNumbers(CONFIG.rosterSheetName, 1, 3, 3),
    [2, 3]
  );
});

test('resolvePrintRowNumbers falls back to all roster rows without a valid selection', () => {
  assert.deepEqual(resolvePrintRowNumbers('読み込み', 1, 1, 4), [2, 3, 4]);
  assert.deepEqual(resolvePrintRowNumbers(CONFIG.rosterSheetName, 1, 1, 4), [2, 3, 4]);
});

test('buildPrintRows keeps selected display data and removes blank or duplicate IDs', () => {
  const rows = buildPrintRows([
    ['1年', 'A組', '安藤', '001'],
    ['1年', 'B組', '佐藤', ''],
    ['2年', 'A組', '鈴木', '002'],
    ['重複', '', '', '001'],
  ], [2, 3, 4, 5]);

  assert.deepEqual(rows, [
    {info: '1年 / A組 / 安藤', id: '001'},
    {info: '2年 / A組 / 鈴木', id: '002'},
  ]);
});

test('buildUniqueSheetName never reuses an unrelated same-named sheet', () => {
  assert.equal(buildUniqueSheetName([], 'バーコード印刷'), 'バーコード印刷');
  assert.equal(
    buildUniqueSheetName(['バーコード印刷'], 'バーコード印刷'),
    'バーコード印刷 (2)'
  );
  assert.equal(
    buildUniqueSheetName(
      ['バーコード印刷', 'バーコード印刷 (2)', 'バーコード印刷 (3)'],
      'バーコード印刷'
    ),
    'バーコード印刷 (4)'
  );
});

test('findDuplicateIds ignores blanks and returns each duplicate once', () => {
  assert.deepEqual(findDuplicateIds(['A', '', ' A ', 'B', 'B']), ['A', 'B']);
});

test('evaluateSubmissionScans marks matches without writing unknown IDs', () => {
  const result = evaluateSubmissionScans(
    ['001', 123, 'missing', '123', '', 'missing'],
    ['001', '123', 'ABC', '']
  );

  assert.deepEqual(result.flags, [1, 1, '', '']);
  assert.equal(result.scanCount, 5);
  assert.equal(result.matchedCount, 2);
  assert.equal(result.duplicateScanCount, 2);
  assert.deepEqual(result.unknownIds, ['missing']);
  assert.deepEqual(result.duplicateRegisteredIds, []);
});

test('evaluateSubmissionScans reports duplicate registered IDs', () => {
  const result = evaluateSubmissionScans(['A'], ['A', ' A ', 'B']);
  assert.deepEqual(result.duplicateRegisteredIds, ['A']);
});

test('evaluateSubmissionScans treats blank input as no scans', () => {
  const result = evaluateSubmissionScans(['', '  ', null], ['A', 'B']);
  assert.equal(result.scanCount, 0);
  assert.deepEqual(result.flags, ['', '']);
  assert.deepEqual(result.unknownIds, []);
});

test('findResultColumn reuses today or appends a new column', () => {
  assert.equal(findResultColumn(['氏名', '2026/07/29'], '2026/07/29'), 2);
  assert.equal(findResultColumn(['氏名', '2026/07/28'], '2026/07/29'), 3);
});

test('findResultColumns returns every existing column for today', () => {
  assert.deepEqual(
    findResultColumns(['氏名', '2026/07/29', '2026/07/28', ' 2026/07/29 '], '2026/07/29'),
    [2, 4]
  );
});

test('collectExistingSubmissionFlags merges duplicate date columns', () => {
  const rows = [
    [1, '', ''],
    ['', '', 1],
    ['', '', ''],
  ];
  assert.deepEqual(collectExistingSubmissionFlags(rows, [0, 2]), [1, 1, '']);
});

test('mergeSubmissionFlags preserves earlier scans and adds new matches', () => {
  assert.deepEqual(
    mergeSubmissionFlags([1, '', '1', ''], ['', 1, '', '']),
    [1, 1, 1, '']
  );
});

test('buildCheckSummary limits the displayed unknown ID list', () => {
  const summary = buildCheckSummary({
    scanCount: 15,
    matchedCount: 3,
    duplicateScanCount: 1,
    unknownIds: Array.from({length: 11}, (_, index) => `X${index}`),
  });

  assert.match(summary, /未登録コード: 11件/);
  assert.match(summary, /ほか1件/);
  assert.doesNotMatch(summary, /X10/);
});
