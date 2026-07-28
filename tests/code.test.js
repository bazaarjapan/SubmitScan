const test = require('node:test');
const assert = require('node:assert/strict');

const {
  normalizeBarcode,
  buildBarcodeRows,
  findDuplicateIds,
  evaluateSubmissionScans,
  findResultColumn,
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
  assert.equal(rows[1].id, '');
  assert.equal(rows[1].formula, '');
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
