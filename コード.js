// **************************************************
// 内容: 提出物管理
// 作成日: 2021/06/21
// 作成者: noboru ando
// **************************************************

const CONFIG = Object.freeze({
  inputSheetName: '読み込み',
  rosterSheetName: 'バーコード作成',
  rosterStartRow: 2,
  barcodeIdColumn: 4,
  barcodeFormulaColumn: 6,
  barcodeXResolution: 3,
  barcodeImageWidth: 306,
  barcodeImageHeight: 90,
  barcodeColumnWidth: 320,
  barcodeRowHeight: 100,
  printSheetName: 'バーコード印刷',
  printSheetMarkerKey: 'SUBMITSCAN_MANAGED_BARCODE_PRINT_SHEET',
  printSheetMarkerValue: 'true',
  printInfoColumnWidth: 240,
  printIdColumnWidth: 160,
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('バーコード')
    .addItem('チェック', 'barcodelabelcheck')
    .addItem('入力初期化', 'syokika')
    .addItem('バーコード作成', 'barcodelabel')
    .addItem('印刷用シートを作成', 'createBarcodePrintSheet')
    .addToUi();
}

function syokika() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.inputSheetName);

  if (!sheet) {
    ui.alert(`「${CONFIG.inputSheetName}」シートが見つかりません。`);
    return;
  }

  const response = ui.alert(
    '入力初期化',
    '読み込んだバーコードを消去します。よろしいですか？',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) {
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    sheet.getRange(1, 1, lastRow, 1).clearContent();
  }
  sheet.activate();
  sheet.getRange('A1').activate();
  ui.alert('初期化しました。半角英数入力モードで読み込んでください。');
}

function barcodelabel() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.rosterSheetName);

  if (!sheet) {
    ui.alert(`「${CONFIG.rosterSheetName}」シートが見つかりません。`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.rosterStartRow) {
    ui.alert('バーコードを作成する名簿データがありません。');
    return;
  }

  const rowCount = lastRow - CONFIG.rosterStartRow + 1;
  const sourceRows = sheet
    .getRange(CONFIG.rosterStartRow, 1, rowCount, 3)
    .getDisplayValues();
  const barcodeRows = buildBarcodeRows(sourceRows, CONFIG.rosterStartRow);
  const duplicateIds = findDuplicateIds(barcodeRows.map(row => row.id));

  if (duplicateIds.length > 0) {
    ui.alert(`重複するバーコードIDがあります。\n${duplicateIds.join('\n')}`);
    return;
  }

  sheet
    .getRange(CONFIG.rosterStartRow, CONFIG.barcodeIdColumn, rowCount, 1)
    .setValues(barcodeRows.map(row => [row.id]));
  sheet
    .getRange(CONFIG.rosterStartRow, CONFIG.barcodeFormulaColumn, rowCount, 1)
    .setValues(barcodeRows.map(row => [row.formula]));
  sheet.setColumnWidth(CONFIG.barcodeFormulaColumn, CONFIG.barcodeColumnWidth);
  sheet.setRowHeightsForced(
    CONFIG.rosterStartRow,
    rowCount,
    CONFIG.barcodeRowHeight
  );

  const generatedCount = barcodeRows.filter(row => row.id !== '').length;
  ui.alert(`${generatedCount}件のバーコードを作成しました。`);
}

function createBarcodePrintSheet() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActive();
  const rosterSheet = spreadsheet.getSheetByName(CONFIG.rosterSheetName);

  if (!rosterSheet) {
    ui.alert(`「${CONFIG.rosterSheetName}」シートが見つかりません。`);
    return;
  }

  const lastRow = rosterSheet.getLastRow();
  if (lastRow < CONFIG.rosterStartRow) {
    ui.alert('印刷する名簿データがありません。');
    return;
  }

  const activeSheet = spreadsheet.getActiveSheet();
  const activeRange = activeSheet ? activeSheet.getActiveRange() : null;
  const selectedRows = resolvePrintRowNumbers(
    activeSheet ? activeSheet.getName() : '',
    activeRange ? activeRange.getRow() : null,
    activeRange ? activeRange.getNumRows() : null,
    lastRow
  );
  const rowCount = lastRow - CONFIG.rosterStartRow + 1;
  const sourceRows = rosterSheet
    .getRange(CONFIG.rosterStartRow, 1, rowCount, CONFIG.barcodeIdColumn)
    .getDisplayValues();
  const printRows = buildPrintRows(sourceRows, selectedRows);

  if (printRows.length === 0) {
    ui.alert('印刷できるバーコードIDがありません。先にバーコードを作成してください。');
    return;
  }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    ui.alert('別の印刷用シート作成処理が実行中です。しばらく待ってから再実行してください。');
    return;
  }

  try {
    const printSheet = getOrCreateManagedPrintSheet(spreadsheet);
    ensureSheetSize(printSheet, printRows.length + 1, 3);
    printSheet.clear();
    printSheet.setHiddenGridlines(true);
    printSheet.setFrozenRows(1);
    printSheet.getRange(1, 1, 1, 3)
      .setValues([['名簿情報', 'バーコードID', 'バーコード']])
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    printSheet.getRange(2, 1, printRows.length, 2)
      .setValues(printRows.map(row => [row.info, row.id]))
      .setVerticalAlignment('middle');
    printSheet.getRange(2, 3, printRows.length, 1)
      .setFormulas(printRows.map((row, index) => [
        buildBarcodeFormula(`B${index + 2}`),
      ]))
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    printSheet.setColumnWidth(1, CONFIG.printInfoColumnWidth);
    printSheet.setColumnWidth(2, CONFIG.printIdColumnWidth);
    printSheet.setColumnWidth(3, CONFIG.barcodeColumnWidth);
    printSheet.setRowHeightsForced(2, printRows.length, CONFIG.barcodeRowHeight);
    printSheet.activate();
    printSheet.getRange('A1').activate();
  } finally {
    lock.releaseLock();
  }

  ui.alert(
    `${printRows.length}件の印刷用シートを作成しました。\n` +
    '「ファイル」→「印刷」から印刷してください。'
  );
}

function barcodelabelcheck() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActive();
  const inputSheet = spreadsheet.getSheetByName(CONFIG.inputSheetName);
  const rosterSheet = spreadsheet.getSheetByName(CONFIG.rosterSheetName);

  if (!inputSheet || !rosterSheet) {
    const missingNames = [
      !inputSheet ? CONFIG.inputSheetName : '',
      !rosterSheet ? CONFIG.rosterSheetName : '',
    ].filter(Boolean);
    ui.alert(`次のシートが見つかりません。\n${missingNames.join('\n')}`);
    return;
  }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    ui.alert('別の処理が実行中です。しばらく待ってから再実行してください。');
    return;
  }

  try {
    const inputLastRow = inputSheet.getLastRow();
    if (inputLastRow < 1) {
      ui.alert('読み込まれたバーコードがありません。');
      return;
    }

    const rosterLastRow = rosterSheet.getLastRow();
    if (rosterLastRow < CONFIG.rosterStartRow) {
      ui.alert('照合する名簿データがありません。');
      return;
    }

    const scannedIds = inputSheet
      .getRange(1, 1, inputLastRow, 1)
      .getDisplayValues()
      .map(row => row[0]);
    const rosterRowCount = rosterLastRow - CONFIG.rosterStartRow + 1;
    const registeredIds = rosterSheet
      .getRange(CONFIG.rosterStartRow, CONFIG.barcodeIdColumn, rosterRowCount, 1)
      .getDisplayValues()
      .map(row => row[0]);
    const result = evaluateSubmissionScans(scannedIds, registeredIds);

    if (result.scanCount === 0) {
      ui.alert('読み込まれたバーコードがありません。');
      return;
    }
    if (result.duplicateRegisteredIds.length > 0) {
      ui.alert(
        `名簿に重複するバーコードIDがあります。処理を中止しました。\n${result.duplicateRegisteredIds.join('\n')}`
      );
      return;
    }

    const today = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy/MM/dd'
    );
    const lastColumn = Math.max(rosterSheet.getLastColumn(), 1);
    const headers = rosterSheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
    const existingResultColumns = findResultColumns(headers, today);
    const resultColumn = existingResultColumns.length > 0
      ? existingResultColumns[0]
      : headers.length + 1;
    let outputFlags = result.flags;

    if (existingResultColumns.length > 0) {
      const firstResultColumn = existingResultColumns[0];
      const lastResultColumn = existingResultColumns[existingResultColumns.length - 1];
      const existingRows = rosterSheet
        .getRange(
          CONFIG.rosterStartRow,
          firstResultColumn,
          rosterRowCount,
          lastResultColumn - firstResultColumn + 1
        )
        .getValues();
      const columnOffsets = existingResultColumns.map(column => column - firstResultColumn);
      const existingFlags = collectExistingSubmissionFlags(existingRows, columnOffsets);
      outputFlags = mergeSubmissionFlags(existingFlags, result.flags);
    }

    rosterSheet.getRange(1, resultColumn).setValue(today);
    rosterSheet
      .getRange(CONFIG.rosterStartRow, resultColumn, rosterRowCount, 1)
      .setValues(outputFlags.map(value => [value]));
    rosterSheet.activate();

    ui.alert(buildCheckSummary(result));
  } finally {
    lock.releaseLock();
  }
}

function normalizeBarcode(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function buildBarcodeRows(sourceRows, startRow) {
  return sourceRows.map((parts, index) => {
    const id = parts.map(normalizeBarcode).join('');
    const sheetRow = startRow + index;
    const formula = id === '' ? '' : buildBarcodeFormula(`D${sheetRow}`);
    return {id, formula};
  });
}

function buildBarcodeFormula(cellReference) {
  return `=IMAGE("https://www.webarcode.com/barcode/image.php?code="&ENCODEURL(${cellReference})&"&type=C128B&xres=${CONFIG.barcodeXResolution}&height=${CONFIG.barcodeImageHeight}&width=${CONFIG.barcodeImageWidth}&font=3&output=png&style=196",4,${CONFIG.barcodeImageHeight},${CONFIG.barcodeImageWidth})`;
}

function resolvePrintRowNumbers(activeSheetName, selectionStartRow, selectionRowCount, lastRow) {
  const allRows = [];
  for (let row = CONFIG.rosterStartRow; row <= lastRow; row += 1) {
    allRows.push(row);
  }
  if (
    activeSheetName !== CONFIG.rosterSheetName ||
    !Number.isInteger(selectionStartRow) ||
    !Number.isInteger(selectionRowCount) ||
    selectionRowCount < 1
  ) {
    return allRows;
  }

  const selectedRows = [];
  const selectionEndRow = selectionStartRow + selectionRowCount - 1;
  for (
    let row = Math.max(selectionStartRow, CONFIG.rosterStartRow);
    row <= Math.min(selectionEndRow, lastRow);
    row += 1
  ) {
    selectedRows.push(row);
  }
  return selectedRows.length > 0 ? selectedRows : allRows;
}

function buildPrintRows(sourceRows, selectedRows) {
  const seenIds = new Set();
  const printRows = [];

  selectedRows.forEach(sheetRow => {
    const sourceRow = sourceRows[sheetRow - CONFIG.rosterStartRow];
    if (!sourceRow) {
      return;
    }
    const id = normalizeBarcode(sourceRow[CONFIG.barcodeIdColumn - 1]);
    if (id === '' || seenIds.has(id)) {
      return;
    }
    seenIds.add(id);
    const info = sourceRow
      .slice(0, CONFIG.barcodeIdColumn - 1)
      .map(normalizeBarcode)
      .filter(Boolean)
      .join(' / ');
    printRows.push({info, id});
  });

  return printRows;
}

function ensureSheetSize(sheet, requiredRows, requiredColumns) {
  const missingRows = requiredRows - sheet.getMaxRows();
  if (missingRows > 0) {
    sheet.insertRowsAfter(sheet.getMaxRows(), missingRows);
  }
  const missingColumns = requiredColumns - sheet.getMaxColumns();
  if (missingColumns > 0) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), missingColumns);
  }
}

function getOrCreateManagedPrintSheet(spreadsheet) {
  const managedSheet = spreadsheet.getSheets().find(sheet => (
    sheet.getDeveloperMetadata().some(metadata => (
      metadata.getKey() === CONFIG.printSheetMarkerKey &&
      metadata.getValue() === CONFIG.printSheetMarkerValue
    ))
  ));
  if (managedSheet) {
    return managedSheet;
  }

  const sheetName = buildUniqueSheetName(
    spreadsheet.getSheets().map(sheet => sheet.getName()),
    CONFIG.printSheetName
  );
  const printSheet = spreadsheet.insertSheet(sheetName);
  printSheet.addDeveloperMetadata(
    CONFIG.printSheetMarkerKey,
    CONFIG.printSheetMarkerValue
  );
  return printSheet;
}

function buildUniqueSheetName(existingNames, baseName) {
  const names = new Set(existingNames);
  if (!names.has(baseName)) {
    return baseName;
  }

  let suffix = 2;
  while (names.has(`${baseName} (${suffix})`)) {
    suffix += 1;
  }
  return `${baseName} (${suffix})`;
}

function findDuplicateIds(ids) {
  const seen = new Set();
  const duplicates = new Set();

  ids.map(normalizeBarcode).filter(Boolean).forEach(id => {
    if (seen.has(id)) {
      duplicates.add(id);
    }
    seen.add(id);
  });

  return [...duplicates];
}

function evaluateSubmissionScans(scannedIds, registeredIds) {
  const normalizedRegisteredIds = registeredIds.map(normalizeBarcode);
  const duplicateRegisteredIds = findDuplicateIds(normalizedRegisteredIds);
  const rowIndexById = new Map();
  normalizedRegisteredIds.forEach((id, index) => {
    if (id !== '' && !rowIndexById.has(id)) {
      rowIndexById.set(id, index);
    }
  });

  const flags = registeredIds.map(() => '');
  const seenScans = new Set();
  const unknownIds = new Set();
  let scanCount = 0;
  let matchedCount = 0;
  let duplicateScanCount = 0;

  scannedIds.forEach(rawId => {
    const id = normalizeBarcode(rawId);
    if (id === '') {
      return;
    }
    scanCount += 1;
    if (seenScans.has(id)) {
      duplicateScanCount += 1;
      return;
    }
    seenScans.add(id);

    const rowIndex = rowIndexById.get(id);
    if (rowIndex === undefined) {
      unknownIds.add(id);
      return;
    }
    flags[rowIndex] = 1;
    matchedCount += 1;
  });

  return {
    flags,
    scanCount,
    matchedCount,
    duplicateScanCount,
    unknownIds: [...unknownIds],
    duplicateRegisteredIds,
  };
}

function findResultColumn(headers, today) {
  const existingColumns = findResultColumns(headers, today);
  return existingColumns.length > 0 ? existingColumns[0] : headers.length + 1;
}

function findResultColumns(headers, today) {
  return headers.reduce((columns, header, index) => {
    if (normalizeBarcode(header) === today) {
      columns.push(index + 1);
    }
    return columns;
  }, []);
}

function collectExistingSubmissionFlags(rows, columnOffsets) {
  return rows.map(row => {
    const wasSubmitted = columnOffsets.some(offset => (
      normalizeBarcode(row[offset]) === '1'
    ));
    return wasSubmitted ? 1 : '';
  });
}

function mergeSubmissionFlags(existingFlags, newFlags) {
  return newFlags.map((newValue, index) => {
    const existingValue = existingFlags[index];
    const wasSubmitted = normalizeBarcode(existingValue) === '1';
    const isSubmitted = normalizeBarcode(newValue) === '1';
    return wasSubmitted || isSubmitted ? 1 : '';
  });
}

function buildCheckSummary(result) {
  const lines = [
    'チェックが完了しました。',
    '',
    `読み込み件数: ${result.scanCount}件`,
    `提出確認: ${result.matchedCount}件`,
    `重複スキャン: ${result.duplicateScanCount}件`,
    `未登録コード: ${result.unknownIds.length}件`,
  ];
  if (result.unknownIds.length > 0) {
    lines.push('', '未登録コード:', ...result.unknownIds.slice(0, 10));
    if (result.unknownIds.length > 10) {
      lines.push(`ほか${result.unknownIds.length - 10}件`);
    }
  }
  return lines.join('\n');
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
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
  };
}
