/**
 * 作業リスト／入力シートの読み取り。
 * 「どの列が何か」はヘッダー名で解決し、以降は名前付きレコードだけを扱う。
 */

/**
 * 作業リストと入力シートの列マップ・レコード配列を一度に作る。
 *
 * @return {{
 *   workSheet: GoogleAppsScript.Spreadsheet.Sheet,
 *   inputSheet: GoogleAppsScript.Spreadsheet.Sheet,
 *   workCols: Object<string, number>,
 *   inputCols: Object<string, number>,
 *   workRows: object[]
 * }}
 */
function loadContext_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workSheet = ss.getSheetByName(CONFIG.workList.sheetName);
  const inputSheet = ss.getSheetByName(CONFIG.input.sheetName);

  if (!workSheet) {
    throw new Error('シートが見つかりません: ' + CONFIG.workList.sheetName);
  }
  if (!inputSheet) {
    throw new Error('シートが見つかりません: ' + CONFIG.input.sheetName);
  }

  const workValues = workSheet.getDataRange().getValues();
  const workCols = resolveColumns_(workValues[CONFIG.workList.headerRow - 1], CONFIG.workList.headers);

  if (!workCols.major || !workCols.mid) {
    throw new Error('作業リストに「大項目」「中項目」ヘッダーが見つかりません。1行目を確認してください。');
  }

  const inputHeaderRow = inputSheet
    .getRange(CONFIG.input.headerRow, 1, 1, Math.max(inputSheet.getLastColumn(), 1))
    .getValues()[0];
  const inputCols = resolveColumns_(inputHeaderRow, CONFIG.input.headers);

  // 入力シートにヘッダーが無い・名前が違う場合のフォールバック（仕様の B9 / C9）
  if (!inputCols.major) {
    inputCols.major = 2;
    Logger.log('%s 入力シートに大項目ヘッダーが無いため B 列を使います', CONFIG.logPrefix);
  }
  if (!inputCols.mid) {
    inputCols.mid = 3;
    Logger.log('%s 入力シートに中項目/作業内容ヘッダーが無いため C 列を使います', CONFIG.logPrefix);
  }

  const workRows = parseWorkList_(workValues, workCols);

  Logger.log(
    '%s loadContext_: 作業リスト行=%s / 作業列=%s / 入力列=%s',
    CONFIG.logPrefix,
    workRows.length,
    JSON.stringify(workCols),
    JSON.stringify(inputCols)
  );

  return {
    workSheet: workSheet,
    inputSheet: inputSheet,
    workCols: workCols,
    inputCols: inputCols,
    workRows: workRows
  };
}

/**
 * ヘッダー行配列から、論理名 → 1 始まりの列番号を作る。
 * 見つからない論理名はマップに載せない（任意列として扱える）。
 *
 * @param {*[]} headerRow
 * @param {Object<string, string[]>} aliasesByKey
 * @return {Object<string, number>}
 */
function resolveColumns_(headerRow, aliasesByKey) {
  const normalizedHeaders = headerRow.map(function (cell) {
    return normalize_(cell);
  });
  const result = {};

  Object.keys(aliasesByKey).forEach(function (logicalName) {
    const aliases = aliasesByKey[logicalName];
    for (let a = 0; a < aliases.length; a++) {
      const alias = normalize_(aliases[a]);
      const idx = normalizedHeaders.indexOf(alias);
      if (idx !== -1) {
        result[logicalName] = idx + 1;
        return;
      }
    }
  });

  return result;
}

/**
 * 作業リストを名前付きレコードの配列にする。列番号はこの後使わない。
 *
 * @param {*[][]} values
 * @param {Object<string, number>} cols
 * @return {object[]}
 */
function parseWorkList_(values, cols) {
  const headerIndex = CONFIG.workList.headerRow - 1;
  const rows = [];

  for (let i = headerIndex + 1; i < values.length; i++) {
    const raw = values[i];
    const major = normalize_(cell_(raw, cols.major));
    const mid = normalize_(cell_(raw, cols.mid));

    if (!major && !mid) {
      continue;
    }

    rows.push({
      sourceIndex: i + 1,
      major: major,
      mid: mid,
      content: cell_(raw, cols.content),
      fee: cell_(raw, cols.fee),
      order: cell_(raw, cols.order),
      partMajor: cell_(raw, cols.partMajor),
      partMid: cell_(raw, cols.partMid),
      qty: cell_(raw, cols.qty),
      unitPrice: cell_(raw, cols.unitPrice)
    });
  }

  return rows;
}
