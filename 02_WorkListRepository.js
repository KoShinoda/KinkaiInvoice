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
  if (loadContext_.memo_) {
    return loadContext_.memo_;
  }

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

  const inputCols = resolveInputCols_(inputSheet);
  const parsed = parseWorkList_(workValues, workCols);
  const ordersPending = rowsHaveEmptyOrder_(parsed);
  assignEmptyOrdersInGroups_(parsed);
  const workRows = sortWorkListRecords_(parsed);
  const index = buildWorkIndex_(workRows);

  log_(
    '%s loadContext_: 作業リスト行=%s / 中項目グループ=%s',
    CONFIG.logPrefix,
    workRows.length,
    Object.keys(index.midsByMajor).length
  );

  const ctx = {
    workSheet: workSheet,
    inputSheet: inputSheet,
    workCols: workCols,
    inputCols: inputCols,
    workRows: workRows,
    ordersPending: ordersPending,
    midsByMajor: index.midsByMajor,
    recordsByMajorMid: index.recordsByMajorMid
  };
  loadContext_.memo_ = ctx;
  return ctx;
}

/**
 * 入力シートの列。fixedCols があれば見出しを読まない。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} inputSheet
 * @return {Object<string, number>}
 */
function resolveInputCols_(inputSheet) {
  if (CONFIG.input.fixedCols) {
    return {
      major: CONFIG.input.fixedCols.major,
      mid: CONFIG.input.fixedCols.mid,
      fee: CONFIG.input.fixedCols.fee
    };
  }

  const inputHeaderRow = inputSheet
    .getRange(CONFIG.input.headerRow, 1, 1, Math.max(inputSheet.getLastColumn(), 1))
    .getValues()[0];
  const inputCols = resolveColumns_(inputHeaderRow, CONFIG.input.headers);
  if (!inputCols.major) {
    inputCols.major = 2;
  }
  if (!inputCols.mid) {
    inputCols.mid = 3;
  }
  if (!inputCols.fee) {
    inputCols.fee = 4;
  }
  return inputCols;
}

/**
 * 大項目 → 中項目一覧、大項目+中項目 → レコード、を一回で作る。
 *
 * @param {object[]} rows
 * @return {{midsByMajor: Object<string, string[]>, recordsByMajorMid: Object<string, object[]>}}
 */
function buildWorkIndex_(rows) {
  const midsByMajor = {};
  const recordsByMajorMid = {};

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!midsByMajor[row.major]) {
      midsByMajor[row.major] = [];
    }
    if (row.mid && midsByMajor[row.major].indexOf(row.mid) === -1) {
      midsByMajor[row.major].push(row.mid);
    }
    const key = row.major + '\t' + row.mid;
    if (!recordsByMajorMid[key]) {
      recordsByMajorMid[key] = [];
    }
    recordsByMajorMid[key].push(row);
  }

  Object.keys(recordsByMajorMid).forEach(function (key) {
    recordsByMajorMid[key] = recordsByMajorMid[key].slice().sort(compareMidGroupRows_);
  });

  return {
    midsByMajor: midsByMajor,
    recordsByMajorMid: recordsByMajorMid
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
  let carryMajor = '';
  let carryMid = '';

  for (let i = headerIndex + 1; i < values.length; i++) {
    const raw = values[i];
    const rawMajor = normalize_(cell_(raw, cols.major));
    const rawMid = normalize_(cell_(raw, cols.mid));
    const content = cell_(raw, cols.content);
    const partMajor = cell_(raw, cols.partMajor);
    const partMid = cell_(raw, cols.partMid);

    if (rawMajor) {
      carryMajor = rawMajor;
      if (!rawMid) {
        carryMid = '';
      }
    }
    if (rawMid) {
      carryMid = rawMid;
    }

    const major = rawMajor || carryMajor;
    const mid = rawMid || carryMid;

    if (!major && !mid && !isFilled_(content) && !normalize_(partMajor) && !normalize_(partMid)) {
      continue;
    }

    rows.push({
      sourceIndex: i + 1,
      major: major,
      mid: mid,
      content: content,
      fee: cell_(raw, cols.fee),
      workerCode: cell_(raw, cols.workerCode),
      order: cell_(raw, cols.order),
      partMajor: cell_(raw, cols.partMajor),
      partMid: cell_(raw, cols.partMid),
      qty: cell_(raw, cols.qty),
      unitPrice: cell_(raw, cols.unitPrice)
    });
  }

  return rows;
}
