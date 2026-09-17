/**
 * 作業リストの読み取り。
 * 「どの列が何か」はヘッダー名で解決し、以降は名前付きレコードだけを扱う。
 */

function invalidateContext_() {
  loadContext_.memo_ = null;
}

/**
 * 作業リストの列マップ・レコード配列を一度に作る。
 *
 * @return {{
 *   workSheet: GoogleAppsScript.Spreadsheet.Sheet,
 *   workCols: Object<string, number>,
 *   workRows: object[]
 * }}
 */
function loadContext_() {
  if (loadContext_.memo_) {
    return loadContext_.memo_;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workSheet = ss.getSheetByName(CONFIG.workList.sheetName);

  if (!workSheet) {
    throw new Error('シートが見つかりません: ' + CONFIG.workList.sheetName);
  }

  const workValues = workSheet.getDataRange().getValues();
  const workCols = resolveColumns_(workValues[CONFIG.workList.headerRow - 1], CONFIG.workList.headers);

  if (!workCols.major || !workCols.mid) {
    throw new Error('作業リストに「大項目」「中項目」ヘッダーが見つかりません。1行目を確認してください。');
  }

  const parsed = parseWorkList_(workValues, workCols);
  promoteMidFeesToWorkContent_(parsed);
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
    workCols: workCols,
    workRows: workRows,
    ordersPending: ordersPending,
    midsByMajor: index.midsByMajor,
    recordsByMajorMid: index.recordsByMajorMid
  };
  loadContext_.memo_ = ctx;
  return ctx;
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
  const majorSet = {};
  for (let i = 0; i < rows.length; i++) {
    const name = normalize_(rows[i].major);
    if (name) {
      majorSet[name] = true;
    }
  }

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!midsByMajor[row.major]) {
      midsByMajor[row.major] = [];
    }
    const key = row.major + '\t' + row.mid;
    if (!recordsByMajorMid[key]) {
      recordsByMajorMid[key] = [];
    }
    recordsByMajorMid[key].push(row);
  }

  Object.keys(recordsByMajorMid).forEach(function (key) {
    recordsByMajorMid[key] = recordsByMajorMid[key].slice().sort(compareMidGroupRows_);
    const parts = key.split('\t');
    const major = parts[0] || '';
    const mid = parts.slice(1).join('\t');
    if (!normalize_(major) || !mid || !midsByMajor[major]) {
      return;
    }
    if (midsByMajor[major].indexOf(mid) !== -1) {
      return;
    }
    const group = recordsByMajorMid[key];
    const midKey = normalize_(mid);
    const otherMajor = majorSet[midKey] && midKey !== normalize_(major);
    if (otherMajor && !workGroupHasSubstance_(group)) {
      return;
    }
    midsByMajor[major].push(mid);
  });

  return {
    midsByMajor: midsByMajor,
    recordsByMajorMid: recordsByMajorMid
  };
}

function workGroupHasSubstance_(group) {
  if (!group || !group.length) {
    return false;
  }
  for (let i = 0; i < group.length; i++) {
    const row = group[i];
    if (hasWorkContent_(row) || hasPartFields_(row) || isFilled_(row.fee)) {
      return true;
    }
  }
  return false;
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
  const majorNames = {};
  for (let i = headerIndex + 1; i < values.length; i++) {
    const name = normalize_(cell_(values[i], cols.major));
    if (name) {
      majorNames[name] = true;
    }
  }
  let carryMajor = '';
  let carryMid = '';

  for (let i = headerIndex + 1; i < values.length; i++) {
    const raw = values[i];
    const rawMajor = normalize_(cell_(raw, cols.major));
    const rawMid = normalize_(cell_(raw, cols.mid));
    const content = cell_(raw, cols.content);
    const partMajor = cell_(raw, cols.partMajor);
    const partMid = cell_(raw, cols.partMid);
    const fee = cell_(raw, cols.fee);

    if (!rawMajor && !rawMid && !isFilled_(content) && !normalize_(partMajor) && !normalize_(partMid) && !isFilled_(fee)) {
      carryMajor = '';
      carryMid = '';
      continue;
    }

    if (!rawMajor && rawMid && majorNames[rawMid] && rawMid !== carryMajor &&
        !isFilled_(content) && !normalize_(partMajor) && !normalize_(partMid) && !isFilled_(fee)) {
      carryMajor = rawMid;
      carryMid = '';
      rows.push({
        sourceIndex: i + 1,
        major: carryMajor,
        mid: '',
        content: content,
        fee: fee,
        workerCode: cell_(raw, cols.workerCode),
        order: cell_(raw, cols.order),
        partMajor: partMajor,
        partMid: partMid,
        qty: cell_(raw, cols.qty),
        unitPrice: cell_(raw, cols.unitPrice)
      });
      continue;
    }

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

    rows.push({
      sourceIndex: i + 1,
      major: major,
      mid: mid,
      content: content,
      fee: fee,
      workerCode: cell_(raw, cols.workerCode),
      order: cell_(raw, cols.order),
      partMajor: partMajor,
      partMid: partMid,
      qty: cell_(raw, cols.qty),
      unitPrice: cell_(raw, cols.unitPrice)
    });
  }

  return rows;
}

function getRecordsForSelection_(ctx, major, mid) {
  if (major) {
    return ctx.recordsByMajorMid[major + '\t' + mid] || [];
  }
  const rows = [];
  const suffix = '\t' + mid;
  const keys = Object.keys(ctx.recordsByMajorMid);
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (key.length >= suffix.length && key.substring(key.length - suffix.length) === suffix) {
      const group = ctx.recordsByMajorMid[key];
      for (let j = 0; j < group.length; j++) {
        rows.push(group[j]);
      }
    }
  }
  return rows;
}

/**
 * 作業内容が空で技術料がある行は、リスト追加と同じく作業内容＝中項目名にする。
 * 技術料は動かさない（その行の作業内容に載せる）。
 *
 * @param {object[]} rows
 * @return {number} 埋めた件数
 */
function promoteMidFeesToWorkContent_(rows) {
  if (!rows || !rows.length) {
    return 0;
  }
  let filled = 0;
  rows.forEach(function (row) {
    const mid = normalize_(row && row.mid);
    if (!mid || hasWorkContent_(row) || !isFilled_(row.fee)) {
      return;
    }
    row.content = mid;
    filled += 1;
  });
  return filled;
}

/**
 * 中項目に対する作業内容（入力アプリの展開用。技術料は作業内容行だけ）。
 */
function resolveMidOutput_(ctx, major, mid) {
  const records = getRecordsForSelection_(ctx, major, mid);
  tagMidGroups_(records);
  const sorted = records.slice().sort(compareMidGroupRows_);
  const workRows = [];
  for (let i = 0; i < sorted.length; i++) {
    if (hasWorkContent_(sorted[i])) {
      workRows.push(sorted[i]);
    }
  }
  return {
    midFee: '',
    workRows: workRows
  };
}
