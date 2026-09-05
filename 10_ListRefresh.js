/**
 * 作業リスト／部品リスト／作業者リストの「更新」チェック。
 * 順番（表示順）の小さい順にデータ行を並べ、同じ順で選択肢を作り直す。
 * 行の挿入はしない。数式（IMAGE など）は残す。
 */

function isListRefreshSheet_(sheetName) {
  return CONFIG.listRefresh.sheets.indexOf(sheetName) !== -1;
}

function listRefreshRangeName_(sheetName) {
  const map = {
    作業リスト: 'KINKAI_REFRESH_WORK',
    部品リスト: 'KINKAI_REFRESH_PARTS',
    作業者リスト: 'KINKAI_REFRESH_WORKERS'
  };
  return map[sheetName] || ('KINKAI_REFRESH_' + String(sheetName).replace(/\s+/g, '_'));
}

function isListMetaHeader_(header) {
  const n = normalize_(header);
  return n === '順番' || n === '表示順' || n === CONFIG.listRefresh.buttonLabel;
}

function isTruthyCheck_(value) {
  return value === true || value === 'TRUE' || value === 'true' || value === 1 || value === '1';
}

function isListRefreshCheckboxEdit_(sheet, range) {
  const named = sheet.getParent().getRangeByName(listRefreshRangeName_(sheet.getName()));
  if (named && named.getSheet().getSheetId() === sheet.getSheetId()) {
    return rangesOverlap_(range, named);
  }
  const headerRow = 1;
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const label = CONFIG.listRefresh.buttonLabel;
  for (let c = 0; c < headers.length; c++) {
    if (normalize_(headers[c]) !== label) {
      continue;
    }
    const box = sheet.getRange(headerRow, c + 2);
    if (rangesOverlap_(range, box)) {
      return true;
    }
  }
  return false;
}

function rangesOverlap_(a, b) {
  return a.getSheet().getSheetId() === b.getSheet().getSheetId()
    && a.getRow() <= b.getLastRow()
    && a.getLastRow() >= b.getRow()
    && a.getColumn() <= b.getLastColumn()
    && a.getLastColumn() >= b.getColumn();
}

/**
 * 3 シートに順番列と「更新」チェックを置く（なければ）。
 */
function ensureListRefreshControls_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  CONFIG.listRefresh.sheets.forEach(function (name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      return;
    }
    ensureOrderColumnOnSheet_(sheet);
    ensureListRefreshControlOnSheet_(sheet);
  });
}

function lastDataHeaderCol_(headers) {
  let last = 1;
  const label = CONFIG.listRefresh.buttonLabel;
  for (let i = 0; i < headers.length; i++) {
    const h = normalize_(headers[i]);
    if (!h || h === label) {
      continue;
    }
    last = i + 1;
  }
  return last;
}

function orderHeaderAliases_() {
  return ['順番', '表示順'];
}

function findOrderCol_(headers) {
  const aliases = orderHeaderAliases_();
  for (let i = 0; i < headers.length; i++) {
    const h = normalize_(headers[i]);
    if (aliases.indexOf(h) !== -1) {
      return i + 1;
    }
  }
  return 0;
}

function ensureOrderColumnOnSheet_(sheet) {
  const headerRow = 1;
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  let orderCol = findOrderCol_(headers);
  if (orderCol) {
    return orderCol;
  }
  const dataCol = lastDataHeaderCol_(headers);
  orderCol = dataCol + 1;
  sheet.getRange(headerRow, orderCol).setValue('順番').setFontWeight('bold');
  return orderCol;
}

function ensureListRefreshControlOnSheet_(sheet) {
  const ss = sheet.getParent();
  const rangeName = listRefreshRangeName_(sheet.getName());
  const headerRow = 1;
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const label = CONFIG.listRefresh.buttonLabel;
  let labelCol = 0;
  for (let c = 0; c < headers.length; c++) {
    if (normalize_(headers[c]) === label) {
      labelCol = c + 1;
      break;
    }
  }
  let created = false;
  if (!labelCol) {
    labelCol = lastDataHeaderCol_(headers) + 2;
    sheet.getRange(headerRow, labelCol).setValue(label).setFontWeight('bold');
    created = true;
  }
  const box = sheet.getRange(headerRow, labelCol + 1);
  box.insertCheckboxes();
  if (created) {
    box.setValue(false);
  }
  const existing = ss.getRangeByName(rangeName);
  if (!existing
      || existing.getSheet().getSheetId() !== sheet.getSheetId()
      || existing.getRow() !== box.getRow()
      || existing.getColumn() !== box.getColumn()) {
    if (existing) {
      ss.removeNamedRange(rangeName);
    }
    ss.setNamedRange(rangeName, box);
  }
}

/**
 * メニュー「リストを更新（順番・選択肢）」
 */
function refreshAllMasterLists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  writeInternal_(function () {
    ensureListRefreshControls_();
    CONFIG.listRefresh.sheets.forEach(function (name) {
      const sheet = ss.getSheetByName(name);
      if (sheet) {
        refreshMasterListSheet_(sheet);
      }
    });
  });
  ss.toast('リストを順番で並べ替え、選択肢を更新しました', '請求書入力', 5);
}

/**
 * 1 シート：空の順番を埋める（全部空のときだけ）→ 行を並べ替え → 作業リストなら候補再生成。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function refreshMasterListSheet_(sheet) {
  const headerRow = 1;
  const orderCol = ensureOrderColumnOnSheet_(sheet);
  ensureListRefreshControlOnSheet_(sheet);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const dataCol = lastDataHeaderCol_(headers);
  fillDefaultOrdersIfEmpty_(sheet, headerRow, orderCol, dataCol, headers);
  sortListDataRows_(sheet, headerRow, dataCol, orderCol, headers);
  invalidateContext_();
  if (sheet.getName() === CONFIG.workList.sheetName) {
    rebuildMidCandidateSheet_();
  }
  log_('%s refreshMasterListSheet_: %s を順番で並べ替えました', CONFIG.logPrefix, sheet.getName());
}

function fillDefaultOrdersIfEmpty_(sheet, headerRow, orderCol, dataCol, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) {
    return;
  }
  const height = lastRow - headerRow;
  const orders = sheet.getRange(headerRow + 1, orderCol, height, 1).getValues();
  const data = sheet.getRange(headerRow + 1, 1, height, dataCol).getValues();
  const formulas = sheet.getRange(headerRow + 1, 1, height, dataCol).getFormulas();
  const name = sheet.getName();
  let out;
  if (name === CONFIG.workList.sheetName) {
    out = fillWorkListGroupOrders_(data, formulas, orders, resolveColumns_(headers, CONFIG.workList.headers));
  } else if (name === CONFIG.parts.sheetName) {
    out = fillWorkListGroupOrders_(data, formulas, orders, resolveColumns_(headers, CONFIG.parts.headers));
  } else {
    let anyOrder = false;
    for (let i = 0; i < orders.length; i++) {
      if (!listRowIsEmpty_(data[i], formulas[i]) && toOrderNumber_(orders[i][0]) !== Number.POSITIVE_INFINITY) {
        anyOrder = true;
        break;
      }
    }
    if (anyOrder) {
      return;
    }
    const step = CONFIG.listRefresh.orderStep || 10;
    let n = 0;
    out = orders.map(function (row, i) {
      if (listRowIsEmpty_(data[i], formulas[i])) {
        return [''];
      }
      n += 1;
      return [n * step];
    });
  }
  sheet.getRange(headerRow + 1, orderCol, height, 1).setValues(out);
}

/**
 * 中項目が 2 行以上で順番が全部空のときだけ埋める。
 * 先頭（作業内容が空、または先頭行）に 1、残りに 2, 3…。
 */
function fillWorkListGroupOrders_(data, formulas, orders, cols) {
  const out = orders.map(function (row) {
    return [row[0]];
  });
  const groups = {};
  const keys = [];
  let carryMajor = '';
  let carryMid = '';
  for (let i = 0; i < data.length; i++) {
    if (listRowIsEmpty_(data[i], formulas[i])) {
      continue;
    }
    const raw = data[i];
    const rawMajor = normalize_(cell_(raw, cols.major));
    const rawMid = normalize_(cell_(raw, cols.mid));
    if (rawMajor) {
      carryMajor = rawMajor;
      if (!rawMid) {
        carryMid = '';
      }
    }
    if (rawMid) {
      carryMid = rawMid;
    }
    const rec = {
      i: i,
      major: rawMajor || carryMajor,
      mid: rawMid || carryMid,
      content: cols.content ? cell_(raw, cols.content) : '',
      order: orders[i][0],
      sourceIndex: i
    };
    const key = rec.major + '\t' + rec.mid;
    if (!groups[key]) {
      groups[key] = [];
      keys.push(key);
    }
    groups[key].push(rec);
  }
  keys.forEach(function (key) {
    const g = groups[key];
    if (g.length < 2) {
      return;
    }
    let any = false;
    for (let k = 0; k < g.length; k++) {
      if (toOrderNumber_(g[k].order) !== Number.POSITIVE_INFINITY) {
        any = true;
        break;
      }
    }
    if (any) {
      return;
    }
    g.sort(function (a, b) {
      return a.sourceIndex - b.sourceIndex;
    });
    let anchorIdx = -1;
    for (let k = 0; k < g.length; k++) {
      if (isMidAnchorRow_(g[k])) {
        anchorIdx = k;
        break;
      }
    }
    if (anchorIdx < 0) {
      anchorIdx = 0;
    }
    out[g[anchorIdx].i] = [1];
    let n = 2;
    for (let k = 0; k < g.length; k++) {
      if (k === anchorIdx) {
        continue;
      }
      out[g[k].i] = [n];
      n += 1;
    }
  });
  return out;
}

function listRowIsEmpty_(values, formulas) {
  const len = Math.max(values.length, formulas ? formulas.length : 0);
  for (let c = 0; c < len; c++) {
    if (formulas && formulas[c]) {
      return false;
    }
    const v = values[c];
    if (v === 0 || v === false) {
      return false;
    }
    if (v !== '' && v != null) {
      return false;
    }
  }
  return true;
}

function sortListDataRows_(sheet, headerRow, dataCol, orderCol, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRow || dataCol < 1) {
    return;
  }
  const height = lastRow - headerRow;
  const range = sheet.getRange(headerRow + 1, 1, height, dataCol);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const formats = range.getNumberFormats();
  const name = sheet.getName();
  const grouped = name === CONFIG.workList.sheetName || name === CONFIG.parts.sheetName;
  const cols = grouped
    ? resolveColumns_(headers, name === CONFIG.workList.sheetName ? CONFIG.workList.headers : CONFIG.parts.headers)
    : {};
  const rows = [];
  let carryMajor = '';
  let carryMid = '';
  for (let i = 0; i < height; i++) {
    const orderVal = orderCol <= dataCol ? values[i][orderCol - 1] : '';
    const rec = {
      order: orderVal,
      sourceIndex: i,
      values: values[i],
      formulas: formulas[i],
      formats: formats[i],
      major: '',
      mid: '',
      content: ''
    };
    if (grouped && !listRowIsEmpty_(values[i], formulas[i])) {
      const rawMajor = normalize_(cell_(values[i], cols.major));
      const rawMid = normalize_(cell_(values[i], cols.mid));
      if (rawMajor) {
        carryMajor = rawMajor;
        if (!rawMid) {
          carryMid = '';
        }
      }
      if (rawMid) {
        carryMid = rawMid;
      }
      rec.major = rawMajor || carryMajor;
      rec.mid = rawMid || carryMid;
      rec.content = cols.content ? cell_(values[i], cols.content) : '';
    }
    rows.push(rec);
  }
  const emptyRows = [];
  const filledRows = [];
  for (let i = 0; i < rows.length; i++) {
    if (listRowIsEmpty_(rows[i].values, rows[i].formulas)) {
      emptyRows.push(rows[i]);
    } else {
      filledRows.push(rows[i]);
    }
  }
  const sorted = grouped ? sortWorkListRecords_(filledRows) : sortByOrder_(filledRows);
  const finalRows = sorted.concat(emptyRows);
  const outValues = [];
  const outFormats = [];
  for (let r = 0; r < finalRows.length; r++) {
    const valueRow = [];
    for (let c = 0; c < dataCol; c++) {
      const f = finalRows[r].formulas[c];
      valueRow.push(f ? f : finalRows[r].values[c]);
    }
    outValues.push(valueRow);
    outFormats.push(finalRows[r].formats);
  }
  range.setValues(outValues);
  range.setNumberFormats(outFormats);
}
