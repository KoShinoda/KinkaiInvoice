/**
 * 作業リスト／部品リストの並べ替えと、欠けている順番・列の補完。
 * 更新は図形ボタンまたはメニューから refreshAllMasterLists を実行する（チェックは置かない）。
 * 既存の順番は上書きしない。
 */

function isListRefreshSheet_(sheetName) {
  return CONFIG.listRefresh.sheets.indexOf(sheetName) !== -1;
}

function isListMetaHeader_(header) {
  const n = normalize_(header);
  return n === '順番' || n === '表示順' || n === CONFIG.listRefresh.buttonLabel;
}

/**
 * 図形のボタンに割り当てる。作業リスト・部品リストを順番で並べ、選択肢を作り直す。
 */
function refreshAllMasterLists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  writeInternal_(function () {
    ensureListMasterSheets_();
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
 * 今見ているリストシートだけ更新する図形用。
 */
function refreshActiveMasterList() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!sheet || !isListRefreshSheet_(sheet.getName())) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '作業リストまたは部品リストを表示してから実行してください',
      '請求書入力',
      5
    );
    return;
  }
  writeInternal_(function () {
    refreshMasterListSheet_(sheet);
  });
  SpreadsheetApp.getActiveSpreadsheet().toast(
    sheet.getName() + ' を順番で並べ替え、選択肢を更新しました',
    '請求書入力',
    5
  );
}

function ensureListMasterSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const work = ss.getSheetByName(CONFIG.workList.sheetName);
  if (work) {
    ensureOrderColumnOnSheet_(work);
    ensureWorkListPartColumns_(work);
    assignMissingListOrders_(work);
  }
  const parts = ss.getSheetByName(CONFIG.parts.sheetName);
  if (parts) {
    ensureOrderColumnOnSheet_(parts);
    assignMissingListOrders_(parts);
  }
  const workers = ss.getSheetByName(CONFIG.workers.sheetName);
  if (workers) {
    removeWorkerOrderColumn_(workers);
    assignMissingWorkerCodes_(workers);
  }
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

function ensureWorkListPartColumns_(sheet) {
  const wanted = [
    { name: '部品_中項目', aliases: CONFIG.workList.headers.partMid },
    { name: '単価', aliases: CONFIG.workList.headers.unitPrice },
    { name: '数量', aliases: CONFIG.workList.headers.qty }
  ];
  wanted.forEach(function (item) {
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const normalized = headers.map(function (h) {
      return normalize_(h);
    });
    let found = false;
    for (let a = 0; a < item.aliases.length; a++) {
      if (normalized.indexOf(normalize_(item.aliases[a])) !== -1) {
        found = true;
        break;
      }
    }
    if (found) {
      return;
    }
    const at = lastDataHeaderCol_(headers) + 1;
    sheet.getRange(1, at).setValue(item.name).setFontWeight('bold');
  });
}

function removeWorkerOrderColumn_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const orderCol = findOrderCol_(headers);
  if (orderCol) {
    sheet.deleteColumn(orderCol);
  }
}

/**
 * 1 シート：欠けている順番を補う → 行を並べ替え → 作業リストなら候補再生成。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function refreshMasterListSheet_(sheet) {
  const headerRow = 1;
  const orderCol = ensureOrderColumnOnSheet_(sheet);
  if (sheet.getName() === CONFIG.workList.sheetName) {
    ensureWorkListPartColumns_(sheet);
  }
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const dataCol = lastDataHeaderCol_(headers);
  assignMissingListOrders_(sheet);
  sortListDataRows_(sheet, headerRow, dataCol, orderCol, headers);
  invalidateContext_();
  if (sheet.getName() === CONFIG.workList.sheetName) {
    rebuildMidCandidateSheet_();
  }
  log_('%s refreshMasterListSheet_: %s を順番で並べ替えました', CONFIG.logPrefix, sheet.getName());
}

/**
 * 空の順番だけ埋める。手動で入っている値は触らない。
 * 中項目が 1 件だけなら 1。追加行は既存の最大 + 1。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {boolean} 書き込んだか
 */
function assignMissingListOrders_(sheet) {
  const headerRow = 1;
  const orderCol = ensureOrderColumnOnSheet_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) {
    return false;
  }
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const dataCol = lastDataHeaderCol_(headers);
  const height = lastRow - headerRow;
  const orders = sheet.getRange(headerRow + 1, orderCol, height, 1).getValues();
  const data = sheet.getRange(headerRow + 1, 1, height, dataCol).getValues();
  const formulas = sheet.getRange(headerRow + 1, 1, height, dataCol).getFormulas();
  const name = sheet.getName();
  const aliases = name === CONFIG.parts.sheetName ? CONFIG.parts.headers : CONFIG.workList.headers;
  const out = fillWorkListGroupOrders_(data, formulas, orders, resolveColumns_(headers, aliases));
  let changed = false;
  for (let i = 0; i < out.length; i++) {
    if (String(out[i][0]) !== String(orders[i][0])) {
      changed = true;
      break;
    }
  }
  if (changed) {
    sheet.getRange(headerRow + 1, orderCol, height, 1).setValues(out);
  }
  return changed;
}

/**
 * 空の順番だけ埋める。既存値は維持する。
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
    const g = groups[key].slice().sort(function (a, b) {
      return a.sourceIndex - b.sourceIndex;
    });
    let maxN = 0;
    g.forEach(function (rec) {
      const n = toOrderNumber_(rec.order);
      if (n !== Number.POSITIVE_INFINITY && n > maxN) {
        maxN = n;
      }
    });
    const empty = g.filter(function (rec) {
      return toOrderNumber_(rec.order) === Number.POSITIVE_INFINITY;
    });
    if (!empty.length) {
      return;
    }
    if (maxN === 0) {
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
      return;
    }
    empty.forEach(function (rec) {
      maxN += 1;
      out[rec.i] = [maxN];
    });
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

function assignMissingWorkerCodes_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    return;
  }
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headerRow = 1;
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  let start = 0;
  const h0 = normalize_(headers[0]);
  if (h0.indexOf('コード') !== -1 || h0.indexOf('作業者') !== -1 || h0.indexOf('担当') !== -1) {
    start = 1;
  }
  const cols = start === 1 ? resolveColumns_(headers, CONFIG.workers.headers) : {};
  const codeCol = cols.code || 1;
  const nameCol = cols.name || 2;
  if (lastRow <= start) {
    return;
  }
  const height = lastRow - start;
  const codes = sheet.getRange(start + 1, codeCol, height, 1).getValues();
  const names = sheet.getRange(start + 1, nameCol, height, 1).getValues();
  const existing = [];
  for (let i = 0; i < codes.length; i++) {
    const c = normalize_(codes[i][0]);
    if (c) {
      existing.push(c);
    }
  }
  let next = nextWorkerCode_(existing);
  const out = [];
  let changed = false;
  for (let i = 0; i < codes.length; i++) {
    const code = normalize_(codes[i][0]);
    const name = normalize_(names[i][0]);
    if (!code && name) {
      out.push([next]);
      existing.push(String(next));
      next = nextWorkerCode_(existing);
      changed = true;
    } else {
      out.push([codes[i][0]]);
    }
  }
  if (changed) {
    sheet.getRange(start + 1, codeCol, height, 1).setValues(out);
  }
}

function nextWorkerCode_(codes) {
  let max = 0;
  let width = 1;
  (codes || []).forEach(function (c) {
    const s = String(c).trim();
    if (/^\d+$/.test(s)) {
      if (s.length > width) {
        width = s.length;
      }
      const n = parseInt(s, 10);
      if (n > max) {
        max = n;
      }
    }
  });
  const n = max + 1;
  const s = String(n);
  if (s.length >= width) {
    return s;
  }
  return ('0000000000' + s).slice(-width);
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
