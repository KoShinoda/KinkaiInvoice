/**
 * 中項目候補マスタの生成と、C 列入力規則の初期設定。
 * 大項目を選ぶ瞬間には動かない。作業リスト更新時とメニューからだけ動く。
 */

function invalidateContext_() {
  loadContext_.memo_ = null;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} name
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/**
 * 作業リストから横持ちの中項目候補を作り、参照用数式と C 列の入力規則を付ける。
 */
function rebuildMidCandidateSheet_() {
  invalidateContext_();
  const ctx = loadContext_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cand = getOrCreateSheet_(ss, CONFIG.candidates.sheetName);
  const lookup = getOrCreateSheet_(ss, CONFIG.lookup.sheetName);

  const majors = uniqueValues_(ctx.workRows.map(function (row) {
    return row.major;
  }));
  const allMids = uniqueValues_(ctx.workRows.map(function (row) {
    return row.mid;
  }));
  const maxRows = Math.max(CONFIG.candidates.maxRows, allMids.length, 20);

  const headers = [CONFIG.candidates.allHeader].concat(majors);
  const width = headers.length;
  const height = maxRows + 1;
  const grid = [];
  for (let r = 0; r < height; r++) {
    grid[r] = [];
    for (let c = 0; c < width; c++) {
      grid[r][c] = '';
    }
  }
  for (let c = 0; c < width; c++) {
    grid[0][c] = headers[c];
  }
  for (let i = 0; i < allMids.length; i++) {
    grid[i + 1][0] = allMids[i];
  }
  for (let m = 0; m < majors.length; m++) {
    const mids = ctx.midsByMajor[majors[m]] || [];
    for (let i = 0; i < mids.length; i++) {
      grid[i + 1][m + 1] = mids[i];
    }
  }

  cand.clear();
  cand.getRange(1, 1, height, width).setValues(grid);
  cand.getRange(1, 1, 1, width).setFontWeight('bold');
  cand.setFrozenRows(1);

  writeLookupFormulas_(lookup, ctx);
  bindMajorValidations_(ctx, cand, majors.length);
  bindAllMidValidations_(ctx, lookup);

  lookup.hideSheet();
  log_('%s 中項目候補を再生成しました。大項目=%s 全中項目=%s', CONFIG.logPrefix, majors.length, allMids.length);
}

/**
 * 中項目_参照の2行目に、入力行ごとの FILTER 数式を横に並べる。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} lookup
 * @param {object} ctx
 */
function writeLookupFormulas_(lookup, ctx) {
  const inputName = CONFIG.input.sheetName.replace(/'/g, "''");
  const candName = CONFIG.candidates.sheetName.replace(/'/g, "''");
  const start = CONFIG.input.dataStartRow;
  const end = start + CONFIG.input.maxSelectRows - 1;
  const maxR = 1 + CONFIG.candidates.maxRows;
  const majorLetter = toA1Col_(ctx.inputCols.major);
  const formula = '=IFERROR(LET(' +
    'maj,INDEX(\'' + inputName + '\'!$' + majorLetter + '$' + start + ':$' + majorLetter + '$' + end + ',COLUMN()),' +
    'col,IF(maj="",1,IFERROR(MATCH(maj,\'' + candName + '\'!$1:$1,0),1)),' +
    'list,INDEX(\'' + candName + '\'!$2:$' + maxR + ',0,col),' +
    'FILTER(list,list<>"")),"")';

  const num = CONFIG.input.maxSelectRows;
  lookup.clear();
  lookup.getRange(1, 1, 1, num).setValues([buildLookupHeaderRow_(num)]);
  lookup.getRange(2, 1, 1, num).setFormula(formula);
}

/**
 * @param {number} num
 * @return {string[]}
 */
function buildLookupHeaderRow_(num) {
  const start = CONFIG.input.dataStartRow;
  const headers = [];
  for (let i = 0; i < num; i++) {
    headers.push('入力' + (start + i) + '行');
  }
  return headers;
}

/**
 * 大項目列は候補シート1行目（（全て）以外）を参照。一度設定すれば再計算だけ。
 *
 * @param {object} ctx
 * @param {GoogleAppsScript.Spreadsheet.Sheet} cand
 * @param {number} majorCount
 */
function bindMajorValidations_(ctx, cand, majorCount) {
  const start = CONFIG.input.dataStartRow;
  const num = CONFIG.input.maxSelectRows;
  const majorRange = ctx.inputSheet.getRange(start, ctx.inputCols.major, num, 1);
  if (majorCount < 1) {
    majorRange.clearDataValidations();
    return;
  }
  const source = cand.getRange(1, 2, 1, majorCount);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(source, true)
    .setAllowInvalid(true)
    .build();
  majorRange.setDataValidation(rule);
}

/**
 * C9 は 中項目_参照 の A 列、C10 は B 列…と対応させる。
 *
 * @param {object} ctx
 * @param {GoogleAppsScript.Spreadsheet.Sheet} lookup
 */
function bindAllMidValidations_(ctx, lookup) {
  const start = CONFIG.input.dataStartRow;
  const num = CONFIG.input.maxSelectRows;
  const midCol = ctx.inputCols.mid;
  const maxR = CONFIG.candidates.maxRows;
  for (let i = 0; i < num; i++) {
    bindMidValidationForRow_(ctx.inputSheet, lookup, start + i, midCol, maxR);
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} inputSheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} lookup
 * @param {number} inputRow
 * @param {number=} midCol
 * @param {number=} maxR
 */
function bindMidValidationForRow_(inputSheet, lookup, inputRow, midCol, maxR) {
  const start = CONFIG.input.dataStartRow;
  if (inputRow < start) {
    return;
  }
  const col = inputRow - start + 1;
  if (col > CONFIG.input.maxSelectRows) {
    return;
  }
  const feeMidCol = midCol || CONFIG.input.fixedCols.mid;
  const height = maxR || CONFIG.candidates.maxRows;
  const source = lookup.getRange(2, col, height, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(source, true)
    .setAllowInvalid(true)
    .build();
  inputSheet.getRange(inputRow, feeMidCol).setDataValidation(rule);
}
