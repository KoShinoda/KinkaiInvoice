/**
 * 入力シートへの書き込み専用。
 * 「何を出すか」は SelectionService、「どう書くか」だけをここで扱う。
 */

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} col
 * @param {number} row
 * @param {string[]} list
 */
function setDropdown_(sheet, col, row, list) {
  setDropdownOnRange_(sheet.getRange(row, col), list);
  Logger.log('%s プルダウン設定: 行%s 列%s 件数=%s', CONFIG.logPrefix, row, col, list.length);
}

/**
 * 同じ候補でよい範囲（大項目列 B9:B208 など）に一度で入力規則を付ける。
 * 中項目は行ごとに候補が違うので、こちらは使わない。
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 * @param {string[]} list
 */
function setDropdownOnRange_(range, list) {
  if (!list.length) {
    range.clearDataValidations();
    Logger.log('%s プルダウン候補が空のため Validation を削除: %s', CONFIG.logPrefix, range.getA1Notation());
    return;
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
  Logger.log('%s プルダウン設定: %s 件数=%s', CONFIG.logPrefix, range.getA1Notation(), list.length);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} col
 * @param {number} row
 */
function clearDropdown_(sheet, col, row) {
  writeInternal_(function () {
    sheet.getRange(row, col).clearDataValidations();
  });
}

/**
 * 技術料列。ヘッダーが無ければ D 列。
 *
 * @param {object} ctx
 * @return {number}
 */
function getFeeCol_(ctx) {
  return ctx.inputCols.fee || 4;
}

/**
 * 下行へ上書きできる最終行（入力枠 200 行と、実際に使っている末尾の大きい方）。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {number}
 */
function getOverwriteEndRow_(sheet) {
  return Math.max(
    CONFIG.input.dataStartRow + CONFIG.input.maxSelectRows - 1,
    sheet.getLastRow()
  );
}

/**
 * 選択行のすぐ下で、大項目（B）が空の連続行数。次の大項目より上だけ上書き可能。
 *
 * @param {object} ctx
 * @param {number} selectRow
 * @return {number}
 */
function countOverwriteCapacity_(ctx, selectRow) {
  const sheet = ctx.inputSheet;
  const end = getOverwriteEndRow_(sheet);
  if (selectRow >= end) {
    return 0;
  }

  const majorCol = ctx.inputCols.major;
  const start = selectRow + 1;
  const num = end - selectRow;
  const majors = sheet.getRange(start, majorCol, num, 1).getValues();
  let count = 0;
  for (let i = 0; i < majors.length; i++) {
    if (normalize_(majors[i][0]) !== '') {
      break;
    }
    count++;
  }
  return count;
}

/**
 * 選択行のすぐ下にある「展開済み作業内容」の行数。
 * 大項目（B）が空で、作業内容（C）か技術料（D）が入っている連続行。
 *
 * @param {object} ctx
 * @param {number} selectRow
 * @return {number}
 */
function countExpandedDetailRows_(ctx, selectRow) {
  const sheet = ctx.inputSheet;
  const end = getOverwriteEndRow_(sheet);
  if (selectRow >= end) {
    return 0;
  }

  const majorCol = ctx.inputCols.major;
  const midCol = ctx.inputCols.mid;
  const feeCol = getFeeCol_(ctx);
  const start = selectRow + 1;
  const num = end - selectRow;
  const minCol = Math.min(majorCol, midCol, feeCol);
  const maxCol = Math.max(majorCol, midCol, feeCol);
  const values = sheet.getRange(start, minCol, num, maxCol - minCol + 1).getValues();

  let count = 0;
  for (let i = 0; i < values.length; i++) {
    const major = values[i][majorCol - minCol];
    const content = values[i][midCol - minCol];
    const fee = values[i][feeCol - minCol];
    if (normalize_(major) !== '') {
      break;
    }
    if (!isFilled_(content) && !isFilled_(fee)) {
      break;
    }
    count++;
  }
  return count;
}

/**
 * 中項目行の直下へ作業内容＋技術料を上書きする。行の挿入・削除はしない。
 * 次の大項目（B に値がある行）は上書きしない。余った旧明細は C・D を空にする。
 *
 * @param {object} ctx
 * @param {number} selectRow
 * @param {object[]} contentRows 作業内容があるマスタ行（順番済み）
 */
function replaceExpandedDetailsBelow_(ctx, selectRow, contentRows) {
  const sheet = ctx.inputSheet;
  const midCol = ctx.inputCols.mid;
  const feeCol = getFeeCol_(ctx);
  const existing = countExpandedDetailRows_(ctx, selectRow);
  const capacity = countOverwriteCapacity_(ctx, selectRow);
  const needed = contentRows.length;
  const writeCount = Math.min(needed, capacity);
  const clearCount = Math.max(existing, writeCount);

  Logger.log(
    '%s replaceExpandedDetailsBelow_: 選択行=%s 既存明細=%s 上書き可能=%s 必要=%s 実際に書く=%s',
    CONFIG.logPrefix,
    selectRow,
    existing,
    capacity,
    needed,
    writeCount
  );

  if (needed > capacity) {
    Logger.log(
      '%s 下行の空きが足りません。%s 件中 %s 件だけ上書きします（次の大項目行は維持）',
      CONFIG.logPrefix,
      needed,
      writeCount
    );
  }

  writeInternal_(function () {
    if (clearCount > 0) {
      const minCol = Math.min(midCol, feeCol);
      const maxCol = Math.max(midCol, feeCol);
      const width = maxCol - minCol + 1;
      const blanks = [];
      for (let i = 0; i < clearCount; i++) {
        const line = new Array(width);
        for (let c = 0; c < width; c++) {
          line[c] = '';
        }
        blanks.push(line);
      }
      sheet.getRange(selectRow + 1, minCol, clearCount, width).setValues(blanks);
      sheet.getRange(selectRow + 1, midCol, clearCount, 1).clearDataValidations();
    }

    if (writeCount === 0) {
      return;
    }

    const minCol = Math.min(midCol, feeCol);
    const maxCol = Math.max(midCol, feeCol);
    const width = maxCol - minCol + 1;
    const values = [];
    for (let i = 0; i < writeCount; i++) {
      const row = contentRows[i];
      const line = new Array(width);
      for (let c = 0; c < width; c++) {
        line[c] = '';
      }
      line[midCol - minCol] = row.content;
      line[feeCol - minCol] = isFilled_(row.fee) ? row.fee : '';
      values.push(line);
    }

    sheet.getRange(selectRow + 1, minCol, writeCount, width).setValues(values);
    sheet.getRange(selectRow + 1, midCol, writeCount, 1).clearDataValidations();

    Logger.log(
      '%s 作業内容を %s 行上書きしました（%s行目から）。先頭=%s',
      CONFIG.logPrefix,
      writeCount,
      selectRow + 1,
      JSON.stringify(values[0])
    );
  });
}

/**
 * @param {object} ctx
 * @param {number} row
 * @param {*} fee
 */
function setFeeOnRow_(ctx, row, fee) {
  const feeCol = getFeeCol_(ctx);
  ctx.inputSheet.getRange(row, feeCol).setValue(isFilled_(fee) ? fee : '');
}
