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
    return;
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
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
 * 入力枠の最終行。getLastRow() は遅いので使わない。
 *
 * @return {number}
 */
function getOverwriteEndRow_() {
  return CONFIG.input.dataStartRow + CONFIG.input.maxSelectRows - 1;
}

/**
 * 中項目行の D 列（技術料）と、直下の作業内容をまとめて 1 回で書く。
 * 下行は needed と旧明細の大きい方だけ読む（200 行全読みしない）。
 *
 * @param {object} ctx
 * @param {number} selectRow
 * @param {*} midFee 中項目行の技術料
 * @param {object[]} contentRows
 */
function writeMidFeeAndDetails_(ctx, selectRow, midFee, contentRows) {
  const sheet = ctx.inputSheet;
  const majorCol = ctx.inputCols.major;
  const midCol = ctx.inputCols.mid;
  const feeCol = getFeeCol_(ctx);
  const needed = contentRows.length;
  const end = getOverwriteEndRow_();
  const available = Math.max(0, end - selectRow);
  const peek = Math.min(available, Math.max(needed, CONFIG.input.detailPeekRows || 80));

  let existing = 0;
  let capacity = available;

  if (peek > 0) {
    const minCol = Math.min(majorCol, midCol, feeCol);
    const maxCol = Math.max(majorCol, midCol, feeCol);
    const block = sheet.getRange(selectRow + 1, minCol, peek, maxCol - minCol + 1).getValues();
    capacity = 0;
    for (let i = 0; i < block.length; i++) {
      if (normalize_(block[i][majorCol - minCol]) !== '') {
        break;
      }
      capacity++;
    }
    existing = 0;
    for (let i = 0; i < capacity; i++) {
      if (!isFilled_(block[i][midCol - minCol]) && !isFilled_(block[i][feeCol - minCol])) {
        break;
      }
      existing++;
    }
  }

  const writeCount = Math.min(needed, capacity);
  const clearCount = Math.max(existing, writeCount);

  writeInternal_(function () {
    sheet.getRange(selectRow, feeCol).setValue(isFilled_(midFee) ? midFee : '');

    if (clearCount <= 0) {
      return;
    }

    const minCol = Math.min(midCol, feeCol);
    const maxCol = Math.max(midCol, feeCol);
    const width = maxCol - minCol + 1;
    const values = [];
    for (let i = 0; i < clearCount; i++) {
      const line = new Array(width);
      for (let c = 0; c < width; c++) {
        line[c] = '';
      }
      if (i < writeCount) {
        const row = contentRows[i];
        line[midCol - minCol] = row.content;
        line[feeCol - minCol] = isFilled_(row.fee) ? row.fee : '';
      }
      values.push(line);
    }
    sheet.getRange(selectRow + 1, minCol, clearCount, width).setValues(values);
  });
}

/**
 * @param {object} ctx
 * @param {number} selectRow
 * @param {object[]} contentRows
 */
function replaceExpandedDetailsBelow_(ctx, selectRow, contentRows) {
  writeMidFeeAndDetails_(ctx, selectRow, '', contentRows);
}
