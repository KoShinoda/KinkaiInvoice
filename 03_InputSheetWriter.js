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
 * 選択行のすぐ下にある「展開済み作業内容」の行数。
 * 大項目（B）が空で、作業内容（C）か技術料（D）が入っている連続行だけを明細とみなす。
 * 次の大項目行や、B/C/D がすべて空の入力枠は残す。
 *
 * @param {object} ctx
 * @param {number} selectRow
 * @return {number}
 */
function countExpandedDetailRows_(ctx, selectRow) {
  const sheet = ctx.inputSheet;
  const last = sheet.getLastRow();
  if (selectRow >= last) {
    return 0;
  }

  const majorCol = ctx.inputCols.major;
  const midCol = ctx.inputCols.mid;
  const feeCol = getFeeCol_(ctx);
  const start = selectRow + 1;
  const num = last - selectRow;
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
 * 中項目行の直下を、作業内容＋技術料の明細で差し替える。
 * 次の大項目行は削除せず、足りなければ行挿入、多ければ明細行だけ削除する。
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
  const needed = contentRows.length;

  Logger.log(
    '%s replaceExpandedDetailsBelow_: 選択行=%s 既存明細=%s 必要=%s',
    CONFIG.logPrefix,
    selectRow,
    existing,
    needed
  );

  writeInternal_(function () {
    if (existing > 0) {
      sheet.deleteRows(selectRow + 1, existing);
      Logger.log('%s 旧明細 %s 行を削除しました（%s行目から）', CONFIG.logPrefix, existing, selectRow + 1);
    }

    if (needed === 0) {
      return;
    }

    sheet.insertRowsAfter(selectRow, needed);

    const minCol = Math.min(midCol, feeCol);
    const maxCol = Math.max(midCol, feeCol);
    const width = maxCol - minCol + 1;
    const values = contentRows.map(function (row) {
      const line = new Array(width);
      for (let i = 0; i < width; i++) {
        line[i] = '';
      }
      line[midCol - minCol] = row.content;
      line[feeCol - minCol] = isFilled_(row.fee) ? row.fee : '';
      return line;
    });

    sheet.getRange(selectRow + 1, minCol, needed, width).setValues(values);
    sheet.getRange(selectRow + 1, midCol, needed, 1).clearDataValidations();

    Logger.log(
      '%s 作業内容を %s 行、%s行目から出力しました。先頭=%s',
      CONFIG.logPrefix,
      needed,
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
