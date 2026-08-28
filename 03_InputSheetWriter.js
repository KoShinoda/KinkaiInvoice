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
  const range = sheet.getRange(row, col);
  if (!list.length) {
    range.clearDataValidations();
    Logger.log('%s プルダウン候補が空のため Validation を削除: 行%s 列%s', CONFIG.logPrefix, row, col);
    return;
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
  Logger.log('%s プルダウン設定: 行%s 列%s 件数=%s', CONFIG.logPrefix, row, col, list.length);
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
 * 部品名：部品_中項目（または部品名）があればそれ、なければ部品_大項目。
 *
 * @param {object} row
 * @return {*}
 */
function pickPartName_(row) {
  if (row.partMid !== '' && row.partMid !== null && row.partMid !== undefined) {
    return row.partMid;
  }
  return row.partMajor;
}

/**
 * 入力シート 1 行分を、存在する列だけ組み立てる。
 *
 * @param {object} ctx
 * @param {{displayName:*, fee:*, partMajor:*, partName:*, qty:*, unitPrice:*}} src
 * @return {Object<number, *>}
 */
function buildOutputLine_(ctx, src) {
  const cols = ctx.inputCols;
  const line = {};

  if (cols.mid) {
    line[cols.mid] = src.displayName;
  }
  if (cols.fee) {
    line[cols.fee] = src.fee;
  }
  if (cols.partMajor) {
    line[cols.partMajor] = src.partMajor;
  }
  if (cols.partName) {
    line[cols.partName] = src.partName;
  }
  if (cols.qty) {
    line[cols.qty] = src.qty;
  }
  if (cols.unitPrice) {
    line[cols.unitPrice] = src.unitPrice;
  }

  return line;
}

/**
 * 作業リスト行から、横並び出力用の 1 行を作る。
 *
 * @param {object} ctx
 * @param {object} row displayName を付けたレコード
 * @return {Object<number, *>}
 */
function buildOutputLineFromRow_(ctx, row) {
  return buildOutputLine_(ctx, {
    displayName: row.displayName,
    fee: row.fee,
    partMajor: row.partMajor,
    partName: pickPartName_(row),
    qty: row.qty,
    unitPrice: row.unitPrice
  });
}

/**
 * 出力列だけをまとめて書き込む。担当者・金額・連番は触らない。
 *
 * @param {object} ctx
 * @param {Array<Object<number, *>>} lines
 */
function writeOutputRows_(ctx, lines) {
  if (!lines.length) {
    return;
  }

  const colNumbers = getOutputColumnNumbers_(ctx);
  if (!colNumbers.length) {
    Logger.log('%s 出力先列が見つかりません。入力シートのヘッダー行を確認してください', CONFIG.logPrefix);
    return;
  }

  const startRow = CONFIG.input.outputStartRow;
  const minCol = Math.min.apply(null, colNumbers);
  const maxCol = Math.max.apply(null, colNumbers);
  const width = maxCol - minCol + 1;
  const values = lines.map(function (line) {
    const row = new Array(width);
    for (let i = 0; i < width; i++) {
      row[i] = '';
    }
    Object.keys(line).forEach(function (colKey) {
      const col = Number(colKey);
      row[col - minCol] = line[col] === undefined || line[col] === null ? '' : line[col];
    });
    return row;
  });

  Logger.log(
    '%s 出力: 開始行=%s 列=%s〜%s 行数=%s 先頭行=%s',
    CONFIG.logPrefix,
    startRow,
    minCol,
    maxCol,
    values.length,
    JSON.stringify(values[0])
  );

  writeInternal_(function () {
    ctx.inputSheet.getRange(startRow, minCol, values.length, width).setValues(values);
  });
}

/**
 * C10 以降の「このスクリプトが書く列」だけクリアする。
 *
 * @param {object} ctx
 */
function clearOutputArea_(ctx) {
  const colNumbers = getOutputColumnNumbers_(ctx);
  if (!colNumbers.length) {
    return;
  }

  const startRow = CONFIG.input.outputStartRow;
  const minCol = Math.min.apply(null, colNumbers);
  const maxCol = Math.max.apply(null, colNumbers);
  const width = maxCol - minCol + 1;
  const height = CONFIG.input.maxOutputRows;

  Logger.log('%s 出力エリアをクリア: %s行 x %s列（%s行目, 列%s〜%s）', CONFIG.logPrefix, height, width, startRow, minCol, maxCol);

  writeInternal_(function () {
    ctx.inputSheet.getRange(startRow, minCol, height, width).clearContent();
  });
}

/**
 * @param {object} ctx
 * @return {number[]}
 */
function getOutputColumnNumbers_(ctx) {
  const keys = ['mid', 'fee', 'partMajor', 'partName', 'qty', 'unitPrice'];
  return keys
    .map(function (key) {
      return ctx.inputCols[key];
    })
    .filter(function (col) {
      return typeof col === 'number' && col > 0;
    });
}
