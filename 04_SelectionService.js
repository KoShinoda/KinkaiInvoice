/**
 * 選択ロジック。
 * 大項目 → その行の中項目セレクトボックス（重複なし）。
 * 中項目 → 空白以外の作業内容を下行へ展開し、技術料を D 列へ載せる。
 */

/**
 * 入力シートで大項目・中項目を置く最終行（dataStartRow から maxSelectRows 行）。
 * 行挿入後はシート末尾まで見る。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet=} sheet
 * @return {number}
 */
function getInputDataEndRow_(sheet) {
  const configured = CONFIG.input.dataStartRow + CONFIG.input.maxSelectRows - 1;
  if (!sheet) {
    return configured;
  }
  return Math.max(configured, sheet.getLastRow());
}

/**
 * @param {number} row
 * @param {GoogleAppsScript.Spreadsheet.Sheet=} sheet
 * @return {boolean}
 */
function isInputDataRow_(row, sheet) {
  return row >= CONFIG.input.dataStartRow && row <= getInputDataEndRow_(sheet);
}

/**
 * 作業リストから、指定大項目に属する中項目を重複なし・出現順で返す。
 * セレクトボックスの候補専用。シートへは書き出さない。
 *
 * @param {object[]} workRows
 * @param {string} major
 * @return {string[]}
 */
function getUniqueMidsByMajor_(workRows, major) {
  const matched = workRows.filter(function (row) {
    return row.major === major;
  });
  const mids = uniqueValues_(matched.map(function (row) {
    return row.mid;
  }));

  Logger.log(
    '%s getUniqueMidsByMajor_: 大項目「%s」 レコード数=%s 中項目(重複なし %s件)=%s',
    CONFIG.logPrefix,
    major,
    matched.length,
    mids.length,
    mids.join(' | ')
  );

  return mids;
}

/**
 * 大項目選択時：同じ行の中項目セルに、重複排除したプルダウンを付ける。
 * 大項目が変わったときは、直下の作業内容展開も消す。
 *
 * @param {object} ctx
 * @param {number} targetRow
 * @param {string} major
 * @param {boolean} resetMidValue true なら中項目の値を消してから付け直す
 */
function applyMidDropdownForRow_(ctx, targetRow, major, resetMidValue) {
  Logger.log(
    '%s applyMidDropdownForRow_: 行=%s 大項目=%s resetMid=%s',
    CONFIG.logPrefix,
    targetRow,
    major,
    resetMidValue
  );

  const midRange = ctx.inputSheet.getRange(targetRow, ctx.inputCols.mid);

  if (!major) {
    Logger.log('%s 大項目が空のため、%s行目の中項目プルダウンと下行の作業内容を外します', CONFIG.logPrefix, targetRow);
    writeInternal_(function () {
      if (resetMidValue) {
        midRange.clearContent();
        setFeeOnRow_(ctx, targetRow, '');
      }
      midRange.clearDataValidations();
    });
    if (resetMidValue) {
      replaceExpandedDetailsBelow_(ctx, targetRow, []);
    }
    return;
  }

  const mids = getUniqueMidsByMajor_(ctx.workRows, major);

  writeInternal_(function () {
    setDropdown_(ctx.inputSheet, ctx.inputCols.mid, targetRow, mids);

    if (resetMidValue) {
      midRange.clearContent();
      setFeeOnRow_(ctx, targetRow, '');
    }

    if (mids.length === 1 && (resetMidValue || !normalize_(midRange.getValue()))) {
      Logger.log('%s 中項目が 1 件のため「%s」を %s行目に自動セットします', CONFIG.logPrefix, mids[0], targetRow);
      midRange.setValue(mids[0]);
    }
  });

  if (resetMidValue) {
    if (mids.length === 1) {
      applyMidSelectionForRow_(ctx, targetRow, major, mids[0]);
    } else {
      replaceExpandedDetailsBelow_(ctx, targetRow, []);
    }
  }
}

/**
 * 互換用：既定行の中項目プルダウンだけ更新する。
 *
 * @param {object} ctx
 * @param {string} major
 * @param {boolean} resetMidValue
 */
function applyMajorSelection_(ctx, major, resetMidValue) {
  applyMidDropdownForRow_(ctx, CONFIG.input.selectRow, major, resetMidValue);
}

/**
 * 中項目選択時：
 * - 作業内容が空でない行 → 選択行のすぐ下へ縦に出す（C=作業内容, D=技術料）
 * - 作業内容が空の行 → その技術料を中項目行の D 列へ載せる
 * 次の大項目行は上書きせず、明細行の挿入・削除で共存する。
 *
 * @param {object} ctx
 * @param {number} selectRow
 * @param {string} major
 * @param {string} mid
 */
function applyMidSelectionForRow_(ctx, selectRow, major, mid) {
  Logger.log('%s applyMidSelectionForRow_: 行=%s 大項目=%s 中項目=%s', CONFIG.logPrefix, selectRow, major, mid);

  if (!major || !mid) {
    Logger.log('%s 大項目または中項目が空のため、技術料と下行の作業内容をクリアします', CONFIG.logPrefix);
    writeInternal_(function () {
      setFeeOnRow_(ctx, selectRow, '');
    });
    replaceExpandedDetailsBelow_(ctx, selectRow, []);
    return;
  }

  const records = ctx.workRows.filter(function (row) {
    return row.major === major && row.mid === mid;
  });
  const sorted = sortByOrder_(records);

  Logger.log(
    '%s 抽出件数=%s / 順番=%s',
    CONFIG.logPrefix,
    sorted.length,
    sorted.map(function (row) {
      return row.order === '' ? '(行' + row.sourceIndex + ')' : row.order;
    }).join(', ')
  );

  const blankContentRows = sorted.filter(function (row) {
    return !hasWorkContent_(row);
  });
  const contentRows = sorted.filter(function (row) {
    return hasWorkContent_(row);
  });
  const midFee = pickMidFee_(blankContentRows);

  Logger.log(
    '%s 作業内容あり=%s / 作業内容なし(中項目技術料用)=%s / 中項目技術料=%s',
    CONFIG.logPrefix,
    contentRows.length,
    blankContentRows.length,
    midFee
  );

  writeInternal_(function () {
    setFeeOnRow_(ctx, selectRow, midFee);
  });
  replaceExpandedDetailsBelow_(ctx, selectRow, contentRows);
}

/**
 * 作業内容が空のマスタ行から、中項目に載せる技術料を 1 つ取る（順番が先で、値が入っているもの）。
 *
 * @param {object[]} blankContentRows
 * @return {*}
 */
function pickMidFee_(blankContentRows) {
  for (let i = 0; i < blankContentRows.length; i++) {
    if (isFilled_(blankContentRows[i].fee)) {
      return blankContentRows[i].fee;
    }
  }
  return '';
}

/**
 * 互換用（既定行）。
 *
 * @param {object} ctx
 * @param {string} major
 * @param {string} mid
 */
function applyMidSelection_(ctx, major, mid) {
  applyMidSelectionForRow_(ctx, CONFIG.input.selectRow, major, mid);
}
