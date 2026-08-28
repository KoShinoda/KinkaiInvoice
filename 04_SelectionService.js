/**
 * 選択ロジック。
 * 大項目 → 中項目プルダウン。中項目 → 作業内容を下行へ上書き。
 */

function getInputDataEndRow_() {
  return CONFIG.input.dataStartRow + CONFIG.input.maxSelectRows - 1;
}

function isInputDataRow_(row) {
  return row >= CONFIG.input.dataStartRow && row <= getInputDataEndRow_();
}

/**
 * @param {object} ctx
 * @param {string} major
 * @return {string[]}
 */
function getUniqueMidsByMajor_(ctx, major) {
  return ctx.midsByMajor[major] || [];
}

/**
 * 大項目選択時：同じ行の中項目プルダウンを付け替える。
 *
 * @param {object} ctx
 * @param {number} targetRow
 * @param {string} major
 * @param {boolean} resetMidValue
 */
function applyMidDropdownForRow_(ctx, targetRow, major, resetMidValue) {
  const midRange = ctx.inputSheet.getRange(targetRow, ctx.inputCols.mid);

  if (!major) {
    writeInternal_(function () {
      if (resetMidValue) {
        midRange.clearContent();
      }
      midRange.clearDataValidations();
    });
    if (resetMidValue) {
      writeMidFeeAndDetails_(ctx, targetRow, '', []);
    }
    return;
  }

  const mids = getUniqueMidsByMajor_(ctx, major);

  writeInternal_(function () {
    setDropdown_(ctx.inputSheet, ctx.inputCols.mid, targetRow, mids);
    if (resetMidValue) {
      if (mids.length === 1) {
        midRange.setValue(mids[0]);
      } else {
        midRange.clearContent();
      }
    }
  });

  if (resetMidValue) {
    if (mids.length === 1) {
      applyMidSelectionForRow_(ctx, targetRow, major, mids[0]);
    } else {
      writeMidFeeAndDetails_(ctx, targetRow, '', []);
    }
  }
}

function applyMajorSelection_(ctx, major, resetMidValue) {
  applyMidDropdownForRow_(ctx, CONFIG.input.selectRow, major, resetMidValue);
}

/**
 * 中項目選択時。作業リストは事前インデックスから取る（全件スキャンしない）。
 *
 * @param {object} ctx
 * @param {number} selectRow
 * @param {string} major
 * @param {string} mid
 */
function applyMidSelectionForRow_(ctx, selectRow, major, mid) {
  if (!major || !mid) {
    writeMidFeeAndDetails_(ctx, selectRow, '', []);
    return;
  }

  const records = ctx.recordsByMajorMid[major + '\t' + mid] || [];
  const sorted = sortByOrder_(records);
  const blankContentRows = [];
  const contentRows = [];
  for (let i = 0; i < sorted.length; i++) {
    if (hasWorkContent_(sorted[i])) {
      contentRows.push(sorted[i]);
    } else {
      blankContentRows.push(sorted[i]);
    }
  }

  writeMidFeeAndDetails_(ctx, selectRow, pickMidFee_(blankContentRows), contentRows);
}

function pickMidFee_(blankContentRows) {
  for (let i = 0; i < blankContentRows.length; i++) {
    if (isFilled_(blankContentRows[i].fee)) {
      return blankContentRows[i].fee;
    }
  }
  return '';
}

function applyMidSelection_(ctx, major, mid) {
  applyMidSelectionForRow_(ctx, CONFIG.input.selectRow, major, mid);
}
