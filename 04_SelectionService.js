/**
 * 選択ロジック。
 * 中項目プルダウンはシート数式側。ここでは値のクリアと明細展開だけ行う。
 */

function getInputDataEndRow_() {
  return CONFIG.input.dataStartRow + CONFIG.input.maxSelectRows - 1;
}

function isInputDataRow_(row) {
  return row >= CONFIG.input.dataStartRow && row <= getInputDataEndRow_();
}

function getUniqueMidsByMajor_(ctx, major) {
  return ctx.midsByMajor[major] || [];
}

/**
 * 大項目が変わったとき。入力規則は触らない（数式が候補を切り替える）。
 * 空にしたときは中項目を残し、全中項目から選び直せるようにする。
 *
 * @param {object} ctx
 * @param {number} targetRow
 * @param {string} major
 */
function applyMajorChangeForRow_(ctx, targetRow, major) {
  const midRange = ctx.inputSheet.getRange(targetRow, ctx.inputCols.mid);
  const lookup = ctx.inputSheet.getParent().getSheetByName(CONFIG.lookup.sheetName);
  if (lookup) {
    bindMidValidationForRow_(ctx.inputSheet, lookup, targetRow, ctx.inputCols.mid);
  }

  if (!major) {
    return;
  }

  const mids = getUniqueMidsByMajor_(ctx, major);
  if (mids.length === 1) {
    writeInternal_(function () {
      midRange.setValue(mids[0]);
    });
    applyMidSelectionForRow_(ctx, targetRow, major, mids[0]);
    return;
  }

  writeInternal_(function () {
    midRange.clearContent();
  });
  writeMidFeeAndDetails_(ctx, targetRow, '', []);
}

function applyMidDropdownForRow_(ctx, targetRow, major, resetMidValue) {
  if (resetMidValue) {
    applyMajorChangeForRow_(ctx, targetRow, major);
  }
}

function applyMajorSelection_(ctx, major, resetMidValue) {
  if (resetMidValue) {
    applyMajorChangeForRow_(ctx, CONFIG.input.selectRow, major);
  }
}

/**
 * @param {object} ctx
 * @param {string} major
 * @param {string} mid
 * @return {object[]}
 */
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
 * 中項目選択時。大項目が空でも、選んだ中項目の作業内容を展開する。
 *
 * @param {object} ctx
 * @param {number} selectRow
 * @param {string} major
 * @param {string} mid
 */
function applyMidSelectionForRow_(ctx, selectRow, major, mid) {
  if (!mid) {
    writeMidFeeAndDetails_(ctx, selectRow, '', []);
    return;
  }
  const resolved = resolveMidOutput_(ctx, major, mid);
  writeMidFeeAndDetails_(ctx, selectRow, resolved.midFee, resolved.workRows);
}

/**
 * 中項目に対する技術料・作業内容。シートへは書かない（Web アプリからも使う）。
 *
 * @param {object} ctx
 * @param {string} major
 * @param {string} mid
 * @return {{midFee: *, workRows: object[]}}
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
    midFee: pickMidFee_(sorted),
    workRows: workRows
  };
}

function pickMidFee_(sortedRows) {
  const anchor = pickMidAnchorRow_(sortedRows);
  if (anchor && isFilled_(anchor.fee)) {
    return anchor.fee;
  }
  for (let i = 0; i < sortedRows.length; i++) {
    if (!hasWorkContent_(sortedRows[i]) && isFilled_(sortedRows[i].fee)) {
      return sortedRows[i].fee;
    }
  }
  for (let i = 0; i < sortedRows.length; i++) {
    if (isFilled_(sortedRows[i].fee)) {
      return sortedRows[i].fee;
    }
  }
  return '';
}

function applyMidSelection_(ctx, major, mid) {
  applyMidSelectionForRow_(ctx, CONFIG.input.selectRow, major, mid);
}
