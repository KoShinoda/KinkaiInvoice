/**
 * 選択ロジック。
 * 大項目 → その行の中項目セレクトボックス（重複なし）。
 * 中項目 → 明細展開（CONFIG.expandDetailsOnMidSelect が true のときだけ）。
 */

/**
 * 入力シートで大項目・中項目を置く最終行（dataStartRow から maxSelectRows 行）。
 *
 * @return {number}
 */
function getInputDataEndRow_() {
  return CONFIG.input.dataStartRow + CONFIG.input.maxSelectRows - 1;
}

/**
 * @param {number} row
 * @return {boolean}
 */
function isInputDataRow_(row) {
  return row >= CONFIG.input.dataStartRow && row <= getInputDataEndRow_();
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
 *
 * - 候補は入力規則（Validation）だけに持つ。作業リストや補助列は増えない。
 * - 大項目が変わったら中項目の値は消す（前の大項目の中項目が残らないようにする）。
 * - 中項目が 1 件だけのときは自動セットする。
 *
 * @param {object} ctx
 * @param {number} targetRow 入力シートの行番号（B9 なら 9）
 * @param {string} major
 * @param {boolean} resetMidValue true なら中項目セルの値を消してから付け直す
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
    Logger.log('%s 大項目が空のため、%s行目の中項目プルダウンを外します', CONFIG.logPrefix, targetRow);
    writeInternal_(function () {
      if (resetMidValue) {
        midRange.clearContent();
      }
      midRange.clearDataValidations();
    });
    return;
  }

  const mids = getUniqueMidsByMajor_(ctx.workRows, major);

  writeInternal_(function () {
    setDropdown_(ctx.inputSheet, ctx.inputCols.mid, targetRow, mids);

    if (resetMidValue) {
      midRange.clearContent();
    }

    if (mids.length === 1) {
      Logger.log('%s 中項目が 1 件のため「%s」を %s行目に自動セットします', CONFIG.logPrefix, mids[0], targetRow);
      midRange.setValue(mids[0]);
    }
  });
}

/**
 * 互換用：既定行（selectRow）の中項目プルダウンだけ更新する。
 * 明細の一括クリアはしない（入力行の選択セルを消さないため）。
 *
 * @param {object} ctx
 * @param {string} major
 * @param {boolean} resetMidValue
 */
function applyMajorSelection_(ctx, major, resetMidValue) {
  applyMidDropdownForRow_(ctx, CONFIG.input.selectRow, major, resetMidValue);
}

/**
 * 中項目変更時の明細展開。expandDetailsOnMidSelect が false なら何もしない。
 *
 * @param {object} ctx
 * @param {string} major
 * @param {string} mid
 */
function applyMidSelection_(ctx, major, mid) {
  if (!CONFIG.input.expandDetailsOnMidSelect) {
    Logger.log('%s applyMidSelection_: 明細展開はオフのためスキップ（中項目プルダウンのみ）', CONFIG.logPrefix);
    return;
  }

  Logger.log('%s applyMidSelection_: major=%s mid=%s', CONFIG.logPrefix, major, mid);

  clearOutputArea_(ctx);

  if (!major || !mid) {
    Logger.log('%s 大項目または中項目が空のため出力しません', CONFIG.logPrefix);
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

  if (sorted.length === 0) {
    Logger.log('%s 一致する作業リスト行がありません', CONFIG.logPrefix);
    return;
  }

  const outputRows = resolveDetailRows_(sorted);
  const lines = outputRows.map(function (row) {
    return buildOutputLineFromRow_(ctx, row);
  });
  writeOutputRows_(ctx, lines);
}

/**
 * 作業内容の有無から、実際に書く明細行を決める。
 *
 * @param {object[]} sorted
 * @return {object[]} displayName を付けた行
 */
function resolveDetailRows_(sorted) {
  const withContent = sorted.filter(function (row) {
    return hasWorkContent_(row);
  });

  if (withContent.length > 0) {
    Logger.log(
      '%s 作業内容あり %s 件。中項目名は出さず、作業内容だけを縦に並べます',
      CONFIG.logPrefix,
      withContent.length
    );
    return withContent.map(function (row) {
      return Object.assign({}, row, { displayName: row.content });
    });
  }

  Logger.log(
    '%s 作業内容なし %s 件。中項目名をそのまま出力します',
    CONFIG.logPrefix,
    sorted.length
  );
  return sorted.map(function (row) {
    return Object.assign({}, row, { displayName: row.mid });
  });
}
