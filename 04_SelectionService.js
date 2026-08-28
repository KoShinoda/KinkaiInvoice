/**
 * 選択ロジック。
 * 大項目 → 中項目候補、中項目 → 明細行、という「何を出すか」だけを決める。
 */

/**
 * 大項目変更時：
 * - C9 の中項目プルダウンを作り直す（重複なし、出現順）
 * - 中項目の値はリセット（大項目が変われば前の中項目は無効なため）
 * - 中項目が 1 件だけのときは自動選択し、続けて明細出力まで行う
 *
 * @param {object} ctx
 * @param {string} major
 * @param {boolean} resetMidValue true なら C9 の値を消す（ユーザーが大項目を変えたとき）
 */
function applyMajorSelection_(ctx, major, resetMidValue) {
  Logger.log('%s applyMajorSelection_: major=%s resetMid=%s', CONFIG.logPrefix, major, resetMidValue);

  clearOutputArea_(ctx);

  if (!major) {
    Logger.log('%s 大項目が空です。中項目プルダウンと出力をクリアします', CONFIG.logPrefix);
    if (resetMidValue) {
      writeInternal_(function () {
        ctx.inputSheet.getRange(CONFIG.input.selectRow, ctx.inputCols.mid).clearContent();
      });
    }
    clearDropdown_(ctx.inputSheet, ctx.inputCols.mid, CONFIG.input.selectRow);
    return;
  }

  const matched = ctx.workRows.filter(function (row) {
    return row.major === major;
  });
  const mids = uniqueValues_(matched.map(function (row) {
    return row.mid;
  }));

  Logger.log(
    '%s 大項目「%s」に紐づくレコード数=%s / 中項目(重複なし)=%s',
    CONFIG.logPrefix,
    major,
    matched.length,
    mids.join(' | ')
  );

  writeInternal_(function () {
    setDropdown_(ctx.inputSheet, ctx.inputCols.mid, CONFIG.input.selectRow, mids);

    if (resetMidValue) {
      ctx.inputSheet.getRange(CONFIG.input.selectRow, ctx.inputCols.mid).clearContent();
    }

    if (mids.length === 1) {
      Logger.log('%s 中項目が 1 件のため「%s」を自動セットします', CONFIG.logPrefix, mids[0]);
      ctx.inputSheet.getRange(CONFIG.input.selectRow, ctx.inputCols.mid).setValue(mids[0]);
    }
  });

  if (mids.length === 1) {
    applyMidSelection_(ctx, major, mids[0]);
  }
}

/**
 * 中項目変更時：
 * - 該当行を順番（なければ行順）で並べる
 * - 作業内容が1件でもあれば、作業内容だけを縦出力する（中項目名は出さない）
 * - 作業内容がすべて空なら、中項目名を明細として出す（中項目に技術料を載せる用途）
 *
 * @param {object} ctx
 * @param {string} major
 * @param {string} mid
 */
function applyMidSelection_(ctx, major, mid) {
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
 * 中項目に値段を付けたい場合：
 *   作業内容が空の行 → 明細名は中項目。技術料・部品はその行の値。
 * 中項目と同じ文言も明細にしたい場合：
 *   作業内容列に中項目と同じ文字列の行を追加する。
 *   作業内容がある時点で中項目名は自動出力しない。
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
