/**
 * 共通ユーティリティ。
 * シートや選択ルールに依存しない処理だけを置く。
 */

/**
 * 比較・プルダウン用に文字列を揃える（全角スペースも半角に）。
 *
 * @param {*} value
 * @return {string}
 */
function normalize_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).replace(/\u3000/g, ' ').trim();
}

/**
 * 作業内容セルに中身があるか。空なら「中項目を明細名にする」判定に使う。
 *
 * @param {object} row
 * @return {boolean}
 */
function hasWorkContent_(row) {
  return isFilled_(row.content);
}

/**
 * 0 は「値あり」。空文字・null だけを空とみなす。
 *
 * @param {*} value
 * @return {boolean}
 */
function isFilled_(value) {
  if (value === 0 || value === '0') {
    return true;
  }
  return normalize_(value) !== '';
}

/**
 * 空を除いた重複なし配列。先に出た順を維持する（プルダウン＝順番ソート後の出現順）。
 *
 * @param {*[]} values
 * @return {string[]}
 */
function uniqueValues_(values) {
  const seen = {};
  const result = [];
  values.forEach(function (v) {
    const key = normalize_(v);
    if (!key || seen[key]) {
      return;
    }
    seen[key] = true;
    result.push(key);
  });
  return result;
}

/**
 * 1 始まり列番号でセル値を取る。列が未定義なら空文字。
 *
 * @param {*[]} raw
 * @param {number=} col1
 * @return {*}
 */
function cell_(raw, col1) {
  if (!col1) {
    return '';
  }
  const v = raw[col1 - 1];
  return v === undefined || v === null ? '' : v;
}

/**
 * 順番昇順。空・非数値は最後、同順はシート上の行番号で安定させる。
 * 作業者リストなど、大項目／中項目で固めない一覧用。
 *
 * @param {object[]} records
 * @return {object[]}
 */
function sortByOrder_(records) {
  return records.slice().sort(function (a, b) {
    const oa = toOrderNumber_(a.order);
    const ob = toOrderNumber_(b.order);
    if (oa !== ob) {
      return oa - ob;
    }
    return a.sourceIndex - b.sourceIndex;
  });
}

function hasPartFields_(row) {
  if (!row) {
    return false;
  }
  return isFilled_(row.partMid) || isFilled_(row.partMajor) || isFilled_(row.qty) || isFilled_(row.unitPrice);
}

/**
 * 中項目ブロックの先頭 1 行。作業内容が空の行を優先。順番 1 の空行があればそれ。
 * 部品セットの追加行（作業内容空）は先頭にしない。
 *
 * @param {object[]} group
 * @return {object|null}
 */
function pickMidAnchorRow_(group) {
  if (!group || !group.length) {
    return null;
  }
  const blanks = group.filter(function (r) {
    return !hasWorkContent_(r);
  });
  const ones = blanks.filter(function (r) {
    return toOrderNumber_(r.order) === 1;
  });
  if (ones.length) {
    ones.sort(function (a, b) {
      return (a.sourceIndex || 0) - (b.sourceIndex || 0);
    });
    return ones[0];
  }
  if (blanks.length) {
    blanks.sort(function (a, b) {
      return (a.sourceIndex || 0) - (b.sourceIndex || 0);
    });
    return blanks[0];
  }
  return group.slice().sort(function (a, b) {
    return (a.sourceIndex || 0) - (b.sourceIndex || 0);
  })[0];
}

/**
 * レーン 0=先頭, 1=作業内容, 2=追加部品セット。
 *
 * @param {object[]} records
 */
function tagMidGroups_(records) {
  const grouped = {};
  (records || []).forEach(function (row) {
    const key = normalize_(row.major) + '\t' + normalize_(row.mid);
    if (!grouped[key]) {
      grouped[key] = [];
    }
    grouped[key].push(row);
    row._isMidAnchor = false;
    row._midLane = 1;
  });
  Object.keys(grouped).forEach(function (key) {
    const g = grouped[key];
    const anchor = pickMidAnchorRow_(g);
    g.forEach(function (row) {
      row._isMidAnchor = !!(anchor && row === anchor);
      if (row._isMidAnchor) {
        row._midLane = 0;
      } else if (hasWorkContent_(row)) {
        row._midLane = 1;
      } else {
        row._midLane = 2;
      }
    });
  });
}

/**
 * 空の順番だけ埋める（レコード上）。シートへは書かない。
 * 更新忘れでも入力アプリの並びを更新後と同じにする。
 *
 * @param {object[]} records
 * @return {boolean} 1件でも埋めたか
 */
function assignEmptyOrdersInGroups_(records) {
  if (!records || !records.length) {
    return false;
  }
  tagMidGroups_(records);
  const grouped = {};
  records.forEach(function (row) {
    const key = normalize_(row.major) + '\t' + normalize_(row.mid);
    if (!grouped[key]) {
      grouped[key] = [];
    }
    grouped[key].push(row);
  });
  let changed = false;
  Object.keys(grouped).forEach(function (key) {
    [0, 1, 2].forEach(function (lane) {
      const rows = grouped[key].filter(function (r) {
        return r._midLane === lane;
      });
      if (!rows.length) {
        return;
      }
      let maxN = 0;
      rows.forEach(function (rec) {
        const n = toOrderNumber_(rec.order);
        if (n !== Number.POSITIVE_INFINITY && n > maxN) {
          maxN = n;
        }
      });
      const empty = rows.filter(function (rec) {
        return toOrderNumber_(rec.order) === Number.POSITIVE_INFINITY;
      }).sort(function (a, b) {
        return (a.sourceIndex || 0) - (b.sourceIndex || 0);
      });
      if (!empty.length) {
        return;
      }
      changed = true;
      if (maxN === 0) {
        if (lane === 0) {
          empty[0].order = 1;
          return;
        }
        let n = 2;
        empty.forEach(function (rec) {
          rec.order = n;
          n += 1;
        });
        return;
      }
      empty.forEach(function (rec) {
        maxN += 1;
        rec.order = maxN;
      });
    });
  });
  return changed;
}

function rowsHaveEmptyOrder_(records) {
  if (!records) {
    return false;
  }
  for (let i = 0; i < records.length; i++) {
    if (toOrderNumber_(records[i].order) === Number.POSITIVE_INFINITY) {
      return true;
    }
  }
  return false;
}

/**
 * 同じ中項目内：先頭 → 作業内容（順番） → 追加部品（順番）。
 *
 * @param {object} a
 * @param {object} b
 * @return {number}
 */
function compareMidGroupRows_(a, b) {
  const la = a._midLane != null ? a._midLane : (isMidAnchorRow_(a) ? 0 : hasWorkContent_(a) ? 1 : 2);
  const lb = b._midLane != null ? b._midLane : (isMidAnchorRow_(b) ? 0 : hasWorkContent_(b) ? 1 : 2);
  if (la !== lb) {
    return la - lb;
  }
  const oa = toOrderNumber_(a.order);
  const ob = toOrderNumber_(b.order);
  if (oa !== ob) {
    return oa - ob;
  }
  return (a.sourceIndex || 0) - (b.sourceIndex || 0);
}

/**
 * 中項目ブロックの先頭行か（タグ付きならそれを使う）。
 *
 * @param {object} row
 * @return {boolean}
 */
function isMidAnchorRow_(row) {
  if (!row) {
    return false;
  }
  if (row._isMidAnchor === true) {
    return true;
  }
  if (row._isMidAnchor === false) {
    return false;
  }
  if (toOrderNumber_(row.order) === 1 && !hasWorkContent_(row)) {
    return true;
  }
  return !hasWorkContent_(row);
}

function groupFirstSeen_(records, keyFn) {
  const map = {};
  records.forEach(function (row) {
    const key = keyFn(row);
    const idx = row.sourceIndex || 0;
    if (map[key] === undefined || idx < map[key]) {
      map[key] = idx;
    }
  });
  return map;
}

/**
 * 大項目で固め、その中で中項目で固め、中項目内だけ順番する。
 * 中項目同士は先頭行の順番の昇順（空は後）。
 *
 * @param {object[]} records
 * @return {object[]}
 */
function sortWorkListRecords_(records) {
  tagMidGroups_(records);
  const majorOrder = groupFirstSeen_(records, function (r) {
    return normalize_(r.major);
  });
  const midKey = {};
  const grouped = {};
  records.forEach(function (row) {
    const key = normalize_(row.major) + '\t' + normalize_(row.mid);
    if (!grouped[key]) {
      grouped[key] = [];
    }
    grouped[key].push(row);
  });
  Object.keys(grouped).forEach(function (key) {
    const g = grouped[key];
    const tagged = g.filter(function (r) {
      return r._isMidAnchor;
    })[0];
    const anchor = tagged || pickMidAnchorRow_(g) || g[0];
    midKey[key] = {
      order: toOrderNumber_(anchor.order),
      sourceIndex: anchor.sourceIndex || 0
    };
  });
  return records.slice().sort(function (a, b) {
    const dMaj = (majorOrder[normalize_(a.major)] || 0) - (majorOrder[normalize_(b.major)] || 0);
    if (dMaj) {
      return dMaj;
    }
    const ka = normalize_(a.major) + '\t' + normalize_(a.mid);
    const kb = normalize_(b.major) + '\t' + normalize_(b.mid);
    const oa = midKey[ka] || { order: Number.POSITIVE_INFINITY, sourceIndex: 0 };
    const ob = midKey[kb] || { order: Number.POSITIVE_INFINITY, sourceIndex: 0 };
    if (oa.order !== ob.order) {
      return oa.order - ob.order;
    }
    if (oa.sourceIndex !== ob.sourceIndex) {
      return oa.sourceIndex - ob.sourceIndex;
    }
    return compareMidGroupRows_(a, b);
  });
}

/**
 * @param {*} value
 * @return {number}
 */
function toOrderNumber_(value) {
  if (value === '' || value === null || value === undefined) {
    return Number.POSITIVE_INFINITY;
  }
  if (typeof value === 'number' && isFinite(value)) {
    return value;
  }
  const n = Number(String(value).replace(/,/g, '').trim());
  return isFinite(n) ? n : Number.POSITIVE_INFINITY;
}

/**
 * 編集範囲が指定列を含むか（複数セル貼り付け対応）。
 *
 * @param {number} startCol
 * @param {number} numCols
 * @param {number} targetCol
 * @return {boolean}
 */
function toA1Col_(n) {
  let s = '';
  let num = n;
  while (num > 0) {
    const m = (num - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    num = Math.floor((num - 1) / 26);
  }
  return s;
}

function columnOverlaps_(startCol, numCols, targetCol) {
  return startCol <= targetCol && targetCol < startCol + numCols;
}

function log_() {
  if (CONFIG.verboseLog === false) {
    return;
  }
  Logger.log.apply(Logger, arguments);
}

/**
 * インストール型 onEdit が、スクリプト自身の書き込みで再発火するのを防ぐ。
 * 簡易 onEdit ではプログラムからの書き込みは再発火しないので、既定ではガードしない（高速化）。
 *
 * @param {Function} fn
 */
function writeInternal_(fn) {
  if (!CONFIG.useWriteGuard) {
    fn();
    return;
  }
  const cache = CacheService.getScriptCache();
  cache.put(CONFIG.internalWriteCacheKey, '1', 15);
  try {
    fn();
  } finally {
    cache.remove(CONFIG.internalWriteCacheKey);
  }
}

/**
 * いま内部書き込み中か（onEdit の再入判定）。
 *
 * @return {boolean}
 */
function isInternalWrite_() {
  if (!CONFIG.useWriteGuard) {
    return false;
  }
  return CacheService.getScriptCache().get(CONFIG.internalWriteCacheKey) === '1';
}
