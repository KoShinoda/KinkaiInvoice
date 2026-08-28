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
 * 空を除いた重複なし配列。先に出た順を維持する（プルダウン＝作業リストの出現順）。
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
