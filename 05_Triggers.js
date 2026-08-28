/**
 * ユーザー操作の入口。
 * シート編集・メニュー・トリガー作成だけを置き、抽出ルールは SelectionService に任せる。
 */

/**
 * スプレッドシートを開いたときにメニューを出す。
 * 簡易トリガーなので、GAS に保存するだけで有効。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('請求書入力')
    .addItem('初期設定（プルダウン作成）', 'setupInputDropdowns')
    .addItem('選択内容を再出力', 'refreshOutputFromSelection')
    .addSeparator()
    .addItem('作業リストの列マップをログ出力', 'logWorkListColumnMap')
    .addItem('編集トリガーを作成', 'createInstallableOnEditTrigger')
    .addToUi();

  Logger.log('%s onOpen: メニューを追加しました', CONFIG.logPrefix);
}

/**
 * セル編集時。入力範囲内の大項目列なら、その行の中項目プルダウンを作り直す。
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEdit(e) {
  handleEdit_(e);
}

/**
 * インストール型 onEdit 用（メニューから作成）。中身は onEdit と同じ。
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEditInstallable(e) {
  handleEdit_(e);
}

/**
 * 初回セットアップ：
 * - 大項目列（B9:B208 など）に、重複なし大項目のプルダウンを一度で付ける
 * - すでに大項目が入っている行だけ、中項目プルダウンを付け直す
 */
function setupInputDropdowns() {
  const startedAt = Date.now();
  Logger.log('%s setupInputDropdowns: 開始', CONFIG.logPrefix);

  const ctx = loadContext_();
  const majors = uniqueValues_(ctx.workRows.map(function (row) {
    return row.major;
  }));

  Logger.log('%s 大項目の候補数=%s / 内容=%s', CONFIG.logPrefix, majors.length, majors.join(' | '));

  const start = CONFIG.input.dataStartRow;
  const num = CONFIG.input.maxSelectRows;
  const majorRange = ctx.inputSheet.getRange(start, ctx.inputCols.major, num, 1);
  setDropdownOnRange_(majorRange, majors);

  const majorValues = majorRange.getValues();
  for (let i = 0; i < majorValues.length; i++) {
    const major = normalize_(majorValues[i][0]);
    if (!major) {
      continue;
    }
    applyMidDropdownForRow_(ctx, start + i, major, false);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('大項目・中項目プルダウンを設定しました', '請求書入力', 5);
  Logger.log('%s setupInputDropdowns: 完了 (%sms)', CONFIG.logPrefix, Date.now() - startedAt);
}

/**
 * 入力範囲の大項目に合わせて、各行の中項目プルダウンを作り直す。
 */
function refreshOutputFromSelection() {
  Logger.log('%s refreshOutputFromSelection: 開始', CONFIG.logPrefix);
  const ctx = loadContext_();
  const start = CONFIG.input.dataStartRow;
  const num = CONFIG.input.maxSelectRows;
  const majorValues = ctx.inputSheet.getRange(start, ctx.inputCols.major, num, 1).getValues();

  for (let i = 0; i < majorValues.length; i++) {
    const major = normalize_(majorValues[i][0]);
    applyMidDropdownForRow_(ctx, start + i, major, false);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('中項目プルダウンを再設定しました', '請求書入力', 5);
  Logger.log('%s refreshOutputFromSelection: 完了', CONFIG.logPrefix);
}

/**
 * デバッグ用：ヘッダー → 列番号をログに出す。
 */
function logWorkListColumnMap() {
  const ctx = loadContext_();
  Logger.log('%s 作業リスト列マップ: %s', CONFIG.logPrefix, JSON.stringify(ctx.workCols));
  Logger.log('%s 入力シート列マップ: %s', CONFIG.logPrefix, JSON.stringify(ctx.inputCols));
  Logger.log('%s 作業リスト件数=%s', CONFIG.logPrefix, ctx.workRows.length);
}

/**
 * インストール型 onEdit を 1 本だけ作る。
 */
function createInstallableOnEditTrigger() {
  const handlerName = 'onEditInstallable';
  const existing = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === handlerName;
  });

  if (existing.length > 0) {
    Logger.log('%s 編集トリガーは既に %s 本あります。追加しません', CONFIG.logPrefix, existing.length);
    SpreadsheetApp.getUi().alert('編集トリガーは既に作成済みです。');
    return;
  }

  ScriptApp.newTrigger(handlerName)
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  Logger.log('%s インストール型 onEdit トリガーを作成しました（%s）', CONFIG.logPrefix, handlerName);
  SpreadsheetApp.getUi().alert('編集トリガーを作成しました。');
}

/**
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function handleEdit_(e) {
  if (!e || !e.range) {
    Logger.log('%s handleEdit_: イベント情報が無いため終了', CONFIG.logPrefix);
    return;
  }

  if (isInternalWrite_()) {
    Logger.log('%s handleEdit_: 内部書き込み中のためスキップ', CONFIG.logPrefix);
    return;
  }

  const sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.input.sheetName) {
    return;
  }

  const startRow = e.range.getRow();
  const numRows = e.range.getNumRows();
  const startCol = e.range.getColumn();
  const numCols = e.range.getNumColumns();
  const firstRow = Math.max(startRow, CONFIG.input.dataStartRow);
  const lastRow = Math.min(startRow + numRows - 1, getInputDataEndRow_(sheet));

  if (firstRow > lastRow) {
    return;
  }

  let ctx;
  try {
    ctx = loadContext_();
  } catch (err) {
    Logger.log('%s handleEdit_: コンテキスト読み込み失敗: %s', CONFIG.logPrefix, err);
    return;
  }

  const majorCol = ctx.inputCols.major;
  const midCol = ctx.inputCols.mid;
  const touchedMajor = columnOverlaps_(startCol, numCols, majorCol);
  const touchedMid = columnOverlaps_(startCol, numCols, midCol);

  if (!touchedMajor && !touchedMid) {
    return;
  }

  Logger.log(
    '%s handleEdit_: シート=%s 範囲=%s 行=%s〜%s touchedMajor=%s touchedMid=%s',
    CONFIG.logPrefix,
    sheet.getName(),
    e.range.getA1Notation(),
    firstRow,
    lastRow,
    touchedMajor,
    touchedMid
  );

  if (touchedMajor) {
    for (let row = firstRow; row <= lastRow; row++) {
      const major = normalize_(sheet.getRange(row, majorCol).getValue());
      applyMidDropdownForRow_(ctx, row, major, true);
    }
    return;
  }

  if (touchedMid) {
    for (let row = firstRow; row <= lastRow; row++) {
      const major = normalize_(sheet.getRange(row, majorCol).getValue());
      if (!major) {
        Logger.log('%s %s行目は大項目が空のため作業内容展開をスキップ（明細行の編集とみなす）', CONFIG.logPrefix, row);
        continue;
      }
      const mid = normalize_(sheet.getRange(row, midCol).getValue());
      applyMidSelectionForRow_(ctx, row, major, mid);
    }
  }
}
