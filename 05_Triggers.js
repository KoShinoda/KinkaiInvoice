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
 * セル編集時。B9（大項目）または C9（中項目）の変更だけ処理する。
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
 * 初回セットアップ：B9 に大項目プルダウンを付ける。
 */
function setupInputDropdowns() {
  const startedAt = Date.now();
  Logger.log('%s setupInputDropdowns: 開始', CONFIG.logPrefix);

  const ctx = loadContext_();
  const majors = uniqueValues_(ctx.workRows.map(function (row) {
    return row.major;
  }));

  Logger.log('%s 大項目の候補数=%s / 内容=%s', CONFIG.logPrefix, majors.length, majors.join(' | '));

  setDropdown_(ctx.inputSheet, ctx.inputCols.major, CONFIG.input.selectRow, majors);

  const currentMajor = normalize_(ctx.inputSheet.getRange(CONFIG.input.selectRow, ctx.inputCols.major).getValue());
  if (currentMajor) {
    Logger.log('%s 既存の大項目「%s」があるため中項目プルダウンも更新します', CONFIG.logPrefix, currentMajor);
    applyMajorSelection_(ctx, currentMajor, false);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('大項目プルダウンを設定しました', '請求書入力', 5);
  Logger.log('%s setupInputDropdowns: 完了 (%sms)', CONFIG.logPrefix, Date.now() - startedAt);
}

/**
 * 今の B9/C9 の値から、C10 以降を作り直す（手動再実行用）。
 */
function refreshOutputFromSelection() {
  Logger.log('%s refreshOutputFromSelection: 開始', CONFIG.logPrefix);
  const ctx = loadContext_();
  const major = normalize_(ctx.inputSheet.getRange(CONFIG.input.selectRow, ctx.inputCols.major).getValue());
  const mid = normalize_(ctx.inputSheet.getRange(CONFIG.input.selectRow, ctx.inputCols.mid).getValue());

  if (!major) {
    Logger.log('%s 大項目が空のため出力をクリアして終了します', CONFIG.logPrefix);
    clearOutputArea_(ctx);
    clearDropdown_(ctx.inputSheet, ctx.inputCols.mid, CONFIG.input.selectRow);
    return;
  }

  applyMajorSelection_(ctx, major, false);
  if (mid) {
    applyMidSelection_(ctx, major, mid);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('再出力しました', '請求書入力', 5);
  Logger.log('%s refreshOutputFromSelection: 完了 major=%s mid=%s', CONFIG.logPrefix, major, mid);
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
  const selectRow = CONFIG.input.selectRow;
  const touchedSelectRow = startRow <= selectRow && selectRow < startRow + numRows;

  if (!touchedSelectRow) {
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

  const major = normalize_(sheet.getRange(selectRow, majorCol).getValue());
  const mid = normalize_(sheet.getRange(selectRow, midCol).getValue());

  Logger.log(
    '%s handleEdit_: シート=%s 範囲=%s 大項目=%s 中項目=%s touchedMajor=%s touchedMid=%s',
    CONFIG.logPrefix,
    sheet.getName(),
    e.range.getA1Notation(),
    major,
    mid,
    touchedMajor,
    touchedMid
  );

  if (touchedMajor) {
    applyMajorSelection_(ctx, major, true);
    return;
  }

  if (touchedMid) {
    applyMidSelection_(ctx, major, mid);
  }
}
