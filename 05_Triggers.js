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
    .addItem('入力アプリを開く', 'openInputApp')
    .addSeparator()
    .addItem('初期設定（候補マスタとプルダウン）', 'setupInputDropdowns')
    .addItem('整備情報シートを整理', 'tidyServiceInfoSheet')
    .addItem('明細テンプレート（サンプル）を用意', 'ensureInvoiceTemplateSheet')
    .addItem('選択内容を再出力', 'refreshOutputFromSelection')
    .addItem('リストを更新（順番・選択肢）', 'refreshAllMasterLists')
    .addSeparator()
    .addItem('作業リストの列マップをログ出力', 'logWorkListColumnMap')
    .addItem('編集トリガーを作成', 'createInstallableOnEditTrigger')
    .addToUi();

  try {
    ensureListMasterSheets_();
  } catch (err) {
    Logger.log('%s onOpen: リスト列の準備に失敗: %s', CONFIG.logPrefix, err);
  }

  Logger.log('%s onOpen: メニューを追加しました', CONFIG.logPrefix);
}

/**
 * セル編集時。
 * 作業リスト変更 → 中項目候補を再生成。
 * 大項目変更 → 入力規則は触らず、値のクリア／自動セットのみ。
 * 中項目変更 → 作業内容を下行へ上書き。
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
 * 初回セットアップ：中項目候補マスタを作り、B/C の入力規則を範囲参照にする。
 */
function rebuildMidCandidates() {
  rebuildMidCandidateSheet_();
}

function setupInputDropdowns() {
  const startedAt = Date.now();
  Logger.log('%s setupInputDropdowns: 開始', CONFIG.logPrefix);
  try {
    ensureListMasterSheets_();
  } catch (err) {
    Logger.log('%s setupInputDropdowns: リスト列の準備に失敗: %s', CONFIG.logPrefix, err);
  }
  rebuildMidCandidateSheet_();
  SpreadsheetApp.getActiveSpreadsheet().toast('中項目候補とプルダウンを設定しました', '請求書入力', 5);
  Logger.log('%s setupInputDropdowns: 完了 (%sms)', CONFIG.logPrefix, Date.now() - startedAt);
}

/**
 * 候補マスタを作り直し、入力規則を張り直す。
 */
function refreshOutputFromSelection() {
  rebuildMidCandidateSheet_();
  SpreadsheetApp.getActiveSpreadsheet().toast('中項目候補を再生成しました', '請求書入力', 5);
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
  const sheetName = sheet.getName();

  if (sheetName === CONFIG.workers.sheetName) {
    writeInternal_(function () {
      assignMissingWorkerCodes_(sheet);
    });
    return;
  }

  if (sheetName === CONFIG.workList.sheetName) {
    if (e.range.getColumn() <= 2) {
      rebuildMidCandidateSheet_();
    }
    return;
  }

  if (sheetName !== CONFIG.input.sheetName) {
    return;
  }

  const startRow = e.range.getRow();
  const numRows = e.range.getNumRows();
  const startCol = e.range.getColumn();
  const numCols = e.range.getNumColumns();
  const firstRow = Math.max(startRow, CONFIG.input.dataStartRow);
  const lastRow = Math.min(startRow + numRows - 1, getInputDataEndRow_());

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

  log_(
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
    const height = lastRow - firstRow + 1;
    const majors = height === 1 && e.value !== undefined
      ? [[e.value]]
      : sheet.getRange(firstRow, majorCol, height, 1).getValues();
    for (let i = 0; i < majors.length; i++) {
      applyMajorChangeForRow_(ctx, firstRow + i, normalize_(majors[i][0]));
    }
    return;
  }

  if (touchedMid) {
    const height = lastRow - firstRow + 1;
    const pairs = sheet.getRange(firstRow, Math.min(majorCol, midCol), height, Math.abs(majorCol - midCol) + 1).getValues();
    const majorOff = majorCol - Math.min(majorCol, midCol);
    const midOff = midCol - Math.min(majorCol, midCol);
    for (let i = 0; i < pairs.length; i++) {
      const major = normalize_(pairs[i][majorOff]);
      const mid = normalize_(pairs[i][midOff]);
      applyMidSelectionForRow_(ctx, firstRow + i, major, mid);
    }
  }
}
