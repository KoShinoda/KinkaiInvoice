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
    .addItem('請求書を検索', 'openInvoiceSearchApp')
    .addSeparator()
    .addItem('リストのプルダウンを設定', 'setupInputDropdowns')
    .addItem('整備情報シートを整理', 'tidyServiceInfoSheet')
    .addItem('明細テンプレート（サンプル）を用意', 'ensureInvoiceTemplateSheet')
    .addItem('請求書保存（サンプル3件）を用意', 'ensureInvoiceSaveSamples')
    .addItem('リストを更新（順番・選択肢）', 'refreshAllMasterLists')
    .addItem('印刷原本（A4・1シート）を作成', 'createPrintOriginalSample')
    .addSeparator()
    .addItem('作業リストの列マップをログ出力', 'logWorkListColumnMap')
    .addItem('編集トリガーを作成', 'createInstallableOnEditTrigger')
    .addToUi();

  Logger.log('%s onOpen: メニューを追加しました', CONFIG.logPrefix);
}

/**
 * セル編集時。
 * 作業リスト・部品・作業者のメンテと、印刷ヘッダーの部門連動。
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
 * 作業リスト／部品リストの入力規則を付ける。
 */
function setupInputDropdowns() {
  const startedAt = Date.now();
  Logger.log('%s setupInputDropdowns: 開始', CONFIG.logPrefix);
  try {
    ensureListMasterSheets_();
  } catch (err) {
    Logger.log('%s setupInputDropdowns: リスト列の準備に失敗: %s', CONFIG.logPrefix, err);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('リストのプルダウンを設定しました', '請求書入力', 5);
  Logger.log('%s setupInputDropdowns: 完了 (%sms)', CONFIG.logPrefix, Date.now() - startedAt);
}

/**
 * デバッグ用：ヘッダー → 列番号をログに出す。
 */
function logWorkListColumnMap() {
  const ctx = loadContext_();
  Logger.log('%s 作業リスト列マップ: %s', CONFIG.logPrefix, JSON.stringify(ctx.workCols));
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

  if (sheetName === CONFIG.invoiceSearch.sheetName) {
    handleInvoiceSearchEdit_(e);
    return;
  }

  if (sheetName === CONFIG.print.sheetName || sheetName === CONFIG.print.sampleSheetName) {
    handlePrintHeaderDeptEdit_(e);
    return;
  }

  if (sheetName === CONFIG.workers.sheetName) {
    writeInternal_(function () {
      assignMissingWorkerCodes_(sheet);
      const work = sheet.getParent().getSheetByName(CONFIG.workList.sheetName);
      if (work) {
        fillWorkListWorkerCodesFromNames_(work);
        applyWorkListOpenDropdowns_(sheet.getParent(), work);
      }
    });
    return;
  }

  if (sheetName === CONFIG.parts.sheetName) {
    writeInternal_(function () {
      const work = sheet.getParent().getSheetByName(CONFIG.workList.sheetName);
      if (work) {
        applyWorkListOpenDropdowns_(sheet.getParent(), work);
      }
    });
    return;
  }

  if (sheetName === CONFIG.workList.sheetName) {
    writeInternal_(function () {
      fillWorkListWorkerCodesFromEdit_(sheet, e.range);
    });
    return;
  }
}
