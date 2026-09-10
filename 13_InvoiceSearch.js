/**
 * シート「請求書検索」。K-No 検索 → 一覧から選択 → 印刷_表示へ描画。
 * 入力アプリの画面は書き換えない。
 */

var INVOICE_SEARCH_KNO_ = 'B3';
var INVOICE_SEARCH_SEARCH_ = 'C3';
var INVOICE_SEARCH_STATUS_ = 'E3';
var INVOICE_SEARCH_PICK_ = 'B4';
var INVOICE_SEARCH_SHOW_ = 'C4';
var INVOICE_SEARCH_HEAD_ROW_ = 6;
var INVOICE_SEARCH_LIST_START_ = 7;
var INVOICE_SEARCH_COLS_ = 10;

function openInvoiceSearchSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ensureInvoiceSearchSheet_(ss);
  ss.setActiveSheet(sh);
  ss.toast('K-No を入れて Enter。一覧から選んで「表示」をオン。', '請求書検索', 6);
}

function handleInvoiceSearchEdit_(e) {
  const sh = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const a1 = e.range.getA1Notation();

  if (a1 === INVOICE_SEARCH_SEARCH_ || (row === 3 && col === 3)) {
    if (!e.range.getValue()) {
      return;
    }
    writeInternal_(function () {
      sh.getRange(INVOICE_SEARCH_SEARCH_).setValue(false);
    });
    runInvoiceSearchOnSheet_(sh);
    return;
  }

  if (a1 === INVOICE_SEARCH_KNO_ || (row === 3 && col === 2)) {
    runInvoiceSearchOnSheet_(sh);
    return;
  }

  if (a1 === INVOICE_SEARCH_SHOW_ || (row === 4 && col === 3)) {
    if (!e.range.getValue()) {
      return;
    }
    writeInternal_(function () {
      sh.getRange(INVOICE_SEARCH_SHOW_).setValue(false);
    });
    showInvoiceFromSearchPick_(sh);
    return;
  }

  if (a1 === INVOICE_SEARCH_PICK_ || (row === 4 && col === 2)) {
    return;
  }

  if (col === 1 && row >= INVOICE_SEARCH_LIST_START_ && e.range.getValue()) {
    const saveId = String(sh.getRange(row, 9).getValue() || '').trim();
    writeInternal_(function () {
      sh.getRange(row, 1).setValue(false);
    });
    if (saveId) {
      showSavedInvoicePrintSafe_(saveId);
    }
  }
}

function runInvoiceSearchOnSheet_(sh) {
  const raw = sh.getRange(INVOICE_SEARCH_KNO_).getValue();
  const kNo = normalizeInvoiceKNo_(raw);
  if (!kNo) {
    writeInternal_(function () {
      sh.getRange(INVOICE_SEARCH_STATUS_).setValue('K-No を入力してください');
    });
    return;
  }
  writeInternal_(function () {
    sh.getRange(INVOICE_SEARCH_KNO_).setNumberFormat('@').setValue(kNo);
    sh.getRange(INVOICE_SEARCH_STATUS_).setValue('検索中…');
  });
  let list = [];
  try {
    list = listInvoiceDrafts(kNo) || [];
  } catch (err) {
    writeInternal_(function () {
      sh.getRange(INVOICE_SEARCH_STATUS_).setValue(String(err.message || err));
    });
    SpreadsheetApp.getActive().toast(String(err.message || err), '請求書検索', 8);
    return;
  }
  writeInternal_(function () {
    fillInvoiceSearchList_(sh, kNo, list);
  });
}

function fillInvoiceSearchList_(sh, kNo, list) {
  const start = INVOICE_SEARCH_LIST_START_;
  const last = Math.max(sh.getLastRow(), start);
  if (last >= start) {
    sh.getRange(start, 1, last - start + 1, INVOICE_SEARCH_COLS_).clearContent();
    sh.getRange(start, 1, last - start + 1, 1).removeCheckboxes();
  }
  sh.getRange(INVOICE_SEARCH_PICK_).clearDataValidations().clearContent();

  const labels = [];
  const used = {};
  const rows = [];
  for (let i = 0; i < list.length; i++) {
    const row = list[i];
    let label = invoiceSearchLabel_(row);
    if (used[label]) {
      label = label + '  (' + String(row.saveId || '').slice(0, 8) + ')';
    }
    used[label] = true;
    labels.push(label);
    rows.push([
      false,
      row.savedAt || '',
      row.userName || '',
      row.plate || '',
      row.billDate || '',
      row.firstMid || '',
      row.lineCount ? row.lineCount + '行' : '',
      row.printSheet || '',
      row.saveId || '',
      label
    ]);
  }

  if (rows.length) {
    sh.getRange(start, 1, rows.length, INVOICE_SEARCH_COLS_).setValues(rows);
    sh.getRange(start, 1, rows.length, 1).insertCheckboxes();
    const labelRange = sh.getRange(start, 10, rows.length, 1);
    sh.getRange(INVOICE_SEARCH_PICK_).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInRange(labelRange, true).setAllowInvalid(false).build()
    );
    sh.getRange(INVOICE_SEARCH_PICK_).setValue(labels[0]);
  }

  sh.getRange(INVOICE_SEARCH_STATUS_).setValue(
    list.length ? ('K-No ' + kNo + ' ／ ' + list.length + ' 件。一覧か上のリストから選んで「表示」。') : ('K-No ' + kNo + ' の保存はありません')
  );
}

function invoiceSearchLabel_(row) {
  return [row.savedAt || '', row.userName || '', row.plate || '', (row.lineCount || 0) + '行'].join('  ').trim();
}

function showInvoiceFromSearchPick_(sh) {
  const label = String(sh.getRange(INVOICE_SEARCH_PICK_).getValue() || '').trim();
  if (!label) {
    SpreadsheetApp.getActive().toast('先に請求書を選んでください', '請求書検索', 6);
    return;
  }
  const start = INVOICE_SEARCH_LIST_START_;
  const last = sh.getLastRow();
  if (last < start) {
    SpreadsheetApp.getActive().toast('先に K-No で検索してください', '請求書検索', 6);
    return;
  }
  const vals = sh.getRange(start, 9, last - start + 1, 2).getValues();
  let saveId = '';
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][1] || '').trim() === label) {
      saveId = String(vals[i][0] || '').trim();
      break;
    }
  }
  if (!saveId) {
    SpreadsheetApp.getActive().toast('選んだ請求書の保存IDが見つかりません。もう一度検索してください。', '請求書検索', 8);
    return;
  }
  showSavedInvoicePrintSafe_(saveId);
}

function showSavedInvoicePrintSafe_(saveId) {
  try {
    showSavedInvoicePrint_(saveId);
  } catch (err) {
    SpreadsheetApp.getActive().toast(String(err.message || err), '請求書検索', 8);
  }
}

function showSavedInvoicePrint_(saveId) {
  const id = String(saveId || '').trim();
  if (!id) {
    throw new Error('保存データが指定されていません。');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const draft = loadInvoiceDraft(id);
  if (!draft || !draft.items || !draft.items.filter(rowHasContent_).length) {
    throw new Error('この保存には表示できる明細がありません。');
  }
  const payload = {
    header: draft.header || {},
    items: draft.items || [],
    summary: draft.summary || {},
    saveId: id
  };
  const viewName = (CONFIG.print && CONFIG.print.viewSheetName) || '印刷_表示';
  ss.toast('保存データから表示しています…', '請求書検索', 5);
  const built = buildInvoicePrintSheet_(ss, viewName, payload);
  setInvoicePrintSheetName_(id, built.sheet.getName());
  ss.setActiveSheet(built.sheet);
  ss.toast(built.sheet.getName() + ' に表示しました（K-No ' + formatPrintKNo_(payload.header.kNo) + '）', '請求書検索', 6);
}

function ensureInvoiceSearchSheet_(ss) {
  const name = CONFIG.invoiceSearch.sheetName;
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  writeInternal_(function () {
    layoutInvoiceSearchSheet_(sh);
  });
  return sh;
}

function layoutInvoiceSearchSheet_(sh) {
  sh.setHiddenGridlines(false);
  sh.setFrozenRows(INVOICE_SEARCH_HEAD_ROW_);
  sh.getRange('A1').setValue('請求書検索').setFontSize(16).setFontWeight('bold');
  sh.getRange('A2').setValue('K-No を入れて Enter（または「検索」をオン）→ リストから選ぶ → 「表示」をオン。一覧の「表示」でも開けます。');
  sh.getRange('A3').setValue('K-No');
  sh.getRange('A4').setValue('請求書');
  sh.getRange('C3').insertCheckboxes();
  sh.getRange('C3').setValue(false);
  sh.getRange('C4').insertCheckboxes();
  sh.getRange('C4').setValue(false);
  sh.getRange('D3').setValue('検索');
  sh.getRange('D4').setValue('表示');
  if (!String(sh.getRange(INVOICE_SEARCH_STATUS_).getValue() || '').trim()) {
    sh.getRange(INVOICE_SEARCH_STATUS_).setValue('');
  }
  sh.getRange(INVOICE_SEARCH_KNO_).setNumberFormat('@');
  sh.getRange('A3:A4').setFontWeight('bold');
  sh.getRange(INVOICE_SEARCH_HEAD_ROW_, 1, 1, 8).setValues([[
    '表示', '保存日時', 'ユーザー', '登録番号', '請求日', '中項目1', '明細', '印刷シート'
  ]]);
  sh.getRange(INVOICE_SEARCH_HEAD_ROW_, 1, 1, 8).setFontWeight('bold');
  sh.setColumnWidth(1, 52);
  sh.setColumnWidth(2, 110);
  sh.setColumnWidth(3, 140);
  sh.setColumnWidth(4, 120);
  sh.setColumnWidth(5, 110);
  sh.setColumnWidth(6, 220);
  sh.setColumnWidth(7, 60);
  sh.setColumnWidth(8, 160);
  sh.setColumnWidth(9, 80);
  sh.setColumnWidth(10, 80);
  sh.hideColumns(9, 2);
}
