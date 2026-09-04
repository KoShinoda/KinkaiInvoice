/**
 * 入力アプリ用の明細テンプレート。
 * シート「明細テンプレート」：同じテンプレート名の行＝1つの明細セット。
 */

function ensureInvoiceTemplateSheet() {
  const sh = ensureInvoiceTemplateSheet_();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'シート「' + sh.getName() + '」を用意しました',
    '請求書入力',
    5
  );
}

function loadInvoiceTemplateNames_() {
  const parsed = parseInvoiceTemplateSheet_();
  return parsed.names;
}

/**
 * 入力アプリから呼ぶ。選んだテンプレートの明細行。
 *
 * @param {string} name
 * @return {object[]}
 */
function getInvoiceTemplateLines(name) {
  const parsed = parseInvoiceTemplateSheet_();
  const key = normalize_(name);
  return parsed.linesByName[key] || [];
}

function ensureInvoiceTemplateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = CONFIG.invoiceTemplate.sheetName;
  let sh = ss.getSheetByName(sheetName);
  if (sh && isInvoiceTemplateLayout_(sh)) {
    return sh;
  }
  if (!sh) {
    sh = ss.insertSheet(sheetName);
  }
  writeInvoiceTemplateSample_(sh);
  return sh;
}

function isInvoiceTemplateLayout_(sheet) {
  if (!sheet) {
    return false;
  }
  const header = sheet.getRange(1, 1, 1, 10).getValues()[0].map(function (v) {
    return normalize_(v);
  });
  return header.indexOf('テンプレート名') !== -1;
}

function parseInvoiceTemplateSheet_() {
  const empty = { names: [], linesByName: {} };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.invoiceTemplate.sheetName);
  if (!sh || !isInvoiceTemplateLayout_(sh)) {
    sh = ensureInvoiceTemplateSheet_();
  }
  const values = sh.getDataRange().getValues();
  if (values.length < 2) {
    return empty;
  }
  const cols = resolveColumns_(values[CONFIG.invoiceTemplate.headerRow - 1], CONFIG.invoiceTemplate.headers);
  if (!cols.name) {
    return empty;
  }
  const names = [];
  const linesByName = {};
  for (let i = CONFIG.invoiceTemplate.headerRow; i < values.length; i++) {
    const raw = values[i];
    const tmpl = normalize_(cell_(raw, cols.name));
    if (!tmpl) {
      continue;
    }
    const line = {
      major: normalize_(cell_(raw, cols.major)),
      mid: normalize_(cell_(raw, cols.mid)),
      fee: cell_(raw, cols.fee),
      workerCode: normalize_(cell_(raw, cols.workerCode)),
      partMajor: normalize_(cell_(raw, cols.partMajor)),
      partMid: normalize_(cell_(raw, cols.partMid)),
      unitPrice: cell_(raw, cols.unitPrice),
      qty: cell_(raw, cols.qty),
      discYen: cell_(raw, cols.discYen)
    };
    if (!line.major && !line.mid && !isFilled_(line.fee) && !line.partMajor && !line.partMid &&
      !isFilled_(line.qty) && !isFilled_(line.unitPrice)) {
      continue;
    }
    if (!linesByName[tmpl]) {
      linesByName[tmpl] = [];
      names.push(tmpl);
    }
    linesByName[tmpl].push(line);
  }
  return { names: names, linesByName: linesByName };
}

function writeInvoiceTemplateSample_(sh) {
  sh.clear();
  const headers = [[
    'テンプレート名', '大項目', '中項目', '技術料', '作業者コード',
    '部品_大項目', '部品_中項目', '単価', '数量', '値引額'
  ]];
  sh.getRange(1, 1, 1, 10).setValues(headers);
  sh.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#e8f0ec');

  const body = defaultInvoiceTemplateRows_().map(function (r) {
    return [
      r.name, r.major, r.mid, r.fee, r.workerCode || '',
      r.partMajor || '', r.partMid || '', r.unitPrice === '' || r.unitPrice == null ? '' : r.unitPrice,
      r.qty === '' || r.qty == null ? '' : r.qty,
      r.discYen === '' || r.discYen == null ? '' : r.discYen
    ];
  });
  sh.getRange(2, 1, body.length, 10).setValues(body);
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 100);
  sh.setColumnWidth(3, 220);
  sh.setColumnWidth(4, 80);
  sh.setColumnWidth(5, 100);
  sh.setColumnWidth(6, 110);
  sh.setColumnWidth(7, 160);
  sh.setColumnWidth(8, 80);
  sh.setColumnWidth(9, 60);
  sh.setColumnWidth(10, 80);
  sh.getRange(1, 1).setNote(
    '同じテンプレート名の行が、入力アプリで選んだときの明細になります。\n' +
    '中項目列には、画面の中項目（作業内容）を書いてください。\n' +
    '部品は同じ行に横並びでも、作業だけの行／部品だけの行に分けても構いません。\n' +
    '合計は数量×単価から自動計算します。値引額は円（空欄可）。'
  );
}

function defaultInvoiceTemplateRows_() {
  return [
    { name: '3カ月定期点検', major: '定期点検', mid: '＊＊　3カ月定期点検　＊＊', fee: 31000 },
    { name: '6カ月定期点検', major: '定期点検', mid: '＊＊　6カ月定期点検　＊＊', fee: 1000 },
    { name: '6カ月定期点検', major: '定期点検', mid: '6か月作業1', fee: 2000 },
    { name: '6カ月定期点検', major: '定期点検', mid: '6か月作業2', fee: 3000 },
    { name: '１２カ月定期点検', major: '定期点検', mid: '＊＊　１２カ月定期点検　＊＊', fee: 35000 },
    { name: '１２カ月定期点検', major: '定期点検', mid: 'シャシ洗浄、グリスアップ', fee: 8000, partMajor: '油脂', partMid: 'ＢＰＷ用ハブＢ／ｇグリス', unitPrice: 7200, qty: 1 },
    { name: '１２カ月定期点検', major: '定期点検', mid: 'シャシグレー塗装', fee: 12000 },
    { name: '１２カ月定期点検', major: '定期点検', mid: 'シャシマスキング', fee: 3000 },
    { name: '１２カ月定期点検', major: '定期点検', mid: '保安確認検査料', fee: 3000 },
    { name: '１２カ月定期点検', major: '定期点検', mid: '代行料', fee: 5000 },
    { name: '部品交換サンプル', major: '', mid: '', fee: '', partMajor: '油脂', partMid: 'カートリッジグリス', unitPrice: 500, qty: 1 },
    { name: '部品交換サンプル', major: '', mid: '', fee: '', partMajor: '油脂', partMid: 'ｼｬｼｸﾞﾚｰ', unitPrice: 15000, qty: 1 },
    { name: '部品交換サンプル', major: '', mid: '', fee: '', partMajor: '油脂', partMid: 'ｽﾓｰﾙ･ﾊﾟｰﾂ', unitPrice: 2500, qty: 1, discYen: 100 }
  ];
}
