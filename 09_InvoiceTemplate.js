/**
 * 入力アプリ用のテンプレート。
 * シート「テンプレートリスト」：同じテンプレート名の行＝1つの明細セット。
 * ヘッダー初期値（ユーザー・登録番号・整備部門など）は同じ名前のうち最初の値。
 */

var INVOICE_TEMPLATE_NAME_HEADER_ = 'テンプレート名';
var INVOICE_TEMPLATE_LINE_HEADERS_ = [
  '大項目', '中項目', '技術料', '作業者コード',
  '部品_大項目', '部品_中項目', '単価', '数量', '値引額'
];
var INVOICE_TEMPLATE_META_HEADERS_ = [
  'ユーザー', '登録番号', '整備部門', '整備種別', '受付担当',
  '入庫日', '出庫日', '請求日', '値引技術%', '値引部品%'
];
var INVOICE_TEMPLATE_META_KEYS_ = [
  'userName', 'plate', 'dept', 'serviceType', 'receptionist',
  'inDate', 'outDate', 'billDate', 'techPct', 'partPct'
];

function invoiceTemplateColumnHeaders_() {
  return [INVOICE_TEMPLATE_NAME_HEADER_].concat(INVOICE_TEMPLATE_META_HEADERS_).concat(INVOICE_TEMPLATE_LINE_HEADERS_);
}

function findInvoiceTemplateSheet_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const names = [CONFIG.invoiceTemplate.sheetName, '明細テンプレート'];
  for (let i = 0; i < names.length; i++) {
    const sh = ss.getSheetByName(names[i]);
    if (sh) {
      return sh;
    }
  }
  return null;
}

function ensureInvoiceTemplateSheet() {
  const sh = ensureInvoiceTemplateSheet_();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'シート「' + sh.getName() + '」を用意しました（ヘッダー初期値の列も含みます）',
    '請求書入力',
    5
  );
}

function getInvoiceTemplateNames() {
  const sh = findInvoiceTemplateSheet_();
  if (sh) {
    ensureInvoiceTemplateHeaderCols_(sh);
  }
  return listInvoiceTemplateNamesFast_();
}

function loadInvoiceTemplateNames_() {
  return listInvoiceTemplateNamesFast_();
}

function listInvoiceTemplateNamesFast_() {
  const sh = findInvoiceTemplateSheet_();
  if (!sh) {
    return [];
  }
  const last = sh.getLastRow();
  if (last < 2) {
    return [];
  }
  const vals = sh.getRange(2, 1, last - 1, 1).getValues();
  return uniqueValues_(vals.map(function (row) {
    return row[0];
  }));
}

/**
 * 入力アプリから呼ぶ。選んだテンプレートの明細とヘッダー初期値。
 *
 * @param {string} name
 * @return {{name: string, lines: object[], header: object, summary: object}}
 */
function getInvoiceTemplateLines(name) {
  const parsed = parseInvoiceTemplateSheet_();
  const key = normalize_(name);
  const header = parsed.headerByName[key] || {};
  return invoiceJsonSafe_({
    name: key,
    lines: parsed.linesByName[key] || [],
    header: templateHeaderForClient_(header),
    summary: templateSummaryForClient_(header)
  });
}

function ensureInvoiceTemplateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = CONFIG.invoiceTemplate.sheetName;
  let sh = findInvoiceTemplateSheet_(ss);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    writeInvoiceTemplateSample_(sh);
    return sh;
  }
  if (sh.getName() !== sheetName) {
    try {
      sh.setName(sheetName);
    } catch (err) {
      // 新しい名前が使えないときは、見つかったシートのまま使う。
    }
  }
  if (sh.getLastRow() < 2 && !isInvoiceTemplateLayout_(sh)) {
    writeInvoiceTemplateSample_(sh);
  } else {
    ensureInvoiceTemplateHeaderCols_(sh);
  }
  return sh;
}

function isInvoiceTemplateLayout_(sheet) {
  if (!sheet) {
    return false;
  }
  const lastCol = Math.max(sheet.getLastColumn(), invoiceTemplateColumnHeaders_().length);
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) {
    return normalize_(v);
  });
  return header.indexOf('テンプレート名') !== -1;
}

function invoiceTemplateHeaderCol_(sh, logicalKey) {
  const last = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, last).getValues()[0];
  const cols = resolveColumns_(headers, CONFIG.invoiceTemplate.headers);
  return cols[logicalKey] || 0;
}

function moveSheetColumn_(sh, from1, to1) {
  if (!from1 || !to1 || from1 === to1) {
    return;
  }
  const rows = sh.getMaxRows();
  const width = sh.getColumnWidth(from1);
  if (from1 > to1) {
    sh.insertColumnBefore(to1);
    sh.getRange(1, from1 + 1, rows, 1).moveTo(sh.getRange(1, to1));
    sh.deleteColumn(from1 + 1);
  } else {
    sh.insertColumnAfter(to1);
    sh.getRange(1, from1, rows, 1).moveTo(sh.getRange(1, to1 + 1));
    sh.deleteColumn(from1);
  }
  try {
    sh.setColumnWidth(to1, width);
  } catch (err) {
    // 列幅は必須ではない。
  }
}

function ensureHeaderTitleAt_(sh, logicalKey, title, dest1, width) {
  let from = invoiceTemplateHeaderCol_(sh, logicalKey);
  if (!from) {
    if (dest1 <= 1) {
      sh.insertColumnBefore(1);
    } else {
      sh.insertColumnAfter(dest1 - 1);
    }
    sh.getRange(1, dest1).setValue(title).setFontWeight('bold').setBackground('#e8f0ec');
    if (width) {
      sh.setColumnWidth(dest1, width);
    }
    return;
  }
  if (from !== dest1) {
    moveSheetColumn_(sh, from, dest1);
  }
  sh.getRange(1, dest1).setFontWeight('bold').setBackground('#e8f0ec');
  if (width) {
    sh.setColumnWidth(dest1, width);
  }
}

function invoiceTemplateMetaInPlace_(sh) {
  if (invoiceTemplateHeaderCol_(sh, 'name') !== 1) {
    return false;
  }
  for (let i = 0; i < INVOICE_TEMPLATE_META_KEYS_.length; i++) {
    if (invoiceTemplateHeaderCol_(sh, INVOICE_TEMPLATE_META_KEYS_[i]) !== i + 2) {
      return false;
    }
  }
  return true;
}

function ensureInvoiceTemplateHeaderCols_(sh) {
  if (!sh || !isInvoiceTemplateLayout_(sh)) {
    return;
  }
  try {
    if (invoiceTemplateMetaInPlace_(sh)) {
      return;
    }
    ensureHeaderTitleAt_(sh, 'name', INVOICE_TEMPLATE_NAME_HEADER_, 1, 160);
    INVOICE_TEMPLATE_META_KEYS_.forEach(function (key, i) {
      const title = INVOICE_TEMPLATE_META_HEADERS_[i];
      const dest = i + 2;
      const width = (key === 'plate' || key === 'userName') ? 120 : 90;
      ensureHeaderTitleAt_(sh, key, title, dest, width);
    });
    sh.getRange(1, 1).setNote(invoiceTemplateSheetNote_());
  } catch (err) {
    // 保護などで列を足せないときは、既存の列だけで読む。
  }
}

function parseInvoiceTemplateSheet_() {
  const empty = { names: [], linesByName: {}, headerByName: {} };
  const sh = findInvoiceTemplateSheet_();
  if (!sh) {
    return empty;
  }
  ensureInvoiceTemplateHeaderCols_(sh);
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
  const headerByName = {};
  for (let i = CONFIG.invoiceTemplate.headerRow; i < values.length; i++) {
    const raw = values[i];
    const tmpl = normalize_(cell_(raw, cols.name));
    if (!tmpl) {
      continue;
    }
    if (!headerByName[tmpl]) {
      headerByName[tmpl] = {};
    }
    mergeTemplateHeaderFirstWins_(headerByName[tmpl], pickTemplateHeader_(raw, cols));
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
    const hasLine = !!(line.major || line.mid || isFilled_(line.fee) || line.partMajor || line.partMid ||
      isFilled_(line.qty) || isFilled_(line.unitPrice) || isFilled_(line.discYen));
    if (!hasLine) {
      if (!linesByName[tmpl] && Object.keys(headerByName[tmpl]).length) {
        names.push(tmpl);
        linesByName[tmpl] = [];
      }
      continue;
    }
    if (!linesByName[tmpl]) {
      linesByName[tmpl] = [];
      names.push(tmpl);
    }
    linesByName[tmpl].push(line);
  }
  return { names: names, linesByName: linesByName, headerByName: headerByName };
}

function pickTemplateHeader_(raw, cols) {
  const h = {};
  takeTemplateText_(h, 'userName', raw, cols);
  takeTemplateText_(h, 'plate', raw, cols);
  takeTemplateText_(h, 'dept', raw, cols);
  takeTemplateText_(h, 'serviceType', raw, cols);
  takeTemplateText_(h, 'receptionist', raw, cols);
  takeTemplateDate_(h, 'inDate', raw, cols);
  takeTemplateDate_(h, 'outDate', raw, cols);
  takeTemplateDate_(h, 'billDate', raw, cols);
  takeTemplatePct_(h, 'techPct', raw, cols);
  takeTemplatePct_(h, 'partPct', raw, cols);
  return h;
}

function takeTemplateText_(h, key, raw, cols) {
  const v = normalize_(cell_(raw, cols[key]));
  if (v) {
    h[key] = v;
  }
}

function takeTemplateDate_(h, key, raw, cols) {
  if (!cols[key]) {
    return;
  }
  const v = formatInvoiceYmd_(cell_(raw, cols[key]));
  if (v) {
    h[key] = v;
  }
}

function takeTemplatePct_(h, key, raw, cols) {
  if (!cols[key]) {
    return;
  }
  const v = cell_(raw, cols[key]);
  if (!isFilled_(v)) {
    return;
  }
  const n = Number(String(v).replace(/,/g, '').replace(/%/g, '').trim());
  if (isFinite(n)) {
    h[key] = n;
  }
}

function mergeTemplateHeaderFirstWins_(dst, src) {
  Object.keys(src || {}).forEach(function (key) {
    if (dst[key] === undefined || dst[key] === '') {
      dst[key] = src[key];
    }
  });
}

function templateHeaderForClient_(header) {
  const h = header || {};
  const out = {};
  ['userName', 'plate', 'dept', 'serviceType', 'receptionist', 'inDate', 'outDate', 'billDate'].forEach(function (key) {
    if (h[key] !== undefined && h[key] !== '') {
      out[key] = h[key];
    }
  });
  return out;
}

function templateSummaryForClient_(header) {
  const h = header || {};
  const out = {};
  if (h.techPct !== undefined && h.techPct !== '') {
    out.techPct = h.techPct;
  }
  if (h.partPct !== undefined && h.partPct !== '') {
    out.partPct = h.partPct;
  }
  return out;
}

function invoiceTemplateSheetNote_() {
  return '同じテンプレート名の行が、入力アプリで選んだときの明細になります。\n' +
    'ヘッダー初期値（ユーザー〜請求日、値引％）は同じ名前のうち最初に入っている値を使います。空欄の項目は画面の値を残します。\n' +
    '中項目列には、画面の中項目（作業内容）を書いてください。\n' +
    '部品は同じ行に横並びでも、作業だけの行／部品だけの行に分けても構いません。\n' +
    '合計は数量×単価から自動計算します。値引額は円（空欄可）。';
}

function writeInvoiceTemplateSample_(sh) {
  sh.clear();
  const headers = [invoiceTemplateColumnHeaders_()];
  const colCount = headers[0].length;
  sh.getRange(1, 1, 1, colCount).setValues(headers);
  sh.getRange(1, 1, 1, colCount).setFontWeight('bold').setBackground('#e8f0ec');

  const body = defaultInvoiceTemplateRows_().map(function (r) {
    return [
      r.name,
      r.userName || '', r.plate || '', r.dept || '', r.serviceType || '', r.receptionist || '',
      r.inDate || '', r.outDate || '', r.billDate || '',
      r.techPct === '' || r.techPct == null ? '' : r.techPct,
      r.partPct === '' || r.partPct == null ? '' : r.partPct,
      r.major || '', r.mid || '', r.fee === '' || r.fee == null ? '' : r.fee, r.workerCode || '',
      r.partMajor || '', r.partMid || '', r.unitPrice === '' || r.unitPrice == null ? '' : r.unitPrice,
      r.qty === '' || r.qty == null ? '' : r.qty,
      r.discYen === '' || r.discYen == null ? '' : r.discYen
    ];
  });
  sh.getRange(2, 1, body.length, colCount).setValues(body);
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 120);
  sh.setColumnWidth(3, 140);
  sh.setColumnWidth(4, 90);
  sh.setColumnWidth(5, 100);
  sh.setColumnWidth(6, 90);
  sh.setColumnWidth(7, 90);
  sh.setColumnWidth(8, 90);
  sh.setColumnWidth(9, 90);
  sh.setColumnWidth(10, 90);
  sh.setColumnWidth(11, 90);
  sh.setColumnWidth(12, 100);
  sh.setColumnWidth(13, 220);
  sh.setColumnWidth(14, 80);
  sh.setColumnWidth(15, 100);
  sh.setColumnWidth(16, 110);
  sh.setColumnWidth(17, 160);
  sh.setColumnWidth(18, 80);
  sh.setColumnWidth(19, 60);
  sh.setColumnWidth(20, 80);
  sh.getRange(1, 1).setNote(invoiceTemplateSheetNote_());
}

function defaultInvoiceTemplateRows_() {
  return [
    { name: '3カ月定期点検', major: '定期点検', mid: '＊＊　3カ月定期点検　＊＊', fee: 31000 },
    { name: '6カ月定期点検', major: '定期点検', mid: '＊＊　6カ月定期点検　＊＊', fee: 1000 },
    { name: '6カ月定期点検', major: '定期点検', mid: '6か月作業1', fee: 2000 },
    { name: '6カ月定期点検', major: '定期点検', mid: '6か月作業2', fee: 3000 },
    { name: '１２カ月定期点検', dept: '大型', serviceType: '点検大型' },
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
