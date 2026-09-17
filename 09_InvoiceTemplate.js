/**
 * 入力アプリ用のテンプレート。
 * シート「テンプレートリスト」：同じテンプレート名の行＝1つの明細セット。
 * ヘッダー初期値は整備部門・整備種別。同じ名前のうち最初の値。
 * 順番列と図形ボタン refreshInvoiceTemplateList で並べ替える。入力画面から保存できる。
 */

var INVOICE_TEMPLATE_NAME_HEADER_ = 'テンプレート名';
var INVOICE_TEMPLATE_LINE_HEADERS_ = [
  '大項目', '中項目', '技術料', '作業者コード',
  '部品_大項目', '部品_中項目', '単価', '数量', '値引額'
];
var INVOICE_TEMPLATE_META_HEADERS_ = [
  '整備部門', '整備種別'
];
var INVOICE_TEMPLATE_META_KEYS_ = [
  'dept', 'serviceType'
];
var INVOICE_TEMPLATE_OBSOLETE_HEADERS_ = [
  'ユーザー', 'ユーザー名', '顧客', '登録番号', 'ナンバー', '受付担当',
  '入庫日', '出庫日', '請求日', '値引技術%', '技術値引%', '値引部品%', '部品値引%'
];

function invoiceTemplateColumnHeaders_() {
  return [INVOICE_TEMPLATE_NAME_HEADER_].concat(INVOICE_TEMPLATE_META_HEADERS_).concat(INVOICE_TEMPLATE_LINE_HEADERS_).concat(['順番']);
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
    'シート「' + sh.getName() + '」を用意しました',
    '請求書入力',
    5
  );
}

function getInvoiceTemplateNames() {
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
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const cols = resolveColumns_(header, CONFIG.invoiceTemplate.headers);
  const nameCol = cols.name || 1;
  const orderCol = cols.order || 0;
  if (!orderCol) {
    const vals = sh.getRange(2, nameCol, last - 1, 1).getValues();
    return uniqueValues_(vals.map(function (row) {
      return row[0];
    }));
  }
  const width = Math.max(nameCol, orderCol);
  const block = sh.getRange(2, 1, last - 1, width).getValues();
  return listInvoiceTemplateNamesFromRows_(block, nameCol, orderCol);
}

function listInvoiceTemplateNamesFromRows_(rows, nameCol, orderCol) {
  const seen = {};
  const meta = [];
  for (let i = 0; i < (rows || []).length; i++) {
    const name = normalize_(rows[i][nameCol - 1]);
    if (!name) {
      continue;
    }
    const ord = orderCol ? toOrderNumber_(rows[i][orderCol - 1]) : i + 1;
    if (!seen[name]) {
      seen[name] = { name: name, order: ord, idx: i };
      meta.push(seen[name]);
    } else if (ord < seen[name].order) {
      seen[name].order = ord;
    }
  }
  meta.sort(function (a, b) {
    if (a.order !== b.order) {
      return a.order - b.order;
    }
    return a.idx - b.idx;
  });
  return meta.map(function (m) {
    return m.name;
  });
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
    applyInvoiceTemplateDropdowns_(sh);
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
  applyInvoiceTemplateDropdowns_(sh);
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

function invoiceTemplateHeaderMap_(sh) {
  const last = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, last).getValues()[0];
  return resolveColumns_(headers, CONFIG.invoiceTemplate.headers);
}

function invoiceTemplateHeaderCol_(sh, logicalKey) {
  return invoiceTemplateHeaderMap_(sh)[logicalKey] || 0;
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
  const cols = invoiceTemplateHeaderMap_(sh);
  if (cols.name !== 1) {
    return false;
  }
  for (let i = 0; i < INVOICE_TEMPLATE_META_KEYS_.length; i++) {
    if (cols[INVOICE_TEMPLATE_META_KEYS_[i]] !== i + 2) {
      return false;
    }
  }
  return true;
}

function removeObsoleteInvoiceTemplateCols_(sh) {
  const drop = {};
  INVOICE_TEMPLATE_OBSOLETE_HEADERS_.forEach(function (title) {
    drop[normalize_(title)] = true;
  });
  const last = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, last).getValues()[0];
  for (let i = headers.length - 1; i >= 0; i--) {
    if (drop[normalize_(headers[i])]) {
      sh.deleteColumn(i + 1);
    }
  }
}

function ensureInvoiceTemplateOrderCol_(sh) {
  const col = ensureOrderColumnOnSheet_(sh);
  try {
    sh.getRange(1, col).setFontWeight('bold').setBackground('#e8f0ec');
    sh.setColumnWidth(col, 64);
  } catch (err) {
    // 列幅は必須ではない。
  }
  return col;
}

function ensureInvoiceTemplateHeaderCols_(sh) {
  if (!sh || !isInvoiceTemplateLayout_(sh)) {
    return;
  }
  try {
    const lastColBefore = sh.getLastColumn();
    removeObsoleteInvoiceTemplateCols_(sh);
    const removed = sh.getLastColumn() < lastColBefore;
    const needLayout = !invoiceTemplateMetaInPlace_(sh);
    if (needLayout) {
      ensureHeaderTitleAt_(sh, 'name', INVOICE_TEMPLATE_NAME_HEADER_, 1, 160);
      INVOICE_TEMPLATE_META_KEYS_.forEach(function (key, i) {
        const title = INVOICE_TEMPLATE_META_HEADERS_[i];
        const dest = i + 2;
        const width = key === 'serviceType' ? 100 : 90;
        ensureHeaderTitleAt_(sh, key, title, dest, width);
      });
      sh.getRange(1, 1).setNote(invoiceTemplateSheetNote_());
    }
    ensureInvoiceTemplateOrderCol_(sh);
    if (removed || needLayout) {
      applyInvoiceTemplateDropdowns_(sh);
    }
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
  const linesByName = {};
  const headerByName = {};
  const lineMeta = {};
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
    const ord = toOrderNumber_(cell_(raw, cols.order));
    if (!hasLine) {
      if (!lineMeta[tmpl] && Object.keys(headerByName[tmpl]).length) {
        lineMeta[tmpl] = [];
        linesByName[tmpl] = [];
      }
      continue;
    }
    if (!lineMeta[tmpl]) {
      lineMeta[tmpl] = [];
      linesByName[tmpl] = [];
    }
    lineMeta[tmpl].push({ order: ord, idx: i, line: line });
  }
  Object.keys(lineMeta).forEach(function (tmpl) {
    lineMeta[tmpl].sort(function (a, b) {
      if (a.order !== b.order) {
        return a.order - b.order;
      }
      return a.idx - b.idx;
    });
    linesByName[tmpl] = lineMeta[tmpl].map(function (row) {
      return row.line;
    });
  });
  return {
    names: listInvoiceTemplateNamesFromRows_(
      values.slice(CONFIG.invoiceTemplate.headerRow),
      cols.name || 1,
      cols.order || 0
    ),
    linesByName: linesByName,
    headerByName: headerByName
  };
}

function pickTemplateHeader_(raw, cols) {
  const h = {};
  takeTemplateText_(h, 'dept', raw, cols);
  takeTemplateText_(h, 'serviceType', raw, cols);
  return h;
}

function takeTemplateText_(h, key, raw, cols) {
  const v = normalize_(cell_(raw, cols[key]));
  if (v) {
    h[key] = v;
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
  ['dept', 'serviceType'].forEach(function (key) {
    if (h[key] !== undefined && h[key] !== '') {
      out[key] = h[key];
    }
  });
  return out;
}

function templateSummaryForClient_(header) {
  return {};
}

function invoiceTemplateSheetNote_() {
  return '同じテンプレート名の行が、入力アプリで選んだときの明細になります。\n' +
    '整備部門・整備種別は同じ名前のうち最初に入っている値をヘッダー初期値にします。空欄は画面の値を残します。\n' +
    '順番列でテンプレート名の並びと、同じ名前の中の行順を決めます。空欄は図形ボタン（refreshInvoiceTemplateList）で 10,20,… と埋まります。手で入れた番号は残します。\n' +
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

  const body = defaultInvoiceTemplateRows_().map(function (r, i) {
    return [
      r.name,
      r.dept || '', r.serviceType || '',
      r.major || '', r.mid || '', r.fee === '' || r.fee == null ? '' : r.fee, r.workerCode || '',
      r.partMajor || '', r.partMid || '', r.unitPrice === '' || r.unitPrice == null ? '' : r.unitPrice,
      r.qty === '' || r.qty == null ? '' : r.qty,
      r.discYen === '' || r.discYen == null ? '' : r.discYen,
      (i + 1) * invoiceTemplateOrderStep_()
    ];
  });
  sh.getRange(2, 1, body.length, colCount).setValues(body);
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 90);
  sh.setColumnWidth(3, 100);
  sh.setColumnWidth(4, 100);
  sh.setColumnWidth(5, 220);
  sh.setColumnWidth(6, 80);
  sh.setColumnWidth(7, 100);
  sh.setColumnWidth(8, 110);
  sh.setColumnWidth(9, 160);
  sh.setColumnWidth(10, 80);
  sh.setColumnWidth(11, 60);
  sh.setColumnWidth(12, 80);
  sh.setColumnWidth(13, 64);
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

function templateServiceTypeChoices_(dept) {
  let info = emptyServiceInfo_();
  try {
    info = loadServiceInfo_() || info;
  } catch (err) {
    Logger.log('%s templateServiceTypeChoices_: %s', CONFIG.logPrefix, err);
  }
  const raw = normalize_(dept);
  const key = raw === '車検' ? '大型' : raw === '一般整備' ? '小型' : raw === '板金塗装' ? 'BP板金' : raw;
  const types = key && info.typesByDept ? info.typesByDept[key] : [];
  if (types && types.length) {
    return types;
  }
  return info.allServiceTypes || [];
}

function templatePartMidChoices_(partMajor) {
  const out = [];
  const seen = {};
  function add(v) {
    const key = normalize_(v);
    if (!key || seen[key] || out.length >= 500) {
      return;
    }
    seen[key] = true;
    out.push(String(v));
  }
  try {
    const drop = loadPartListDropdowns_();
    const mids = partMajor && drop.partMidsByMajor && drop.partMidsByMajor[partMajor]
      ? drop.partMidsByMajor[partMajor]
      : (drop.allPartMids || []);
    (mids || []).forEach(add);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.parts.sheetName);
    if (sh) {
      parsePartsSheetValues_(sh.getDataRange().getValues()).rows.forEach(function (row) {
        if (partMajor && row.major && row.major !== partMajor) {
          return;
        }
        if (row.set) {
          add(row.set);
        }
      });
    }
  } catch (err) {
    Logger.log('%s templatePartMidChoices_: %s', CONFIG.logPrefix, err);
  }
  return out;
}

function applyInvoiceTemplateDropdowns_(sh) {
  sh = sh || findInvoiceTemplateSheet_();
  if (!sh) {
    return;
  }
  const cols = invoiceTemplateHeaderMap_(sh);
  const lastData = Math.max(sh.getLastRow(), 2);
  const rows = Math.max(lastData + 80, 200) - 1;
  let service = emptyServiceInfo_();
  let partsDrop = { partMajors: [], allPartMids: [] };
  try {
    service = loadServiceInfo_() || service;
  } catch (err) {
    Logger.log('%s applyInvoiceTemplateDropdowns_ service: %s', CONFIG.logPrefix, err);
  }
  try {
    partsDrop = loadPartListDropdowns_() || partsDrop;
  } catch (err) {
    Logger.log('%s applyInvoiceTemplateDropdowns_ parts: %s', CONFIG.logPrefix, err);
  }
  if (cols.dept) {
    applyOpenListValidation_(sh.getRange(2, cols.dept, rows, 1), service.departments || []);
  }
  if (cols.serviceType) {
    applyOpenListValidation_(sh.getRange(2, cols.serviceType, rows, 1), service.allServiceTypes || []);
  }
  if (cols.workerCode) {
    const labels = workerCodeLabelsForTemplateDropdown_();
    applyOpenListValidation_(sh.getRange(2, cols.workerCode, rows, 1), labels);
    coerceInvoiceTemplateWorkerCodes_(sh, sh.getRange(2, cols.workerCode, rows, 1));
  }
  if (cols.partMajor) {
    applyOpenListValidation_(sh.getRange(2, cols.partMajor, rows, 1), partsDrop.partMajors || []);
  }
  if (cols.partMid) {
    applyOpenListValidation_(sh.getRange(2, cols.partMid, rows, 1), templatePartMidChoices_(''));
  }
}

function workerCodeLabelForTemplate_(w) {
  const code = normalize_(w && w.code);
  const name = normalize_(w && w.name);
  if (!code) {
    return '';
  }
  return name ? (code + '：' + name) : code;
}

function workerCodeLabelsForTemplateDropdown_() {
  const seen = {};
  const out = [];
  try {
    (loadWorkers_() || []).forEach(function (w) {
      const label = workerCodeLabelForTemplate_(w);
      const key = normalize_(label);
      if (!key || seen[key] || out.length >= 500) {
        return;
      }
      seen[key] = true;
      out.push(label);
    });
  } catch (err) {
    Logger.log('%s workerCodeLabelsForTemplateDropdown_: %s', CONFIG.logPrefix, err);
  }
  return out;
}

function parseTemplateWorkerCodeCell_(raw) {
  if (raw === '' || raw == null) {
    return '';
  }
  let s = '';
  if (typeof raw === 'number' && isFinite(raw)) {
    s = String(Math.round(raw));
  } else {
    s = normalize_(raw);
  }
  if (!s) {
    return '';
  }
  let code = '';
  const labeled = s.match(/^(\d+)\s*[：:]\s*(.*)$/);
  if (labeled) {
    code = labeled[1];
  } else if (/^\d+$/.test(s)) {
    code = s;
  } else {
    try {
      const byName = workerCodeByNameMap_();
      if (byName[s] != null) {
        return String(byName[s]);
      }
    } catch (err) {
      Logger.log('%s parseTemplateWorkerCodeCell_: %s', CONFIG.logPrefix, err);
    }
    return '';
  }
  try {
    const workers = loadWorkers_() || [];
    for (let i = 0; i < workers.length; i++) {
      if (normalize_(workers[i].code) === code) {
        return String(workers[i].code);
      }
    }
  } catch (err) {
    Logger.log('%s parseTemplateWorkerCodeCell_ workers: %s', CONFIG.logPrefix, err);
  }
  return code;
}

function templateWorkerCodeWriteValue_(code) {
  const s = String(code == null ? '' : code).trim();
  if (/^\d+$/.test(s) && s === String(Number(s))) {
    return Number(s);
  }
  return s;
}

function coerceInvoiceTemplateWorkerCodes_(sh, range) {
  const cols = invoiceTemplateHeaderMap_(sh);
  coerceWorkerCodeColumn_(sh, range, cols && cols.workerCode);
}

/**
 * 選択肢は「コード：作業者」。セルには作業コードだけ残す。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 * @param {number} col
 */
function coerceWorkerCodeColumn_(sh, range, col) {
  if (!sh || !range || !col) {
    return;
  }
  const c1 = range.getColumn();
  const c2 = range.getLastColumn();
  if (col < c1 || col > c2) {
    return;
  }
  const from = Math.max(range.getRow(), 2);
  const to = range.getLastRow();
  if (to < from) {
    return;
  }
  const cells = sh.getRange(from, col, to - from + 1, 1);
  const vals = cells.getValues();
  const labels = workerCodeLabelsForTemplateDropdown_();
  let changed = false;
  for (let i = 0; i < vals.length; i++) {
    const parsed = parseTemplateWorkerCodeCell_(vals[i][0]);
    if (!parsed) {
      continue;
    }
    const writeVal = templateWorkerCodeWriteValue_(parsed);
    const cur = vals[i][0];
    const same = (typeof cur === 'number' && isFinite(cur) && Number(cur) === Number(parsed)) ||
      normalize_(cur) === parsed;
    if (!same) {
      vals[i][0] = writeVal;
      changed = true;
    }
    applyOpenListValidation_(sh.getRange(from + i, col), labels.concat([parsed]));
  }
  if (changed) {
    cells.setValues(vals);
  }
}

function applyInvoiceTemplateRowDropdowns_(sh, fromRow, toRow) {
  if (!sh) {
    return;
  }
  const cols = invoiceTemplateHeaderMap_(sh);
  const last = Math.max(sh.getLastRow(), 2);
  const from = Math.max(fromRow || 2, 2);
  const to = Math.min(toRow || last, last);
  if (to < from) {
    return;
  }
  const height = to - from + 1;
  const depts = cols.dept ? sh.getRange(from, cols.dept, height, 1).getValues() : [];
  const partMajors = cols.partMajor ? sh.getRange(from, cols.partMajor, height, 1).getValues() : [];
  for (let i = 0; i < height; i++) {
    const row = from + i;
    if (cols.serviceType) {
      applyOpenListValidation_(
        sh.getRange(row, cols.serviceType),
        templateServiceTypeChoices_(depts[i] ? depts[i][0] : '')
      );
    }
    if (cols.partMid) {
      applyOpenListValidation_(
        sh.getRange(row, cols.partMid),
        templatePartMidChoices_(partMajors[i] ? partMajors[i][0] : '')
      );
    }
  }
}

function invoiceTemplateOrderStep_() {
  return (CONFIG.listRefresh && CONFIG.listRefresh.orderStep) || 10;
}

/**
 * 図形のボタンに割り当てる。空の順番を埋め、同じテンプレート名を固めて並べ替える。
 */
function refreshInvoiceTemplateList() {
  const sh = findInvoiceTemplateSheet_();
  if (!sh) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'シート「' + CONFIG.invoiceTemplate.sheetName + '」がありません',
      '請求書入力',
      5
    );
    return;
  }
  writeInternal_(function () {
    refreshInvoiceTemplateList_(sh);
  });
  SpreadsheetApp.getActiveSpreadsheet().toast(
    sh.getName() + ' を順番で並べ替えました',
    '請求書入力',
    5
  );
}

function refreshInvoiceTemplateList_(sh) {
  if (!sh) {
    return;
  }
  ensureInvoiceTemplateHeaderCols_(sh);
  assignMissingTemplateOrders_(sh);
  sortTemplateListRows_(sh);
  applyInvoiceTemplateDropdowns_(sh);
}

function assignMissingTemplateOrders_(sh) {
  const orderCol = ensureInvoiceTemplateOrderCol_(sh);
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) {
    return false;
  }
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const cols = resolveColumns_(headers, CONFIG.invoiceTemplate.headers);
  const nameCol = cols.name || 1;
  const height = lastRow - 1;
  const data = sh.getRange(2, 1, height, lastCol).getValues();
  const formulas = sh.getRange(2, 1, height, lastCol).getFormulas();
  const orders = sh.getRange(2, orderCol, height, 1).getValues();
  const step = invoiceTemplateOrderStep_();
  const groups = {};
  const keys = [];
  let globalMax = 0;
  for (let i = 0; i < height; i++) {
    const n = toOrderNumber_(orders[i][0]);
    if (isFinite(n) && n > globalMax) {
      globalMax = n;
    }
    if (listRowIsEmpty_(data[i], formulas[i])) {
      continue;
    }
    const name = normalize_(data[i][nameCol - 1]);
    if (!name) {
      continue;
    }
    if (!groups[name]) {
      groups[name] = [];
      keys.push(name);
    }
    groups[name].push(i);
  }
  const out = orders.map(function (row) {
    return [row[0]];
  });
  let changed = false;
  keys.forEach(function (name) {
    const idxs = groups[name];
    let maxOrd = 0;
    let hasFilled = false;
    idxs.forEach(function (i) {
      const n = toOrderNumber_(out[i][0]);
      if (isFinite(n)) {
        hasFilled = true;
        if (n > maxOrd) {
          maxOrd = n;
        }
      }
    });
    let next = hasFilled ? maxOrd + step : (globalMax > 0 ? globalMax + step : step);
    idxs.forEach(function (i) {
      const n = toOrderNumber_(out[i][0]);
      if (isFinite(n)) {
        if (n > globalMax) {
          globalMax = n;
        }
        return;
      }
      out[i][0] = next;
      changed = true;
      if (next > globalMax) {
        globalMax = next;
      }
      next += step;
    });
  });
  if (changed) {
    sh.getRange(2, orderCol, height, 1).setValues(out);
  }
  return changed;
}

function sortTemplateListRows_(sh) {
  const headerRow = 1;
  const lastRow = sh.getLastRow();
  if (lastRow <= headerRow) {
    return;
  }
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const dataCol = lastDataHeaderCol_(headers);
  const cols = resolveColumns_(headers, CONFIG.invoiceTemplate.headers);
  const nameCol = cols.name || 1;
  const orderCol = findOrderCol_(headers) || dataCol;
  const height = lastRow - headerRow;
  const range = sh.getRange(headerRow + 1, 1, height, dataCol);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const formats = range.getNumberFormats();
  const emptyRows = [];
  const groups = {};
  const groupKeys = [];
  for (let i = 0; i < height; i++) {
    const rec = {
      order: orderCol <= dataCol ? values[i][orderCol - 1] : '',
      sourceIndex: i,
      values: values[i],
      formulas: formulas[i],
      formats: formats[i],
      name: normalize_(values[i][nameCol - 1])
    };
    if (listRowIsEmpty_(values[i], formulas[i])) {
      emptyRows.push(rec);
      continue;
    }
    const key = rec.name || ('\0' + i);
    if (!groups[key]) {
      groups[key] = [];
      groupKeys.push(key);
    }
    groups[key].push(rec);
  }
  const metas = groupKeys.map(function (key) {
    const rows = groups[key];
    let minOrder = Number.POSITIVE_INFINITY;
    let minIdx = rows[0].sourceIndex;
    rows.forEach(function (r) {
      const o = toOrderNumber_(r.order);
      if (o < minOrder) {
        minOrder = o;
      }
      if (r.sourceIndex < minIdx) {
        minIdx = r.sourceIndex;
      }
    });
    rows.sort(function (a, b) {
      const oa = toOrderNumber_(a.order);
      const ob = toOrderNumber_(b.order);
      if (oa !== ob) {
        return oa - ob;
      }
      return a.sourceIndex - b.sourceIndex;
    });
    return { order: minOrder, idx: minIdx, rows: rows };
  });
  metas.sort(function (a, b) {
    if (a.order !== b.order) {
      return a.order - b.order;
    }
    return a.idx - b.idx;
  });
  const finalRows = [];
  metas.forEach(function (m) {
    m.rows.forEach(function (r) {
      finalRows.push(r);
    });
  });
  emptyRows.forEach(function (r) {
    finalRows.push(r);
  });
  const outValues = [];
  const outFormats = [];
  for (let r = 0; r < finalRows.length; r++) {
    const valueRow = [];
    for (let c = 0; c < dataCol; c++) {
      const f = finalRows[r].formulas[c];
      valueRow.push(f ? f : finalRows[r].values[c]);
    }
    outValues.push(valueRow);
    outFormats.push(finalRows[r].formats);
  }
  range.setValues(outValues);
  range.setNumberFormats(outFormats);
}

function invoiceTemplateItemHasContent_(it) {
  if (!it) {
    return false;
  }
  return !!(normalize_(it.major) || normalize_(it.mid) || isFilled_(it.fee) ||
    normalize_(it.workerCode) || normalize_(it.partMajor) || normalize_(it.partMid) ||
    isFilled_(it.qty) || isFilled_(it.unitPrice));
}

function invoiceTemplateNumericCell_(v) {
  if (v === '' || v == null) {
    return '';
  }
  if (typeof v === 'number' && isFinite(v)) {
    return v;
  }
  const n = Number(String(v).replace(/,/g, '').trim());
  return isFinite(n) ? n : '';
}

function invoiceTemplateSetCell_(row, col1, val) {
  if (!col1) {
    return;
  }
  row[col1 - 1] = val == null ? '' : val;
}

function invoiceTemplateWorkerWriteFast_(raw) {
  if (raw === '' || raw == null) {
    return '';
  }
  if (typeof raw === 'number' && isFinite(raw)) {
    return Math.round(raw);
  }
  const s = normalize_(raw);
  if (!s) {
    return '';
  }
  const m = s.match(/^(\d+)/);
  return m ? templateWorkerCodeWriteValue_(m[1]) : s;
}

function invoiceTemplateScanNameOrders_(block, cols) {
  const nameCol = cols.name || 1;
  const orderCol = cols.order || 0;
  const byName = {};
  let maxOrder = 0;
  const emptyIdx = [];
  for (let i = 0; i < (block || []).length; i++) {
    const name = normalize_(block[i][nameCol - 1]);
    if (!name) {
      continue;
    }
    const o = orderCol ? toOrderNumber_(block[i][orderCol - 1]) : Number.POSITIVE_INFINITY;
    if (isFinite(o)) {
      if (o > maxOrder) {
        maxOrder = o;
      }
    } else {
      emptyIdx.push(i);
    }
    if (!byName[name]) {
      byName[name] = { rows: [], minOrder: Number.POSITIVE_INFINITY };
    }
    byName[name].rows.push(i + 2);
    if (o < byName[name].minOrder) {
      byName[name].minOrder = o;
    }
  }
  return { byName: byName, maxOrder: maxOrder, emptyIdx: emptyIdx };
}

function invoiceTemplateFillEmptyOrdersInBlock_(block, cols, emptyIdx, maxOrder) {
  const orderCol = cols.order;
  const step = invoiceTemplateOrderStep_();
  if (!orderCol || !emptyIdx.length) {
    return maxOrder;
  }
  let next = maxOrder > 0 ? maxOrder + step : step;
  emptyIdx.forEach(function (i) {
    block[i][orderCol - 1] = next;
    if (next > maxOrder) {
      maxOrder = next;
    }
    next += step;
  });
  return maxOrder;
}

function writeInvoiceTemplateBlock_(sh, existingRows, body, width) {
  const nNew = body.length;
  if (!existingRows || !existingRows.length) {
    let start = Math.max(sh.getLastRow(), 1) + 1;
    if (start < 2) {
      start = 2;
    }
    sh.getRange(start, 1, nNew, width).setValues(body);
    return;
  }
  const first = existingRows[0];
  const last = existingRows[existingRows.length - 1];
  const nOld = existingRows.length;
  let contiguous = last - first + 1 === nOld;
  if (contiguous) {
    for (let i = 1; i < nOld; i++) {
      if (existingRows[i] !== first + i) {
        contiguous = false;
        break;
      }
    }
  }
  if (contiguous) {
    if (nNew > nOld) {
      sh.insertRowsAfter(last, nNew - nOld);
    } else if (nNew < nOld) {
      sh.deleteRows(first + nNew, nOld - nNew);
    }
    sh.getRange(first, 1, nNew, width).setValues(body);
    return;
  }
  for (let i = nOld - 1; i >= 0; i--) {
    sh.deleteRow(existingRows[i]);
  }
  const after = sh.getLastRow();
  if (first <= after) {
    sh.insertRowsAfter(first - 1, nNew);
  }
  sh.getRange(first, 1, nNew, width).setValues(body);
}

function invoiceTemplateBodyRows_(cols, lastCol, name, header, items, startOrder, step) {
  const lines = items && items.length ? items : [{}];
  const dept = normalize_(header && header.dept);
  const serviceType = normalize_(header && header.serviceType);
  return lines.map(function (item, i) {
    const row = [];
    for (let c = 0; c < lastCol; c++) {
      row.push('');
    }
    invoiceTemplateSetCell_(row, cols.name, name);
    if (i === 0) {
      invoiceTemplateSetCell_(row, cols.dept, dept);
      invoiceTemplateSetCell_(row, cols.serviceType, serviceType);
    }
    invoiceTemplateSetCell_(row, cols.major, normalize_(item.major));
    invoiceTemplateSetCell_(row, cols.mid, normalize_(item.mid));
    invoiceTemplateSetCell_(row, cols.fee, invoiceTemplateNumericCell_(item.fee));
    invoiceTemplateSetCell_(row, cols.workerCode, invoiceTemplateWorkerWriteFast_(item.workerCode));
    invoiceTemplateSetCell_(row, cols.partMajor, normalize_(item.partMajor));
    invoiceTemplateSetCell_(row, cols.partMid, normalize_(item.partMid));
    invoiceTemplateSetCell_(row, cols.unitPrice, invoiceTemplateNumericCell_(item.unitPrice));
    invoiceTemplateSetCell_(row, cols.qty, invoiceTemplateNumericCell_(item.qty));
    const disc = invoiceTemplateNumericCell_(item.discYen);
    invoiceTemplateSetCell_(row, cols.discYen, disc === 0 ? '' : disc);
    invoiceTemplateSetCell_(row, cols.order, startOrder + i * step);
    return row;
  });
}

/**
 * 入力アプリから呼ぶ。現在の明細と整備部門・整備種別をテンプレートリストへ書く。
 *
 * @param {{name: string, header: object, items: object[]}} payload
 * @return {{ok: boolean, name: string, overwritten: boolean, names: string[]}}
 */
function saveInvoiceTemplate(payload) {
  payload = payload || {};
  const name = normalize_(payload.name);
  if (!name) {
    throw new Error('テンプレート名を入力してください。');
  }
  const header = payload.header || {};
  const items = (payload.items || []).filter(invoiceTemplateItemHasContent_);
  let sh = findInvoiceTemplateSheet_();
  if (!sh) {
    sh = ensureInvoiceTemplateSheet_();
  }
  let overwritten = false;
  let names = [];
  writeInternal_(function () {
    const lastCol = Math.max(sh.getLastColumn(), 1);
    let headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    let cols = resolveColumns_(headers, CONFIG.invoiceTemplate.headers);
    if (!cols.order) {
      ensureInvoiceTemplateOrderCol_(sh);
      headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
      cols = resolveColumns_(headers, CONFIG.invoiceTemplate.headers);
    }
    const width = lastDataHeaderCol_(headers);
    const last = sh.getLastRow();
    const block = last >= 2 ? sh.getRange(2, 1, last - 1, width).getValues() : [];
    let scan = invoiceTemplateScanNameOrders_(block, cols);
    if (scan.emptyIdx.length && cols.order) {
      invoiceTemplateFillEmptyOrdersInBlock_(block, cols, scan.emptyIdx, scan.maxOrder);
      sh.getRange(2, cols.order, block.length, 1).setValues(block.map(function (row) {
        return [row[cols.order - 1]];
      }));
      scan = invoiceTemplateScanNameOrders_(block, cols);
    }
    const found = scan.byName[name] || { rows: [], minOrder: Number.POSITIVE_INFINITY };
    overwritten = found.rows.length > 0;
    const step = invoiceTemplateOrderStep_();
    const startOrder = overwritten && isFinite(found.minOrder)
      ? found.minOrder
      : (scan.maxOrder > 0 ? scan.maxOrder + step : step);
    const body = invoiceTemplateBodyRows_(cols, width, name, header, items, startOrder, step);
    writeInvoiceTemplateBlock_(sh, found.rows, body, width);
    names = listInvoiceTemplateNamesFast_();
  });
  return invoiceJsonSafe_({
    ok: true,
    name: name,
    overwritten: overwritten,
    names: names
  });
}

