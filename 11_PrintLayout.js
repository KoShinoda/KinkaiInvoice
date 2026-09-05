/**
 * A4 印刷原本。車検原紙は使わず、1 シートにページを縦積みする。
 * 1 枚目だけヘッダー、最終枚だけフッター、各ページ右上に No.。
 */

var PRINT_COL_WIDTHS_ = [36, 210, 78, 46, 196, 44, 72, 86];
var PRINT_COL_HEADERS_ = ['No', '作業内容', '技術料', '作業者', '部品', '数量', '単価', '金額'];
var PRINT_YEN_FORMAT_ = '#,##0';

/**
 * 車検_入力保存後の印刷シート。常に「印刷」1 枚。
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {object} payload
 * @param {string=} sheetName
 * @return {{pageCount: number, sheetNames: string[]}}
 */
function writePrintSheets_(ss, payload, sheetName) {
  const name = sheetName || CONFIG.print.sheetName;
  cleanupPrintSheets_(ss, name);
  const built = buildInvoicePrintSheet_(ss, name, payload);
  ss.setActiveSheet(built.sheet);
  return {
    pageCount: built.pageCount,
    sheetNames: [name]
  };
}

/**
 * メニュー「印刷原本（A4・1シート）を作成」。
 * 暫定明細つき。完成後にシート「印刷原本」は消してよい。
 */
function createPrintOriginalSample() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const payload = makePrintSamplePayload_();
  const name = CONFIG.print.sampleSheetName;
  cleanupPrintSheets_(ss, name);
  const built = buildInvoicePrintSheet_(ss, name, payload);
  ss.setActiveSheet(built.sheet);
  ss.toast(
    name + ' を作成しました（A4 ' + built.pageCount + ' 枚分を 1 シート）。1枚目＝ヘッダー、最終＝フッター。暫定データです。',
    '請求書入力',
    8
  );
}

function makePrintSamplePayload_() {
  const items = [];
  let techSub = 0;
  let partSub = 0;
  const mids = [
    '＊＊　１２カ月定期点検　＊＊',
    'シャシ洗浄、グリスアップ',
    'シャシグレー塗装',
    'シャシマスキング',
    '保安確認検査料',
    '代行料'
  ];
  const parts = ['部品1', '部品2', '部品3', 'カートリッジグリス', 'ｼｬｼｸﾞﾚｰ', 'ｽﾓｰﾙ･ﾊﾟｰﾂ'];
  for (let i = 0; i < 72; i++) {
    const fee = i % 7 === 0 ? 0 : 800 + (i % 12) * 200;
    const qty = 1;
    const unitPrice = 400 + (i % 9) * 50;
    const amount = qty * unitPrice;
    techSub += fee;
    partSub += amount;
    items.push({
      mid: mids[i % mids.length] + (i >= 6 ? '（暫定' + (i + 1) + '）' : ''),
      fee: fee || '',
      workerCode: String((i % 9) + 1),
      partMid: parts[i % parts.length],
      qty: qty,
      unitPrice: unitPrice,
      amount: amount
    });
  }
  const techDisc = Math.round(techSub * 0.03);
  const partDisc = Math.round(partSub * 0.1);
  return {
    header: {
      kNo: 'K-9999',
      plate: '仮-5331',
      receptionist: 'サンプル',
      inDate: '2026/09/01',
      outDate: '2026/09/05',
      billDate: '2026/09/05'
    },
    items: items,
    summary: {
      techSub: techSub,
      techDisc: techDisc,
      techTotal: techSub - techDisc,
      partSub: partSub,
      partDisc: partDisc,
      partTotal: partSub - partDisc,
      grand: techSub - techDisc + partSub - partDisc
    }
  };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} keepName
 */
function cleanupPrintSheets_(ss, keepName) {
  const prefixes = [CONFIG.print.sheetNamePrefix, CONFIG.print.samplePrefix];
  const sheets = ss.getSheets();
  const toDelete = [];
  for (let i = 0; i < sheets.length; i++) {
    const n = sheets[i].getName();
    if (n === keepName) {
      continue;
    }
    for (let p = 0; p < prefixes.length; p++) {
      if (prefixes[p] && n.indexOf(prefixes[p]) === 0) {
        toDelete.push(sheets[i]);
        break;
      }
    }
  }
  for (let i = 0; i < toDelete.length; i++) {
    if (ss.getSheets().length <= 1) {
      break;
    }
    ss.deleteSheet(toDelete[i]);
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetName
 * @param {object} payload
 * @return {{sheet: GoogleAppsScript.Spreadsheet.Sheet, pageCount: number}}
 */
function buildInvoicePrintSheet_(ss, sheetName, payload) {
  const items = (payload.items || []).filter(function (it) {
    return rowHasContent_(it);
  });
  const per = CONFIG.print.linesPerPage;
  const pageCount = Math.max(1, Math.ceil(items.length / per) || 1);
  const sheet = replacePrintSheet_(ss, sheetName);

  sheet.setHiddenGridlines(true);
  sheet.setTabColor('#1a365d');

  for (let c = 0; c < CONFIG.print.colCount; c++) {
    sheet.setColumnWidth(c + 1, PRINT_COL_WIDTHS_[c]);
  }

  for (let p = 0; p < pageCount; p++) {
    const start = 1 + p * CONFIG.print.pageRows;
    const slice = items.slice(p * per, p * per + per);
    fillPrintPage_(sheet, start, payload.header || {}, slice, {
      page: p + 1,
      pageCount: pageCount,
      serialOffset: p * per,
      showHeader: p === 0,
      showFooter: p === pageCount - 1,
      summary: payload.summary || {}
    });
  }

  trimPrintSheet_(sheet, pageCount * CONFIG.print.pageRows);
  applyA4PageSetup_(sheet, pageCount);
  for (let p = 1; p < pageCount; p++) {
    sheet.setRowPageBreak(p * CONFIG.print.pageRows, true);
  }
  return { sheet: sheet, pageCount: pageCount };
}

function replacePrintSheet_(ss, name) {
  const old = ss.getSheetByName(name);
  const temp = name + '_tmp';
  let n = temp;
  let i = 1;
  while (ss.getSheetByName(n)) {
    n = temp + i;
    i++;
  }
  const sheet = ss.insertSheet(n);
  if (old && ss.getSheets().length > 1) {
    ss.deleteSheet(old);
  }
  sheet.setName(name);
  return sheet;
}

function trimPrintSheet_(sheet, lastRow) {
  const maxR = sheet.getMaxRows();
  if (maxR > lastRow) {
    sheet.deleteRows(lastRow + 1, maxR - lastRow);
  }
  const maxC = sheet.getMaxColumns();
  const cols = CONFIG.print.colCount;
  if (maxC > cols) {
    sheet.deleteColumns(cols + 1, maxC - cols);
  }
}

function applyA4PageSetup_(sheet, pageCount) {
  try {
    const ps = sheet.getPageSetup();
    ps.setPaperSize(SpreadsheetApp.PaperSize.A4);
    ps.setOrientation(SpreadsheetApp.PageOrientation.PORTRAIT);
    ps.setPrintGridlines(false);
    ps.setTopMargin(0.39);
    ps.setBottomMargin(0.39);
    ps.setLeftMargin(0.39);
    ps.setRightMargin(0.39);
    ps.setFitToWidth(1);
    ps.setFitToHeight(pageCount);
  } catch (err) {
    Logger.log('%s A4 page setup: %s', CONFIG.logPrefix, err);
  }
}

function fillPrintPage_(sheet, start, header, lines, opts) {
  const L = CONFIG.print.layout;
  const cols = CONFIG.print.colCount;
  const per = CONFIG.print.linesPerPage;
  const pageEnd = start + CONFIG.print.pageRows - 1;

  sheet.getRange(start, 1, CONFIG.print.pageRows, cols)
    .setFontFamily('Yu Gothic')
    .setFontSize(9)
    .setVerticalAlignment('middle')
    .setWrap(true);

  applyPrintPageHeights_(sheet, start);
  mergePrintPage_(sheet, start, opts.showHeader, opts.showFooter);

  const titleRow = start + L.title;
  sheet.getRange(titleRow, 1).setValue(opts.showHeader ? CONFIG.print.title : '');
  sheet.getRange(titleRow, 1, 1, 5)
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');
  sheet.getRange(titleRow, 6).setValue('No.' + opts.page + '／' + opts.pageCount);
  sheet.getRange(titleRow, 6, 1, 3)
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('right');

  if (opts.showHeader) {
    fillPrintHeader_(sheet, start, header);
  }

  const headRow = start + L.colHead;
  sheet.getRange(headRow, 1, 1, cols)
    .setValues([PRINT_COL_HEADERS_])
    .setBackground('#2d3748')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setWrap(false);

  const first = start + L.firstLine;
  const body = [];
  for (let i = 0; i < per; i++) {
    const it = lines[i];
    if (!it) {
      body.push(['', '', '', '', '', '', '', '']);
      continue;
    }
    const qty = toNumberOrBlank_(it.qty);
    const price = toNumberOrBlank_(it.unitPrice);
    const amount = lineAmount_(it, qty, price);
    body.push([
      opts.serialOffset + i + 1,
      it.mid || it.name || '',
      it.fee === undefined || it.fee === null || it.fee === '' ? '' : it.fee,
      it.workerCode || '',
      it.partMid || it.part || '',
      qty,
      price,
      amount
    ]);
  }
  const bodyRange = sheet.getRange(first, 1, per, cols);
  bodyRange.setValues(body);
  bodyRange.setBorder(true, true, true, true, true, true, '#4a5568', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(headRow, 1, 1, cols)
    .setBorder(true, true, true, true, true, true, '#2d3748', SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(first, 1, per, 1).setHorizontalAlignment('center').setWrap(false);
  sheet.getRange(first, 3, per, 1).setNumberFormat(PRINT_YEN_FORMAT_).setHorizontalAlignment('right');
  sheet.getRange(first, 4, per, 1).setHorizontalAlignment('center').setWrap(false);
  sheet.getRange(first, 6, per, 1).setHorizontalAlignment('center').setWrap(false);
  sheet.getRange(first, 7, per, 2).setNumberFormat(PRINT_YEN_FORMAT_).setHorizontalAlignment('right');

  if (opts.showFooter) {
    fillPrintFooterBlock_(sheet, start, opts.summary || {});
  }

  sheet.getRange(start, 1, pageEnd - start + 1, cols).setVerticalAlignment('middle');
}

function applyPrintPageHeights_(sheet, start) {
  const L = CONFIG.print.layout;
  sheet.setRowHeight(start + L.title, 34);
  sheet.setRowHeight(start + L.meta1, 22);
  sheet.setRowHeight(start + L.meta2, 22);
  sheet.setRowHeight(start + L.spacer, 10);
  sheet.setRowHeight(start + L.colHead, 24);
  sheet.setRowHeights(start + L.firstLine, CONFIG.print.linesPerPage, 18);
  const footerStart = start + L.footerStart;
  const footerRows = CONFIG.print.pageRows - L.footerStart;
  sheet.setRowHeights(footerStart, footerRows, 20);
}

function mergePrintPage_(sheet, start, showHeader, showFooter) {
  const L = CONFIG.print.layout;
  sheet.getRange(start + L.title, 1, 1, 5).merge();
  sheet.getRange(start + L.title, 6, 1, 3).merge();
  if (showHeader) {
    sheet.getRange(start + L.meta1, 2, 1, 2).merge();
    sheet.getRange(start + L.meta1, 5, 1, 2).merge();
    sheet.getRange(start + L.meta2, 2, 1, 2).merge();
    sheet.getRange(start + L.meta2, 5, 1, 2).merge();
  }
  if (showFooter) {
    const f = start + L.footerStart;
    sheet.getRange(f, 1, 1, 3).merge();
    sheet.getRange(f, 4, 1, 2).merge();
    sheet.getRange(f, 6, 1, 3).merge();
    sheet.getRange(f + 4, 1, 1, 6).merge();
  }
}

function fillPrintHeader_(sheet, start, header) {
  const L = CONFIG.print.layout;
  const r1 = start + L.meta1;
  const r2 = start + L.meta2;
  const labelBg = '#edf2f7';
  sheet.getRange(r1, 1).setValue('K-No').setBackground(labelBg).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(r1, 2).setValue(header.kNo || '');
  sheet.getRange(r1, 4).setValue('登録番号').setBackground(labelBg).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(r1, 5).setValue(header.plate || '');
  sheet.getRange(r1, 7).setValue('受付').setBackground(labelBg).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(r1, 8).setValue(header.receptionist || header.staff || '');

  sheet.getRange(r2, 1).setValue('入庫日').setBackground(labelBg).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(r2, 2).setValue(header.inDate || '');
  sheet.getRange(r2, 4).setValue('出庫日').setBackground(labelBg).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(r2, 5).setValue(header.outDate || header.doneDate || '');
  sheet.getRange(r2, 7).setValue('請求日').setBackground(labelBg).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(r2, 8).setValue(header.billDate || '');

  sheet.getRange(r1, 1, 2, 8)
    .setBorder(true, true, true, true, true, true, '#a0aec0', SpreadsheetApp.BorderStyle.SOLID);
}

function fillPrintFooterBlock_(sheet, start, summary) {
  const f = start + CONFIG.print.layout.footerStart;
  const labelBg = '#edf2f7';
  sheet.getRange(f, 1).setValue('技術料').setFontWeight('bold').setHorizontalAlignment('center').setBackground(labelBg);
  sheet.getRange(f, 4).setValue('部品').setFontWeight('bold').setHorizontalAlignment('center').setBackground(labelBg);
  sheet.getRange(f, 6).setValue('ご請求').setFontWeight('bold').setHorizontalAlignment('center').setBackground(labelBg);

  sheet.getRange(f + 1, 1).setValue('小計').setBackground(labelBg);
  sheet.getRange(f + 1, 2).setValue(blankIfEmpty_(summary.techSub));
  sheet.getRange(f + 1, 4).setValue(blankIfEmpty_(summary.partSub));
  sheet.getRange(f + 2, 1).setValue('値引').setBackground(labelBg);
  sheet.getRange(f + 2, 2).setValue(blankIfEmpty_(summary.techDisc));
  sheet.getRange(f + 2, 4).setValue(blankIfEmpty_(summary.partDisc));
  sheet.getRange(f + 3, 1).setValue('計').setFontWeight('bold').setBackground(labelBg);
  sheet.getRange(f + 3, 2).setValue(blankIfEmpty_(summary.techTotal));
  sheet.getRange(f + 3, 4).setValue(blankIfEmpty_(summary.partTotal));
  sheet.getRange(f + 4, 1).setValue('ご請求金額').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('right');
  sheet.getRange(f + 4, 7).setValue(blankIfEmpty_(summary.grand));
  sheet.getRange(f + 4, 7, 1, 2)
    .merge()
    .setFontSize(14)
    .setFontWeight('bold')
    .setNumberFormat('"¥"#,##0')
    .setHorizontalAlignment('right');

  sheet.getRange(f + 1, 2, 3, 1).setNumberFormat(PRINT_YEN_FORMAT_).setHorizontalAlignment('right');
  sheet.getRange(f + 1, 4, 3, 1).setNumberFormat(PRINT_YEN_FORMAT_).setHorizontalAlignment('right');
  sheet.getRange(f, 1, 5, 8)
    .setBorder(true, true, true, true, true, true, '#4a5568', SpreadsheetApp.BorderStyle.SOLID);
}

function blankIfEmpty_(value) {
  if (value === undefined || value === null || value === '') {
    return '';
  }
  return value;
}
