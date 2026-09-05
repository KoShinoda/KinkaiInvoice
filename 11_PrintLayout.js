/**
 * A4 印刷原本。1 シートにページを縦積みする。
 * 1 枚目だけヘッダー、最終枚だけフッター、各ページ右上に No.。
 * 大量印刷向けに色は使わない。
 */

/** 列幅（px）。A4 印刷幅に収まる合計。
 * No / 作業者 / 数量 … 999 まで
 * 技術料 / 単価 / 金額 … 999,999 まで
 * 残りは作業内容を優先し、次に部品。
 */
var PRINT_COL_WIDTHS_ = [32, 270, 70, 36, 132, 32, 70, 70];
var PRINT_COL_HEADERS_ = ['No', '作業内容', '技術料', '作業者', '部品', '数量', '単価', '金額'];
var PRINT_YEN_FORMAT_ = '#,##0';
var PRINT_BLACK_ = '#000000';
/** A4 縦（余白込み）に近いピクセル。明細行に余りを振る。 */
var PRINT_PAGE_PX_ = 1000;
var PRINT_FONT_MAX_ = 12;
var PRINT_FONT_MIN_ = 6;

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
    name + ' を作成しました（A4 ' + built.pageCount + ' 枚分を 1 シート）。暫定データです。',
    '請求書入力',
    8
  );
}

function makePrintSamplePayload_() {
  const pack = collectPrintSampleCatalog_();
  const mids = pack.mids;
  const parts = pack.parts;
  const items = [];
  let techSub = 0;
  let partSub = 0;
  const n = 60;
  for (let i = 0; i < n; i++) {
    const fee = i % 7 === 0 ? 0 : 800 + (i % 12) * 200;
    const qty = 1 + (i % 3 === 0 ? 1 : 0);
    const unitPrice = 400 + (i % 15) * 80;
    const amount = qty * unitPrice;
    techSub += fee;
    partSub += amount;
    items.push({
      mid: mids[i % mids.length],
      fee: fee || '',
      workerCode: String((i % 9) + 1),
      partMid: parts[i % parts.length],
      qty: qty,
      unitPrice: unitPrice,
      amount: amount
    });
  }
  const techPct = 3;
  const partPct = 10;
  const techDisc = Math.round(techSub * techPct / 100);
  const partDisc = Math.round(partSub * partPct / 100);
  return {
    header: {
      userName: '近海請求書',
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
      techPct: techPct,
      techTotal: techSub - techDisc,
      partSub: partSub,
      partDisc: partDisc,
      partPct: partPct,
      partTotal: partSub - partDisc,
      grand: techSub - techDisc + partSub - partDisc
    }
  };
}

function collectPrintSampleCatalog_() {
  const mids = [];
  const parts = [];
  const seenMid = {};
  const seenPart = {};
  function addMid(v) {
    const s = String(v || '').trim();
    if (!s || seenMid[s]) {
      return;
    }
    seenMid[s] = true;
    mids.push(s);
  }
  function addPart(v) {
    const s = String(v || '').trim();
    if (!s || seenPart[s]) {
      return;
    }
    seenPart[s] = true;
    parts.push(s);
  }
  try {
    const ctx = loadContext_();
    (ctx.workRows || []).forEach(function (r) {
      addMid(r.mid);
      if (r.content && String(r.content).trim() !== String(r.mid || '').trim()) {
        addMid(r.content);
      }
      addPart(r.partMid);
    });
  } catch (err) {
    Logger.log('%s sample catalog: %s', CONFIG.logPrefix, err);
  }
  PRINT_SAMPLE_MIDS_.forEach(addMid);
  PRINT_SAMPLE_PARTS_.forEach(addPart);
  if (!mids.length) {
    PRINT_SAMPLE_MIDS_.forEach(addMid);
  }
  if (!parts.length) {
    PRINT_SAMPLE_PARTS_.forEach(addPart);
  }
  return { mids: mids, parts: parts };
}

var PRINT_SAMPLE_MIDS_ = [
  '＊＊　１２カ月定期点検　＊＊',
  'シャシ洗浄、グリスアップ',
  'シャシグレー塗装',
  'シャシマスキング',
  '保安確認検査料',
  '代行料',
  '構造変更分解整備',
  'リレーバルブＡＳＳＹ脱着',
  'エアーカプラゴム交換',
  'ホイールナット規定トルク締め付一式',
  'タイヤ空気圧点検、調整（９ｋｇｆ）',
  'エアサス廻り点検締め付け',
  '左右ランディングASSY交換',
  'ウイングシャワーテスト',
  'ウイング開閉点検',
  '左右車幅灯交換',
  'バックランプ交換',
  'ナンバー灯交換',
  '左右Ｒ２エアサスブラケット当板補強修理',
  'リヤバンパー上部リフレクター取付ブラケット製作交換',
  '左右サイドバンパー製作交換',
  '荷台アオリ支柱（４本）補強修理',
  '＊＊　１２カ月定期点検　＊＊\nシャシ廻り一式',
  'ランディング高さ違い点検修理\n（ギヤ灯高・低切り替え不良）',
  '左右メーンフレーム当板修理およびリヤロッカーレール切断曲がり修理（６か所）当板補強一式'
];

var PRINT_SAMPLE_PARTS_ = [
  'カートリッジグリス',
  'ｼｬｼｸﾞﾚｰ',
  'ｽﾓｰﾙ･ﾊﾟｰﾂ',
  '産業廃棄物処理料',
  'ＢＰＷ用ハブＢ／ｇグリス',
  'ハブパッキン',
  'ハブＢ／ｇグリス',
  'エアカプラーゴム',
  'リレーバルブインナーキット',
  'リレーバルブアッパーカバー（６H）',
  'エアーパイプジョイント',
  '左ランディングASSY',
  '右ランディングＡＳＳＹ',
  'サンドシュー',
  'ランバーメイト',
  'ナンバープレートブラケット',
  'アオリスケットASSY',
  'カーゴロック',
  '車幅灯（LED）',
  'サイドウインカーランプ',
  'テールランプレンズ',
  'リヤウインカーランプレンズ',
  'リフレクター',
  'バックランプ',
  'バックブザー',
  'マーカーランプ',
  'ナンバー灯',
  '三角リフレクター',
  '後部大型反射器',
  'コーキング',
  'アルミリベット',
  '鋼材一式',
  '10×65ボルト',
  '16×35ボルト'
];

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
  const slotH = printDataRowHeight_();

  sheet.setHiddenGridlines(true);

  for (let c = 0; c < CONFIG.print.colCount; c++) {
    sheet.setColumnWidth(c + 1, PRINT_COL_WIDTHS_[c]);
  }

  let cursor = 1;
  const breakRows = [];
  for (let p = 0; p < pageCount; p++) {
    const slice = items.slice(p * per, p * per + per);
    const used = fillPrintPage_(sheet, cursor, payload.header || {}, slice, {
      page: p + 1,
      pageCount: pageCount,
      serialOffset: p * per,
      showHeader: p === 0,
      showFooter: p === pageCount - 1,
      summary: payload.summary || {},
      slotH: slotH
    });
    const end = cursor + used - 1;
    if (p < pageCount - 1) {
      breakRows.push(end);
    }
    cursor = end + 1;
  }

  trimPrintSheet_(sheet, cursor - 1);
  applyA4PageSetup_(sheet, pageCount);
  applyPrintPageBreaksAt_(sheet, breakRows);
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
    ps.setTopMargin(0.28);
    ps.setBottomMargin(0.28);
    ps.setLeftMargin(0.32);
    ps.setRightMargin(0.32);
    if (typeof ps.setScale === 'function') {
      ps.setScale(100);
    }
  } catch (err) {
    Logger.log('%s A4 page setup: %s', CONFIG.logPrefix, err);
  }
}

/**
 * 一部の GAS には setRowPageBreak が無い。あるときだけ使い、無ければ FitToHeight で枚数を合わせる。
 */
function applyPrintPageBreaksAt_(sheet, breakRows) {
  if (typeof sheet.setRowPageBreak !== 'function') {
    return;
  }
  for (let i = 0; i < breakRows.length; i++) {
    sheet.setRowPageBreak(breakRows[i], true);
  }
}

function printPageLayout_(showHeader, showFooter) {
  let r = 0;
  const L = { title: r };
  r++;
  if (showHeader) {
    L.metaL1 = r;
    r++;
    L.metaV1 = r;
    r++;
    L.metaL2 = r;
    r++;
    L.metaV2 = r;
    r++;
    L.spacer = r;
    r++;
  }
  L.colHead = r;
  r++;
  L.firstLine = r;
  r += CONFIG.print.linesPerPage;
  if (showFooter) {
    L.footerStart = r;
    r += 5;
  }
  L.pageRows = r;
  return L;
}

function printDataRowHeight_() {
  const titleH = 32;
  const metaH = 18 + 24 + 18 + 24;
  const spacerH = 4;
  const colHeadH = 22;
  const footerH = 20 * 5;
  const chrome = titleH + metaH + spacerH + colHeadH + footerH;
  return Math.max(18, Math.floor((PRINT_PAGE_PX_ - chrome) / CONFIG.print.linesPerPage));
}

function fillPrintPage_(sheet, start, header, lines, opts) {
  const showHeader = !!opts.showHeader;
  const showFooter = !!opts.showFooter;
  const L = printPageLayout_(showHeader, showFooter);
  const cols = CONFIG.print.colCount;
  const slotH = opts.slotH || printDataRowHeight_();

  sheet.getRange(start, 1, L.pageRows, cols)
    .setFontFamily('Yu Gothic')
    .setFontSize(PRINT_FONT_MAX_)
    .setFontColor(PRINT_BLACK_)
    .setVerticalAlignment('middle')
    .setWrap(false);

  applyPrintPageHeights_(sheet, start, L, showHeader, showFooter, slotH);
  mergePrintPage_(sheet, start, L, showHeader, showFooter);

  const titleRow = start + L.title;
  const title = String(header.userName || CONFIG.print.title || '近海請求書').trim() || '近海請求書';
  sheet.getRange(titleRow, 1).setValue(showHeader ? title : '');
  sheet.getRange(titleRow, 1, 1, 4)
    .setFontSize(20)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');
  sheet.getRange(titleRow, 5).setValue('No.' + opts.page + '／' + opts.pageCount);
  sheet.getRange(titleRow, 5, 1, 4)
    .setFontSize(12)
    .setFontWeight('bold')
    .setHorizontalAlignment('right');

  if (showHeader) {
    fillPrintHeader_(sheet, start, L, header);
  }

  const headRow = start + L.colHead;
  sheet.getRange(headRow, 1, 1, cols)
    .setValues([PRINT_COL_HEADERS_])
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, PRINT_BLACK_, SpreadsheetApp.BorderStyle.SOLID);

  const first = start + L.firstLine;
  fillPrintBody_(sheet, first, lines, opts.serialOffset || 0);

  if (showFooter) {
    fillPrintFooterBlock_(sheet, start + L.footerStart, opts.summary || {});
  }

  sheet.getRange(start, 1, L.pageRows, cols).setVerticalAlignment('middle');
  fillPrintBodyFonts_(sheet, first, lines);
  return L.pageRows;
}

function fillPrintBody_(sheet, first, lines, serialOffset) {
  const per = CONFIG.print.linesPerPage;
  const cols = CONFIG.print.colCount;
  const empty = ['', '', '', '', '', '', '', ''];
  const body = [];
  for (let i = 0; i < per; i++) {
    const it = lines[i];
    if (!it) {
      body.push(empty.slice());
      continue;
    }
    const qty = toNumberOrBlank_(it.qty);
    const price = toNumberOrBlank_(it.unitPrice);
    const amount = lineAmount_(it, qty, price);
    body.push([
      serialOffset + i + 1,
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
  bodyRange.setBorder(true, true, true, true, true, true, PRINT_BLACK_, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(first, 1, per, 1).setHorizontalAlignment('center').setWrap(false);
  sheet.getRange(first, 2, per, 1).setWrap(false);
  sheet.getRange(first, 3, per, 1).setNumberFormat(PRINT_YEN_FORMAT_).setHorizontalAlignment('right');
  sheet.getRange(first, 4, per, 1).setHorizontalAlignment('center').setWrap(false);
  sheet.getRange(first, 5, per, 1).setWrap(false);
  sheet.getRange(first, 6, per, 1).setHorizontalAlignment('center').setWrap(false);
  sheet.getRange(first, 7, per, 2).setNumberFormat(PRINT_YEN_FORMAT_).setHorizontalAlignment('right');
  fillPrintBodyFonts_(sheet, first, lines);
}

function fillPrintBodyFonts_(sheet, first, lines) {
  const per = CONFIG.print.linesPerPage;
  const baseH = sheet.getRowHeight(first);
  for (let i = 0; i < per; i++) {
    const it = lines[i];
    if (!it) {
      continue;
    }
    const work = it.mid || it.name || '';
    const part = it.partMid || it.part || '';
    const workBreak = hasPrintNewline_(work);
    const partBreak = hasPrintNewline_(part);
    const workCell = sheet.getRange(first + i, 2);
    const partCell = sheet.getRange(first + i, 5);
    workCell.setFontSize(fitPrintFont_(work, PRINT_COL_WIDTHS_[1], workBreak ? newlineCount_(work) : 1));
    partCell.setFontSize(fitPrintFont_(part, PRINT_COL_WIDTHS_[4], partBreak ? newlineCount_(part) : 1));
    workCell.setWrap(workBreak);
    partCell.setWrap(partBreak);
    const linesN = Math.max(workBreak ? newlineCount_(work) : 1, partBreak ? newlineCount_(part) : 1);
    if (linesN > 1) {
      sheet.setRowHeight(first + i, Math.round(baseH * 1.4));
    }
  }
}

function hasPrintNewline_(text) {
  return /[\r\n]/.test(String(text || ''));
}

function newlineCount_(text) {
  const s = String(text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  if (!s) {
    return 1;
  }
  return s.split('\n').length;
}

function printCharPx_(fontPt) {
  return fontPt * (96 / 72);
}

function printColCapacity_(colWidth, fontPt) {
  const inner = Math.max(8, colWidth - 12);
  return inner / printCharPx_(fontPt);
}

function displayUnits_(text) {
  const s = String(text || '').replace(/\n/g, '');
  let u = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c <= 0x7f) {
      u += 0.72;
    } else if (c >= 0xff61 && c <= 0xff9f) {
      u += 0.7;
    } else {
      u += 1;
    }
  }
  return u;
}

function fitPrintFont_(text, colWidth, lines) {
  const raw = String(text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  if (!raw) {
    return PRINT_FONT_MAX_;
  }
  const parts = raw.split('\n');
  let units = 0;
  for (let i = 0; i < parts.length; i++) {
    units = Math.max(units, displayUnits_(parts[i]));
  }
  const lineCount = lines > 1 ? lines : 1;
  if (units < 0.5) {
    return PRINT_FONT_MAX_;
  }
  units += 0.35;
  for (let pt = PRINT_FONT_MAX_; pt >= PRINT_FONT_MIN_; pt--) {
    if (units <= printColCapacity_(colWidth, pt) * lineCount) {
      return pt;
    }
  }
  return PRINT_FONT_MIN_;
}

function applyPrintPageHeights_(sheet, start, L, showHeader, showFooter, slotH) {
  const titleH = 32;
  const metaLH = 18;
  const metaVH = 24;
  const spacerH = 4;
  const colHeadH = 22;
  const footerH = 20;
  sheet.setRowHeight(start + L.title, titleH);
  if (showHeader) {
    sheet.setRowHeight(start + L.metaL1, metaLH);
    sheet.setRowHeight(start + L.metaV1, metaVH);
    sheet.setRowHeight(start + L.metaL2, metaLH);
    sheet.setRowHeight(start + L.metaV2, metaVH);
    sheet.setRowHeight(start + L.spacer, spacerH);
  }
  sheet.setRowHeight(start + L.colHead, colHeadH);
  sheet.setRowHeights(start + L.firstLine, CONFIG.print.linesPerPage, slotH);
  if (showFooter) {
    sheet.setRowHeights(start + L.footerStart, 5, footerH);
  }
}

function mergePrintPage_(sheet, start, L, showHeader, showFooter) {
  sheet.getRange(start + L.title, 1, 1, 4).merge();
  sheet.getRange(start + L.title, 5, 1, 4).merge();
  if (showHeader) {
    [L.metaL1, L.metaV1, L.metaL2, L.metaV2].forEach(function (off) {
      sheet.getRange(start + off, 1, 1, 2).merge();
      sheet.getRange(start + off, 3, 1, 2).merge();
      sheet.getRange(start + off, 5, 1, 2).merge();
    });
  }
  if (showFooter) {
    const f = start + L.footerStart;
    sheet.getRange(f, 1, 1, 4).merge();
    sheet.getRange(f, 5, 1, 4).merge();
    for (let i = 1; i <= 3; i++) {
      sheet.getRange(f + i, 1, 1, 2).merge();
      sheet.getRange(f + i, 3, 1, 2).merge();
      sheet.getRange(f + i, 5, 1, 2).merge();
      sheet.getRange(f + i, 7, 1, 2).merge();
    }
    sheet.getRange(f + 4, 1, 1, 6).merge();
    sheet.getRange(f + 4, 7, 1, 2).merge();
  }
}

function fillPrintHeader_(sheet, start, L, header) {
  const l1 = start + L.metaL1;
  const v1 = start + L.metaV1;
  const l2 = start + L.metaL2;
  const v2 = start + L.metaV2;

  sheet.getRange(l1, 1).setValue('K-No').setFontSize(12).setFontWeight('bold');
  sheet.getRange(l1, 3).setValue('登録番号').setFontSize(12).setFontWeight('bold');
  sheet.getRange(l1, 5).setValue('受付').setFontSize(12).setFontWeight('bold');
  sheet.getRange(v1, 1).setValue(header.kNo || '').setFontSize(12);
  sheet.getRange(v1, 3).setValue(header.plate || '').setFontSize(12);
  sheet.getRange(v1, 5).setValue(header.receptionist || header.staff || '').setFontSize(12);

  sheet.getRange(l2, 1).setValue('入庫日').setFontSize(12).setFontWeight('bold');
  sheet.getRange(l2, 3).setValue('出庫日').setFontSize(12).setFontWeight('bold');
  sheet.getRange(l2, 5).setValue('請求日').setFontSize(12).setFontWeight('bold');
  sheet.getRange(v2, 1).setValue(header.inDate || '').setFontSize(12);
  sheet.getRange(v2, 3).setValue(header.outDate || header.doneDate || '').setFontSize(12);
  sheet.getRange(v2, 5).setValue(header.billDate || '').setFontSize(12);

  sheet.getRange(l1, 1, 4, 6)
    .setHorizontalAlignment('left')
    .setBorder(true, true, true, true, false, false, PRINT_BLACK_, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(v1, 1, 1, 6)
    .setBorder(null, null, true, null, false, false, PRINT_BLACK_, SpreadsheetApp.BorderStyle.SOLID);
}

function fillPrintFooterBlock_(sheet, f, summary) {
  const techPct = printPct_(summary.techSub, summary.techDisc, summary.techPct);
  const partPct = printPct_(summary.partSub, summary.partDisc, summary.partPct);

  sheet.getRange(f, 1).setValue('技術').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  sheet.getRange(f, 5).setValue('部品').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

  sheet.getRange(f + 1, 1).setValue('合計').setFontWeight('bold');
  sheet.getRange(f + 1, 3).setValue(blankIfEmpty_(summary.techSub)).setNumberFormat(PRINT_YEN_FORMAT_).setHorizontalAlignment('right');
  sheet.getRange(f + 1, 5).setValue('合計').setFontWeight('bold');
  sheet.getRange(f + 1, 7).setValue(blankIfEmpty_(summary.partSub)).setNumberFormat(PRINT_YEN_FORMAT_).setHorizontalAlignment('right');

  sheet.getRange(f + 2, 1).setValue(discLabel_(techPct)).setFontWeight('bold');
  sheet.getRange(f + 2, 3).setValue(blankIfEmpty_(summary.techDisc)).setNumberFormat(PRINT_YEN_FORMAT_).setHorizontalAlignment('right');
  sheet.getRange(f + 2, 5).setValue(discLabel_(partPct)).setFontWeight('bold');
  sheet.getRange(f + 2, 7).setValue(blankIfEmpty_(summary.partDisc)).setNumberFormat(PRINT_YEN_FORMAT_).setHorizontalAlignment('right');

  sheet.getRange(f + 3, 1).setValue('値引後').setFontWeight('bold');
  sheet.getRange(f + 3, 3).setValue(blankIfEmpty_(summary.techTotal)).setNumberFormat(PRINT_YEN_FORMAT_).setHorizontalAlignment('right');
  sheet.getRange(f + 3, 5).setValue('値引後').setFontWeight('bold');
  sheet.getRange(f + 3, 7).setValue(blankIfEmpty_(summary.partTotal)).setNumberFormat(PRINT_YEN_FORMAT_).setHorizontalAlignment('right');

  sheet.getRange(f + 4, 1).setValue('合計').setFontWeight('bold').setFontSize(18).setHorizontalAlignment('right');
  sheet.getRange(f + 4, 7)
    .setValue(blankIfEmpty_(summary.grand))
    .setFontSize(18)
    .setFontWeight('bold')
    .setNumberFormat(PRINT_YEN_FORMAT_)
    .setHorizontalAlignment('right');

  sheet.getRange(f, 1, 5, 8)
    .setBorder(true, true, true, true, true, true, PRINT_BLACK_, SpreadsheetApp.BorderStyle.SOLID);
}

function printPct_(sub, disc, given) {
  if (given !== undefined && given !== null && given !== '') {
    const n = Number(given);
    return isFinite(n) ? n : 0;
  }
  const s = Number(sub);
  const d = Number(disc);
  if (!s || !isFinite(s) || !isFinite(d)) {
    return 0;
  }
  return Math.round(d / s * 100);
}

function discLabel_(pct) {
  if (pct === undefined || pct === null || pct === '') {
    return '値引額（％）';
  }
  return '値引額（' + pct + '％）';
}

function blankIfEmpty_(value) {
  if (value === undefined || value === null || value === '') {
    return '';
  }
  return value;
}
