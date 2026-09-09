/**
 * A4 印刷原本。1 シートにページを縦積みする。
 * 1 枚目だけヘッダー、最終枚だけフッター、各ページ右下に No.。
 * 明細と No. の間の空行でページ高さを合わせる（余白は標準）。
 * 大量印刷向けに色は使わない。
 */

/** 列幅（px）。A4・余白標準の印刷幅に合わせて縮小する。
 * No / 作業者 / 数量 … 999 まで
 * 技術料 / 単価 / 金額 … 999,999 まで
 * 残りは作業内容を優先し、次に部品。
 */
var PRINT_COL_WIDTHS_ = [32, 270, 70, 36, 132, 32, 70, 70];
var PRINT_COL_HEADERS_ = ['No', '作業内容', '技術料', '作業者', '部品', '数量', '単価', '金額'];
var PRINT_YEN_FORMAT_ = '#,##0';
var PRINT_BLACK_ = '#000000';
/** A4 縦。余白は印刷ダイアログの「標準」に合わせる（インチ）。 */
var PRINT_MARGIN_IN_ = { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7 };
/** シート／PDF のヘッダー・フッター余白。0 以外だと本文が次ページへ落ちやすい。 */
var PRINT_HF_MARGIN_IN_ = 0;
/** 明細と No. のあいだの最小空行。余りは半分だけ使う（全部使うと No. が次ページへ落ちる）。 */
var PRINT_PAD_MIN_ = 8;
var PRINT_PAD_FILL_ = 0.5;
var PRINT_PX_PER_IN_ = 96;
var PRINT_FONT_MAX_ = 12;
var PRINT_FONT_MIN_ = 6;
var PRINT_PAGE_NO_H_ = 16;
/** 明細列見出し（印刷シート 9 行目）。折返し禁止で高さを固定する。 */
var PRINT_COL_HEAD_H_ = 14;
var PRINT_TITLE_H_ = 26;
var PRINT_META_LABEL_H_ = 16;
var PRINT_META_VALUE_H_ = 22;
var PRINT_SPACER_H_ = 2;
var PRINT_FOOTER_H_ = 18;

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
    name + ' を作成しました（A4縦 ' + built.pageCount + ' 枚分を 1 シート）。印刷は「縦向き」で。暫定データです。',
    '請求書入力',
    8
  );
}

function makePrintSamplePayload_() {
  const pack = collectPrintSampleCatalog_();
  const mids = pack.mids;
  const parts = pack.parts;
  const twoLine = [
    {
      mid: '左右ウイングアッパーレール\nコーキング修理',
      fee: 18000,
      workerCode: '1',
      partMid: 'リレーバルブインナーキット\n（新品）',
      qty: 1,
      unitPrice: 7400,
      amount: 7400
    },
    {
      mid: 'ランディング高さ違い点検修理\n（ギヤ灯高・低切り替え不良）',
      fee: 20000,
      workerCode: '2',
      partMid: '左ランディングASSY\n右ランディングＡＳＳＹ',
      qty: 1,
      unitPrice: 108510,
      amount: 108510
    }
  ];
  const items = twoLine.slice();
  let techSub = 0;
  let partSub = 0;
  twoLine.forEach(function (it) {
    techSub += Number(it.fee) || 0;
    partSub += Number(it.amount) || 0;
  });
  const n = 58;
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
      plate: '苫小牧888あ5555',
      dept: '大型',
      serviceType: '',
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
  printWorkerValue_.maps_ = null;
  const sheet = replacePrintSheet_(ss, sheetName);
  const slotH = printDataRowHeight_();

  sheet.setHiddenGridlines(true);

  applyPrintColumnWidths_(sheet);

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
  applyPrintPageBreaksAt_(sheet, breakRows);
  applyA4PageSetup_(sheet, pageCount);
  ss.setActiveSheet(sheet);
  applyA4PageSetup_(sheet, pageCount);
  hideHelperSheet_(ss, ss.getSheetByName('_印刷A4縦'));
  ss.setActiveSheet(sheet);
  return { sheet: sheet, pageCount: pageCount };
}

function replacePrintSheet_(ss, name) {
  const tpl = ensurePrintSetupSheet_(ss);
  const copy = tpl.copyTo(ss);
  try {
    copy.showSheet();
  } catch (err) {}
  hideHelperSheet_(ss, tpl);
  const old = ss.getSheetByName(name);
  if (old && old.getSheetId() !== copy.getSheetId() && ss.getSheets().length > 1) {
    ss.deleteSheet(old);
  }
  copy.setName(name);
  copy.clear();
  applyA4PageSetup_(copy, 1);
  return copy;
}

function ensurePrintSetupSheet_(ss) {
  const name = '_印刷A4縦';
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1).setValue('印刷設定用（非表示）');
    applyA4PageSetup_(sh, 1);
  } else {
    applyA4PageSetup_(sh, 1);
  }
  hideHelperSheet_(ss, sh);
  return sh;
}

/**
 * アクティブなシートは hideSheet できない。別シートへ移してから隠す。
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function hideHelperSheet_(ss, sh) {
  if (!sh) {
    return;
  }
  try {
    if (sh.isSheetHidden()) {
      return;
    }
  } catch (err) {}
  const sheets = ss.getSheets();
  let other = null;
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === sh.getSheetId()) {
      continue;
    }
    try {
      if (!sheets[i].isSheetHidden()) {
        other = sheets[i];
        break;
      }
    } catch (err2) {
      other = sheets[i];
      break;
    }
  }
  if (!other) {
    return;
  }
  try {
    ss.setActiveSheet(other);
    sh.hideSheet();
  } catch (err) {
    Logger.log('%s hideHelperSheet_: %s', CONFIG.logPrefix, err);
  }
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
    const m = PRINT_MARGIN_IN_;
    ps.setPaperSize(SpreadsheetApp.PaperSize.A4);
    ps.setPrintGridlines(false);
    if (typeof ps.setFitToPage === 'function') {
      ps.setFitToPage(false);
    }
    if (typeof ps.setScale === 'function') {
      ps.setScale(100);
    }
    if (typeof ps.setTopMargin === 'function') {
      ps.setTopMargin(m.top);
      ps.setBottomMargin(m.bottom);
      ps.setLeftMargin(m.left);
      ps.setRightMargin(m.right);
    }
    if (typeof ps.setHeaderMargin === 'function') {
      ps.setHeaderMargin(PRINT_HF_MARGIN_IN_);
      ps.setFooterMargin(PRINT_HF_MARGIN_IN_);
    }
    ps.setOrientation(SpreadsheetApp.PageOrientation.PORTRAIT);
  } catch (err) {
    Logger.log('%s A4 page setup: %s', CONFIG.logPrefix, err);
  }
  SpreadsheetApp.flush();
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
  const L = {};
  if (showHeader) {
    L.title = r;
    r++;
    L.metaL1 = r;
    r++;
    L.metaV1 = r;
    r++;
    L.metaL2 = r;
    r++;
    L.metaV2 = r;
    r++;
    L.metaL3 = r;
    r++;
    L.metaV3 = r;
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
  L.pad = r;
  r++;
  L.pageNo = r;
  r++;
  L.pageRows = r;
  return L;
}

function printInnerWidthPx_() {
  const m = PRINT_MARGIN_IN_;
  return Math.max(480, Math.floor((210 / 25.4 - m.left - m.right) * PRINT_PX_PER_IN_) - 12);
}

function printTargetInnerPx_() {
  const m = PRINT_MARGIN_IN_;
  const raw = (297 / 25.4 - m.top - m.bottom) * PRINT_PX_PER_IN_;
  return Math.max(600, Math.floor(raw) - 48);
}

function applyPrintColumnWidths_(sheet) {
  const src = PRINT_COL_WIDTHS_;
  const n = src.length;
  for (let c = 0; c < n; c++) {
    sheet.setColumnWidth(c + 1, printScaledColWidth_(c));
  }
}

function printScaledColWidth_(i) {
  const src = PRINT_COL_WIDTHS_;
  const sum = src.reduce(function (a, b) { return a + b; }, 0);
  const inner = printInnerWidthPx_();
  if (i === src.length - 1) {
    let used = 0;
    for (let c = 0; c < src.length - 1; c++) {
      used += Math.max(18, Math.floor(src[c] * inner / sum));
    }
    return Math.max(18, inner - used);
  }
  return Math.max(18, Math.floor(src[i] * inner / sum));
}

function printChromePx_(showHeader, showFooter, includeMinPad) {
  let h = PRINT_COL_HEAD_H_ + PRINT_PAGE_NO_H_;
  if (includeMinPad) {
    h += PRINT_PAD_MIN_;
  }
  if (showHeader) {
    h += PRINT_TITLE_H_ + (PRINT_META_LABEL_H_ + PRINT_META_VALUE_H_) * 3 + PRINT_SPACER_H_;
  }
  if (showFooter) {
    h += PRINT_FOOTER_H_ * 5;
  }
  return h;
}

function printDataRowHeight_() {
  const inner = printTargetInnerPx_();
  const chrome = printChromePx_(true, true, true);
  return Math.max(21, Math.floor((inner - chrome) / CONFIG.print.linesPerPage));
}

function printPadHeight_(showHeader, showFooter, slotH) {
  const inner = printTargetInnerPx_();
  const used = printChromePx_(showHeader, showFooter, false) + CONFIG.print.linesPerPage * slotH;
  const leftover = inner - used;
  return Math.max(PRINT_PAD_MIN_, Math.floor(leftover * PRINT_PAD_FILL_));
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

  if (showHeader && L.title != null) {
    const titleRow = start + L.title;
    const title = String(header.userName || CONFIG.print.title || '近海請求書').trim() || '近海請求書';
    sheet.getRange(titleRow, 2).setValue(title);
    sheet.getRange(titleRow, 2, 1, 7)
      .setFontSize(20)
      .setFontWeight('bold')
      .setHorizontalAlignment('left');
    fillPrintHeader_(sheet, start, L, header);
  }

  const headRow = start + L.colHead;
  sheet.getRange(headRow, 1, 1, cols)
    .setValues([PRINT_COL_HEADERS_])
    .setFontWeight('bold')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(false)
    .setBorder(true, true, true, true, true, true, PRINT_BLACK_, SpreadsheetApp.BorderStyle.SOLID);

  const first = start + L.firstLine;
  fillPrintBody_(sheet, first, lines, opts.serialOffset || 0);

  if (showFooter) {
    fillPrintFooterBlock_(sheet, start + L.footerStart, opts.summary || {});
  }

  const noRow = start + L.pageNo;
  sheet.getRange(noRow, 5).setValue('No.' + opts.page + '／' + opts.pageCount);
  sheet.getRange(noRow, 5, 1, 4)
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');

  sheet.getRange(start, 1, L.pageRows, cols).setVerticalAlignment('middle');
  fillPrintBodyFonts_(sheet, first, lines);
  return L.pageRows;
}

/**
 * 印刷の作業者列は狭いのでコード優先。リスト未登録の手入力はそのまま出す。
 */
function printWorkerValue_(it) {
  if (!it) {
    return '';
  }
  const maps = printWorkerValue_.maps_ || (printWorkerValue_.maps_ = workerPrintMaps_());
  const code = normalize_(it.workerCode);
  const name = normalize_(it.workerName);
  if (code && maps.byCode[code]) {
    return maps.byCode[code];
  }
  if (code && maps.byName[code]) {
    return maps.byName[code];
  }
  if (name && maps.byName[name]) {
    return maps.byName[name];
  }
  return '';
}

function workerPrintMaps_() {
  const maps = { byCode: {}, byName: {} };
  try {
    (loadWorkers_() || []).forEach(function (w) {
      const codeKey = normalize_(w.code);
      const nameKey = normalize_(w.name);
      const shown = String(w.code == null ? '' : w.code).trim();
      if (codeKey && !maps.byCode[codeKey]) {
        maps.byCode[codeKey] = shown;
      }
      if (nameKey && shown && !maps.byName[nameKey]) {
        maps.byName[nameKey] = shown;
      }
    });
  } catch (err) {
    Logger.log('%s workerPrintMaps_: %s', CONFIG.logPrefix, err);
  }
  return maps;
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
      printWorkerValue_(it),
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
  sheet.getRange(first, 7, per, 1).setNumberFormat(PRINT_YEN_FORMAT_).setHorizontalAlignment('right');
  sheet.getRange(first, 8, per, 1).setNumberFormat(PRINT_YEN_FORMAT_).setHorizontalAlignment('right').setFontSize(PRINT_FONT_MAX_ - 1);
  fillPrintBodyFonts_(sheet, first, lines);
}

function fillPrintBodyFonts_(sheet, first, lines) {
  const per = CONFIG.print.linesPerPage;
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
    workCell.setFontSize(fitPrintFont_(work, printScaledColWidth_(1)));
    partCell.setFontSize(fitPrintFont_(part, printScaledColWidth_(4)));
    workCell.setWrap(workBreak);
    partCell.setWrap(partBreak);
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

function fitPrintFont_(text, colWidth) {
  const raw = String(text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  if (!raw) {
    return PRINT_FONT_MAX_;
  }
  const parts = raw.split('\n');
  let units = 0;
  for (let i = 0; i < parts.length; i++) {
    units = Math.max(units, displayUnits_(parts[i]));
  }
  if (units < 0.5) {
    return PRINT_FONT_MAX_;
  }
  units += 0.5;
  for (let pt = PRINT_FONT_MAX_; pt >= PRINT_FONT_MIN_; pt--) {
    if (units <= printColCapacity_(colWidth, pt)) {
      return pt;
    }
  }
  return PRINT_FONT_MIN_;
}

function applyPrintPageHeights_(sheet, start, L, showHeader, showFooter, slotH) {
  const padH = printPadHeight_(showHeader, showFooter, slotH);
  if (showHeader && L.title != null) {
    sheet.setRowHeight(start + L.title, PRINT_TITLE_H_);
    sheet.setRowHeight(start + L.metaL1, PRINT_META_LABEL_H_);
    sheet.setRowHeight(start + L.metaV1, PRINT_META_VALUE_H_);
    sheet.setRowHeight(start + L.metaL2, PRINT_META_LABEL_H_);
    sheet.setRowHeight(start + L.metaV2, PRINT_META_VALUE_H_);
    sheet.setRowHeight(start + L.metaL3, PRINT_META_LABEL_H_);
    sheet.setRowHeight(start + L.metaV3, PRINT_META_VALUE_H_);
    sheet.setRowHeight(start + L.spacer, PRINT_SPACER_H_);
  }
  sheet.setRowHeight(start + L.colHead, PRINT_COL_HEAD_H_);
  sheet.setRowHeights(start + L.firstLine, CONFIG.print.linesPerPage, slotH);
  if (showFooter) {
    sheet.setRowHeights(start + L.footerStart, 5, PRINT_FOOTER_H_);
  }
  sheet.setRowHeight(start + L.pad, padH);
  sheet.setRowHeight(start + L.pageNo, PRINT_PAGE_NO_H_);
}

function mergePrintPage_(sheet, start, L, showHeader, showFooter) {
  if (showHeader && L.title != null) {
    sheet.getRange(start + L.title, 2, 1, 7).merge();
  }
  if (showHeader) {
    [L.metaL1, L.metaV1, L.metaL2, L.metaV2, L.metaL3, L.metaV3].forEach(function (off) {
      sheet.getRange(start + off, 3, 1, 3).merge();
      sheet.getRange(start + off, 6, 1, 3).merge();
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
  sheet.getRange(start + L.pageNo, 5, 1, 4).merge();
}

function fillPrintHeader_(sheet, start, L, header) {
  const l1 = start + L.metaL1;
  const v1 = start + L.metaV1;
  const l2 = start + L.metaL2;
  const v2 = start + L.metaV2;
  const l3 = start + L.metaL3;
  const v3 = start + L.metaV3;

  sheet.getRange(l1, 2).setValue('ユーザー').setFontSize(12).setFontWeight('bold');
  sheet.getRange(l1, 3).setValue('K-No').setFontSize(12).setFontWeight('bold');
  sheet.getRange(l1, 6).setValue('登録番号').setFontSize(12).setFontWeight('bold');
  sheet.getRange(v1, 2).setValue(header.userName || '').setFontSize(12);
  sheet.getRange(v1, 3).setValue(header.kNo || '').setFontSize(12);
  sheet.getRange(v1, 6).setValue(header.plate || '').setFontSize(12);

  sheet.getRange(l2, 2).setValue('整備部門').setFontSize(12).setFontWeight('bold');
  sheet.getRange(l2, 3).setValue('整備種別').setFontSize(12).setFontWeight('bold');
  sheet.getRange(l2, 6).setValue('受付').setFontSize(12).setFontWeight('bold');
  sheet.getRange(v2, 2).setValue(header.dept || '').setFontSize(12);
  sheet.getRange(v2, 3).setValue(header.serviceType || '').setFontSize(12);
  sheet.getRange(v2, 6).setValue(header.receptionist || header.staff || '').setFontSize(12);
  applyPrintDeptDropdown_(sheet.getRange(v2, 2), header.dept || '');
  applyPrintServiceTypeDropdown_(sheet.getRange(v2, 3), header.dept || '', header.serviceType || '');
  applyPrintReceptionistDropdown_(sheet.getRange(v2, 6), header.receptionist || header.staff || '');

  sheet.getRange(l3, 2).setValue('入庫日').setFontSize(12).setFontWeight('bold');
  sheet.getRange(l3, 3).setValue('出庫日').setFontSize(12).setFontWeight('bold');
  sheet.getRange(l3, 6).setValue('請求日').setFontSize(12).setFontWeight('bold');
  sheet.getRange(v3, 2).setValue(header.inDate || '').setFontSize(12);
  sheet.getRange(v3, 3).setValue(header.outDate || header.doneDate || '').setFontSize(12);
  sheet.getRange(v3, 6).setValue(header.billDate || '').setFontSize(12);

  sheet.getRange(l1, 2, 6, 1)
    .setHorizontalAlignment('left')
    .setBorder(true, true, true, true, false, true, PRINT_BLACK_, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(l1, 3, 6, 3)
    .setHorizontalAlignment('left')
    .setBorder(true, true, true, true, false, true, PRINT_BLACK_, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(l1, 6, 6, 3)
    .setHorizontalAlignment('left')
    .setBorder(true, true, true, true, false, true, PRINT_BLACK_, SpreadsheetApp.BorderStyle.SOLID);
}

function applyPrintDeptDropdown_(cell, current) {
  let names = [];
  try {
    names = serviceDepartments_().slice();
  } catch (err) {
    names = ['大型', '小型', 'BP板金', '部品販売'];
  }
  applyPrintOpenDropdown_(cell, names, current);
}

function applyPrintServiceTypeDropdown_(cell, dept, current) {
  let names = [];
  try {
    names = typesForServiceDept_(dept);
    if (!dept) {
      names = (loadServiceInfo_().allServiceTypes || []).slice();
    }
  } catch (err) {
    Logger.log('%s applyPrintServiceTypeDropdown_: %s', CONFIG.logPrefix, err);
  }
  applyPrintOpenDropdown_(cell, names, current);
}

function applyPrintReceptionistDropdown_(cell, current) {
  let names = [];
  try {
    names = (loadServiceInfo_().receptionists || []).slice();
  } catch (err) {
    Logger.log('%s applyPrintReceptionistDropdown_: %s', CONFIG.logPrefix, err);
  }
  applyPrintOpenDropdown_(cell, names, current);
}

function applyPrintOpenDropdown_(cell, names, current) {
  const list = (names || []).slice();
  const cur = String(current == null ? '' : current).trim();
  if (cur) {
    const key = normalize_(cur);
    let found = false;
    for (let i = 0; i < list.length; i++) {
      if (normalize_(list[i]) === key) {
        found = true;
        break;
      }
    }
    if (!found) {
      list.push(cur);
    }
  }
  applyOpenListValidation_(cell, list);
}

function handlePrintHeaderDeptEdit_(e) {
  const sheet = e.range.getSheet();
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) {
    return;
  }
  const finder = sheet.createTextFinder('整備部門').matchEntireCell(true).matchCase(false);
  const labels = finder.findAll();
  if (!labels || !labels.length) {
    return;
  }
  for (let i = 0; i < labels.length; i++) {
    const label = labels[i];
    if (label.getColumn() !== 2) {
      continue;
    }
    const valueRow = label.getRow() + 1;
    if (e.range.getRow() !== valueRow || e.range.getColumn() !== 2) {
      continue;
    }
    const dept = String(e.range.getValue() || '').trim();
    const typeCell = sheet.getRange(valueRow, 3);
    const cur = String(typeCell.getValue() || '').trim();
    const types = typesForServiceDept_(dept);
    let keep = '';
    for (let t = 0; t < types.length; t++) {
      if (normalize_(types[t]) === normalize_(cur)) {
        keep = types[t];
        break;
      }
    }
    if (dept === '部品販売' && types.indexOf('部品販売') !== -1) {
      keep = '部品販売';
    }
    if ((dept === 'BP板金' || dept === '板金塗装') && types.indexOf('板金塗装') !== -1 && !keep) {
      keep = '板金塗装';
    }
    typeCell.setValue(keep);
    applyPrintServiceTypeDropdown_(typeCell, dept, keep);
    return;
  }
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
