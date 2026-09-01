/**
 * 車検入力の Web アプリ。
 * メニュー「入力アプリを開く」または「デプロイ → ウェブアプリ」。
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('入力アプリ')
    .setTitle('車検 請求入力')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function openInputApp() {
  const html = HtmlService.createHtmlOutputFromFile('入力アプリ')
    .setWidth(1040)
    .setHeight(740);
  SpreadsheetApp.getUi().showModalDialog(html, '車検 請求入力');
}

/**
 * 画面初期データ（大項目・中項目・展開に使うマスタ）。
 */
function getInvoiceMaster() {
  invalidateContext_();
  const ctx = loadContext_();
  const majors = uniqueValues_(ctx.workRows.map(function (row) {
    return row.major;
  }));
  const allMids = uniqueValues_(ctx.workRows.map(function (row) {
    return row.mid;
  }));
  return {
    majors: majors,
    allMids: allMids,
    midsByMajor: ctx.midsByMajor,
    linesPerPage: CONFIG.print.linesPerPage
  };
}

/**
 * 中項目に紐づく作業内容（空白以外）と、中項目技術料。
 */
function expandMidSelection(major, mid) {
  invalidateContext_();
  const ctx = loadContext_();
  const resolved = resolveMidOutput_(ctx, normalize_(major), normalize_(mid));
  return {
    midFee: resolved.midFee,
    works: resolved.workRows.map(function (row) {
      return {
        name: row.content,
        fee: row.fee,
        part: pickPartDisplay_(row),
        qty: row.qty,
        unitPrice: row.unitPrice
      };
    })
  };
}

function pickPartDisplay_(row) {
  if (isFilled_(row.partMid)) {
    return row.partMid;
  }
  return row.partMajor || '';
}

/**
 * 車検_入力へ保存し、原紙を 30 行ずつコピーして印刷シートを作る。
 *
 * @param {object} payload
 * @return {{pageCount: number, lineCount: number, sheetNames: string[]}}
 */
function publishInvoices(payload) {
  if (!payload || !payload.items || !payload.items.length) {
    throw new Error('明細がありません。中項目を追加してください。');
  }

  invalidateContext_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(CONFIG.input.sheetName);
  if (!inputSheet) {
    throw new Error('入力シートがありません: ' + CONFIG.input.sheetName);
  }

  writeInputSheetFromApp_(inputSheet, payload);
  const names = writePrintSheets_(ss, payload);

  return {
    pageCount: names.length,
    lineCount: payload.items.length,
    sheetNames: names
  };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {object} payload
 */
function writeInputSheetFromApp_(sheet, payload) {
  const start = CONFIG.input.dataStartRow;
  const maxRows = CONFIG.input.maxSelectRows;
  const items = payload.items;
  const width = 10;
  const values = [];
  for (let i = 0; i < Math.max(items.length, maxRows); i++) {
    values.push(['', '', '', '', '', '', '', '', '', '']);
  }
  for (let i = 0; i < items.length && i < maxRows; i++) {
    const it = items[i];
    const qty = toNumberOrBlank_(it.qty);
    const price = toNumberOrBlank_(it.unitPrice);
    const amount = qty !== '' && price !== '' ? Number(qty) * Number(price) : '';
    values[i][0] = i + 1;
    values[i][1] = it.kind === 'mid' ? (it.major || '') : '';
    values[i][2] = it.name || '';
    values[i][3] = it.fee === undefined || it.fee === null ? '' : it.fee;
    values[i][4] = (payload.header && payload.header.staff) || '';
    values[i][5] = it.part || '';
    values[i][6] = it.part || '';
    values[i][7] = qty;
    values[i][8] = price;
    values[i][9] = amount;
  }
  sheet.getRange(start, 1, values.length, width).setValues(values);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {object} payload
 * @return {string[]}
 */
function writePrintSheets_(ss, payload) {
  const templateName = CONFIG.print.templateSheetName;
  const template = ss.getSheetByName(templateName);
  if (!template) {
    throw new Error('印刷テンプレートがありません: ' + templateName);
  }

  const per = CONFIG.print.linesPerPage;
  const items = payload.items;
  const pageCount = Math.ceil(items.length / per);
  const created = [];

  deleteOldPrintSheets_(ss);

  for (let p = 0; p < pageCount; p++) {
    const name = CONFIG.print.sheetNamePrefix + (p + 1);
    const copy = template.copyTo(ss);
    copy.setName(name);
    const slice = items.slice(p * per, p * per + per);
    fillPrintSheet_(copy, payload.header || {}, slice);
    created.push(name);
  }

  if (created.length) {
    ss.setActiveSheet(ss.getSheetByName(created[0]));
  }
  return created;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function deleteOldPrintSheets_(ss) {
  const prefix = CONFIG.print.sheetNamePrefix;
  const toDelete = [];
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().indexOf(prefix) === 0) {
      toDelete.push(sheets[i]);
    }
  }
  for (let i = 0; i < toDelete.length; i++) {
    ss.deleteSheet(toDelete[i]);
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {object} header
 * @param {object[]} lines
 */
function fillPrintSheet_(sheet, header, lines) {
  const map = CONFIG.print.header;
  setIfMapped_(sheet, map.kNo, header.kNo);
  setIfMapped_(sheet, map.plate, header.plate);
  setIfMapped_(sheet, map.staff, header.staff);
  setIfMapped_(sheet, map.inDate, header.inDate);
  setIfMapped_(sheet, map.doneDate, header.doneDate);
  setIfMapped_(sheet, map.billDate, header.billDate);

  const cols = CONFIG.print.cols;
  const first = CONFIG.print.firstLineRow;
  const per = CONFIG.print.linesPerPage;
  const serials = [];
  const works = [];
  const fees = [];
  const staffs = [];
  const parts = [];
  const qtys = [];
  const prices = [];
  const amounts = [];

  for (let i = 0; i < per; i++) {
    const it = lines[i];
    if (!it) {
      serials.push(['']);
      works.push(['']);
      fees.push(['']);
      staffs.push(['']);
      parts.push(['']);
      qtys.push(['']);
      prices.push(['']);
      amounts.push(['']);
      continue;
    }
    const qty = toNumberOrBlank_(it.qty);
    const price = toNumberOrBlank_(it.unitPrice);
    const amount = qty !== '' && price !== '' ? Number(qty) * Number(price) : '';
    serials.push([i + 1]);
    works.push([it.name || '']);
    fees.push([it.fee === undefined || it.fee === null ? '' : it.fee]);
    staffs.push([header.staff || '']);
    parts.push([it.part || '']);
    qtys.push([qty]);
    prices.push([price]);
    amounts.push([amount]);
  }

  sheet.getRange(first, cols.serial, per, 1).setValues(serials);
  sheet.getRange(first, cols.work, per, 1).setValues(works);
  sheet.getRange(first, cols.fee, per, 1).setValues(fees);
  sheet.getRange(first, cols.staff, per, 1).setValues(staffs);
  sheet.getRange(first, cols.part, per, 1).setValues(parts);
  sheet.getRange(first, cols.qty, per, 1).setValues(qtys);
  sheet.getRange(first, cols.unitPrice, per, 1).setValues(prices);
  sheet.getRange(first, cols.amount, per, 1).setValues(amounts);
}

function setIfMapped_(sheet, a1, value) {
  if (!a1 || value === undefined || value === null || value === '') {
    return;
  }
  sheet.getRange(a1).setValue(value);
}

function toNumberOrBlank_(value) {
  if (value === '' || value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'number' && isFinite(value)) {
    return value;
  }
  const n = Number(String(value).replace(/,/g, '').trim());
  return isFinite(n) ? n : '';
}
