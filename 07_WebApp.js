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
  notifyPendingListOrders_();
  const html = HtmlService.createHtmlOutputFromFile('入力アプリ')
    .setWidth(1640)
    .setHeight(860);
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
  const workLines = ctx.workRows.map(function (row) {
    return {
      major: row.major,
      mid: row.mid,
      content: normalize_(row.content),
      fee: row.fee,
      order: row.order,
      sourceIndex: row.sourceIndex,
      partMajor: normalize_(row.partMajor),
      partMid: normalize_(row.partMid),
      qty: row.qty,
      unitPrice: row.unitPrice
    };
  });
  const parts = loadPartCatalog_(ctx.workRows);
  const service = loadServiceInfo_();
  const workerSheet = loadWorkerSheet_();
  return {
    majors: majors,
    allMids: allMids,
    midsByMajor: ctx.midsByMajor,
    workLines: workLines,
    partMajors: parts.partMajors,
    allPartMids: parts.allPartMids,
    partMidsByMajor: parts.partMidsByMajor,
    partLines: parts.partLines,
    workers: workerSheet.rows,
    workerHeaders: workerSheet.headers,
    workerCodes: workerSheet.rows.map(function (w) {
      return w.name ? w.code + ' ' + w.name : w.code;
    }),
    departments: service.departments,
    typesByDept: service.typesByDept,
    typeSlotsByDept: service.typeSlotsByDept || {},
    allServiceTypes: service.allServiceTypes || [],
    receptionists: service.receptionists,
    lineCount: CONFIG.input.appRows || 120,
    linesPerPage: CONFIG.print.linesPerPage,
    ordersPending: !!ctx.ordersPending
  };
}

function imageUrlFromCell_(value, formula) {
  const f = String(formula || '');
  let m = f.match(/IMAGE\s*\(\s*"([^"]+)"/i);
  if (!m) m = f.match(/IMAGE\s*\(\s*'([^']+)'/i);
  if (m) return m[1];
  const v = String(value == null ? '' : value).trim();
  if (/^https?:\/\//i.test(v)) return v;
  return '';
}

function loadWorkerSheet_() {
  const empty = { headers: ['コード', '名前'], rows: [] };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.workers.sheetName);
  if (!sh) {
    return empty;
  }
  const range = sh.getDataRange();
  const vals = range.getValues();
  const formulas = range.getFormulas();
  if (!vals.length) {
    return empty;
  }
  let start = 0;
  const h0 = normalize_(vals[0][0]);
  if (h0.indexOf('コード') !== -1 || h0.indexOf('作業者') !== -1 || h0.indexOf('担当') !== -1) {
    start = 1;
  }
  const colCount = vals[0].length;
  const headers = [];
  for (let c = 0; c < colCount; c++) {
    headers.push(start === 1 ? normalize_(vals[0][c]) : '');
  }
  if (!headers[0]) headers[0] = 'コード';
  if (headers.length > 1 && !headers[1]) headers[1] = '名前';

  const cols = start === 1 ? resolveColumns_(vals[0], CONFIG.workers.headers) : {};
  const codeCol0 = cols.code ? cols.code - 1 : 0;
  const nameCol0 = cols.name ? cols.name - 1 : 1;
  const keepIdx = [];
  let skipNext = false;
  const refreshLabel = CONFIG.listRefresh.buttonLabel;
  for (let c = 0; c < colCount; c++) {
    if (skipNext) {
      skipNext = false;
      continue;
    }
    if (normalize_(headers[c]) === refreshLabel) {
      skipNext = true;
      continue;
    }
    if (isListMetaHeader_(headers[c])) {
      continue;
    }
    keepIdx.push(c);
  }
  const showHeaders = keepIdx.map(function (c) {
    return headers[c];
  });
  if (!showHeaders[0]) showHeaders[0] = 'コード';
  if (showHeaders.length > 1 && !showHeaders[1]) showHeaders[1] = '名前';

  const rows = [];
  const seen = {};
  for (let i = start; i < vals.length; i++) {
    const code = normalize_(vals[i][codeCol0]);
    const name = vals[i].length > nameCol0 ? normalize_(vals[i][nameCol0]) : '';
    if (!code && !name) {
      continue;
    }
    const key = code + '\t' + name;
    if (seen[key]) {
      continue;
    }
    seen[key] = true;
    const cells = [];
    const images = [];
    for (let k = 0; k < keepIdx.length; k++) {
      const c = keepIdx[k];
      const raw = vals[i][c];
      const formula = formulas[i] ? formulas[i][c] : '';
      const img = imageUrlFromCell_(raw, formula);
      images.push(img);
      cells.push(img ? '' : normalize_(raw));
    }
    rows.push({
      code: code,
      name: name,
      cells: cells,
      images: images,
      sourceIndex: i + 1
    });
  }
  return { headers: showHeaders, rows: rows };
}

function loadWorkers_() {
  return loadWorkerSheet_().rows;
}

function loadWorkerCodes_() {
  return loadWorkers_().map(function (w) {
    return w.name ? w.code + ' ' + w.name : w.code;
  });
}

function loadPartCatalog_(workRows) {
  const partLines = [];
  workRows.forEach(function (row) {
    const major = normalize_(row.partMajor);
    const mid = normalize_(row.partMid);
    if (!major && !mid) {
      return;
    }
    partLines.push({
      major: major,
      mid: mid,
      content: normalize_(row.content),
      order: row.order,
      sourceIndex: row.sourceIndex,
      qty: row.qty,
      unitPrice: row.unitPrice
    });
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.parts.sheetName);
  if (sh) {
    const vals = sh.getDataRange().getValues();
    let start = 0;
    if (vals.length) {
      const h = normalize_(vals[0][0]) + (vals[0].length > 1 ? normalize_(vals[0][1]) : '');
      if (h.indexOf('大項目') !== -1 || h.indexOf('部品') !== -1 || h.indexOf('中項目') !== -1) {
        start = 1;
      }
    }
    const cols = start === 1 && vals.length ? resolveColumns_(vals[0], CONFIG.parts.headers) : {};
    for (let i = start; i < vals.length; i++) {
      const c0 = normalize_(vals[i][0]);
      const c1 = vals[i].length > 1 ? normalize_(vals[i][1]) : '';
      let major = cols.major ? normalize_(cell_(vals[i], cols.major)) : c0;
      let mid = cols.mid ? normalize_(cell_(vals[i], cols.mid)) : c1;
      if (!major && !mid) {
        continue;
      }
      let content = '';
      let qty = vals[i].length > 2 ? vals[i][2] : '';
      let unitPrice = vals[i].length > 3 ? vals[i][3] : '';
      if (!mid) {
        major = major || c0;
        mid = major;
        qty = vals[i].length > 1 && isProbablyNumber_(vals[i][1]) ? vals[i][1] : '';
        unitPrice = vals[i].length > 2 ? vals[i][2] : '';
      } else if (!cols.major && vals[i].length > 2 && !isProbablyNumber_(vals[i][2])) {
        content = normalize_(vals[i][2]);
        qty = vals[i].length > 3 ? vals[i][3] : '';
        unitPrice = vals[i].length > 4 ? vals[i][4] : '';
      }
      partLines.push({
        major: major,
        mid: mid,
        content: content,
        qty: qty,
        unitPrice: unitPrice,
        order: cols.order ? cell_(vals[i], cols.order) : '',
        sourceIndex: i + 1
      });
    }
  }

  assignEmptyOrdersInGroups_(partLines);
  const seenPart = {};
  const uniquePartLines = [];
  sortWorkListRecords_(partLines).forEach(function (r) {
    const key = [r.major, r.mid, r.content, r.qty, r.unitPrice].join('\t');
    if (seenPart[key]) {
      return;
    }
    seenPart[key] = true;
    uniquePartLines.push(r);
  });

  const partMajors = uniqueValues_(uniquePartLines.map(function (r) {
    return r.major;
  }));
  const allPartMids = uniqueValues_(uniquePartLines.map(function (r) {
    return r.mid;
  }));
  const partMidsByMajor = {};
  uniquePartLines.forEach(function (r) {
    if (!r.major || !r.mid) {
      return;
    }
    if (!partMidsByMajor[r.major]) {
      partMidsByMajor[r.major] = [];
    }
    if (partMidsByMajor[r.major].indexOf(r.mid) === -1) {
      partMidsByMajor[r.major].push(r.mid);
    }
  });
  return {
    partLines: uniquePartLines,
    partMajors: partMajors,
    allPartMids: allPartMids,
    partMidsByMajor: partMidsByMajor
  };
}

function isProbablyNumber_(value) {
  if (value === '' || value === null || value === undefined) {
    return false;
  }
  if (typeof value === 'number') {
    return isFinite(value);
  }
  const n = Number(String(value).replace(/,/g, '').trim());
  return String(value).replace(/,/g, '').trim() !== '' && isFinite(n);
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
    throw new Error('明細がありません。');
  }
  const filled = payload.items.filter(function (it) {
    return rowHasContent_(it);
  });
  if (!filled.length) {
    throw new Error('1行以上入力してください。');
  }

  invalidateContext_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(CONFIG.input.sheetName);
  if (!inputSheet) {
    throw new Error('入力シートがありません: ' + CONFIG.input.sheetName);
  }

  writeInputSheetFromApp_(inputSheet, payload);
  writeSummaryToInput_(inputSheet, payload.summary);
  const names = writePrintSheets_(ss, payload);

  return {
    pageCount: names.length,
    lineCount: filled.length,
    sheetNames: names
  };
}

function rowHasContent_(it) {
  if (!it) {
    return false;
  }
  return isFilled_(it.major) || isFilled_(it.mid) || isFilled_(it.name) || isFilled_(it.fee) ||
    isFilled_(it.partMajor) || isFilled_(it.partMid) || isFilled_(it.part) ||
    isFilled_(it.qty) || isFilled_(it.unitPrice) || isFilled_(it.amount) || isFilled_(it.workerCode);
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
    const amount = lineAmount_(it, qty, price);
    values[i][0] = i + 1;
    values[i][1] = it.major || '';
    values[i][2] = it.mid || it.name || '';
    values[i][3] = it.fee === undefined || it.fee === null ? '' : it.fee;
    values[i][4] = it.workerCode || (payload.header && payload.header.staff) || '';
    values[i][5] = it.partMajor || '';
    values[i][6] = it.partMid || it.part || '';
    values[i][7] = qty;
    values[i][8] = price;
    values[i][9] = amount;
  }
  sheet.getRange(start, 1, values.length, width).setValues(values);
}

function writeSummaryToInput_(sheet, summary) {
  if (!summary || !CONFIG.summary) {
    return;
  }
  const map = CONFIG.summary;
  setIfMapped_(sheet, map.techSub, summary.techSub);
  setIfMapped_(sheet, map.techDisc, summary.techDisc);
  setIfMapped_(sheet, map.techTotal, summary.techTotal);
  setIfMapped_(sheet, map.partSub, summary.partSub);
  setIfMapped_(sheet, map.partDisc, summary.partDisc);
  setIfMapped_(sheet, map.grand, summary.grand);
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
  const items = payload.items.filter(function (it) {
    return rowHasContent_(it);
  });
  const pageCount = Math.max(1, Math.ceil(items.length / per));
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
  setIfMapped_(sheet, map.receptionist, header.receptionist || header.staff);
  setIfMapped_(sheet, map.staff, header.receptionist || header.staff);
  setIfMapped_(sheet, map.inDate, header.inDate);
  setIfMapped_(sheet, map.outDate, header.outDate || header.doneDate);
  setIfMapped_(sheet, map.doneDate, header.outDate || header.doneDate);
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
    const amount = lineAmount_(it, qty, price);
    serials.push([i + 1]);
    works.push([it.mid || it.name || '']);
    fees.push([it.fee === undefined || it.fee === null ? '' : it.fee]);
    staffs.push([it.workerCode || header.receptionist || header.staff || '']);
    parts.push([it.partMid || it.part || '']);
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

function lineAmount_(it, qty, price) {
  const typed = toNumberOrBlank_(it.amount);
  if (typed !== '') {
    return typed;
  }
  if (qty !== '' && price !== '') {
    return Number(qty) * Number(price);
  }
  return '';
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
