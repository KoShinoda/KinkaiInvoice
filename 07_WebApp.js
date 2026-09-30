/**
 * 車検入力の Web アプリ。
 * メニュー「入力アプリを開く」または「デプロイ → ウェブアプリ」。
 */

function invoiceAppOutput_(startView) {
  const view = startView === 'search' ? 'search' : 'input';
  const email = String(workJobUserKey_() || '');
  const raw = HtmlService.createHtmlOutputFromFile('入力アプリ').getContent();
  const html = raw
    .replace("var START_VIEW = 'input'; // kinkai:startView", "var START_VIEW = " + JSON.stringify(view) + "; // kinkai:startView")
    .replace("var OPERATOR_EMAIL = ''; // kinkai:operatorEmail", "var OPERATOR_EMAIL = " + JSON.stringify(email) + "; // kinkai:operatorEmail");
  return HtmlService.createHtmlOutput(html)
    .setTitle('車検 請求入力')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doGet(e) {
  const page = e && e.parameter ? String(e.parameter.page || '') : '';
  return invoiceAppOutput_(page === 'search' ? 'search' : 'input');
}

function openInputApp() {
  notifyPendingListOrders_();
  const html = invoiceAppOutput_('input')
    .setWidth(1900)
    .setHeight(860);
  SpreadsheetApp.getUi().showModalDialog(html, '車検 請求入力');
}

function openInvoiceSearchApp() {
  notifyPendingListOrders_();
  const html = invoiceAppOutput_('search')
    .setWidth(1900)
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
      workerCode: row.workerCode,
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
    plateAreas: (service.plateAreas && service.plateAreas.length)
      ? service.plateAreas
      : ['苫小牧', '室蘭', '北九州'],
    defaultReceptionist: receptionistForOperator_(service, workJobUserKey_()),
    defaultTechPct: clampInvoicePct_(service.techPct, 3),
    defaultPartPct: clampInvoicePct_(service.partPct, 10),
    operatorEmail: workJobUserKey_(),
    spreadsheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl(),
    templateNames: listInvoiceTemplateNamesFast_(),
    lineCount: CONFIG.app.lineCount || 120,
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
    const partMajor = normalize_(row.partMajor);
    const partMid = normalize_(row.partMid);
    const content = normalize_(row.content);
    if (partMid) {
      partLines.push(coercePartMeasure_({
        major: partMajor,
        mid: partMid,
        content: content,
        order: row.order,
        sourceIndex: row.sourceIndex,
        qty: row.qty,
        unitPrice: row.unitPrice
      }));
    }
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.parts.sheetName);
  if (sh) {
    ensurePartsSetHeader_(sh);
    const parsed = parsePartsSheetValues_(sh.getDataRange().getValues());
    parsed.rows.forEach(function (row) {
      row.fromParts = true;
      partLines.push(coercePartMeasure_(row));
    });
  }

  assignEmptyOrdersInGroups_(partLines);
  const seenPart = {};
  const uniquePartLines = [];
  sortWorkListRecords_(partLines).forEach(function (r) {
    const key = [r.major, r.mid, r.set || '', r.content, r.qty, r.unitPrice].join('\t');
    if (seenPart[key]) {
      return;
    }
    seenPart[key] = true;
    uniquePartLines.push(r);
  });

  const drop = loadPartListDropdowns_();
  return {
    partLines: uniquePartLines,
    partMajors: drop.partMajors,
    allPartMids: drop.allPartMids,
    partMidsByMajor: drop.partMidsByMajor
  };
}

/** 入力の部品大項目・中項目は部品リストのみ（作業リストは展開用 partLines に残す）。 */
function loadPartListDropdowns_() {
  const majors = [];
  const allMids = [];
  const byMajor = {};
  function addMajor(v) {
    if (!v || majors.indexOf(v) !== -1) {
      return;
    }
    majors.push(v);
  }
  function addMid(major, mid) {
    if (!mid) {
      return;
    }
    if (allMids.indexOf(mid) === -1) {
      allMids.push(mid);
    }
    if (!major) {
      return;
    }
    if (!byMajor[major]) {
      byMajor[major] = [];
    }
    if (byMajor[major].indexOf(mid) === -1) {
      byMajor[major].push(mid);
    }
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.parts.sheetName);
  if (!sh) {
    return { partMajors: majors, allPartMids: allMids, partMidsByMajor: byMajor };
  }
  ensurePartsSetHeader_(sh);
  const parsed = parsePartsSheetValues_(sh.getDataRange().getValues());
  parsed.rows.forEach(function (row) {
    addMajor(row.major);
    addMid(row.major, row.mid);
  });
  return { partMajors: majors, allPartMids: allMids, partMidsByMajor: byMajor };
}

function ensurePartsSetHeader_(sheet) {
  if (!sheet) {
    return;
  }
  const last = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, last).getValues()[0];
  for (let c = 0; c < headers.length; c++) {
    if (normalize_(headers[c]) === '部品') {
      sheet.getRange(1, c + 1).setValue('部品_セット');
    }
  }
}

function parsePartsSheetValues_(vals) {
  const rows = [];
  if (!vals || !vals.length) {
    return { rows: rows };
  }
  const joined = (vals[0] || []).map(function (v) {
    return normalize_(v);
  }).join(' ');
  const start = (joined.indexOf('大項目') !== -1 || joined.indexOf('部品') !== -1 ||
    joined.indexOf('中項目') !== -1 || joined.indexOf('セット') !== -1) ? 1 : 0;
  const cols = start === 1 ? resolveColumns_(vals[0], CONFIG.parts.headers) : {};
  const carry = { major: '', set: '', group: '' };
  for (let i = start; i < vals.length; i++) {
    const rec = readPartsListRow_(vals[i], cols, carry, i + 1);
    if (!rec.major && !rec.mid) {
      continue;
    }
    rows.push(rec);
  }
  promotePartSetNames_(rows);
  return { rows: rows, cols: cols };
}

/**
 * 部品_セットが空で単価・数量がある行は、セット名に中項目名を入れる。
 *
 * @param {object[]} rows
 * @return {number}
 */
function promotePartSetNames_(rows) {
  if (!rows || !rows.length) {
    return 0;
  }
  let filled = 0;
  rows.forEach(function (row) {
    const mid = normalize_(row && row.mid);
    if (!mid) {
      return;
    }
    if (normalize_(row.set)) {
      return;
    }
    if (normalize_(row.content) && !isNumericCell_(row.content)) {
      row.set = normalize_(row.content);
      filled += 1;
      return;
    }
    if (!isFilled_(row.qty) && !isFilled_(row.unitPrice)) {
      return;
    }
    row.set = mid;
    row.content = mid;
    filled += 1;
  });
  return filled;
}

function readPartsListRow_(raw, cols, carry, sourceIndex) {
  const rawMajor = cols.major ? normalize_(cell_(raw, cols.major)) : normalize_(raw[0]);
  const rawSet = cols.set ? normalize_(cell_(raw, cols.set)) : '';
  let rawGroup = '';
  if (cols.mid) {
    rawGroup = normalize_(cell_(raw, cols.mid));
  } else if (!cols.set && raw.length > 1) {
    rawGroup = normalize_(raw[1]);
  }
  if (rawMajor) {
    carry.major = rawMajor;
  }
  if (rawGroup) {
    if (rawGroup !== carry.group && !rawSet) {
      carry.set = '';
    }
    carry.group = rawGroup;
  }
  if (rawSet) {
    carry.set = rawSet;
  }
  const major = rawMajor || carry.major;
  const group = rawGroup || carry.group;
  const setName = rawSet || carry.set;
  const mid = group || setName;
  let content = cols.name ? normalize_(cell_(raw, cols.name)) : '';
  let qty = cols.qty ? cell_(raw, cols.qty) : '';
  let unitPrice = cols.unitPrice ? cell_(raw, cols.unitPrice) : '';
  if (content && isProbablyNumber_(content)) {
    if (!isFilled_(unitPrice)) {
      unitPrice = content;
    }
    content = '';
  }
  if (!cols.name && !cols.qty && !cols.unitPrice) {
    const setIdx = cols.set ? cols.set - 1 : (cols.mid ? cols.mid : 1);
    if (raw.length > setIdx + 1 && !isProbablyNumber_(raw[setIdx + 1])) {
      content = content || normalize_(raw[setIdx + 1]);
      qty = raw.length > setIdx + 2 ? raw[setIdx + 2] : '';
      unitPrice = raw.length > setIdx + 3 ? raw[setIdx + 3] : '';
    } else {
      qty = qty !== '' && qty != null ? qty : (raw.length > setIdx + 1 ? raw[setIdx + 1] : '');
      unitPrice = unitPrice !== '' && unitPrice != null ? unitPrice : (raw.length > setIdx + 2 ? raw[setIdx + 2] : '');
    }
  }
  if (setName) {
    content = setName;
  }
  return {
    major: major,
    mid: mid,
    set: setName,
    content: content,
    qty: qty,
    unitPrice: unitPrice,
    order: cols.order ? cell_(raw, cols.order) : '',
    sourceIndex: sourceIndex,
    fromParts: true
  };
}

function isProbablyNumber_(value) {
  return isNumericCell_(value);
}

/**
 * 中項目に紐づく作業内容（空白以外）。技術料は作業内容行だけ。
 */
function expandMidSelection(major, mid) {
  invalidateContext_();
  const ctx = loadContext_();
  const resolved = resolveMidOutput_(ctx, normalize_(major), normalize_(mid));
  return {
    midFee: '',
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

function publishPrintSheet(payload) {
  return publishInvoicePdf(payload);
}

function publishInvoicePdf(payload) {
  if (!payload || !payload.header || !normalizeInvoiceKNo_(payload.header.kNo)) {
    throw new Error('K-No を入力してください。');
  }
  const saved = saveInvoiceDraft_(payload);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let prev = null;
  try {
    prev = ss.getActiveSheet();
  } catch (err) {
    prev = null;
  }
  let tmp = null;
  try {
    const printed = publishInvoices(payload, pdfTempSheetName_());
    const tmpName = (printed.sheetNames && printed.sheetNames[0]) || '';
    tmp = printed.sheet || (tmpName ? ss.getSheetByName(tmpName) : null);
    if (!tmp) {
      throw new Error('PDF 用の帳票を作れませんでした。');
    }
    try {
      tmp.showSheet();
    } catch (errShow) {}
    SpreadsheetApp.flush();
    const pdfBase64 = exportSheetPdf_(ss, tmp);
    return invoiceJsonSafe_({
      pageCount: printed.pageCount,
      lineCount: printed.lineCount,
      saveId: saved.saveId,
      savedAt: saved.savedAt,
      overwritten: !!saved.overwritten,
      kNo: saved.kNo,
      pdfName: invoicePdfFileName_(payload),
      pdfBase64: pdfBase64,
      blankDebug: printed.blankDebug || ''
    });
  } finally {
    if (tmp) {
      try {
        ss.deleteSheet(tmp);
      } catch (err2) {}
    }
    if (prev) {
      try {
        ss.setActiveSheet(prev);
      } catch (err3) {}
    }
  }
}

/**
 * @param {object} payload
 * @param {string=} sheetName
 * @return {{pageCount: number, lineCount: number, sheetNames: string[]}}
 */
function publishInvoices(payload, sheetName) {
  if (!payload || !payload.items || !payload.items.length) {
    throw new Error('明細がありません。');
  }
  const filled = payload.items.filter(function (it) {
    return rowHasContent_(it);
  });
  if (!filled.length) {
    throw new Error('1行以上入力してください。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const printed = writePrintSheets_(ss, payload, sheetName);

  return {
    pageCount: printed.pageCount,
    lineCount: filled.length,
    sheetNames: printed.sheetNames,
    sheet: printed.sheet,
    blankDebug: invoiceBlankRowDebug_(payload.items)
  };
}

/** スペースだけ。改行だけの欄は文字が無いので空とみなす。 */
function isSpaceOnly_(value) {
  const s = String(value == null ? '' : value).replace(/\u3000/g, ' ');
  if (!s || normalize_(s) !== '') {
    return false;
  }
  return s.indexOf(' ') !== -1;
}

function isZeroNumber_(value) {
  if (value === 0 || value === '0') {
    return true;
  }
  if (value === '' || value == null) {
    return false;
  }
  const s = String(value).replace(/,/g, '').replace(/\u3000/g, ' ').trim();
  if (!s) {
    return false;
  }
  const n = Number(s);
  return isFinite(n) && n === 0;
}

function rowHasContent_(it) {
  if (!it) {
    return false;
  }
  if (isFilled_(it.mid) || isFilled_(it.name) || isSpaceOnly_(it.mid) || isSpaceOnly_(it.name)) {
    return true;
  }
  if (isFilled_(it.partMid) || isFilled_(it.part) || isSpaceOnly_(it.partMid) || isSpaceOnly_(it.part)) {
    return true;
  }
  return isFilled_(it.fee) && !isZeroNumber_(it.fee);
}

function debugFieldText_(value) {
  const s = String(value == null ? '' : value);
  if (!s) {
    return '（空）';
  }
  const codes = [];
  for (let i = 0; i < s.length && codes.length < 8; i++) {
    const c = s.charCodeAt(i);
    if (c <= 32 || c === 0x3000 || c === 0xFEFF || c === 0x200B) {
      codes.push('U+' + c.toString(16).toUpperCase());
    }
  }
  const shown = s.replace(/\r/g, '').replace(/\n/g, '↵').replace(/ /g, '·');
  const short = shown.length > 24 ? shown.slice(0, 24) + '…' : shown;
  return short + (codes.length ? ' [' + codes.join(' ') + ']' : '');
}

function rowKeepReasons_(it) {
  const reasons = [];
  if (!it) {
    return reasons;
  }
  if (isFilled_(it.mid)) {
    reasons.push('中項目=' + debugFieldText_(it.mid));
  }
  if (isFilled_(it.name)) {
    reasons.push('name=' + debugFieldText_(it.name));
  }
  if (isSpaceOnly_(it.mid)) {
    reasons.push('中項目はスペースのみ ' + debugFieldText_(it.mid));
  }
  if (isSpaceOnly_(it.name)) {
    reasons.push('nameはスペースのみ ' + debugFieldText_(it.name));
  }
  if (isFilled_(it.partMid)) {
    reasons.push('部品中項目=' + debugFieldText_(it.partMid));
  }
  if (isFilled_(it.part)) {
    reasons.push('part=' + debugFieldText_(it.part));
  }
  if (isSpaceOnly_(it.partMid)) {
    reasons.push('部品中項目はスペースのみ ' + debugFieldText_(it.partMid));
  }
  if (isSpaceOnly_(it.part)) {
    reasons.push('partはスペースのみ ' + debugFieldText_(it.part));
  }
  if (isFilled_(it.fee) && !isZeroNumber_(it.fee)) {
    reasons.push('技術料=' + it.fee);
  }
  return reasons;
}

function debugBlankPrintLines_(text) {
  const raw = String(text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  if (!raw) {
    return 0;
  }
  let n = 0;
  raw.split('\n').forEach(function (line) {
    if (normalize_(line) === '') {
      n++;
    }
  });
  return n;
}

/** 部品欄が空に見えるのに残した行と、改行で高くなった行。 */
function invoiceBlankRowDebug_(items) {
  const lines = [];
  (items || []).forEach(function (it, i) {
    const reasons = rowKeepReasons_(it);
    if (!reasons.length) {
      return;
    }
    const no = (i + 1) + '行目';
    const partText = isFilled_(it.partMid) || isFilled_(it.part);
    if (!partText) {
      lines.push(no + ' 部品欄は空。残した理由: ' + reasons.join(' / '));
    }
    const work = printFittedText_(it.mid || it.name || '', printScaledColWidth_(1));
    const part = printFittedText_(printPartSource_(it), printScaledColWidth_(4));
    const workBlank = debugBlankPrintLines_(work);
    const partBlank = debugBlankPrintLines_(part);
    const tall = Math.max(newlineCount_(work), newlineCount_(part));
    if (tall > 1 || workBlank || partBlank) {
      lines.push(no + ' この明細の高さ' + tall + '行（作業の空行' + workBlank + '、部品の空行' + partBlank + '）');
    }
  });
  if (!lines.length) {
    return '部品欄が空なのに残した行はありません。';
  }
  const shown = lines.slice(0, 12);
  if (lines.length > shown.length) {
    shown.push('他 ' + (lines.length - shown.length) + ' 件');
  }
  return shown.join('\n');
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
  if (typeof value === 'number') {
    return isFinite(value) ? value : '';
  }
  const s = String(value).replace(/,/g, '').replace(/\u3000/g, ' ').trim();
  if (!s) {
    return '';
  }
  const n = Number(s);
  return isFinite(n) ? n : '';
}
