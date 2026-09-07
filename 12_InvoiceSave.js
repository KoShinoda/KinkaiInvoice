/**
 * 請求書の保存・呼び出し（キーは K-No 4桁。同一番号で複数件）。
 * 一覧は 1 件 1 行、明細は年シート。印刷は A4 縦・余白狭の PDF。
 */

var INVOICE_INDEX_HEADERS_ = [
  '保存ID', 'K-No', '保存日時', '年', 'ユーザー', '登録番号', '請求日',
  '入庫日', '出庫日', '整備部門', '整備種別', '受付', '値引技術%', '値引部品%',
  '明細件数', '明細開始行', '明細シート', '中項目1'
];

var INVOICE_DETAIL_HEADERS_ = [
  '保存ID', '行番号', '大項目', '中項目', '技術料', '作業コード', '作業者名',
  '部品大項目', '部品中項目', '数量', '単価', '金額', '値引額', '種別'
];

function saveInvoiceDraft(payload) {
  return saveInvoiceDraft_(payload);
}

function listInvoiceDrafts(kNo) {
  const key = normalizeInvoiceKNo_(kNo);
  if (!key) {
    throw new Error('K-No を入力してください。');
  }
  const sh = ensureInvoiceSaveIndexSheet_(true);
  const last = sh.getLastRow();
  if (last < 2) {
    return [];
  }
  const vals = sh.getRange(2, 1, last - 1, INVOICE_INDEX_HEADERS_.length).getValues();
  const list = [];
  for (let i = 0; i < vals.length; i++) {
    if (normalizeInvoiceKNo_(vals[i][1]) !== key) {
      continue;
    }
    list.push({
      saveId: String(vals[i][0] || '').trim(),
      kNo: String(vals[i][1] || ''),
      savedAt: formatInvoiceYmd_(vals[i][2]),
      userName: invoicePlain_(vals[i][4]),
      plate: invoicePlain_(vals[i][5]),
      billDate: formatInvoiceYmd_(vals[i][6]),
      lineCount: Number(vals[i][14]) || 0,
      firstMid: invoiceFirstMid_(vals[i][17]),
      detailStart: Number(vals[i][15]) || 0,
      detailSheet: String(vals[i][16] || '')
    });
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  list.forEach(function (row) {
    if (row.firstMid || !row.detailStart || !row.detailSheet) {
      delete row.detailStart;
      delete row.detailSheet;
      return;
    }
    try {
      const dsh = ss.getSheetByName(row.detailSheet);
      if (dsh) {
        row.firstMid = invoiceFirstMid_(dsh.getRange(row.detailStart, 4).getValue());
      }
    } catch (err) {}
    delete row.detailStart;
    delete row.detailSheet;
  });
  list.sort(function (a, b) {
    return String(b.savedAt).localeCompare(String(a.savedAt));
  });
  return invoiceJsonSafe_(list);
}

function loadInvoiceDraft(saveId) {
  const id = String(saveId || '').trim();
  if (!id) {
    throw new Error('保存データが指定されていません。');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const index = ensureInvoiceSaveIndexSheet_(true);
  const last = index.getLastRow();
  if (last < 2) {
    throw new Error('保存データがありません。');
  }
  const width = INVOICE_INDEX_HEADERS_.length;
  const vals = index.getRange(2, 1, last - 1, width).getValues();
  let meta = null;
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim() === id) {
      meta = vals[i];
      break;
    }
  }
  if (!meta) {
    throw new Error('保存データが見つかりません。');
  }
  const header = {
    kNo: invoicePlain_(meta[1]),
    userName: invoicePlain_(meta[4]),
    plate: invoicePlain_(meta[5]),
    billDate: formatInvoiceYmd_(meta[6]),
    inDate: formatInvoiceYmd_(meta[7]),
    outDate: formatInvoiceYmd_(meta[8]),
    doneDate: formatInvoiceYmd_(meta[8]),
    dept: invoicePlain_(meta[9]),
    serviceType: invoicePlain_(meta[10]),
    receptionist: invoicePlain_(meta[11]),
    staff: invoicePlain_(meta[11])
  };
  const techPct = meta[12] === '' || meta[12] == null ? 3 : Number(meta[12]);
  const partPct = meta[13] === '' || meta[13] == null ? 10 : Number(meta[13]);
  const items = loadInvoiceSaveLines_(ss, {
    saveId: id,
    lineCount: Number(meta[14]) || 0,
    detailStart: Number(meta[15]) || 0,
    detailSheet: String(meta[16] || '')
  });
  return invoiceJsonSafe_({
    header: header,
    items: items,
    summary: {
      techPct: techPct,
      partPct: partPct
    }
  });
}

function publishInvoicePdf(payload) {
  const printed = publishInvoices(payload);
  let saved = null;
  if (payload.header && normalizeInvoiceKNo_(payload.header.kNo)) {
    saved = saveInvoiceDraft_(payload);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = printed.sheetNames && printed.sheetNames[0];
  const sh = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
  const blob = exportPrintSheetPdf_(ss, sh, pdfFileName_(payload));
  return invoiceJsonSafe_({
    pageCount: printed.pageCount,
    lineCount: printed.lineCount,
    sheetNames: printed.sheetNames,
    saveId: saved ? saved.saveId : '',
    savedAt: saved ? saved.savedAt : '',
    filename: blob.getName(),
    pdfBase64: Utilities.base64Encode(blob.getBytes())
  });
}

function saveInvoiceDraft_(payload, savedAtOpt) {
  if (!payload || !payload.header) {
    throw new Error('保存する請求書がありません。');
  }
  const kNo = normalizeInvoiceKNo_(payload.header.kNo);
  if (!kNo) {
    throw new Error('K-No を入力してください。');
  }
  const items = (payload.items || []).filter(function (it) {
    return rowHasContent_(it);
  });
  const index = ensureInvoiceSaveIndexSheet_(true);
  const saveId = Utilities.getUuid();
  const savedAt = savedAtOpt || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const year = invoiceSaveYear_(savedAt);
  const detail = ensureInvoiceSaveDetailSheet_(year);
  const h = payload.header;
  const sum = payload.summary || {};
  const start = detail.getLastRow() + 1;
  const firstMid = items.length ? invoiceFirstMid_(items[0].mid || items[0].name) : '';
  if (items.length) {
    const lines = items.map(function (it, i) {
      return invoiceSaveDetailRow_(saveId, i + 1, it);
    });
    detail.getRange(start, 1, lines.length, INVOICE_DETAIL_HEADERS_.length).setValues(lines);
  }
  index.appendRow([
    saveId, kNo, savedAt, year,
    h.userName || '', h.plate || '', h.billDate || '', h.inDate || '', h.outDate || '',
    h.dept || '', h.serviceType || '', h.receptionist || h.staff || '',
    sum.techPct == null ? '' : sum.techPct,
    sum.partPct == null ? '' : sum.partPct,
    items.length,
    items.length ? start : '',
    detail.getName(),
    firstMid
  ]);
  return { saveId: saveId, savedAt: formatInvoiceYmd_(savedAt), lineCount: items.length, kNo: kNo };
}

function invoiceSaveDetailRow_(saveId, lineNo, it) {
  return [
    saveId, lineNo,
    it.major || '', it.mid || it.name || '', it.fee == null ? '' : it.fee,
    it.workerCode || '', it.workerName || '',
    it.partMajor || '', it.partMid || it.part || '',
    it.qty == null ? '' : it.qty,
    it.unitPrice == null ? '' : it.unitPrice,
    it.amount == null ? '' : it.amount,
    it.discYen == null ? '' : it.discYen,
    it.kind || ''
  ];
}

function loadInvoiceSaveLines_(ss, meta) {
  const n = Number(meta.lineCount) || 0;
  if (n < 1) {
    return [];
  }
  const sh = ss.getSheetByName(meta.detailSheet);
  if (!sh) {
    return [];
  }
  const start = Number(meta.detailStart) || 0;
  let vals = [];
  if (start >= 2 && start + n - 1 <= sh.getMaxRows()) {
    vals = sh.getRange(start, 1, n, INVOICE_DETAIL_HEADERS_.length).getValues();
    if (String(vals[0][0] || '').trim() !== meta.saveId) {
      vals = [];
    }
  }
  if (!vals.length) {
    vals = findInvoiceSaveLinesById_(sh, meta.saveId);
  }
  return vals.map(function (row) {
    return {
      major: invoicePlain_(row[2]),
      mid: invoicePlain_(row[3]),
      fee: invoicePlain_(row[4]),
      workerCode: invoicePlain_(row[5]),
      workerName: invoicePlain_(row[6]),
      partMajor: invoicePlain_(row[7]),
      partMid: invoicePlain_(row[8]),
      qty: invoicePlain_(row[9]),
      unitPrice: invoicePlain_(row[10]),
      amount: invoicePlain_(row[11]),
      discYen: invoicePlain_(row[12]),
      kind: invoicePlain_(row[13])
    };
  });
}

function findInvoiceSaveLinesById_(sh, saveId) {
  const last = sh.getLastRow();
  if (last < 2) {
    return [];
  }
  const vals = sh.getRange(2, 1, last - 1, INVOICE_DETAIL_HEADERS_.length).getValues();
  const out = [];
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim() === saveId) {
      out.push(vals[i]);
    }
  }
  return out;
}

/** google.script.run は Date を含む戻り値で成功コールバックが来ないことがある。 */
function invoiceJsonSafe_(obj) {
  return JSON.parse(JSON.stringify(obj));
}

function invoicePlain_(value) {
  if (value === true || value === false) {
    return value;
  }
  if (value instanceof Date && !isNaN(value.getTime())) {
    return formatInvoiceYmd_(value);
  }
  if (typeof value === 'number') {
    return isFinite(value) ? value : '';
  }
  if (value == null) {
    return '';
  }
  return String(value);
}

function formatInvoiceYmd_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy/MM/dd');
  }
  const s = String(value == null ? '' : value).trim();
  if (!s) {
    return '';
  }
  const m = s.match(/(\d{4})[\/\-年\.](\d{1,2})[\/\-月\.](\d{1,2})/);
  if (!m) {
    return s;
  }
  const mm = ('0' + Number(m[2])).slice(-2);
  const dd = ('0' + Number(m[3])).slice(-2);
  return m[1] + '/' + mm + '/' + dd;
}

function invoiceFirstMid_(text) {
  return String(text == null ? '' : text).replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n')[0].trim();
}

function ensureInvoiceSaveIndexSheet_(skipSeed) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = CONFIG.invoiceSave.sheetName;
  let sh = ss.getSheetByName(name);
  let created = false;
  if (!sh) {
    sh = ss.insertSheet(name);
    created = true;
    writeInvoiceIndexHeader_(sh);
  }
  migrateInvoiceSaveLegacy_(sh);
  sh = ss.getSheetByName(name) || sh;
  if (!skipSeed && (created || sh.getLastRow() < 2)) {
    writeInvoiceSaveSamples_();
    sh = ss.getSheetByName(name) || sh;
  }
  return sh;
}

function writeInvoiceIndexHeader_(sh) {
  sh.getRange(1, 1, 1, INVOICE_INDEX_HEADERS_.length).setValues([INVOICE_INDEX_HEADERS_]);
  sh.getRange(1, 1, 1, INVOICE_INDEX_HEADERS_.length).setFontWeight('bold').setBackground('#e8f0ec');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 80);
  sh.setColumnWidth(2, 70);
  sh.setColumnWidth(3, 150);
}

function ensureInvoiceSaveDetailSheet_(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = (CONFIG.invoiceSave.detailPrefix || '請求書保存明細_') + year;
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, INVOICE_DETAIL_HEADERS_.length).setValues([INVOICE_DETAIL_HEADERS_]);
    sh.getRange(1, 1, 1, INVOICE_DETAIL_HEADERS_.length).setFontWeight('bold').setBackground('#e8f0ec');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 80);
  }
  return sh;
}

function invoiceSaveYear_(savedAt) {
  const m = String(savedAt || '').match(/(\d{4})/);
  if (m) {
    return m[1];
  }
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy');
}

/**
 * K-1001 / k1001 / 1 → 1001 / 0001。数字以外だけなら空。
 */
function normalizeInvoiceKNo_(value) {
  let s = String(value == null ? '' : value).replace(/\u3000/g, ' ').trim();
  if (!s) {
    return '';
  }
  s = s.replace(/^K[-ー－–—−ｰ\s]*/i, '').trim();
  const digits = s.replace(/\D/g, '');
  if (!digits) {
    return '';
  }
  if (digits.length >= 4) {
    return digits;
  }
  return ('0000' + digits).slice(-4);
}

function migrateInvoiceSaveLegacy_(sh) {
  if (!sh || sh.getLastRow() < 1) {
    return;
  }
  const header = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  if (normalize_(header[1]) !== '行種別') {
    return;
  }
  const last = sh.getLastRow();
  const width = Math.max(sh.getLastColumn(), 1);
  const vals = last >= 2 ? sh.getRange(2, 1, last - 1, width).getValues() : [];
  const byId = {};
  const order = [];
  for (let i = 0; i < vals.length; i++) {
    const saveId = String(vals[i][0] || '').trim();
    if (!saveId) {
      continue;
    }
    if (!byId[saveId]) {
      byId[saveId] = { head: null, lines: [] };
      order.push(saveId);
    }
    if (String(vals[i][1] || '').trim() === 'HEAD') {
      byId[saveId].head = vals[i];
    } else if (String(vals[i][1] || '').trim() === 'LINE') {
      byId[saveId].lines.push(vals[i]);
    }
  }
  const ss = sh.getParent();
  const oldName = sh.getName();
  let n = 1;
  let archived = oldName + '_旧形式';
  while (ss.getSheetByName(archived)) {
    n += 1;
    archived = oldName + '_旧形式' + n;
  }
  sh.setName(archived);
  const index = ss.insertSheet(oldName);
  writeInvoiceIndexHeader_(index);
  order.forEach(function (saveId) {
    const pack = byId[saveId];
    const head = pack.head || pack.lines[0];
    if (!head) {
      return;
    }
    const savedAt = String(head[4] || '');
    const year = invoiceSaveYear_(savedAt);
    const detail = ensureInvoiceSaveDetailSheet_(year);
    const kNo = normalizeInvoiceKNo_(head[3]);
    const start = detail.getLastRow() + 1;
    if (pack.lines.length) {
      const lines = pack.lines.map(function (row, i) {
        return [
          saveId, i + 1,
          row[15] || '', row[16] || '', row[17],
          row[18] || '', row[19] || '',
          row[20] || '', row[21] || '',
          row[22], row[23], row[24], row[25], row[26] || ''
        ];
      });
      detail.getRange(start, 1, lines.length, INVOICE_DETAIL_HEADERS_.length).setValues(lines);
    }
    index.appendRow([
      saveId, kNo, savedAt, year,
      head[5] || '', head[6] || '', head[7] || '', head[8] || '', head[9] || '',
      head[10] || '', head[11] || '', head[12] || '',
      head[13], head[14],
      pack.lines.length,
      pack.lines.length ? start : '',
      detail.getName()
    ]);
  });
  try {
    sh.hideSheet();
  } catch (err) {}
}

function ensureInvoiceSaveSamples() {
  const sh = ensureInvoiceSaveIndexSheet_(true);
  if (invoiceSaveHasSampleK_(sh)) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'サンプル（1001 / 2088）は既にあります。入力アプリの呼び出しで確認できます。',
      '請求書入力',
      6
    );
    return;
  }
  writeInvoiceSaveSamples_();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '請求書保存にサンプル 3 件を追加しました。1001 は 2 件あるので呼び出しで選べます。',
    '請求書入力',
    8
  );
}

function invoiceSaveHasSampleK_(sh) {
  const last = sh.getLastRow();
  if (last < 2) {
    return false;
  }
  const vals = sh.getRange(2, 2, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    const k = normalizeInvoiceKNo_(vals[i][0]);
    if (k === '1001' || k === '2088') {
      return true;
    }
  }
  return false;
}

function writeInvoiceSaveSamples_() {
  const samples = makeInvoiceSaveSamplePayloads_();
  samples.forEach(function (s) {
    saveInvoiceDraft_(s.payload, s.savedAt);
  });
}

function makeInvoiceSaveSamplePayloads_() {
  return [
    {
      savedAt: '2026/08/10 10:15:00',
      payload: invoicePayloadFromTemplate_('3カ月定期点検', {
        userName: 'サンプル運輸',
        kNo: '1001',
        plate: '苫小牧800あ1111',
        receptionist: '佐藤',
        inDate: '2026/08/08',
        outDate: '2026/08/10',
        billDate: '2026/08/10',
        dept: '',
        serviceType: ''
      }, 3, 10)
    },
    {
      savedAt: '2026/09/01 14:30:00',
      payload: invoicePayloadFromTemplate_('１２カ月定期点検', {
        userName: 'サンプル運輸',
        kNo: '1001',
        plate: '苫小牧800あ1111',
        receptionist: '佐藤',
        inDate: '2026/08/28',
        outDate: '2026/09/01',
        billDate: '2026/09/01',
        dept: '',
        serviceType: ''
      }, 3, 10)
    },
    {
      savedAt: '2026/09/05 09:20:00',
      payload: invoicePayloadFromTemplate_('6カ月定期点検', {
        userName: '北港商事',
        kNo: '2088',
        plate: '札幌330い2222',
        receptionist: '鈴木',
        inDate: '2026/09/03',
        outDate: '2026/09/05',
        billDate: '2026/09/05',
        dept: '',
        serviceType: ''
      }, 3, 8)
    }
  ];
}

function invoicePayloadFromTemplate_(templateName, header, techPct, partPct) {
  const items = [];
  let techSub = 0;
  let partSub = 0;
  (defaultInvoiceTemplateRows_() || []).forEach(function (src) {
    if (src.name !== templateName) {
      return;
    }
    const qty = src.qty == null || src.qty === '' ? '' : src.qty;
    const price = src.unitPrice == null || src.unitPrice === '' ? '' : src.unitPrice;
    const list = (qty !== '' && price !== '') ? Number(qty) * Number(price) : 0;
    const disc = src.discYen == null || src.discYen === '' ? 0 : Number(src.discYen);
    const amount = list ? Math.round(list - disc) : '';
    const fee = src.fee == null || src.fee === '' ? '' : src.fee;
    techSub += Number(fee) || 0;
    partSub += Number(amount) || 0;
    items.push({
      major: src.major || '',
      mid: src.mid || '',
      fee: fee,
      workerCode: src.workerCode || '',
      workerName: '',
      partMajor: src.partMajor || '',
      partMid: src.partMid || '',
      qty: qty,
      unitPrice: price,
      amount: amount,
      discYen: disc || 0,
      kind: ''
    });
  });
  const techDisc = Math.round(techSub * techPct / 100);
  const partDisc = Math.round(partSub * partPct / 100);
  return {
    header: header,
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

function pdfFileName_(payload) {
  const kNo = normalizeInvoiceKNo_(payload && payload.header && payload.header.kNo) || '請求書';
  const when = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  return '請求書_' + kNo + '_' + when + '.pdf';
}

function exportPrintSheetPdf_(ss, sheet, filename) {
  SpreadsheetApp.flush();
  const id = ss.getId();
  const gid = sheet.getSheetId();
  const url = 'https://docs.google.com/spreadsheets/d/' + id + '/export'
    + '?exportFormat=pdf&format=pdf'
    + '&gid=' + gid
    + '&size=A4'
    + '&portrait=true'
    + '&scale=1'
    + '&top_margin=0.25'
    + '&bottom_margin=0.25'
    + '&left_margin=0.25'
    + '&right_margin=0.25'
    + '&gridlines=false'
    + '&printnotes=false'
    + '&printtitle=false'
    + '&sheetnames=false'
    + '&pagenumbers=false'
    + '&fzr=false'
    + '&horizontal_alignment=LEFT'
    + '&vertical_alignment=TOP';
  const token = ScriptApp.getOAuthToken();
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
    followRedirects: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('PDFの作成に失敗しました（' + res.getResponseCode() + '）。権限の承認をやり直してください。');
  }
  return res.getBlob().setName(filename || '請求書.pdf');
}
