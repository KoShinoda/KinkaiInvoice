/**
 * 請求書の保存・呼び出し（キーは K-No。同一 K-No で複数件）。
 * 印刷は A4 縦・余白狭の PDF。ダイアログからの自動印刷はブラウザ制限あり。
 */

var INVOICE_SAVE_HEADERS_ = [
  '保存ID', '行種別', '行番号', 'K-No', '保存日時',
  'ユーザー', '登録番号', '請求日', '入庫日', '出庫日',
  '整備部門', '整備種別', '受付', '値引技術%', '値引部品%',
  '大項目', '中項目', '技術料', '作業コード', '作業者名',
  '部品大項目', '部品中項目', '数量', '単価', '金額', '値引額', '種別'
];

function saveInvoiceDraft(payload) {
  return saveInvoiceDraft_(payload);
}

function listInvoiceDrafts(kNo) {
  const key = normalize_(kNo);
  if (!key) {
    throw new Error('K-No を入力してください。');
  }
  const sh = ensureInvoiceSaveSheet_();
  const last = sh.getLastRow();
  if (last < 2) {
    return [];
  }
  const vals = sh.getRange(2, 1, last - 1, 8).getValues();
  const byId = {};
  const order = [];
  for (let i = 0; i < vals.length; i++) {
    const saveId = String(vals[i][0] || '').trim();
    const rowType = String(vals[i][1] || '').trim();
    const rowK = normalize_(vals[i][3]);
    if (!saveId || rowK !== key) {
      continue;
    }
    if (!byId[saveId]) {
      byId[saveId] = {
        saveId: saveId,
        kNo: String(vals[i][3] || ''),
        savedAt: String(vals[i][4] || ''),
        userName: '',
        plate: '',
        billDate: '',
        lineCount: 0
      };
      order.push(saveId);
    }
    if (rowType === 'HEAD') {
      byId[saveId].userName = String(vals[i][5] || '');
      byId[saveId].plate = String(vals[i][6] || '');
      byId[saveId].billDate = String(vals[i][7] || '');
    } else if (rowType === 'LINE') {
      byId[saveId].lineCount += 1;
    }
  }
  const list = order.map(function (id) {
    return byId[id];
  });
  list.sort(function (a, b) {
    return String(b.savedAt).localeCompare(String(a.savedAt));
  });
  return list;
}

function loadInvoiceDraft(saveId) {
  const id = String(saveId || '').trim();
  if (!id) {
    throw new Error('保存データが指定されていません。');
  }
  const sh = ensureInvoiceSaveSheet_();
  const last = sh.getLastRow();
  if (last < 2) {
    throw new Error('保存データがありません。');
  }
  const width = INVOICE_SAVE_HEADERS_.length;
  const vals = sh.getRange(2, 1, last - 1, width).getValues();
  let header = null;
  const items = [];
  let techPct = 3;
  let partPct = 10;
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim() !== id) {
      continue;
    }
    const rowType = String(vals[i][1] || '').trim();
    if (rowType === 'HEAD') {
      header = {
        userName: vals[i][5] || '',
        kNo: vals[i][3] || '',
        plate: vals[i][6] || '',
        billDate: vals[i][7] || '',
        inDate: vals[i][8] || '',
        outDate: vals[i][9] || '',
        doneDate: vals[i][9] || '',
        dept: vals[i][10] || '',
        serviceType: vals[i][11] || '',
        receptionist: vals[i][12] || '',
        staff: vals[i][12] || ''
      };
      techPct = vals[i][13] === '' || vals[i][13] == null ? 3 : vals[i][13];
      partPct = vals[i][14] === '' || vals[i][14] == null ? 10 : vals[i][14];
    } else if (rowType === 'LINE') {
      items.push({
        major: vals[i][15] || '',
        mid: vals[i][16] || '',
        fee: vals[i][17],
        workerCode: vals[i][18] || '',
        workerName: vals[i][19] || '',
        partMajor: vals[i][20] || '',
        partMid: vals[i][21] || '',
        qty: vals[i][22],
        unitPrice: vals[i][23],
        amount: vals[i][24],
        discYen: vals[i][25],
        kind: vals[i][26] || ''
      });
    }
  }
  if (!header) {
    throw new Error('保存データが見つかりません。');
  }
  return {
    header: header,
    items: items,
    summary: {
      techPct: techPct,
      partPct: partPct
    }
  };
}

function publishInvoicePdf(payload) {
  const printed = publishInvoices(payload);
  let saved = null;
  if (payload.header && normalize_(payload.header.kNo)) {
    saved = saveInvoiceDraft_(payload);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = printed.sheetNames && printed.sheetNames[0];
  const sh = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
  const blob = exportPrintSheetPdf_(ss, sh, pdfFileName_(payload));
  return {
    pageCount: printed.pageCount,
    lineCount: printed.lineCount,
    sheetNames: printed.sheetNames,
    saveId: saved ? saved.saveId : '',
    savedAt: saved ? saved.savedAt : '',
    filename: blob.getName(),
    pdfBase64: Utilities.base64Encode(blob.getBytes())
  };
}

function saveInvoiceDraft_(payload) {
  if (!payload || !payload.header) {
    throw new Error('保存する請求書がありません。');
  }
  const kNo = normalize_(payload.header.kNo);
  if (!kNo) {
    throw new Error('K-No を入力してください。');
  }
  const items = (payload.items || []).filter(function (it) {
    return rowHasContent_(it);
  });
  const sh = ensureInvoiceSaveSheet_();
  const saveId = Utilities.getUuid();
  const savedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const h = payload.header;
  const sum = payload.summary || {};
  const rows = [];
  rows.push(invoiceSaveHeadRow_(saveId, savedAt, h, sum, kNo));
  for (let i = 0; i < items.length; i++) {
    rows.push(invoiceSaveLineRow_(saveId, savedAt, h, kNo, i + 1, items[i]));
  }
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, INVOICE_SAVE_HEADERS_.length).setValues(rows);
  return { saveId: saveId, savedAt: savedAt, lineCount: items.length };
}

function invoiceSaveHeadRow_(saveId, savedAt, h, sum, kNo) {
  return [
    saveId, 'HEAD', 0, kNo, savedAt,
    h.userName || '', h.plate || '', h.billDate || '', h.inDate || '', h.outDate || '',
    h.dept || '', h.serviceType || '', h.receptionist || h.staff || '',
    sum.techPct == null ? '' : sum.techPct,
    sum.partPct == null ? '' : sum.partPct,
    '', '', '', '', '', '', '', '', '', '', '', ''
  ];
}

function invoiceSaveLineRow_(saveId, savedAt, h, kNo, lineNo, it) {
  return [
    saveId, 'LINE', lineNo, kNo, savedAt,
    h.userName || '', h.plate || '', h.billDate || '', '', '',
    '', '', '', '', '',
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

function ensureInvoiceSaveSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = CONFIG.invoiceSave.sheetName;
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, INVOICE_SAVE_HEADERS_.length).setValues([INVOICE_SAVE_HEADERS_]);
    sh.getRange(1, 1, 1, INVOICE_SAVE_HEADERS_.length).setFontWeight('bold').setBackground('#e8f0ec');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 80);
    sh.setColumnWidth(4, 100);
    sh.setColumnWidth(5, 150);
  }
  return sh;
}

function pdfFileName_(payload) {
  const kNo = normalize_(payload && payload.header && payload.header.kNo) || '請求書';
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
