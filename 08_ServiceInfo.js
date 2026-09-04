/**
 * 整備情報マスタ（整備部門・整備種別・受付担当）。
 * レイアウト：
 *   A 整備部門 / B〜E 整備種別1〜4
 *   F 受付担当
 */

function loadServiceInfo_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.serviceInfo.sheetName);
  const parsed = parseServiceInfoSheet_(sh);
  if (!sh || !isServiceInfoLayout_(sh)) {
    sh = rebuildServiceInfoSheet_(ss, parsed);
  }
  return parseServiceInfoSheet_(sh);
}

function isServiceInfoLayout_(sheet) {
  if (!sheet) {
    return false;
  }
  return normalize_(sheet.getRange(1, 1).getValue()) === '整備部門' &&
    normalize_(sheet.getRange(1, 6).getValue()) === '受付担当';
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet=} sheet
 * @return {{departments: string[], typesByDept: Object<string, string[]>, allServiceTypes: string[], receptionists: string[], rows: object[]}}
 */
function parseServiceInfoSheet_(sheet) {
  const empty = { departments: [], typesByDept: {}, typeSlotsByDept: {}, allServiceTypes: [], receptionists: [], rows: [] };
  if (!sheet) {
    return empty;
  }
  const values = sheet.getDataRange().getValues();
  if (!values.length) {
    return empty;
  }

  let deptCol = 0;
  let typeStart = 1;
  let typeEnd = 4;
  let recvCol = 5;
  let dataStart = 1;
  const header = values[0].map(function (v) {
    return normalize_(v);
  });

  const deptIdx = header.indexOf('整備部門');
  const recvIdx = header.indexOf('受付担当');
  if (deptIdx >= 0) {
    deptCol = deptIdx;
    dataStart = 1;
  }
  if (recvIdx >= 0) {
    recvCol = recvIdx;
  }

  const typesByDept = {};
  const typeSlotsByDept = {};
  const departments = [];
  const rows = [];
  const receptionists = [];

  for (let i = dataStart; i < values.length; i++) {
    const dept = normalize_(values[i][deptCol]);
    const slots = ['', '', '', ''];
    const types = [];
    if (dept) {
      for (let s = 0; s < 4; s++) {
        const c = typeStart + s;
        if (c === recvCol || c >= values[i].length) {
          continue;
        }
        slots[s] = normalize_(values[i][c]);
        if (slots[s] && types.indexOf(slots[s]) === -1) {
          types.push(slots[s]);
        }
      }
      if (dept === '部品販売') {
        slots[3] = '部品販売';
        if (types.indexOf('部品販売') === -1) {
          types.push('部品販売');
        }
      }
      if ((dept === 'BP板金' || dept === '板金塗装') && !slots[2] && types.indexOf('板金塗装') === -1) {
        slots[2] = '板金塗装';
        types.push('板金塗装');
      }
      if (departments.indexOf(dept) === -1) {
        departments.push(dept);
      }
      if (!typesByDept[dept]) {
        typesByDept[dept] = [];
      }
      types.forEach(function (t) {
        if (typesByDept[dept].indexOf(t) === -1) {
          typesByDept[dept].push(t);
        }
      });
      typeSlotsByDept[dept] = slots;
      rows.push({ dept: dept, types: types, slots: slots });
    }
    if (recvCol < values[i].length) {
      const rec = normalize_(values[i][recvCol]);
      if (rec && rec !== '受付担当' && receptionists.indexOf(rec) === -1) {
        receptionists.push(rec);
      }
    }
    [5, 6].forEach(function (c) {
      if (c === recvCol || c >= values[i].length) {
        return;
      }
      const rec = normalize_(values[i][c]);
      if (rec && rec !== '受付担当' && receptionists.indexOf(rec) === -1) {
        receptionists.push(rec);
      }
    });
  }

  return {
    departments: departments,
    typesByDept: typesByDept,
    typeSlotsByDept: typeSlotsByDept,
    allServiceTypes: uniqueValues_(Object.keys(typesByDept).reduce(function (acc, dept) {
      return acc.concat(typesByDept[dept]);
    }, [])),
    receptionists: receptionists,
    rows: rows
  };
}

function rebuildServiceInfoSheet_(ss, parsed) {
  let sh = ss.getSheetByName(CONFIG.serviceInfo.sheetName);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.serviceInfo.sheetName);
  }

  const rows = (parsed && parsed.rows && parsed.rows.length)
    ? parsed.rows
    : defaultServiceInfoRows_();
  const receptionists = (parsed && parsed.receptionists && parsed.receptionists.length)
    ? parsed.receptionists
    : [];

  const height = Math.max(rows.length, receptionists.length, 8) + 2;
  sh.clear();

  sh.getRange(1, 1, 1, 6).setValues([['整備部門', '整備種別1', '整備種別2', '整備種別3', '整備種別4', '受付担当']]);
  sh.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#e8f0ec');

  const body = [];
  for (let i = 0; i < rows.length; i++) {
    const slots = rows[i].slots || [];
    const types = rows[i].types || [];
    body.push([
      rows[i].dept,
      slots[0] || types[0] || '',
      slots[1] || types[1] || '',
      slots[2] || types[2] || '',
      slots[3] || types[3] || ''
    ]);
  }
  if (body.length) {
    sh.getRange(2, 1, body.length, 5).setValues(body);
  }
  if (receptionists.length) {
    const recVals = receptionists.map(function (name) {
      return [name];
    });
    sh.getRange(2, 6, recVals.length, 1).setValues(recVals);
  }

  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 120);
  sh.setColumnWidth(2, 120);
  sh.setColumnWidth(3, 120);
  sh.setColumnWidth(4, 120);
  sh.setColumnWidth(5, 120);
  sh.setColumnWidth(6, 120);
  sh.getRange(1, 1).setNote(
    'A列＝整備部門。B〜E列＝その部門の整備種別（空欄可。列位置を維持する）。\n' +
    '部門「BP板金」は整備種別3「板金塗装」、部門「部品販売」は整備種別4「部品販売」を入力アプリで自動セットします。\n' +
    'F列＝受付担当（行ごとに1名）。'
  );
  sh.getRange(1, 6).setNote('受付担当を縦に並べます。');
  return sh;
}

function defaultServiceInfoRows_() {
  return [
    { dept: '車検', types: [] },
    { dept: '一般整備', types: [] },
    { dept: '部品販売', types: ['部品販売'], slots: ['', '', '', '部品販売'] },
    { dept: 'BP板金', types: ['板金塗装'], slots: ['', '', '板金塗装', ''] }
  ];
}

function tidyServiceInfoSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.serviceInfo.sheetName);
  const parsed = parseServiceInfoSheet_(sh);
  rebuildServiceInfoSheet_(ss, parsed);
  SpreadsheetApp.getActiveSpreadsheet().toast('整備情報シートを整理しました', '請求書入力', 5);
}
