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
  const empty = { departments: [], typesByDept: {}, allServiceTypes: [], receptionists: [], rows: [] };
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
  const departments = [];
  const rows = [];
  const receptionists = [];

  for (let i = dataStart; i < values.length; i++) {
    const dept = normalize_(values[i][deptCol]);
    const types = [];
    if (dept) {
      for (let c = typeStart; c <= typeEnd && c < values[i].length; c++) {
        if (c === recvCol) {
          continue;
        }
        const t = normalize_(values[i][c]);
        if (t && types.indexOf(t) === -1) {
          types.push(t);
        }
      }
      if (dept === '部品販売' && !types.length) {
        types.push('部品販売');
      }
      if (dept === '板金塗装' && !types.length) {
        types.push('BP板金');
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
      rows.push({ dept: dept, types: types });
    }
    if (recvCol < values[i].length) {
      const rec = normalize_(values[i][recvCol]);
      if (rec && rec !== '受付担当' && receptionists.indexOf(rec) === -1) {
        receptionists.push(rec);
      }
    }
  }

  return {
    departments: departments,
    typesByDept: typesByDept,
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
    const types = rows[i].types || [];
    body.push([
      rows[i].dept,
      types[0] || '',
      types[1] || '',
      types[2] || '',
      types[3] || ''
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
    'A列＝整備部門。B〜E列＝その部門の整備種別（空欄可）。1つだけの部門は入力アプリで種別を自動セットします。\n' +
    '部品販売は種別「部品販売」、板金塗装は種別「BP板金」を推奨。\n' +
    'F列＝受付担当（行ごとに1名）。'
  );
  sh.getRange(1, 6).setNote('受付担当を縦に並べます。');
  return sh;
}

function defaultServiceInfoRows_() {
  return [
    { dept: '車検', types: [] },
    { dept: '一般整備', types: [] },
    { dept: '部品販売', types: ['部品販売'] },
    { dept: '板金塗装', types: ['BP板金'] }
  ];
}

function tidyServiceInfoSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.serviceInfo.sheetName);
  const parsed = parseServiceInfoSheet_(sh);
  rebuildServiceInfoSheet_(ss, parsed);
  SpreadsheetApp.getActiveSpreadsheet().toast('整備情報シートを整理しました', '請求書入力', 5);
}
