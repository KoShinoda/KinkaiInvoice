/**
 * 整備情報マスタ。
 * レイアウト（縦持ち）:
 *   A 全て（部門未選択時の整備種別。上からこの順）
 *   B 整備種別1＝大型 / C 整備種別2＝小型 / D 整備種別3＝BP板金 / E 整備種別4＝部品販売
 *   F 受付担当
 */

var SERVICE_ALL_HEADER_ = '全て';
var SERVICE_RECV_HEADER_ = '受付担当';

var SERVICE_DEPT_COLS_ = [
  { header: '整備種別1', dept: '大型' },
  { header: '整備種別2', dept: '小型' },
  { header: '整備種別3', dept: 'BP板金', fallback: '板金塗装' },
  { header: '整備種別4', dept: '部品販売', fallback: '部品販売' }
];

/** 部門未選択時の整備種別。シート「全て」列の初期順。 */
var SERVICE_ALL_TYPES_DEFAULT_ = [
  '車検大型',
  '点検大型',
  '一般大型',
  '車検小型',
  '点検小型',
  '一般小型',
  '構造変更',
  '板金塗装',
  '部品販売',
  '特装',
  '諸経費'
];

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
  return normalize_(sheet.getRange(1, 1).getValue()) === SERVICE_ALL_HEADER_ &&
    normalize_(sheet.getRange(1, 6).getValue()) === SERVICE_RECV_HEADER_;
}

function serviceDeptColumns_() {
  return SERVICE_DEPT_COLS_;
}

function serviceDepartments_() {
  return SERVICE_DEPT_COLS_.map(function (c) {
    return c.dept;
  });
}

function deptToServiceColIndex_(dept) {
  const n = normalize_(dept);
  if (n === '小型' || n === '一般整備') {
    return 1;
  }
  if (n === 'bp板金' || n === '板金塗装') {
    return 2;
  }
  if (n === '部品販売') {
    return 3;
  }
  return 0;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet=} sheet
 * @return {{departments: string[], typesByDept: Object<string, string[]>, typeSlotsByDept: Object<string, string[]>, allServiceTypes: string[], receptionists: string[], rows: object[]}}
 */
function parseServiceInfoSheet_(sheet) {
  const empty = emptyServiceInfo_();
  if (!sheet) {
    return empty;
  }
  if (isServiceInfoLayout_(sheet)) {
    return parseServiceInfoVertical_(sheet);
  }
  if (normalize_(sheet.getRange(1, 1).getValue()) === '整備種別1') {
    return parseServiceInfoOldVertical_(sheet);
  }
  return parseServiceInfoLegacy_(sheet);
}

function emptyServiceInfo_() {
  const typesByDept = {};
  const typeSlotsByDept = {};
  serviceDepartments_().forEach(function (dept) {
    typesByDept[dept] = [];
    typeSlotsByDept[dept] = ['', '', '', ''];
  });
  return {
    departments: serviceDepartments_(),
    typesByDept: typesByDept,
    typeSlotsByDept: typeSlotsByDept,
    allServiceTypes: SERVICE_ALL_TYPES_DEFAULT_.slice(),
    receptionists: [],
    rows: []
  };
}

function isServiceInfoSkipLabel_(value) {
  const t = normalize_(value);
  if (!t) {
    return true;
  }
  if (t === SERVICE_ALL_HEADER_ || t === SERVICE_RECV_HEADER_ || t === '未選択') {
    return true;
  }
  if (t === '大型' || t === '小型' || t === 'BP板金') {
    return true;
  }
  for (let i = 0; i < SERVICE_DEPT_COLS_.length; i++) {
    if (t === SERVICE_DEPT_COLS_[i].header) {
      return true;
    }
  }
  return false;
}

function collectVerticalCol_(values, colIndex, startRow) {
  const types = [];
  if (colIndex < 0) {
    return types;
  }
  for (let i = startRow; i < values.length; i++) {
    if (colIndex >= values[i].length) {
      continue;
    }
    const raw = String(values[i][colIndex] == null ? '' : values[i][colIndex]).replace(/\u3000/g, ' ').trim();
    if (!raw || isServiceInfoSkipLabel_(raw) || types.indexOf(raw) !== -1) {
      continue;
    }
    types.push(raw);
  }
  return types;
}

function parseServiceInfoVertical_(sheet) {
  const out = emptyServiceInfo_();
  const last = Math.max(sheet.getLastRow(), 1);
  const width = Math.max(sheet.getLastColumn(), 6);
  const values = sheet.getRange(1, 1, last, width).getValues();
  const header = values[0].map(function (v) {
    return normalize_(v);
  });
  const allCol = header.indexOf(SERVICE_ALL_HEADER_);
  const recvCol = header.indexOf(SERVICE_RECV_HEADER_);
  const cols = SERVICE_DEPT_COLS_;
  for (let c = 0; c < cols.length; c++) {
    const dept = cols[c].dept;
    const colIndex = header.indexOf(cols[c].header);
    const types = collectVerticalCol_(values, colIndex, 2);
    if (cols[c].fallback && types.indexOf(cols[c].fallback) === -1) {
      types.push(cols[c].fallback);
    }
    out.typesByDept[dept] = types;
    out.typeSlotsByDept[dept] = types.slice();
    out.rows.push({ dept: dept, types: types, slots: types.slice() });
  }
  out.allServiceTypes = mergeAllServiceTypes_(collectVerticalCol_(values, allCol, 2), null);
  if (recvCol >= 0) {
    for (let i = 2; i < values.length; i++) {
      const rec = normalize_(values[i][recvCol]);
      if (rec && rec !== SERVICE_RECV_HEADER_ && out.receptionists.indexOf(rec) === -1) {
        out.receptionists.push(String(values[i][recvCol]).replace(/\u3000/g, ' ').trim());
      }
    }
  }
  return out;
}

function parseServiceInfoOldVertical_(sheet) {
  const out = emptyServiceInfo_();
  const last = Math.max(sheet.getLastRow(), 1);
  const width = Math.max(sheet.getLastColumn(), 5);
  const values = sheet.getRange(1, 1, last, width).getValues();
  const header = values[0].map(function (v) {
    return normalize_(v);
  });
  const recvCol = header.indexOf(SERVICE_RECV_HEADER_) >= 0 ? header.indexOf(SERVICE_RECV_HEADER_) : 4;
  SERVICE_DEPT_COLS_.forEach(function (col, c) {
    const idx = header.indexOf(col.header);
    const types = collectVerticalCol_(values, idx >= 0 ? idx : c, 1);
    if (col.fallback && types.indexOf(col.fallback) === -1) {
      types.push(col.fallback);
    }
    out.typesByDept[col.dept] = types;
    out.typeSlotsByDept[col.dept] = types.slice();
    out.rows.push({ dept: col.dept, types: types, slots: types.slice() });
  });
  out.allServiceTypes = mergeAllServiceTypes_([], out);
  for (let i = 1; i < values.length; i++) {
    if (recvCol >= values[i].length) {
      continue;
    }
    const rec = normalize_(values[i][recvCol]);
    if (rec && rec !== SERVICE_RECV_HEADER_ && out.receptionists.indexOf(rec) === -1) {
      out.receptionists.push(String(values[i][recvCol]).replace(/\u3000/g, ' ').trim());
    }
  }
  return out;
}

function parseServiceInfoLegacy_(sheet) {
  const out = emptyServiceInfo_();
  const values = sheet.getDataRange().getValues();
  if (!values.length) {
    return out;
  }
  const header = values[0].map(function (v) {
    return normalize_(v);
  });
  let deptCol = 0;
  let typeStart = 1;
  let recvCol = 5;
  const deptIdx = header.indexOf('整備部門');
  const recvIdx = header.indexOf(SERVICE_RECV_HEADER_);
  if (deptIdx >= 0) {
    deptCol = deptIdx;
  }
  if (recvIdx >= 0) {
    recvCol = recvIdx;
  }
  for (let i = 1; i < values.length; i++) {
    const dept = normalize_(values[i][deptCol]);
    if (dept) {
      const col = deptToServiceColIndex_(dept);
      const mapped = SERVICE_DEPT_COLS_[col].dept;
      for (let s = 0; s < 4; s++) {
        const c = typeStart + s;
        if (c === recvCol || c >= values[i].length) {
          continue;
        }
        const t = normalize_(values[i][c]);
        if (t && out.typesByDept[mapped].indexOf(t) === -1) {
          out.typesByDept[mapped].push(String(values[i][c]).replace(/\u3000/g, ' ').trim());
        }
      }
    }
    if (recvCol < values[i].length) {
      const rec = normalize_(values[i][recvCol]);
      if (rec && rec !== SERVICE_RECV_HEADER_ && out.receptionists.indexOf(rec) === -1) {
        out.receptionists.push(String(values[i][recvCol]).replace(/\u3000/g, ' ').trim());
      }
    }
  }
  SERVICE_DEPT_COLS_.forEach(function (col) {
    if (col.fallback && out.typesByDept[col.dept].indexOf(col.fallback) === -1) {
      out.typesByDept[col.dept].push(col.fallback);
    }
    out.typeSlotsByDept[col.dept] = out.typesByDept[col.dept].slice();
    out.rows.push({
      dept: col.dept,
      types: out.typesByDept[col.dept].slice(),
      slots: out.typesByDept[col.dept].slice()
    });
  });
  out.allServiceTypes = mergeAllServiceTypes_([], out);
  return out;
}

/** 既定11件を先に、シートや部門列の追加分を後ろへ。 */
function mergeAllServiceTypes_(fromAllCol, parsed) {
  const out = [];
  const seen = {};
  function add(v) {
    const s = String(v || '').replace(/\u3000/g, ' ').trim();
    if (!s || seen[s] || isServiceInfoSkipLabel_(s)) {
      return;
    }
    seen[s] = true;
    out.push(s);
  }
  SERVICE_ALL_TYPES_DEFAULT_.forEach(add);
  (fromAllCol || []).forEach(add);
  if (parsed && parsed.typesByDept) {
    serviceDepartments_().forEach(function (dept) {
      (parsed.typesByDept[dept] || []).forEach(add);
    });
  }
  return out;
}

function rebuildServiceInfoSheet_(ss, parsed) {
  let sh = ss.getSheetByName(CONFIG.serviceInfo.sheetName);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.serviceInfo.sheetName);
  }

  const lists = SERVICE_DEPT_COLS_.map(function () {
    return [];
  });
  const srcTypes = (parsed && parsed.typesByDept) || {};
  Object.keys(srcTypes).forEach(function (dept) {
    const col = deptToServiceColIndex_(dept);
    (srcTypes[dept] || []).forEach(function (t) {
      const v = String(t || '').trim();
      if (v && lists[col].indexOf(v) === -1) {
        lists[col].push(v);
      }
    });
  });
  SERVICE_DEPT_COLS_.forEach(function (col, i) {
    if (col.fallback && lists[i].indexOf(col.fallback) === -1) {
      lists[i].push(col.fallback);
    }
  });

  const allList = mergeAllServiceTypes_((parsed && parsed.allServiceTypes) || [], parsed);
  const receptionists = (parsed && parsed.receptionists && parsed.receptionists.length)
    ? parsed.receptionists.slice()
    : [];

  const height = Math.max(
    allList.length,
    lists.reduce(function (n, a) { return Math.max(n, a.length); }, 0),
    receptionists.length,
    8
  ) + 2;
  sh.clear();
  sh.getRange(1, 1, 1, 6).setValues([[
    SERVICE_ALL_HEADER_, '整備種別1', '整備種別2', '整備種別3', '整備種別4', SERVICE_RECV_HEADER_
  ]]);
  sh.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#e8f0ec');
  sh.getRange(2, 1, 1, 6).setValues([['未選択', '大型', '小型', 'BP板金', '部品販売', '']]);
  sh.getRange(2, 1, 1, 5).setFontColor('#5b6570').setFontSize(10);

  const bodyH = Math.max(height - 2, 1);
  const body = [];
  for (let i = 0; i < bodyH; i++) {
    body.push([
      allList[i] || '',
      lists[0][i] || '',
      lists[1][i] || '',
      lists[2][i] || '',
      lists[3][i] || '',
      receptionists[i] || ''
    ]);
  }
  sh.getRange(3, 1, body.length, 6).setValues(body);

  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 140);
  sh.setColumnWidth(2, 140);
  sh.setColumnWidth(3, 140);
  sh.setColumnWidth(4, 140);
  sh.setColumnWidth(5, 140);
  sh.setColumnWidth(6, 120);
  sh.getRange(1, 1).setNote(
    'A列「全て」＝整備部門が未選択のときの整備種別（上からこの順）。\n' +
    '既定: 車検大型→点検大型→一般大型→車検小型→点検小型→一般小型→構造変更→板金塗装→部品販売→特装→諸経費。\n' +
    'B〜E＝部門別（2行目は部門名。消さない）。F＝受付担当。'
  );
  sh.getRange(1, 6).setNote('受付担当を縦に並べます。');
  return sh;
}

function typesForServiceDept_(dept) {
  const info = loadServiceInfo_();
  if (!normalize_(dept)) {
    return (info.allServiceTypes || SERVICE_ALL_TYPES_DEFAULT_).slice();
  }
  const mapped = SERVICE_DEPT_COLS_[deptToServiceColIndex_(dept)].dept;
  const types = (info.typesByDept[mapped] || []).slice();
  if (mapped === '部品販売' && types.indexOf('部品販売') === -1) {
    types.push('部品販売');
  }
  if (mapped === 'BP板金' && types.indexOf('板金塗装') === -1) {
    types.push('板金塗装');
  }
  return types;
}

function tidyServiceInfoSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.serviceInfo.sheetName);
  const parsed = parseServiceInfoSheet_(sh);
  rebuildServiceInfoSheet_(ss, parsed);
  SpreadsheetApp.getActiveSpreadsheet().toast('整備情報シートを整理しました（全て列＋種別1〜4）', '請求書入力', 5);
}
