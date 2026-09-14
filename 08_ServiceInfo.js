/**
 * 整備情報マスタ。
 * レイアウト（縦持ち）:
 *   A 全て（部門未選択時の整備種別。上からこの順）
 *   B 整備種別1＝大型 / C 整備種別2＝小型 / D 整備種別3＝BP板金 / E 整備種別4＝部品販売
 *   F 受付メール / G 受付担当（同じ行。Googleアカウントと突き合わせて新規の初期値）
 *   H 値引技術% / I 値引部品%（2行目が新規入力の初期値。保存済みには使わない）
 *   J 登録地名（請求入力の登録番号コンボ）
 */

var SERVICE_ALL_HEADER_ = '全て';
var SERVICE_RECV_MAIL_COL_ = 6;
var SERVICE_RECV_COL_ = 7;
var SERVICE_DISC_COL_ = 8;
var SERVICE_PLATE_COL_ = 10;
var SERVICE_RECV_HEADER_ = '受付担当';
var SERVICE_RECV_MAIL_HEADER_ = '受付メール';
var SERVICE_RECV_MAIL_ALIASES_ = ['受付メール', 'メール', 'googleアカウント', 'アカウント'];
var SERVICE_PLATE_HEADER_ = '登録地名';
var SERVICE_PLATE_ALIASES_ = ['登録地名', '地名', 'ナンバー地名', '運輸支局'];
var SERVICE_PLATE_AREAS_DEFAULT_ = ['苫小牧', '室蘭', '北九州'];
var SERVICE_TECH_PCT_HEADER_ = '値引技術%';
var SERVICE_PART_PCT_HEADER_ = '値引部品%';
var SERVICE_TECH_PCT_ALIASES_ = ['値引技術%', '技術値引%', '値引技術'];
var SERVICE_PART_PCT_ALIASES_ = ['値引部品%', '部品値引%', '値引部品'];

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
  try {
    if (!sh || !isServiceInfoLayout_(sh) || !isServiceInfoRecvPairLayout_(sh)) {
      sh = rebuildServiceInfoSheet_(ss, parsed);
    } else {
      ensureServiceInfoDiscCols_(sh, parsed);
      ensureServiceInfoPlateCol_(sh, parsed);
      ensureServiceInfoRecvMailCol_(sh, parsed);
    }
    return parseServiceInfoSheet_(sh);
  } catch (err) {
    return parsed;
  }
}

function isServiceInfoLayout_(sheet) {
  if (!sheet) {
    return false;
  }
  if (normalize_(sheet.getRange(1, 1).getValue()) !== SERVICE_ALL_HEADER_) {
    return false;
  }
  const f = normalize_(sheet.getRange(1, SERVICE_RECV_MAIL_COL_).getValue());
  const g = normalize_(sheet.getRange(1, SERVICE_RECV_COL_).getValue());
  return f === SERVICE_RECV_HEADER_ || f === SERVICE_RECV_MAIL_HEADER_ || g === SERVICE_RECV_HEADER_;
}

function isServiceInfoRecvPairLayout_(sheet) {
  if (!sheet) {
    return false;
  }
  return normalize_(sheet.getRange(1, SERVICE_RECV_MAIL_COL_).getValue()) === SERVICE_RECV_MAIL_HEADER_ &&
    normalize_(sheet.getRange(1, SERVICE_RECV_COL_).getValue()) === SERVICE_RECV_HEADER_;
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
    receptionistMails: [],
    plateAreas: SERVICE_PLATE_AREAS_DEFAULT_.slice(),
    rows: [],
    techPct: defaultDiscPcts_().techPct,
    partPct: defaultDiscPcts_().partPct
  };
}

function defaultDiscPcts_() {
  const cfg = CONFIG.serviceInfo || {};
  return {
    techPct: clampInvoicePct_(cfg.defaultTechPct, 3),
    partPct: clampInvoicePct_(cfg.defaultPartPct, 10)
  };
}

function clampInvoicePct_(v, fallback) {
  const fb = fallback == null ? 0 : fallback;
  if (v === '' || v == null) {
    return fb;
  }
  const n = Number(String(v).replace(/[%％]/g, '').replace(/,/g, '').trim());
  if (!isFinite(n)) {
    return fb;
  }
  return Math.max(0, Math.min(100, Math.round(n)));
}

function headerIndexAny_(header, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const idx = header.indexOf(normalize_(aliases[i]));
    if (idx >= 0) {
      return idx;
    }
  }
  return -1;
}

function firstPctInCol_(values, col, startRow) {
  if (col < 0 || !values) {
    return '';
  }
  for (let i = startRow; i < values.length; i++) {
    if (col >= values[i].length) {
      continue;
    }
    const raw = values[i][col];
    if (raw === '' || raw == null) {
      continue;
    }
    const n = Number(String(raw).replace(/[%％]/g, '').replace(/,/g, '').trim());
    if (isFinite(n)) {
      return n;
    }
  }
  return '';
}

function applyDiscPctsFromValues_(out, values) {
  const d = defaultDiscPcts_();
  out.techPct = d.techPct;
  out.partPct = d.partPct;
  if (!values || !values.length) {
    return out;
  }
  const header = values[0].map(function (v) {
    return normalize_(v);
  });
  const t = firstPctInCol_(values, headerIndexAny_(header, SERVICE_TECH_PCT_ALIASES_), 1);
  const p = firstPctInCol_(values, headerIndexAny_(header, SERVICE_PART_PCT_ALIASES_), 1);
  if (t !== '') {
    out.techPct = clampInvoicePct_(t, d.techPct);
  }
  if (p !== '') {
    out.partPct = clampInvoicePct_(p, d.partPct);
  }
  return out;
}

function loadDiscPctDefaults_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(CONFIG.serviceInfo.sheetName);
    const parsed = parseServiceInfoSheet_(sh);
    return {
      techPct: clampInvoicePct_(parsed && parsed.techPct, defaultDiscPcts_().techPct),
      partPct: clampInvoicePct_(parsed && parsed.partPct, defaultDiscPcts_().partPct)
    };
  } catch (err) {
    return defaultDiscPcts_();
  }
}

function writeServiceInfoDiscCols_(sh, pct) {
  const d = pct || defaultDiscPcts_();
  const col = SERVICE_DISC_COL_;
  sh.getRange(1, col, 1, 2).setValues([[SERVICE_TECH_PCT_HEADER_, SERVICE_PART_PCT_HEADER_]]);
  sh.getRange(1, col, 1, 2).setFontWeight('bold').setBackground('#e8f0ec');
  sh.getRange(2, col, 1, 2).setValues([[d.techPct, d.partPct]]);
  sh.getRange(2, col, 1, 2).setNumberFormat('0').setFontSize(11);
  sh.setColumnWidth(col, 110);
  sh.setColumnWidth(col + 1, 110);
  sh.getRange(1, col).setNote('新規の請求入力に使う値引％の初期値。保存済みの請求書には影響しません。');
}

function applyPlateAreasFromValues_(out, values, header, startRow) {
  if (!out || !values || !values.length) {
    return;
  }
  const heads = header && header.length
    ? header
    : values[0].map(function (v) { return normalize_(v); });
  const col = headerIndexAny_(heads, SERVICE_PLATE_ALIASES_);
  if (col < 0) {
    return;
  }
  const list = collectVerticalCol_(values, col, startRow == null ? 1 : startRow);
  if (list.length) {
    out.plateAreas = list;
  }
}

function applyReceptionistsFromValues_(out, values, header, startRow) {
  if (!out || !values || !values.length) {
    return;
  }
  const heads = header && header.length
    ? header
    : values[0].map(function (v) { return normalize_(v); });
  const recvCol = heads.indexOf(SERVICE_RECV_HEADER_);
  if (recvCol < 0) {
    return;
  }
  const mailCol = headerIndexAny_(heads, SERVICE_RECV_MAIL_ALIASES_);
  const from = startRow == null ? 1 : startRow;
  for (let i = from; i < values.length; i++) {
    if (recvCol >= values[i].length) {
      continue;
    }
    const name = String(values[i][recvCol] == null ? '' : values[i][recvCol]).replace(/\u3000/g, ' ').trim();
    const key = normalize_(name);
    if (!name || key === SERVICE_RECV_HEADER_ || isServiceInfoSkipLabel_(name)) {
      continue;
    }
    if (out.receptionists.indexOf(name) !== -1) {
      continue;
    }
    out.receptionists.push(name);
    let mail = '';
    if (mailCol >= 0 && mailCol < values[i].length) {
      mail = String(values[i][mailCol] == null ? '' : values[i][mailCol]).replace(/\u3000/g, ' ').trim().toLowerCase();
    }
    out.receptionistMails.push(mail);
  }
}

function receptionistForOperator_(parsed, email) {
  const mail = String(email || '').trim().toLowerCase();
  if (!mail) {
    return '';
  }
  const local = mail.split('@')[0] || '';
  const names = (parsed && parsed.receptionists) || [];
  const mails = (parsed && parsed.receptionistMails) || [];
  let i;
  for (i = 0; i < names.length; i++) {
    const m = String(mails[i] || '').trim().toLowerCase();
    if (!m) {
      continue;
    }
    const mLocal = m.indexOf('@') >= 0 ? m.split('@')[0] : m;
    if (m === mail || m === local || mLocal === local) {
      return names[i];
    }
  }
  for (i = 0; i < names.length; i++) {
    const n = normalize_(names[i]);
    if (n && (n === normalize_(local) || n === normalize_(mail))) {
      return names[i];
    }
  }
  return '';
}

function writeServiceInfoRecvMailCol_(sh, names, mails) {
  const col = SERVICE_RECV_MAIL_COL_;
  sh.getRange(1, col).setValue(SERVICE_RECV_MAIL_HEADER_).setFontWeight('bold').setBackground('#e8f0ec');
  sh.getRange(1, col).setNote(
    'G列の受付担当と同じ行に、請求入力を開くGoogleアカウントを書きます。\n' +
    'メール全体でも @ より前だけでも可。一致した名前が受付担当の初期値になります。'
  );
  sh.setColumnWidth(col, 200);
  const list = names || [];
  const acc = mails || [];
  const last = Math.max(sh.getLastRow(), list.length + 2);
  if (last >= 2) {
    sh.getRange(2, col, last - 1, 1).clearContent();
  }
  if (!list.length) {
    return;
  }
  const body = list.map(function (name, i) {
    return [acc[i] || ''];
  });
  sh.getRange(3, col, body.length, 1).setValues(body);
}

function serviceInfoHasRecvMailCol_(sheet) {
  if (!sheet) {
    return false;
  }
  const header = normalize_(sheet.getRange(1, SERVICE_RECV_MAIL_COL_).getValue());
  return SERVICE_RECV_MAIL_ALIASES_.indexOf(header) >= 0;
}

function ensureServiceInfoRecvMailCol_(sheet, parsed) {
  if (serviceInfoHasRecvMailCol_(sheet)) {
    return;
  }
  try {
    writeServiceInfoRecvMailCol_(
      sheet,
      parsed && parsed.receptionists,
      parsed && parsed.receptionistMails
    );
  } catch (err) {}
}

function writeServiceInfoPlateCol_(sh, areas) {
  const list = (areas && areas.length) ? areas.slice() : SERVICE_PLATE_AREAS_DEFAULT_.slice();
  const col = SERVICE_PLATE_COL_;
  sh.getRange(1, col).setValue(SERVICE_PLATE_HEADER_).setFontWeight('bold').setBackground('#e8f0ec');
  sh.getRange(1, col).setNote('請求入力の登録番号の地名候補。上からこの順。リストにない地名も手入力できます。');
  sh.setColumnWidth(col, 120);
  const last = Math.max(sh.getLastRow(), list.length + 2);
  if (last >= 2) {
    sh.getRange(2, col, last - 1, 1).clearContent();
  }
  if (list.length) {
    const body = list.map(function (v) { return [v]; });
    sh.getRange(3, col, body.length, 1).setValues(body);
  }
}

function serviceInfoHasPlateCol_(sheet) {
  if (!sheet) {
    return false;
  }
  const header = normalize_(sheet.getRange(1, SERVICE_PLATE_COL_).getValue());
  return SERVICE_PLATE_ALIASES_.indexOf(header) >= 0;
}

function ensureServiceInfoPlateCol_(sheet, parsed) {
  if (serviceInfoHasPlateCol_(sheet)) {
    const last = Math.max(sheet.getLastRow(), 2);
    const vals = sheet.getRange(2, SERVICE_PLATE_COL_, last - 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0] || '').replace(/\u3000/g, ' ').trim()) {
        return;
      }
    }
  }
  try {
    writeServiceInfoPlateCol_(sheet, parsed && parsed.plateAreas);
  } catch (err) {}
}

function serviceInfoHasDiscCols_(sheet) {
  if (!sheet) {
    return false;
  }
  const header = sheet.getRange(1, SERVICE_DISC_COL_, 1, 2).getValues()[0].map(function (v) {
    return normalize_(v);
  });
  return headerIndexAny_(header, SERVICE_TECH_PCT_ALIASES_) >= 0 &&
    headerIndexAny_(header, SERVICE_PART_PCT_ALIASES_) >= 0;
}

function ensureServiceInfoDiscCols_(sheet, parsed) {
  if (serviceInfoHasDiscCols_(sheet)) {
    return;
  }
  try {
    writeServiceInfoDiscCols_(sheet, parsed || defaultDiscPcts_());
  } catch (err) {}
}

function isServiceInfoSkipLabel_(value) {
  const t = normalize_(value);
  if (!t) {
    return true;
  }
  if (t === SERVICE_ALL_HEADER_ || t === SERVICE_RECV_HEADER_ || t === SERVICE_PLATE_HEADER_ ||
      t === SERVICE_RECV_MAIL_HEADER_ || t === '未選択') {
    return true;
  }
  if (SERVICE_RECV_MAIL_ALIASES_.indexOf(t) !== -1) {
    return true;
  }
  if (SERVICE_PLATE_ALIASES_.indexOf(t) !== -1) {
    return true;
  }
  if (SERVICE_TECH_PCT_ALIASES_.indexOf(t) !== -1 || SERVICE_PART_PCT_ALIASES_.indexOf(t) !== -1) {
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
  const width = Math.max(sheet.getLastColumn(), 10);
  const values = sheet.getRange(1, 1, last, width).getValues();
  const header = values[0].map(function (v) {
    return normalize_(v);
  });
  const allCol = header.indexOf(SERVICE_ALL_HEADER_);
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
  applyReceptionistsFromValues_(out, values, header, 2);
  applyDiscPctsFromValues_(out, values);
  applyPlateAreasFromValues_(out, values, header, 1);
  return out;
}

function parseServiceInfoOldVertical_(sheet) {
  const out = emptyServiceInfo_();
  const last = Math.max(sheet.getLastRow(), 1);
  const width = Math.max(sheet.getLastColumn(), 10);
  const values = sheet.getRange(1, 1, last, width).getValues();
  const header = values[0].map(function (v) {
    return normalize_(v);
  });
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
  applyReceptionistsFromValues_(out, values, header, 1);
  applyDiscPctsFromValues_(out, values);
  applyPlateAreasFromValues_(out, values, header, 1);
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
  applyReceptionistsFromValues_(out, values, header, 1);
  applyDiscPctsFromValues_(out, values);
  applyPlateAreasFromValues_(out, values, header, 1);
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
  const receptionistMails = (parsed && parsed.receptionistMails) ? parsed.receptionistMails.slice() : [];
  const plateAreas = (parsed && parsed.plateAreas && parsed.plateAreas.length)
    ? parsed.plateAreas.slice()
    : SERVICE_PLATE_AREAS_DEFAULT_.slice();

  const height = Math.max(
    allList.length,
    lists.reduce(function (n, a) { return Math.max(n, a.length); }, 0),
    receptionists.length,
    plateAreas.length,
    8
  ) + 2;
  sh.clear();
  sh.getRange(1, 1, 1, 7).setValues([[
    SERVICE_ALL_HEADER_, '整備種別1', '整備種別2', '整備種別3', '整備種別4',
    SERVICE_RECV_MAIL_HEADER_, SERVICE_RECV_HEADER_
  ]]);
  sh.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e8f0ec');
  sh.getRange(2, 1, 1, 7).setValues([['未選択', '大型', '小型', 'BP板金', '部品販売', '', '']]);
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
      receptionistMails[i] || '',
      receptionists[i] || ''
    ]);
  }
  sh.getRange(3, 1, body.length, 7).setValues(body);

  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 140);
  sh.setColumnWidth(2, 140);
  sh.setColumnWidth(3, 140);
  sh.setColumnWidth(4, 140);
  sh.setColumnWidth(5, 140);
  sh.setColumnWidth(SERVICE_RECV_MAIL_COL_, 200);
  sh.setColumnWidth(SERVICE_RECV_COL_, 120);
  writeServiceInfoDiscCols_(sh, parsed);
  writeServiceInfoPlateCol_(sh, plateAreas);
  sh.getRange(1, 1).setNote(
    'A列「全て」＝整備部門が未選択のときの整備種別（上からこの順）。\n' +
    '既定: 車検大型→点検大型→一般大型→車検小型→点検小型→一般小型→構造変更→板金塗装→部品販売→特装→諸経費。\n' +
    'B〜E＝部門別（2行目は部門名。消さない）。F＝受付メール。G＝受付担当。H・I＝新規入力の値引％初期値。J＝登録地名。'
  );
  sh.getRange(1, SERVICE_RECV_MAIL_COL_).setNote(
    'G列の受付担当と同じ行に、請求入力を開くGoogleアカウントを書きます。\n' +
    'メール全体でも @ より前だけでも可。一致した名前が受付担当の初期値になります。'
  );
  sh.getRange(1, SERVICE_RECV_COL_).setNote('受付担当を縦に並べます。左のF列に同じ行のGoogleアカウントを書くと、請求入力の初期値になります。');
  sh.getRange(1, SERVICE_PLATE_COL_).setNote('請求入力の登録番号の地名候補。上からこの順。リストにない地名も手入力できます。');
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
  SpreadsheetApp.getActiveSpreadsheet().toast('整備情報シートを整理しました（全て列＋種別1〜4＋値引％＋登録地名＋受付メール）', '請求書入力', 5);
}
