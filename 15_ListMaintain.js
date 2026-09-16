/**
 * 入力アプリ「リスト追加」タブから、作業リスト／部品リスト／作業者リストへ書く。
 * シート全体の並べ替えやプルダウン再生成はしない（図形ボタンの更新に任せる）。
 */

function listMaintainNumeric_(v) {
  if (v === '' || v == null) {
    return '';
  }
  if (typeof v === 'number' && isFinite(v)) {
    return v;
  }
  const n = Number(String(v).replace(/,/g, '').trim());
  return isFinite(n) ? n : '';
}

function listMaintainSetCell_(row, col1, val) {
  if (!col1) {
    return;
  }
  row[col1 - 1] = val == null ? '' : val;
}

function listMaintainBlankRow_(width) {
  const row = [];
  for (let i = 0; i < width; i++) {
    row.push('');
  }
  return row;
}

function listMaintainDeleteRows_(sheet, rowNums) {
  const rows = (rowNums || []).slice().filter(function (n) {
    return n >= 2;
  }).sort(function (a, b) {
    return a - b;
  });
  let i = rows.length - 1;
  while (i >= 0) {
    const end = rows[i];
    let start = end;
    while (i > 0 && rows[i - 1] === start - 1) {
      i--;
      start = rows[i];
    }
    sheet.deleteRows(start, end - start + 1);
    i--;
  }
}

function listMaintainAppendRows_(sheet, body, width) {
  if (!body || !body.length) {
    return;
  }
  const start = Math.max(sheet.getLastRow(), 1) + 1;
  sheet.getRange(start, 1, body.length, width).setValues(body);
}

function listMaintainPreserveLabel_(sheet, values, col1, keepRows, label) {
  if (!col1 || !keepRows || !keepRows.length || !label) {
    return;
  }
  const has = keepRows.some(function (r) {
    return normalize_(values[r - 1][col1 - 1]) === label;
  });
  if (!has) {
    sheet.getRange(keepRows[0], col1).setValue(label);
  }
}

function listMaintainClientMaster_() {
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
    ordersPending: !!ctx.ordersPending
  };
}

function listMaintainWorkLineHasContent_(line) {
  if (!line) {
    return false;
  }
  return !!(normalize_(line.content) || isFilled_(line.fee) || normalize_(line.workerCode) ||
    normalize_(line.partMajor) || normalize_(line.partMid) ||
    isFilled_(line.qty) || isFilled_(line.unitPrice));
}

function listMaintainWorkRow_(cols, width, major, mid, line) {
  const row = listMaintainBlankRow_(width);
  listMaintainSetCell_(row, cols.major, major);
  listMaintainSetCell_(row, cols.mid, mid);
  listMaintainSetCell_(row, cols.content, normalize_(line && line.content));
  listMaintainSetCell_(row, cols.fee, listMaintainNumeric_(line && line.fee));
  const worker = invoiceTemplateWorkerWriteFast_(line && line.workerCode);
  listMaintainSetCell_(row, cols.workerCode, worker);
  listMaintainSetCell_(row, cols.partMajor, normalize_(line && line.partMajor));
  listMaintainSetCell_(row, cols.partMid, normalize_(line && line.partMid));
  listMaintainSetCell_(row, cols.qty, listMaintainNumeric_(line && line.qty));
  listMaintainSetCell_(row, cols.unitPrice, listMaintainNumeric_(line && line.unitPrice));
  return row;
}

/**
 * @param {{major: string, mid: string, lines: object[], overwrite: boolean}} payload
 */
function saveWorkListPack(payload) {
  payload = payload || {};
  const major = normalize_(payload.major);
  const mid = normalize_(payload.mid);
  if (!major) {
    throw new Error('大項目を入力してください。');
  }
  if (!mid) {
    throw new Error('中項目を入力してください。');
  }
  const lines = (payload.lines || []).filter(listMaintainWorkLineHasContent_);
  const bodyLines = lines.length ? lines : [{}];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.workList.sheetName);
  if (!sh) {
    throw new Error('シート「' + CONFIG.workList.sheetName + '」がありません。');
  }
  let overwritten = false;
  writeInternal_(function () {
    const lastCol = Math.max(sh.getLastColumn(), 1);
    const values = sh.getDataRange().getValues();
    const headers = values[CONFIG.workList.headerRow - 1] || [];
    const cols = resolveColumns_(headers, CONFIG.workList.headers);
    if (!cols.major || !cols.mid) {
      throw new Error('作業リストに「大項目」「中項目」ヘッダーが見つかりません。');
    }
    const width = lastDataHeaderCol_(headers);
    const parsed = parseWorkList_(values, cols);
    const packRows = [];
    const keepMajorRows = [];
    parsed.forEach(function (row) {
      if (normalize_(row.major) !== major) {
        return;
      }
      if (normalize_(row.mid) === mid) {
        packRows.push(row.sourceIndex);
      } else {
        keepMajorRows.push(row.sourceIndex);
      }
    });
    overwritten = packRows.length > 0;
    if (overwritten && !payload.overwrite) {
      throw new Error('作業「' + major + ' / ' + mid + '」は既にあります。上書きする場合は確認のうえ保存してください。');
    }
    const body = bodyLines.map(function (line) {
      return listMaintainWorkRow_(cols, width, major, mid, line);
    });
    if (overwritten) {
      keepMajorRows.sort(function (a, b) {
        return a - b;
      });
      listMaintainPreserveLabel_(sh, values, cols.major, keepMajorRows, major);
      listMaintainDeleteRows_(sh, packRows);
    }
    listMaintainAppendRows_(sh, body, width);
  });
  return invoiceJsonSafe_({
    ok: true,
    overwritten: overwritten,
    major: major,
    mid: mid,
    master: listMaintainClientMaster_()
  });
}

function listMaintainPartsLineHasContent_(line) {
  if (!line) {
    return false;
  }
  return !!(normalize_(line.set) || normalize_(line.content) ||
    isFilled_(line.qty) || isFilled_(line.unitPrice));
}

function listMaintainPartsRow_(cols, width, major, mid, line) {
  const row = listMaintainBlankRow_(width);
  const setName = normalize_(line && (line.set || line.content));
  listMaintainSetCell_(row, cols.major, major);
  listMaintainSetCell_(row, cols.mid, mid);
  listMaintainSetCell_(row, cols.set, setName);
  listMaintainSetCell_(row, cols.name, setName);
  listMaintainSetCell_(row, cols.qty, listMaintainNumeric_(line && line.qty));
  listMaintainSetCell_(row, cols.unitPrice, listMaintainNumeric_(line && line.unitPrice));
  return row;
}

/**
 * @param {{major: string, mid: string, lines: object[], overwrite: boolean}} payload
 */
function savePartsListPack(payload) {
  payload = payload || {};
  const major = normalize_(payload.major);
  const mid = normalize_(payload.mid);
  if (!major) {
    throw new Error('部品の大項目を入力してください。');
  }
  if (!mid) {
    throw new Error('部品の中項目を入力してください。');
  }
  const lines = (payload.lines || []).filter(listMaintainPartsLineHasContent_);
  if (!lines.length) {
    throw new Error('セット（部品名）か単価・数量を1行以上入力してください。');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.parts.sheetName);
  if (!sh) {
    throw new Error('シート「' + CONFIG.parts.sheetName + '」がありません。');
  }
  let overwritten = false;
  writeInternal_(function () {
    ensurePartsSetHeader_(sh);
    const lastCol = Math.max(sh.getLastColumn(), 1);
    const values = sh.getDataRange().getValues();
    const parsed = parsePartsSheetValues_(values);
    const cols = parsed.cols && Object.keys(parsed.cols).length
      ? parsed.cols
      : resolveColumns_(values[0] || [], CONFIG.parts.headers);
    if (!cols.major && !cols.mid) {
      throw new Error('部品リストに「大項目」「中項目」ヘッダーが見つかりません。');
    }
    const width = lastDataHeaderCol_(values[0] || []);
    const packRows = [];
    const keepMajorRows = [];
    parsed.rows.forEach(function (row) {
      if (normalize_(row.major) !== major) {
        return;
      }
      if (normalize_(row.mid) === mid) {
        packRows.push(row.sourceIndex);
      } else {
        keepMajorRows.push(row.sourceIndex);
      }
    });
    overwritten = packRows.length > 0;
    if (overwritten && !payload.overwrite) {
      throw new Error('部品「' + major + ' / ' + mid + '」は既にあります。上書きする場合は確認のうえ保存してください。');
    }
    const body = lines.map(function (line) {
      return listMaintainPartsRow_(cols, width, major, mid, line);
    });
    if (overwritten) {
      keepMajorRows.sort(function (a, b) {
        return a - b;
      });
      listMaintainPreserveLabel_(sh, values, cols.major, keepMajorRows, major);
      listMaintainDeleteRows_(sh, packRows);
    }
    listMaintainAppendRows_(sh, body, width);
  });
  return invoiceJsonSafe_({
    ok: true,
    overwritten: overwritten,
    major: major,
    mid: mid,
    master: listMaintainClientMaster_()
  });
}

function listMaintainWorkerSheet_(ss) {
  const sh = ss.getSheetByName(CONFIG.workers.sheetName);
  if (!sh) {
    throw new Error('シート「' + CONFIG.workers.sheetName + '」がありません。');
  }
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const lastRow = sh.getLastRow();
  const vals = lastRow ? sh.getRange(1, 1, lastRow, lastCol).getValues() : [['コード', '名前']];
  let start = 0;
  const h0 = normalize_(vals[0][0]);
  if (h0.indexOf('コード') !== -1 || h0.indexOf('作業者') !== -1 || h0.indexOf('担当') !== -1) {
    start = 1;
  }
  const cols = start === 1 ? resolveColumns_(vals[0], CONFIG.workers.headers) : {};
  return {
    sheet: sh,
    vals: vals,
    start: start,
    codeCol: cols.code || 1,
    nameCol: cols.name || 2,
    lastCol: lastCol
  };
}

/**
 * @param {{code: string, name: string, overwrite: boolean}} payload
 */
function saveWorkerListItem(payload) {
  payload = payload || {};
  const name = normalize_(payload.name);
  if (!name) {
    throw new Error('作業者の氏名を入力してください。');
  }
  let code = normalize_(payload.code);
  if (code) {
    const digits = String(code).match(/^(\d+)/);
    code = digits ? digits[1] : code;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let overwritten = false;
  let wroteCode = code;
  writeInternal_(function () {
    const ctx = listMaintainWorkerSheet_(ss);
    const sh = ctx.sheet;
    let matchRow = 0;
    let matchBy = '';
    for (let i = ctx.start; i < ctx.vals.length; i++) {
      const rowCode = normalize_(ctx.vals[i][ctx.codeCol - 1]);
      const rowName = normalize_(ctx.vals[i][ctx.nameCol - 1]);
      if (code && rowCode && rowCode === code) {
        matchRow = i + 1;
        matchBy = 'code';
        break;
      }
      if (!matchRow && rowName && rowName === name) {
        matchRow = i + 1;
        matchBy = 'name';
      }
    }
    overwritten = matchRow > 0;
    if (overwritten && !payload.overwrite) {
      throw new Error('作業者「' + (matchBy === 'code' ? code : name) +
        '」は既にあります。上書きする場合は確認のうえ保存してください。');
    }
    if (overwritten) {
      sh.getRange(matchRow, ctx.nameCol).setValue(name);
      if (code) {
        const n = Number(code);
        sh.getRange(matchRow, ctx.codeCol).setValue(isFinite(n) && String(n) === code ? n : code);
        wroteCode = code;
      } else {
        wroteCode = normalize_(ctx.vals[matchRow - 1][ctx.codeCol - 1]);
      }
    } else {
      const start = Math.max(sh.getLastRow(), ctx.start) + 1;
      const writeCode = code
        ? (isFinite(Number(code)) && String(Number(code)) === code ? Number(code) : code)
        : '';
      sh.getRange(start, ctx.codeCol).setValue(writeCode);
      sh.getRange(start, ctx.nameCol).setValue(name);
      if (!code) {
        assignMissingWorkerCodes_(sh);
        wroteCode = normalize_(sh.getRange(start, ctx.codeCol).getValue());
      } else {
        wroteCode = code;
      }
    }
  });
  const master = (function () {
    const workerSheet = loadWorkerSheet_();
    return {
      workers: workerSheet.rows,
      workerHeaders: workerSheet.headers,
      workerCodes: workerSheet.rows.map(function (w) {
        return w.name ? w.code + ' ' + w.name : w.code;
      })
    };
  })();
  return invoiceJsonSafe_({
    ok: true,
    overwritten: overwritten,
    code: wroteCode,
    name: name,
    master: master
  });
}
