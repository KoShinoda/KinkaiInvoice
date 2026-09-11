/**
 * 入力アプリの作業中タブ。Googleアカウントごとに保持する。
 * 確定の請求書保存・K-No検索は共用のまま。
 */

var WORK_JOB_INDEX_HEADERS_ = [
  '作業ID', '保存ID', 'K-No', '更新日時', 'ユーザー', '登録番号', '請求日',
  '入庫日', '出庫日', '整備部門', '整備種別', '受付', '値引技術%', '値引部品%',
  '明細件数', '明細開始行', 'テンプレート', 'アクティブ', '作成者'
];

function listWorkJobs() {
  const user = workJobUserKey_();
  if (!user) {
    return invoiceJsonSafe_({
      activeId: '',
      jobs: [],
      operator: '',
      isolated: false
    });
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const index = ensureWorkJobIndexSheet_();
  const last = index.getLastRow();
  if (last < 2) {
    return invoiceJsonSafe_({ activeId: '', jobs: [], operator: user, isolated: true });
  }
  const width = WORK_JOB_INDEX_HEADERS_.length;
  const vals = index.getRange(2, 1, last - 1, width).getValues();
  const disc = loadDiscPctDefaults_();
  const jobs = [];
  let activeId = '';
  for (let i = 0; i < vals.length; i++) {
    const owner = String(vals[i][18] || '').trim().toLowerCase();
    if (owner !== user) {
      continue;
    }
    const id = String(vals[i][0] || '').trim();
    if (!id) {
      continue;
    }
    const kNo = normalizeInvoiceKNo_(vals[i][2]);
    const items = loadWorkJobLines_(ss, id, Number(vals[i][14]) || 0, Number(vals[i][15]) || 0);
    jobs.push({
      id: id,
      saveId: String(vals[i][1] || '').trim(),
      boundKNo: kNo,
      templateName: invoicePlain_(vals[i][16]),
      payload: {
        header: {
          kNo: kNo,
          userName: invoicePlain_(vals[i][4]),
          plate: invoicePlain_(vals[i][5]),
          billDate: formatInvoiceYmd_(vals[i][6]),
          inDate: formatInvoiceYmd_(vals[i][7]),
          outDate: formatInvoiceYmd_(vals[i][8]),
          doneDate: formatInvoiceYmd_(vals[i][8]),
          dept: invoicePlain_(vals[i][9]),
          serviceType: invoicePlain_(vals[i][10]),
          receptionist: invoicePlain_(vals[i][11]),
          staff: invoicePlain_(vals[i][11])
        },
        items: items,
        summary: {
          techPct: vals[i][12] === '' || vals[i][12] == null ? disc.techPct : Number(vals[i][12]),
          partPct: vals[i][13] === '' || vals[i][13] == null ? disc.partPct : Number(vals[i][13])
        }
      }
    });
    if (String(vals[i][17] || '').trim()) {
      activeId = id;
    }
  }
  if (!activeId && jobs.length) {
    activeId = jobs[0].id;
  }
  return invoiceJsonSafe_({
    activeId: activeId,
    jobs: jobs,
    operator: user,
    isolated: true
  });
}

function saveWorkJobs(state) {
  const user = workJobUserKey_();
  if (!user) {
    return invoiceJsonSafe_({ ok: false, isolated: false, count: 0, operator: '' });
  }
  const incoming = (state && state.jobs) || [];
  const activeId = String((state && state.activeId) || '').trim();
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const keepJobs = [];
  for (let i = 0; i < incoming.length && keepJobs.length < 8; i++) {
    const job = incoming[i];
    if (!job || !String(job.id || '').trim()) {
      continue;
    }
    if (!workJobWorthSaving_(job)) {
      continue;
    }
    keepJobs.push(job);
  }
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    writeInternal_(function () {
      replaceWorkJobsForUser_(keepJobs, activeId, user, now);
    });
  } finally {
    lock.releaseLock();
  }
  return invoiceJsonSafe_({ ok: true, isolated: true, count: keepJobs.length, operator: user });
}

function workJobWorthSaving_(job) {
  const header = (job.payload && job.payload.header) || {};
  if (normalizeInvoiceKNo_(header.kNo || job.boundKNo)) {
    return true;
  }
  const items = (job.payload && job.payload.items) || [];
  for (let i = 0; i < items.length; i++) {
    if (rowHasContent_(items[i])) {
      return true;
    }
  }
  return !!(String(header.userName || '').trim() || String(header.plate || '').trim());
}

function replaceWorkJobsForUser_(jobs, activeId, user, now) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const index = ensureWorkJobIndexSheet_();
  const detail = ensureWorkJobDetailSheet_();
  const width = WORK_JOB_INDEX_HEADERS_.length;
  const last = index.getLastRow();
  const kept = [];
  const dropIds = {};
  if (last >= 2) {
    const vals = index.getRange(2, 1, last - 1, width).getValues();
    for (let i = 0; i < vals.length; i++) {
      const id = String(vals[i][0] || '').trim();
      const owner = String(vals[i][18] || '').trim().toLowerCase();
      if (owner === user) {
        if (id) {
          dropIds[id] = true;
        }
      } else {
        kept.push(vals[i]);
      }
    }
  }
  clearWorkJobLinesByIds_(detail, dropIds);
  const newRows = [];
  jobs.forEach(function (job) {
    const id = String(job.id || '').trim();
    const payload = job.payload || {};
    const header = payload.header || {};
    const summary = payload.summary || {};
    const items = (payload.items || []).filter(function (it) {
      return rowHasContent_(it);
    });
    const kNo = normalizeInvoiceKNo_(header.kNo || job.boundKNo);
    let start = '';
    if (items.length) {
      start = detail.getLastRow() + 1;
      const lines = items.map(function (it, n) {
        return invoiceSaveDetailRow_(id, n + 1, it);
      });
      detail.getRange(start, 1, lines.length, INVOICE_DETAIL_HEADERS_.length).setValues(lines);
    }
    newRows.push([
      id,
      String(job.saveId || '').trim(),
      kNo,
      now,
      header.userName || '',
      header.plate || '',
      header.billDate || '',
      header.inDate || '',
      header.outDate || '',
      header.dept || '',
      header.serviceType || '',
      header.receptionist || header.staff || '',
      summary.techPct == null ? '' : summary.techPct,
      summary.partPct == null ? '' : summary.partPct,
      items.length,
      items.length ? start : '',
      job.templateName || '',
      id === activeId ? '1' : '',
      user
    ]);
  });
  const all = kept.concat(newRows);
  if (last >= 2) {
    index.getRange(2, 1, last - 1, width).clearContent();
  }
  if (all.length) {
    index.getRange(2, 1, all.length, width).setValues(all);
  }
  index.getRange('C:C').setNumberFormat('@');
}

function loadWorkJobLines_(ss, jobId, lineCount, detailStart) {
  const sh = ss.getSheetByName((CONFIG.invoiceWork && CONFIG.invoiceWork.detailSheetName) || '請求書作業中明細');
  if (!sh) {
    return [];
  }
  return loadInvoiceSaveLines_(ss, {
    saveId: jobId,
    lineCount: lineCount,
    detailStart: detailStart,
    detailSheet: sh.getName()
  });
}

function clearWorkJobLinesByIds_(sh, idMap) {
  const last = sh.getLastRow();
  if (last < 2) {
    return;
  }
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  let i = 0;
  while (i < ids.length) {
    const id = String(ids[i][0] || '').trim();
    if (!idMap[id]) {
      i += 1;
      continue;
    }
    let j = i + 1;
    while (j < ids.length && idMap[String(ids[j][0] || '').trim()]) {
      j += 1;
    }
    sh.getRange(i + 2, 1, j - i, INVOICE_DETAIL_HEADERS_.length).clearContent();
    i = j;
  }
}

function ensureWorkJobIndexSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = (CONFIG.invoiceWork && CONFIG.invoiceWork.sheetName) || '請求書作業中';
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, WORK_JOB_INDEX_HEADERS_.length).setValues([WORK_JOB_INDEX_HEADERS_]);
    sh.getRange(1, 1, 1, WORK_JOB_INDEX_HEADERS_.length).setFontWeight('bold').setBackground('#e8f0ec');
    sh.setFrozenRows(1);
    sh.getRange('C:C').setNumberFormat('@');
    sh.hideSheet();
  }
  return sh;
}

function ensureWorkJobDetailSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = (CONFIG.invoiceWork && CONFIG.invoiceWork.detailSheetName) || '請求書作業中明細';
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, INVOICE_DETAIL_HEADERS_.length).setValues([INVOICE_DETAIL_HEADERS_]);
    sh.getRange(1, 1, 1, INVOICE_DETAIL_HEADERS_.length).setFontWeight('bold').setBackground('#e8f0ec');
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

function workJobUserKey_() {
  try {
    const active = Session.getActiveUser() && Session.getActiveUser().getEmail();
    if (active) {
      return String(active).trim().toLowerCase();
    }
    const effective = Session.getEffectiveUser() && Session.getEffectiveUser().getEmail();
    return String(effective || '').trim().toLowerCase();
  } catch (err) {
    return '';
  }
}

/** 同じ保存IDを他の人が作業中なら一覧を返す。確定保存の上書きロックとは別。 */
function listWorkJobConflicts(saveId) {
  const id = String(saveId || '').trim();
  if (!id) {
    return [];
  }
  const me = workJobUserKey_();
  const index = ensureWorkJobIndexSheet_();
  const last = index.getLastRow();
  if (last < 2) {
    return invoiceJsonSafe_([]);
  }
  const vals = index.getRange(2, 1, last - 1, WORK_JOB_INDEX_HEADERS_.length).getValues();
  const seen = {};
  const out = [];
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][1] || '').trim() !== id) {
      continue;
    }
    const owner = String(vals[i][18] || '').trim().toLowerCase();
    if (!owner || owner === me || seen[owner]) {
      continue;
    }
    seen[owner] = true;
    out.push({
      operator: owner,
      kNo: normalizeInvoiceKNo_(vals[i][2])
    });
  }
  return invoiceJsonSafe_(out);
}
