// ============================================================
// 工場進捗管理 GAS サーバー v5.1
// ============================================================

function doGet(e) {
  try {
    const action = e.parameter.action || '';
    if (action === 'ping') return ok({ message: 'pong' });
    if (action === 'getAll') return getAll_();
    if (action === 'getShipping') return getShipping_();
    return errRes('unknown action: ' + action);
  } catch (ex) {
    return errRes(ex.message);
  }
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    const action = p.action || '';
    if (action === 'saveItemMaster')      return saveItemMaster_(p);
    if (action === 'savePlansProgress')   return savePlansProgress_(p);
    if (action === 'saveShiftAttendance') return saveShiftAttendance_(p);
    if (action === 'saveOperationLog')    return saveOperationLog_(p);
    if (action === 'saveAll')             return saveAll_(p);
    if (action === 'saveShipping')        return saveShipping_(p);
    if (action === 'saveShippingWork')    return saveShippingWork_(p);
    if (action === 'deleteShipping')      return deleteShipping_(p);
    if (action === 'getSlipNo')           return getSlipNo_(p);
    if (action === 'saveStock')           return saveStock_(p);
    if (action === 'archiveOldMonths')    return archiveOldMonths_(p);
    if (action === 'archiveOperationLog') return archiveOperationLog_(p);
    return errRes('unknown action: ' + action);
  } catch (ex) {
    return errRes(ex.message);
  }
}

// ============================================================
// getAll
// ============================================================
function getAll_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ok({
    itemMaster:   readJsonSheet_(ss, 'itemMaster'),
    plans:        readPlansSheet_(ss),
    progress:     readProgressSheet_(ss),
    shiftAttend:  readShiftAttend_(ss),
    operationLog: readJsonSheet_(ss, 'operationLog'),
    stock:        readStockSheet_(ss)
  });
}

// ============================================================
// itemMaster
// ============================================================
function saveItemMaster_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  writeJsonSheet_(ss, 'itemMaster', p.itemMaster || []);
  return ok({});
}

// ============================================================
// plans + progress（v5.1: progressはマージ保存）
// ============================================================
function savePlansProgress_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return errRes('サーバー混雑中です。少し待って再試行してください（lock timeout）');
  }
  try {
    // plans は全件上書き（計画管理は1端末で行うため）
    writePlansSheet_(ss, p.plans || []);

    // progress はマージ保存（複数端末の同時入力を保護）
    const current = readProgressSheet_(ss);
    const incoming = p.progress || {};

    Object.keys(incoming).forEach(function(planId) {
      var val = incoming[planId];
      if (!current[planId]) {
        // 新規planId → そのまま追加
        current[planId] = {
          preDone:  val.preDone || 0,
          postDone: val.postDone || {}
        };
      } else {
        // 既存planId → preDoneは上書き、postDoneは端末名キー単位でマージ
        current[planId].preDone = val.preDone || 0;
        var incomingPostDone = val.postDone || {};
        Object.keys(incomingPostDone).forEach(function(deviceName) {
          current[planId].postDone[deviceName] = incomingPostDone[deviceName];
        });
      }
    });

    writeProgressSheet_(ss, current);
    return ok({});
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// shiftAttendance (v5.0: startTime/endTime対応)
// ============================================================
function saveShiftAttendance_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var shName = 'shiftAttendance';
  var sheet = ss.getSheetByName(shName);
  if (!sheet) sheet = ss.insertSheet(shName);
  sheet.clearContents();

  var rows = [
    ['baseShift',    JSON.stringify(p.baseShift    || [])],
    ['weeklyAttend', JSON.stringify(p.weeklyAttend || {})]
  ];
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  return ok({});
}

// ============================================================
// operationLog (v5.0新規)
// ============================================================
function saveOperationLog_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var shName = 'operationLog';
  var sheet = ss.getSheetByName(shName);
  if (!sheet) {
    sheet = ss.insertSheet(shName);
    sheet.getRange(1,1,1,8).setValues([['timestamp','deviceName','planId','itemName','color','proc','prevVal','newVal']]);
  }
  var log = p.log;
  if (!log) return ok({});
  sheet.appendRow([
    log.timestamp || new Date().toISOString(),
    log.deviceName || '',
    log.planId     || '',
    log.itemName   || '',
    log.color      || '',
    log.proc       || '',
    log.prevVal !== undefined ? log.prevVal : '',
    log.newVal  !== undefined ? log.newVal  : ''
  ]);
  return ok({});
}

// ============================================================
// saveAll
// ============================================================
function saveAll_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return errRes('サーバー混雑中です。少し待って再試行してください（lock timeout）');
  }
  try {
    if (p.itemMaster) writeJsonSheet_(ss, 'itemMaster', p.itemMaster);
    if (Array.isArray(p.plans)) writePlansSheet_(ss, p.plans); // 空配列も正しく上書き
    if (p.progress) {
      // saveAllもマージ保存
      const current = readProgressSheet_(ss);
      const incoming = p.progress;
      Object.keys(incoming).forEach(function(planId) {
        var val = incoming[planId];
        if (!current[planId]) {
          current[planId] = { preDone: val.preDone || 0, postDone: val.postDone || {} };
        } else {
          current[planId].preDone = val.preDone || 0;
          var incomingPostDone = val.postDone || {};
          Object.keys(incomingPostDone).forEach(function(deviceName) {
            current[planId].postDone[deviceName] = incomingPostDone[deviceName];
          });
        }
      });
      writeProgressSheet_(ss, current);
    }
    if (p.baseShift !== undefined || p.weeklyAttend !== undefined) {
      saveShiftAttendance_({ baseShift: p.baseShift || [], weeklyAttend: p.weeklyAttend || {} });
    }
    return ok({});
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// stock（ストックタブ：計画とは独立した品目別ストック数量）
// ============================================================
function saveStock_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 最大10秒待機
  } catch (e) {
    return errRes('サーバー混雑中です。少し待って再試行してください（lock timeout）');
  }
  try {
    writeStockSheet_(ss, p.stock || {});
    return ok({});
  } finally {
    lock.releaseLock();
  }
}

// stockシート: itemId,qty
function readStockSheet_(ss) {
  var sheet = ss.getSheetByName('stock');
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};
  var result = {};
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    result[String(r[0])] = Number(r[1] || 0);
  }
  return result;
}

function writeStockSheet_(ss, stock) {
  var sheet = ss.getSheetByName('stock');
  if (!sheet) sheet = ss.insertSheet('stock');
  sheet.clearContents();
  var rows = [['itemId','qty']];
  Object.keys(stock).forEach(function(itemId) {
    rows.push([itemId, stock[itemId]]);
  });
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
}

// ============================================================
// シート読み書きヘルパー
// ============================================================

// plansシート: planId,date(YYYY-MM-DD),itemId,qty,memo,priority
function readPlansSheet_(ss) {
  var sheet = ss.getSheetByName('plans');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    result.push({
      id:       String(r[0]),
      date:     toDateString_(r[1]),
      itemId:   String(r[2] || ''),
      qty:      Number(r[3] || 0),
      memo:     String(r[4] || ''),
      priority: r[5] === true || r[5] === 'TRUE' || r[5] === 1
    });
  }
  return result;
}

function writePlansSheet_(ss, plans) {
  var sheet = ss.getSheetByName('plans');
  if (!sheet) sheet = ss.insertSheet('plans');
  sheet.clearContents();
  var rows = [['id','date','itemId','qty','memo','priority']];
  for (var i = 0; i < plans.length; i++) {
    var p = plans[i];
    rows.push([p.id, p.date, p.itemId, p.qty, p.memo || '', p.priority ? 'TRUE' : 'FALSE']);
  }
  if (rows.length > 1) sheet.getRange(1, 1, rows.length, 6).setValues(rows);
  else sheet.getRange(1,1,1,6).setValues(rows);
}

// progressシート: planId,preDone,postDone(JSON)
function readProgressSheet_(ss) {
  var sheet = ss.getSheetByName('progress');
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};
  var result = {};
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    var postDone = {};
    try { postDone = JSON.parse(String(r[2] || '{}')); } catch(e) { postDone = {}; }
    result[String(r[0])] = {
      preDone:  Number(r[1] || 0),
      postDone: postDone
    };
  }
  return result;
}

function writeProgressSheet_(ss, progress) {
  var sheet = ss.getSheetByName('progress');
  if (!sheet) sheet = ss.insertSheet('progress');
  sheet.clearContents();
  var rows = [['planId','preDone','postDone']];
  var keys = Object.keys(progress);
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    var v = progress[k];
    // 旧キー（端末名のみ）と新キー（端末名_日付）が混在する場合、新キーを優先して旧キーを削除
    var cleanedPostDone = cleanPostDone_(v.postDone || {});
    rows.push([k, v.preDone || 0, JSON.stringify(cleanedPostDone)]);
  }
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
}

// postDoneの旧キー（端末名のみ）を新キー（端末名_日付）が存在する場合に削除
function cleanPostDone_(postDone) {
  var entries = Object.keys(postDone).map(function(k) { return [k, postDone[k]]; });
  // 新キーが存在する端末名を収集
  var hasNewKey = {};
  entries.forEach(function(entry) {
    var m = entry[0].match(/^(.+)_(\d{4}-\d{2}-\d{2})$/);
    if (m) hasNewKey[m[1]] = true;
  });
  // 旧キー（端末名のみ）は新キーが存在する場合に除外
  var cleaned = {};
  entries.forEach(function(entry) {
    var key = entry[0];
    var val = entry[1];
    var isOldKey = !key.match(/^(.+)_(\d{4}-\d{2}-\d{2})$/);
    if (isOldKey && hasNewKey[key]) return;
    cleaned[key] = val;
  });
  return cleaned;
}

// shiftAttendance (key-value形式)
function readShiftAttend_(ss) {
  var sheet = ss.getSheetByName('shiftAttendance');
  if (!sheet) return { baseShift: [], weeklyAttend: {} };
  var data = sheet.getDataRange().getValues();
  var kv = {};
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) kv[String(data[i][0])] = String(data[i][1] || '');
  }
  var baseShift = [];
  var weeklyAttend = {};
  try { baseShift = JSON.parse(kv['baseShift'] || '[]'); } catch(e) {}
  try { weeklyAttend = JSON.parse(kv['weeklyAttend'] || '{}'); } catch(e) {}
  return { baseShift: baseShift, weeklyAttend: weeklyAttend };
}

// JSON単列シート (itemMaster, operationLog)
function readJsonSheet_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    try { result.push(JSON.parse(String(data[i][0]))); } catch(e) {}
  }
  return result;
}

function writeJsonSheet_(ss, sheetName, arr) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clearContents();
  if (!arr || !arr.length) return;
  var rows = [['json']];
  for (var i = 0; i < arr.length; i++) rows.push([JSON.stringify(arr[i])]);
  sheet.getRange(1, 1, rows.length, 1).setValues(rows);
}

// 日付→YYYY-MM-DD文字列
function toDateString_(v) {
  if (!v) return '';
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var d = new Date(String(v));
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}

// レスポンスヘルパー
// ============================================================
// メンテナンス：月次アーカイブ（plans/progress）・古いoperationLog削除
// ============================================================

// 当月より前の月のplans/progressを月別アーカイブシートへ退避し、元シートから削除する
function archiveOldMonths_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return errRes('サーバー混雑中です。少し待って再試行してください（lock timeout）');
  }
  try {
    const plans = readPlansSheet_(ss);
    const progress = readProgressSheet_(ss);
    const currentMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');

    // 当月より前の月ごとにplansをグループ化
    const byMonth = {};
    plans.forEach(function (pl) {
      const month = (pl.date || '').slice(0, 7); // 'yyyy-MM'
      if (!month || month >= currentMonth) return; // 当月・日付不明は対象外
      if (!byMonth[month]) byMonth[month] = [];
      byMonth[month].push(pl);
    });

    const archivedMonths = Object.keys(byMonth).sort();
    if (archivedMonths.length === 0) {
      return ok({ message: 'アーカイブ対象の月はありませんでした（当月より前の計画データなし）', archivedMonths: [] });
    }

    archivedMonths.forEach(function (month) {
      const monthPlans = byMonth[month];
      const planIds = {};
      monthPlans.forEach(function (pl) { planIds[pl.id] = true; });

      appendPlansArchive_(ss, month, monthPlans);

      const monthProgress = {};
      Object.keys(progress).forEach(function (planId) {
        if (planIds[planId]) monthProgress[planId] = progress[planId];
      });
      appendProgressArchive_(ss, month, monthProgress);
    });

    // 元のplans/progressから、アーカイブ済みの月のデータを削除
    const remainingPlans = plans.filter(function (pl) {
      const month = (pl.date || '').slice(0, 7);
      return !month || month >= currentMonth;
    });
    writePlansSheet_(ss, remainingPlans);

    const remainingPlanIds = {};
    remainingPlans.forEach(function (pl) { remainingPlanIds[pl.id] = true; });
    const remainingProgress = {};
    Object.keys(progress).forEach(function (planId) {
      if (remainingPlanIds[planId]) remainingProgress[planId] = progress[planId];
    });
    writeProgressSheet_(ss, remainingProgress);

    return ok({ message: 'アーカイブ完了: ' + archivedMonths.join(', '), archivedMonths: archivedMonths });
  } finally {
    lock.releaseLock();
  }
}

// plansアーカイブシートへ追記（月ごとに1シート、既存があれば末尾に追加）
function appendPlansArchive_(ss, month, plansArr) {
  const shName = 'plans_archive_' + month;
  let sheet = ss.getSheetByName(shName);
  if (!sheet) {
    sheet = ss.insertSheet(shName);
    sheet.getRange(1, 1, 1, 6).setValues([['id', 'date', 'itemId', 'qty', 'memo', 'priority']]);
  }
  if (plansArr.length === 0) return;
  const rows = plansArr.map(function (pl) {
    return [pl.id, pl.date, pl.itemId, pl.qty, pl.memo || '', pl.priority ? 'TRUE' : 'FALSE'];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
}

// progressアーカイブシートへ追記（月ごとに1シート、既存があれば末尾に追加）
function appendProgressArchive_(ss, month, progressObj) {
  const shName = 'progress_archive_' + month;
  let sheet = ss.getSheetByName(shName);
  if (!sheet) {
    sheet = ss.insertSheet(shName);
    sheet.getRange(1, 1, 1, 3).setValues([['planId', 'preDone', 'postDone']]);
  }
  const keys = Object.keys(progressObj);
  if (keys.length === 0) return;
  const rows = keys.map(function (k) {
    return [k, progressObj[k].preDone || 0, JSON.stringify(progressObj[k].postDone || {})];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
}

// operationLogの古い行を、単一のアーカイブシートへ移動する（削除ではなく退避）
// ※ 月ごとにシートを分けず1枚に集約：行数は多くなるが、タブ数が際限なく増えるのを防ぐため
//    アーカイブシートは通常の読み込み処理(getAll_)の対象外なので、読み込み速度には影響しない
var OPERATION_LOG_KEEP_MONTHS = 3; // 保持期間（月数）。変更したい場合はここを編集する

function archiveOperationLog_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return errRes('サーバー混雑中です。少し待って再試行してください（lock timeout）');
  }
  try {
    const sheet = ss.getSheetByName('operationLog');
    if (!sheet) return ok({ message: 'operationLogシートがありません', archivedCount: 0 });

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return ok({ message: '対象がありませんでした', archivedCount: 0 });

    const header = data[0];
    const keepMonths = (p && p.keepMonths) || OPERATION_LOG_KEEP_MONTHS;
    const cutoff = new Date();
    cutoff.setMonth(cutoff.getMonth() - keepMonths);

    const kept = [header];
    const toArchive = [];
    for (let i = 1; i < data.length; i++) {
      const ts = data[i][0];
      const d = ts ? new Date(ts) : null;
      if (d && !isNaN(d.getTime()) && d < cutoff) {
        toArchive.push(data[i]);
      } else {
        kept.push(data[i]);
      }
    }

    if (toArchive.length === 0) {
      return ok({ message: 'アーカイブ対象はありませんでした', archivedCount: 0 });
    }

    // アーカイブ先シートへ追記（1枚のみ、月ごとに分けない）
    let archiveSheet = ss.getSheetByName('operationLog_archive');
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet('operationLog_archive');
      archiveSheet.getRange(1, 1, 1, header.length).setValues([header]);
    }
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, toArchive.length, header.length).setValues(toArchive);

    // 元のoperationLogから、アーカイブした行を除去
    sheet.clearContents();
    sheet.getRange(1, 1, kept.length, header.length).setValues(kept);

    return ok({ message: toArchive.length + '件をoperationLog_archiveへ退避しました', archivedCount: toArchive.length });
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// ok/errRes ヘルパー
// ============================================================
function ok(data) {
  data.ok = true;
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
function errRes(msg) {
  return ContentService.createTextOutput(JSON.stringify({ ok: false, error: msg })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 事務用（factory_shipping.html）
// ============================================================
function getShipping_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('deliveryHistory');
    if (!sheet) {
      sheet = ss.insertSheet('deliveryHistory');
      sheet.appendRow(['id','timestamp','deviceName','deliveryNo','memo','palletsJson','shippingDate']);
    }
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const records = rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        // 日付型はYYYY-MM-DD文字列に変換
        if (row[i] instanceof Date) {
          obj[h] = Utilities.formatDate(row[i], 'Asia/Tokyo', 'yyyy-MM-dd');
        } else {
          obj[h] = row[i];
        }
      });
      return obj;
    });
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, deliveryHistory: records }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function saveShipping_(p) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('deliveryHistory');
    if (!sheet) {
      sheet = ss.insertSheet('deliveryHistory');
      sheet.appendRow(['id','timestamp','deviceName','deliveryNo','memo','palletsJson','shippingDate']);
    }
    const records = p.records || [];
    records.forEach(r => {
      sheet.appendRow([
        r.id || '',
        r.timestamp || '',
        r.deviceName || '',
        r.deliveryNo || '',
        r.memo || '',
        r.palletsJson || '',
        r.shippingDate || ''
      ]);
    });
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


function saveShippingWork_(p) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('shippingWork');
    if (!sheet) {
      sheet = ss.insertSheet('shippingWork');
      sheet.appendRow(['json']);
    }
    if (sheet.getLastRow() < 2) sheet.appendRow(['']);
    sheet.getRange(2, 1).setValue(JSON.stringify(p.workData || {}));
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function deleteShipping_(p) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('deliveryHistory');
    if (!sheet) return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'シートが見つかりません' }))
      .setMimeType(ContentService.MimeType.JSON);

    const rows = sheet.getDataRange().getValues();
    for(let i = rows.length - 1; i >= 1; i--){
      if(String(rows[i][0]) === String(p.id)){
        sheet.deleteRow(i + 1);
        return ContentService
          .createTextOutput(JSON.stringify({ ok: true }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: '該当IDが見つかりません' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSlipNo_(p) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const date = String(p.date || '').slice(0, 10);
    const sheet = ss.getSheetByName('deliveryHistory');
    
    let maxNo = 0;
    if (sheet && sheet.getLastRow() >= 2) {
      const rows = sheet.getDataRange().getValues();
      const headers = rows[0];
      const dateIdx   = headers.indexOf('shippingDate');
      const noIdx     = headers.indexOf('deliveryNo');
      const jsonIdx   = headers.indexOf('palletsJson');
      
      rows.slice(1).forEach(row => {
        // 日付型・文字列型どちらにも対応
        let rowDate = '';
        const rawDate = row[dateIdx];
        if (rawDate instanceof Date) {
          rowDate = Utilities.formatDate(rawDate, 'Asia/Tokyo', 'yyyy-MM-dd');
        } else {
          rowDate = String(rawDate || '').slice(0, 10);
        }
        if (rowDate !== date) return;
        
        // deliveryNo列から取得
        const no = String(row[noIdx] || '');
        const m1 = no.match(/(\d+)$/);
        if (m1) { const n = parseInt(m1[1]); if (n > maxNo) maxNo = n; }
        
        // palletsJsonのslipsから全slipNoを取得
        try {
          const pj = JSON.parse(String(row[jsonIdx] || '{}'));
          (pj.slips || []).forEach(s => {
            const m2 = String(s.slipNo || '').match(/(\d+)$/);
            if (m2) { const n = parseInt(m2[1]); if (n > maxNo) maxNo = n; }
          });
        } catch(e2) {}
      });
    }
    
    const startNo = maxNo + 1;
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, startNo: startNo }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message, startNo: 1 }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}