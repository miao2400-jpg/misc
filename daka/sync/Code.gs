/**
 * 打卡小幫手：跨裝置同步（Google Apps Script）
 *
 * 用法：在你自己的 Google 試算表 → 擴充功能 → Apps Script，貼上整份程式，
 * 把下面的 TOKEN 換成你自己的同步密碼，再部署成「網頁應用程式」。
 * 執行身分：我；誰可以存取：所有人。部署後的網址（結尾 /exec）貼到打卡網頁的設定頁。
 *
 * 資料只存在這張試算表：
 *   「打卡紀錄」分頁：每筆打卡一列（最後一欄是完整資料，請勿手動修改）
 *   「設定」分頁：A1 更新時間、A2 設定內容（含表單網址與 Email，不要分享這張試算表）
 *
 * 忘記打卡提醒（給 iPhone 捷徑用）：
 *   網址/exec?token=同步密碼&check=in  → 今天打過上班卡了沒
 *   網址/exec?token=同步密碼&check=out → 今天打過下班卡了沒
 *   網址/exec?token=同步密碼&check=auto → 自己判斷現在該看上班還是下班（一個捷徑就夠，時間點想加幾個都可以）
 *   回傳一行文字；需要提醒時文字裡會有「還沒」。週末、設定裡的放假日、請假時段會回「不用提醒」。
 *   下班檢查會看今天的上班卡：上班時間加上一天的工時（設定裡的下班減上班，預設 9 小時）之前不提醒。
 *   更新這份程式後要重新部署：部署 → 管理部署作業 → 編輯 → 版本選「新版本」→ 部署（網址不會變）。
 */
const TOKEN = 'CHANGE_ME';

const LOG_SHEET = '打卡紀錄';
const CFG_SHEET = '設定';
const LOG_HEAD = ['id', '日期', '時間', '類型', '地點', '狀態', '裝置', '已刪除', '更新時間', 'json'];
const TYPE_LABEL = { in: '上班', out: '下班', trip: '上班時間公出' };
const TZ = 'Asia/Taipei';

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.check) return ContentService.createTextOutput(check_(p));
  return out_({ ok: true, msg: 'daka sync is running' });
}

function check_(p) {
  if (TOKEN === 'CHANGE_ME' || p.token !== TOKEN) return '同步密碼不對，請檢查捷徑裡的網址';
  const now = new Date();
  const today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd'), hm = Utilities.formatDate(now, TZ, 'HH:mm');
  if (Number(Utilities.formatDate(now, TZ, 'u')) >= 6) return '今天是週末，不用提醒';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cfg = null;
  const cs = ss.getSheetByName(CFG_SHEET);
  if (cs) { try { cfg = JSON.parse(cs.getRange('A2').getValue() || 'null'); } catch (err) { cfg = null; } }
  cfg = cfg || {};
  if ((Array.isArray(cfg.offDays) ? cfg.offDays : []).indexOf(today) >= 0) return '今天放假，不用提醒';
  const ws = (cfg.work && cfg.work.start) || '09:00', we = (cfg.work && cfg.work.end) || '18:00';
  const leaves = (Array.isArray(cfg.leaveDays) ? cfg.leaveDays : []).filter(function (l) { return l && l.date === today; });
  const todays = mergeLog_(sheet_(ss, LOG_SHEET, LOG_HEAD), []).filter(function (r) {
    return r.date === today && !r.deleted;
  });
  const has = function (type) { return todays.some(function (r) { return r.type === type; }); };
  const leaveCovers = function (type) {
    return leaves.some(function (l) { return type === 'in' ? l.start <= ws : l.end >= we; });
  };
  // 彈性上下班：今天有上班卡、又沒有請假時，做滿工時（含午休）才提醒下班卡
  const outDue = function () {
    const ins = todays.filter(function (r) { return r.type === 'in' && /^\d\d:\d\d$/.test(r.time || ''); })
      .map(function (r) { return r.time; }).sort();
    return (ins.length && !leaves.length) ? hm_(min_(ins[0]) + min_(we) - min_(ws)) : we;
  };
  const one = function (type) {
    if (leaveCovers(type)) return '今天這段有請假，不用提醒';
    if (has(type)) return '今天已打' + TYPE_LABEL[type] + '卡';
    if (type === 'out' && hm < outDue()) return '預計 ' + outDue() + ' 下班，現在不用提醒';
    return '還沒打' + TYPE_LABEL[type] + '卡！';
  };
  if (p.check !== 'auto') return one(p.check === 'out' ? 'out' : 'in');
  // check=auto：同一個網址自己判斷。上班時間前半小時起看上班卡，打了上班卡（或這段請假）就看下班卡
  if (hm < hm_(min_(ws) - 30)) return '上班時間前，不用提醒';
  if (!has('in') && !leaveCovers('in')) return one('in');
  return one('out');
}

function min_(hm) { const p = hm.split(':'); return Number(p[0]) * 60 + Number(p[1]); }
function hm_(m) { m = Math.min(m, 23 * 60 + 59); return ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + (m % 60)).slice(-2); }

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad json' }); }
  if (TOKEN === 'CHANGE_ME' || !req || req.token !== TOKEN) return out_({ ok: false, error: 'unauthorized' });
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logs = mergeLog_(sheet_(ss, LOG_SHEET, LOG_HEAD), Array.isArray(req.upsert) ? req.upsert : []);
    const config = mergeConfig_(sheet_(ss, CFG_SHEET, null), req.config);
    return out_({ ok: true, log: logs, config: config });
  } finally {
    lock.releaseLock();
  }
}

function mergeLog_(sh, incoming) {
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, LOG_HEAD.length).getValues() : [];
  const index = {};
  const all = [];
  rows.forEach(function (row, i) {
    let entry;
    try { entry = JSON.parse(row[LOG_HEAD.length - 1]); } catch (err) { return; }
    if (!entry || !entry.id) return;
    index[entry.id] = { rowNo: i + 2, entry: entry };
    all.push(entry);
  });
  incoming.forEach(function (entry) {
    if (!entry || typeof entry !== 'object' || !entry.id) return;
    const cur = index[entry.id];
    if (cur && (cur.entry.updatedAt || 0) >= (entry.updatedAt || 0)) return;
    const values = [toRow_(entry)];
    if (cur) {
      sh.getRange(cur.rowNo, 1, 1, LOG_HEAD.length).setValues(values);
      all[all.indexOf(cur.entry)] = entry;
      cur.entry = entry;
    } else {
      sh.appendRow(values[0]);
      index[entry.id] = { rowNo: sh.getLastRow(), entry: entry };
      all.push(entry);
    }
  });
  return all;
}

function toRow_(r) {
  return [
    String(r.id), String(r.date || ''), String(r.time || ''), TYPE_LABEL[r.type] || String(r.type || ''),
    String(r.place || ''), String(r.status || ''), String(r.device || ''), r.deleted ? '是' : '',
    r.updatedAt ? new Date(r.updatedAt) : '', JSON.stringify(r)
  ];
}

function mergeConfig_(sh, incoming) {
  let stored = null;
  try { stored = JSON.parse(sh.getRange('A2').getValue() || 'null'); } catch (err) { stored = null; }
  if (incoming && typeof incoming === 'object' && (incoming.updatedAt || 0) > ((stored && stored.updatedAt) || 0)) {
    sh.getRange('A1').setValue(new Date(incoming.updatedAt));
    sh.getRange('A2').setValue(JSON.stringify(incoming));
    stored = incoming;
  }
  return stored;
}

function sheet_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (head) { sh.appendRow(head); sh.setFrozenRows(1); }
  }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
