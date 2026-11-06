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
 *
 * 推播提醒（ntfy，取代捷徑）：
 *   第一次使用：上方函式選單選「setupReminders」→ 執行 → 照畫面授權。之後每 5 分鐘自動檢查一次。
 *   要停用：函式選單選「stopReminders」→ 執行。
 *   規則：上班時間到了還沒上班卡、應下班時間到了還沒下班卡，就推播；每 30 分鐘一次，提醒到 22:00。
 *   週末與放假日不提醒，除非那天有加班申請（照加班時間提醒）；平日有加班申請時，下班提醒延到加班結束。
 *   請假時段不提醒。通知上可按「1 小時後再提醒」或「今天不用再提醒」。
 *   推播主題和按鈕用的鑰匙在打卡 App 的設定裡產生，會同步到「設定」分頁；同步密碼不會出現在通知裡。
 */
const TOKEN = 'CHANGE_ME';

const LOG_SHEET = '打卡紀錄';
const CFG_SHEET = '設定';
const LOG_HEAD = ['id', '日期', '時間', '類型', '地點', '狀態', '裝置', '已刪除', '更新時間', 'json'];
const TYPE_LABEL = { in: '上班', out: '下班', trip: '上班時間公出' };
const TZ = 'Asia/Taipei';

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.act) return ContentService.createTextOutput(act_(p));
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

/* ───────── 推播提醒 ───────── */
const REMIND_EVERY_MIN = 30, REMIND_UNTIL = '22:00';

function setupReminders() {
  stopReminders();
  ScriptApp.newTrigger('remindTick').timeBased().everyMinutes(5).create();
  return '已啟用：每 5 分鐘檢查一次';
}
function stopReminders() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'remindTick') ScriptApp.deleteTrigger(t); });
}
function readCfg_(ss) {
  const cs = ss.getSheetByName(CFG_SHEET);
  try { return (cs && JSON.parse(cs.getRange('A2').getValue() || 'null')) || {}; } catch (err) { return {}; }
}
function remindState_(today) {
  const props = PropertiesService.getScriptProperties();
  let st = null;
  try { st = JSON.parse(props.getProperty('remind') || 'null'); } catch (err) { st = null; }
  if (!st || st.date !== today) st = { date: today, sent: {}, snoozeUntil: 0, muted: false };
  return { st: st, save: function () { props.setProperty('remind', JSON.stringify(st)); } };
}
function pushReady_(cfg) {
  const pu = cfg.push || {};
  return /^[A-Za-z0-9_-]{12,64}$/.test(pu.topic || '') && /^[A-Za-z0-9]{16,64}$/.test(pu.key || '') ? pu : null;
}

/* 今天該提醒什麼：回傳 { want: 'in'|'out'|null, due, why } */
function plan_(cfg, todays, now) {
  const today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd'), hm = Utilities.formatDate(now, TZ, 'HH:mm');
  const ws = (cfg.work && cfg.work.start) || '09:00', we = (cfg.work && cfg.work.end) || '18:00';
  const weekend = Number(Utilities.formatDate(now, TZ, 'u')) >= 6;
  const off = weekend || (Array.isArray(cfg.offDays) ? cfg.offDays : []).indexOf(today) >= 0;
  const ots = (Array.isArray(cfg.otDays) ? cfg.otDays : []).filter(function (o) { return o && o.date === today && o.start && o.end; });
  const otStart = ots.map(function (o) { return o.start; }).sort()[0];
  const otEnd = ots.map(function (o) { return o.end; }).sort().slice(-1)[0];
  const leaves = (Array.isArray(cfg.leaveDays) ? cfg.leaveDays : []).filter(function (l) { return l && l.date === today; });
  const has = function (t) { return todays.some(function (r) { return r.type === t; }); };
  if (off && !ots.length) return { want: null, why: weekend ? '週末' : '放假日' };
  const inDue = off ? otStart : ws;
  const inLeave = !off && leaves.some(function (l) { return l.start <= ws; });
  if (!has('in') && !inLeave) return { want: hm >= inDue ? 'in' : null, due: inDue, why: '上班' };
  if (has('out')) return { want: null, why: '已打下班卡' };
  if (!off && leaves.some(function (l) { return l.end >= we; })) return { want: null, why: '下午請假' };
  let due;
  if (off) due = otEnd;
  else {
    const ins = todays.filter(function (r) { return r.type === 'in' && /^\d\d:\d\d$/.test(r.time || ''); }).map(function (r) { return r.time; }).sort();
    due = (ins.length && !leaves.length) ? hm_(min_(ins[0]) + min_(we) - min_(ws)) : we;
    if (otEnd && otEnd > due) due = otEnd;
  }
  return { want: hm >= due ? 'out' : null, due: due, why: '下班', ot: !!otEnd };
}

function remindTick() {
  const now = new Date(), hm = Utilities.formatDate(now, TZ, 'HH:mm');
  if (hm > REMIND_UNTIL) return 'after hours';
  const ss = SpreadsheetApp.getActiveSpreadsheet(), cfg = readCfg_(ss), pu = pushReady_(cfg);
  if (!pu) return 'push not set';
  const today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
  const todays = mergeLog_(sheet_(ss, LOG_SHEET, LOG_HEAD), []).filter(function (r) { return r.date === today && !r.deleted; });
  const pl = plan_(cfg, todays, now);
  if (!pl.want) return 'quiet: ' + pl.why;
  const rs = remindState_(today), st = rs.st;
  if (st.muted) return 'muted';
  if (now.getTime() < (st.snoozeUntil || 0)) return 'snoozed';
  if (now.getTime() - (st.sent[pl.want] || 0) < (REMIND_EVERY_MIN - 2) * 60000) return 'sent recently';
  const label = TYPE_LABEL[pl.want];
  const msg = pl.want === 'in'
    ? '現在 ' + hm + '，今天還沒有上班卡。打開打卡 App 補打。'
    : (pl.ot ? '加班到 ' : '應下班 ') + pl.due + '，現在 ' + hm + '。打開打卡 App 打下班卡。';
  const code = push_(pu, '還沒打' + label + '卡', msg, cfg.sync && cfg.sync.url);
  if (code >= 200 && code < 300) { st.sent[pl.want] = now.getTime(); rs.save(); }
  return 'pushed ' + pl.want + ' ' + code;
}

function push_(pu, title, message, execUrl) {
  const body = { topic: pu.topic, title: title, message: message, priority: 4, tags: ['alarm_clock'] };
  if (execUrl && /^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(execUrl)) {
    const base = execUrl + '?key=' + encodeURIComponent(pu.key) + '&act=';
    body.actions = [
      { action: 'http', label: '1 小時後再提醒', url: base + 'snooze', method: 'GET', clear: true },
      { action: 'http', label: '今天不用再提醒', url: base + 'mute', method: 'GET', clear: true }
    ];
  }
  // ntfy.sh 對沒登入的發送者有每日上限；Google 的伺服器是很多人共用的，萬一常收到 429，
  // 到 ntfy.sh 註冊免費帳號產生存取權杖，放進「專案設定 → 指令碼屬性」，名稱 NTFY_TOKEN
  const tk = PropertiesService.getScriptProperties().getProperty('NTFY_TOKEN');
  const res = UrlFetchApp.fetch('https://ntfy.sh/', {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true,
    headers: tk ? { Authorization: 'Bearer ' + tk } : {}
  });
  return res.getResponseCode();
}

/* 通知上的按鈕：用 App 產生的提醒鑰匙，不是同步密碼 */
function act_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), pu = pushReady_(readCfg_(ss));
  if (!pu || p.key !== pu.key) return '鑰匙不對';
  const now = new Date(), rs = remindState_(Utilities.formatDate(now, TZ, 'yyyy-MM-dd'));
  if (p.act === 'snooze') { rs.st.snoozeUntil = now.getTime() + 60 * 60000; rs.save(); return '好，1 小時後再提醒'; }
  if (p.act === 'mute') { rs.st.muted = true; rs.save(); return '好，今天不再提醒'; }
  return '不認得這個動作';
}

function min_(hm) { const p = hm.split(':'); return Number(p[0]) * 60 + Number(p[1]); }
function hm_(m) { m = Math.min(m, 23 * 60 + 59); return ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + (m % 60)).slice(-2); }

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad json' }); }
  if (TOKEN === 'CHANGE_ME' || !req || req.token !== TOKEN) return out_({ ok: false, error: 'unauthorized' });
  if (req.op === 'testPush') {
    const pu = pushReady_({ push: req.push || {} });
    if (!pu) return out_({ ok: false, error: 'bad topic' });
    const code = push_(pu, '測試通知', '收到這則，推播提醒就設定好了。', req.execUrl);
    const triggers = ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === 'remindTick'; }).length;
    return out_({ ok: code >= 200 && code < 300, pushed: code, reminders: triggers > 0, error: code >= 300 ? 'ntfy ' + code : '' });
  }
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
