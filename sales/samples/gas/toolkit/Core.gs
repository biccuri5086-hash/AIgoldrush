/**
 * ツールキット共通部分: ログ・失敗通知・二重実行防止。
 * 目的: 「動いているか分からない」「壊れても気づけない」を無くす。
 * 納品先ごとに TK を書き換える。サンプルのメールアドレスは実在しません。
 */
const TK = {
  LOG_SHEET: '実行ログ',
  NOTIFY_TO: 'example@example.com',   // ★失敗通知の送り先に変更する
  MAX_LOG_ROWS: 1000,                 // ログが増えすぎないよう古いものから削除
  NOTIFY_MIN_GAP_MIN: 60,             // 同じ処理の失敗通知は、この分数に1回まで
};

/** ログ1行を作る(純粋関数) */
function formatLogRow_(now, name, level, message) {
  return [now, name, level, String(message == null ? '' : message).slice(0, 500)];
}

/** 失敗通知を送ってよいか(同じ処理の通知が短時間に連続しないようにする) */
function shouldNotify_(lastMs, nowMs, minGapMin) {
  if (lastMs == null || isNaN(lastMs)) return true;
  return (nowMs - lastMs) >= minGapMin * 60 * 1000;
}

function logEvent_(name, level, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TK.LOG_SHEET) || ss.insertSheet(TK.LOG_SHEET);
  if (sh.getLastRow() === 0) sh.appendRow(['日時', '処理', '種別', '内容']);
  sh.appendRow(formatLogRow_(new Date(), name, level, message));
  const extra = sh.getLastRow() - 1 - TK.MAX_LOG_ROWS;
  if (extra > 0) sh.deleteRows(2, extra);
}

function notifyFailure_(name, message) {
  const props = PropertiesService.getScriptProperties();
  const key = 'lastNotify_' + name;
  const last = Number(props.getProperty(key));
  if (!shouldNotify_(last || null, Date.now(), TK.NOTIFY_MIN_GAP_MIN)) return;
  MailApp.sendEmail(TK.NOTIFY_TO, '【自動処理が失敗しました】' + name,
    '処理名: ' + name + '\n内容: ' + message + '\n\n「' + TK.LOG_SHEET + '」シートに詳細があります。');
  props.setProperty(key, String(Date.now()));
}

/** 処理を安全に実行する。二重実行を避け、結果をログに残し、失敗時は通知する。 */
function runSafely_(name, fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { logEvent_(name, '警告', '他の処理が実行中のためスキップ'); return null; }
  try {
    const r = fn();
    logEvent_(name, '成功', r === undefined ? '' : JSON.stringify(r));
    return r;
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    logEvent_(name, '失敗', msg);
    try { notifyFailure_(name, msg); } catch (_) { /* 通知自体の失敗は握りつぶす(ログには残っている) */ }
    return null;
  } finally {
    lock.releaseLock();
  }
}

/** シートの値(2次元配列、1行目=見出し)をオブジェクトの配列にする(純粋関数) */
function rowsToObjects_(values) {
  if (!values || values.length < 1) return { headers: [], items: [] };
  const headers = values[0].map(function (h) { return String(h).trim(); });
  const items = values.slice(1).filter(function (r) {
    return r.some(function (c) { return c !== '' && c != null; });
  }).map(function (r) {
    const o = {}; headers.forEach(function (h, i) { o[h] = r[i] === undefined ? '' : r[i]; }); return o;
  });
  return { headers: headers, items: items };
}

/** オブジェクトの配列を、指定した見出し順の2次元配列にする(純粋関数) */
function objectsToRows_(items, headers) {
  return items.map(function (o) {
    return headers.map(function (h) { return o[h] === undefined ? '' : o[h]; });
  });
}
