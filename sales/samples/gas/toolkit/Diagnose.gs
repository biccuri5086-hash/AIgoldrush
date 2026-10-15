/**
 * 診断: 「ちゃんと動いているか」を、専門知識がなくても確認できるようにする。
 * 実行すると「診断」シートに 要対応 / 注意 / OK を書き出す。
 */
const DIAG_CFG = {
  REQUIRED_SHEETS: ['フォーム回答', '実行ログ'],       // ★必要なシート名に変更
  REQUIRED_HANDLERS: ['routeFormResponses'],           // ★必要なトリガーの関数名に変更
  MAX_IDLE_HOURS: 48,                                  // この時間、成功ログがなければ注意
  MIN_MAIL_QUOTA: 20,                                  // 残りメール送信数がこれ未満なら注意
};

/** 診断結果を作る(純粋関数)。info の内容は diagnose() を参照。 */
function evaluateDiagnosis_(info, cfg) {
  const out = [];
  const add = function (level, item, msg) { out.push({ level: level, item: item, message: msg }); };
  cfg.REQUIRED_SHEETS.forEach(function (s) {
    if (info.sheets.indexOf(s) < 0) add('要対応', 'シート', '「' + s + '」シートがありません。名前を変えたり削除したりしていませんか。');
  });
  cfg.REQUIRED_HANDLERS.forEach(function (h) {
    if (info.handlers.indexOf(h) < 0) add('要対応', 'トリガー', '「' + h + '」の自動実行(トリガー)が設定されていません。設定し直してください。');
  });
  if (info.lastSuccessMs == null) add('注意', '実行ログ', '成功の記録がまだありません。初回実行前か、処理が一度も成功していない可能性があります。');
  else {
    const hours = (info.nowMs - info.lastSuccessMs) / 3600000;
    if (hours > cfg.MAX_IDLE_HOURS) add('注意', '実行ログ', '最後の成功から約' + Math.floor(hours) + '時間たっています。自動実行が止まっていないか確認してください。');
  }
  if (info.failures24h > 0) add('要対応', '実行ログ', '直近24時間に' + info.failures24h + '件の失敗があります。「実行ログ」シートの内容を確認してください。');
  if (info.mailQuota != null && info.mailQuota < cfg.MIN_MAIL_QUOTA) add('注意', 'メール送信数', '本日の残りメール送信数が' + info.mailQuota + '通です。送信数の上限に近づいています。');
  if (!out.length) add('OK', '全体', '問題は見つかりませんでした。');
  return out;
}

function diagnose() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map(function (s) { return s.getName(); });
  const handlers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  let lastSuccessMs = null, failures24h = 0;
  const log = ss.getSheetByName(TK.LOG_SHEET);
  const now = Date.now();
  if (log && log.getLastRow() > 1) {
    log.getRange(2, 1, log.getLastRow() - 1, 3).getValues().forEach(function (r) {
      const t = r[0] instanceof Date ? r[0].getTime() : NaN;
      if (isNaN(t)) return;
      if (r[2] === '成功' && (lastSuccessMs == null || t > lastSuccessMs)) lastSuccessMs = t;
      if (r[2] === '失敗' && now - t <= 24 * 3600000) failures24h++;
    });
  }
  let mailQuota = null;
  try { mailQuota = MailApp.getRemainingDailyQuota(); } catch (_) { /* 取得できなければ判定しない */ }
  const res = evaluateDiagnosis_({ sheets: sheets, handlers: handlers, lastSuccessMs: lastSuccessMs,
    nowMs: now, failures24h: failures24h, mailQuota: mailQuota }, DIAG_CFG);
  const out = ss.getSheetByName('診断') || ss.insertSheet('診断');
  out.clearContents();
  const table = [['判定', '項目', '内容']].concat(res.map(function (x) { return [x.level, x.item, x.message]; }));
  out.getRange(1, 1, table.length, 3).setValues(table);
  return res;
}
