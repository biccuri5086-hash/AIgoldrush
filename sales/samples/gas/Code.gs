/**
 * サンプル: フォーム回答の転記・月次集計・メール通知(Google Apps Script)
 * 使い方は README.md を参照。サンプルデータはすべてダミーです。
 */
const CONFIG = {
  SOURCE_SHEET: 'フォーム回答',   // フォームの回答先シート名
  SUMMARY_SHEET: '月次集計',      // 集計結果を書き出すシート名
  NOTIFY_TO: 'example@example.com', // ★通知先メールアドレスに変更する
};

/** 集計(純粋関数)。rows: [[日時(Date), 氏名, 商品, 数量, 単価], ...]。ヘッダー行は含めない。
 *  不正な行(日時・数量・単価が不正、数量が整数でない)は除外し、件数を返す。 */
function aggregateOrders_(rows) {
  const totals = {};   // "yyyy-MM|商品" -> {qty, amount}
  let skipped = 0;
  rows.forEach(function (r) {
    const d = r[0], item = String(r[2] || '').trim();
    const qty = Number(r[3]), price = Number(r[4]);
    const ok = d instanceof Date && !isNaN(d.getTime()) && item !== '' &&
      Number.isInteger(qty) && qty > 0 && Number.isInteger(price) && price >= 0;
    if (!ok) { skipped++; return; }
    const key = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '|' + item;
    if (!totals[key]) totals[key] = { qty: 0, amount: 0 };
    totals[key].qty += qty;
    totals[key].amount += qty * price;   // 整数どうしの計算(小数の誤差なし)
  });
  const list = Object.keys(totals).sort().map(function (k) {
    const p = k.split('|');
    return { month: p[0], item: p[1], qty: totals[k].qty, amount: totals[k].amount };
  });
  return { list: list, skipped: skipped };
}

/** フォーム回答シートを集計して「月次集計」シートに書き出す。 */
function updateSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  if (!src) throw new Error('シートが見つかりません: ' + CONFIG.SOURCE_SHEET);
  const last = src.getLastRow();
  const rows = last >= 2 ? src.getRange(2, 1, last - 1, 5).getValues() : [];
  const res = aggregateOrders_(rows);

  const out = ss.getSheetByName(CONFIG.SUMMARY_SHEET) || ss.insertSheet(CONFIG.SUMMARY_SHEET);
  out.clearContents();
  const table = [['年月', '商品', '数量', '金額(円)']].concat(
    res.list.map(function (x) { return [x.month, x.item, x.qty, x.amount]; }));
  out.getRange(1, 1, table.length, 4).setValues(table);
  out.getRange(1, 6).setValue('除外した行数');
  out.getRange(1, 7).setValue(res.skipped);
  return res;
}

/** 集計結果をメールで送る(月次のトリガーから呼ぶ想定)。 */
function sendSummaryMail() {
  const res = updateSummary();
  const total = res.list.reduce(function (s, x) { return s + x.amount; }, 0);
  const lines = res.list.map(function (x) {
    return x.month + ' ' + x.item + ' ' + x.qty + '個 ' + x.amount + '円';
  });
  const body = lines.join('\n') + '\n\n合計: ' + total + '円\n除外した行: ' + res.skipped + '行';
  MailApp.sendEmail(CONFIG.NOTIFY_TO, '月次集計のお知らせ', body);
}

/** フォーム送信時に集計を更新する(トリガー: フォーム送信時)。 */
function onFormSubmit() { updateSummary(); }

/** 初回に1度だけ実行: 送信時トリガーと月次メールのトリガーを設定する。 */
function setupTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('sendSummaryMail').timeBased().onMonthDay(1).atHour(9).create();
}
