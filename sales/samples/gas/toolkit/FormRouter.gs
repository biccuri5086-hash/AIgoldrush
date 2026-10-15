/**
 * Googleフォーム → 日程別シートへの振り分け/アンケートとの照合。
 * 痒い所: ①「6/30」「6月30日(月)」「2026/6/30」など表記ゆれを同じ日として扱う
 *         ②複数日程を選んだ回答を各日のシートへ入れる ③読み取れない日付は捨てず「要確認」へ
 *         ④全角/半角・空白・大文字小文字の違いでも照合できる ⑤照合できなかった回答を一覧にする
 */

/** 照合用の正規化(全角半角・空白・大文字小文字・前後の空白を吸収する。純粋関数) */
function normalizeKey_(s) {
  return String(s == null ? '' : s).normalize('NFKC').replace(/[\s　]+/g, '').toLowerCase();
}

/** 日付の表記ゆれを「M月D日」のシート名にそろえる。読み取れなければ null(純粋関数) */
function dateSheetName_(s) {
  const t = String(s == null ? '' : s).normalize('NFKC');
  const m = t.match(/(?:\d{4}\s*[\/年.-]\s*)?(\d{1,2})\s*[\/月.-]\s*(\d{1,2})\s*日?/);
  if (!m) return null;
  const mo = Number(m[1]), d = Number(m[2]);
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  return mo + '月' + d + '日';
}

/** 「6/30, 7/1」のような複数選択を日付ごとに分ける(純粋関数) */
function splitDates_(s) {
  const t = String(s == null ? '' : s).normalize('NFKC');
  const parts = t.split(/[,、\n]/).map(function (x) { return x.trim(); }).filter(Boolean);
  const names = [], unknown = [];
  parts.forEach(function (p) {
    const n = dateSheetName_(p);
    if (n) { if (names.indexOf(n) < 0) names.push(n); } else unknown.push(p);
  });
  return { names: names, unknown: unknown };
}

/** 回答行を日程別に振り分ける。rows: [[...]]、dateIndex: 日程の列番号(0始まり)。純粋関数。
 *  戻り値: { sheets: {シート名: [行,...]}, review: [要確認の行,...] } */
function routeByDate_(rows, dateIndex) {
  const sheets = {}, review = [];
  rows.forEach(function (r) {
    const sp = splitDates_(r[dateIndex]);
    if (!sp.names.length) { review.push(r); return; }
    sp.names.forEach(function (n) { (sheets[n] = sheets[n] || []).push(r); });
    if (sp.unknown.length) review.push(r);   // 一部だけ読み取れなかった場合も人が確認できるようにする
  });
  return { sheets: sheets, review: review };
}

/** アンケート回答を参加者一覧と照合して「済」を付ける(純粋関数。元の配列は変更しない)。
 *  participants/surveys: [{email, name}]。メールアドレス優先、なければ氏名で照合。
 *  戻り値: { participants: [... , 済], unmatched: [照合できなかったアンケート] } */
function markSurveyDone_(participants, surveys) {
  const out = participants.map(function (p) { return Object.assign({}, p); });
  const unmatched = [];
  surveys.forEach(function (sv) {
    const em = normalizeKey_(sv.email), nm = normalizeKey_(sv.name);
    let idx = em ? out.findIndex(function (p) { return normalizeKey_(p.email) === em; }) : -1;
    if (idx < 0 && nm) idx = out.findIndex(function (p) { return normalizeKey_(p.name) === nm; });
    if (idx >= 0) out[idx]['済'] = '済'; else unmatched.push(sv);
  });
  return { participants: out, unmatched: unmatched };
}

/** フォーム送信時トリガー用: 回答を日程別シートへ転記する。 */
const FORM_CFG = {
  SOURCE_SHEET: 'フォーム回答',
  DATE_COLUMN: 3,              // ★「参加希望日」の列(1始まり)
  REVIEW_SHEET: '要確認',
};
function routeFormResponses() {
  return runSafely_('フォーム振り分け', function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const src = ss.getSheetByName(FORM_CFG.SOURCE_SHEET);
    if (!src) throw new Error('シートが見つかりません: ' + FORM_CFG.SOURCE_SHEET);
    const values = src.getDataRange().getValues();
    const header = values[0], rows = values.slice(1);
    const res = routeByDate_(rows, FORM_CFG.DATE_COLUMN - 1);
    let n = 0;
    Object.keys(res.sheets).forEach(function (name) {
      const sh = ss.getSheetByName(name) || ss.insertSheet(name);
      sh.clearContents();                       // 毎回作り直す(重複転記を防ぐ)
      sh.getRange(1, 1, 1, header.length).setValues([header]);
      sh.getRange(2, 1, res.sheets[name].length, header.length).setValues(res.sheets[name]);
      n += res.sheets[name].length;
    });
    const rv = ss.getSheetByName(FORM_CFG.REVIEW_SHEET) || ss.insertSheet(FORM_CFG.REVIEW_SHEET);
    rv.clearContents();
    rv.getRange(1, 1, 1, header.length).setValues([header]);
    if (res.review.length) rv.getRange(2, 1, res.review.length, header.length).setValues(res.review);
    return { 転記: n, 要確認: res.review.length, 日程数: Object.keys(res.sheets).length };
  });
}
