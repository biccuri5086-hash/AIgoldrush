/**
 * メール → スプレッドシート。予約・応募・注文などの通知メールを解析して一覧にする。
 * 痒い所: ①同じメールを二重に処理しない ②キャンセルメールで該当行を更新する
 *         ③項目が読み取れなかったメールは捨てずに「要確認」として残す
 */
const MAIL_CFG = {
  SHEET: '予約一覧',
  QUERY: 'newer_than:30d (subject:予約 OR subject:キャンセル)',   // ★対象メールの検索条件に変更
  PROCESSED_SHEET: '処理済みID',
  KEY_FIELD: '予約番号',
  FIELDS: [   // ★メール本文の書式に合わせて正規表現を調整する
    { name: '予約番号', re: /予約番号[:：]\s*([A-Za-z0-9-]+)/ },
    { name: '氏名', re: /お名前[:：]\s*(.+)/ },
    { name: '日時', re: /日時[:：]\s*(.+)/ },
  ],
  CANCEL_WORDS: ['キャンセル', '取消', '取り消し'],
};

/** メール1通を解析する(純粋関数) */
function parseMail_(subject, body, cfg) {
  const text = String(body == null ? '' : body).replace(/\r/g, '');
  const record = {}, missing = [];
  cfg.FIELDS.forEach(function (f) {
    const m = text.match(f.re);
    if (m) record[f.name] = m[1].trim(); else missing.push(f.name);
  });
  const subj = String(subject == null ? '' : subject);
  const isCancel = cfg.CANCEL_WORDS.some(function (w) { return subj.indexOf(w) >= 0; });
  return { record: record, missing: missing, type: isCancel ? 'cancel' : 'new' };
}

/** 解析結果を一覧(オブジェクト配列)へ反映する(純粋関数。元の配列は変更しない) */
function applyMailEvent_(table, parsed, keyField) {
  const out = table.map(function (o) { return Object.assign({}, o); });
  if (parsed.missing.indexOf(keyField) >= 0) {
    return { table: out, action: 'error', note: keyField + 'を読み取れませんでした' };
  }
  const key = parsed.record[keyField];
  const idx = out.findIndex(function (o) { return String(o[keyField]) === key; });
  if (parsed.type === 'cancel') {
    if (idx >= 0) { out[idx]['状態'] = 'キャンセル'; return { table: out, action: 'canceled' }; }
    const o = Object.assign({}, parsed.record, { '状態': 'キャンセル(元の予約なし)' });
    out.push(o); return { table: out, action: 'canceled_orphan' };
  }
  const note = parsed.missing.length ? '要確認: ' + parsed.missing.join(',') + 'が空' : '';
  if (idx >= 0) {
    Object.assign(out[idx], parsed.record, { '状態': '有効', '備考': note });
    return { table: out, action: 'updated' };
  }
  out.push(Object.assign({}, parsed.record, { '状態': '有効', '備考': note }));
  return { table: out, action: 'added' };
}

/** 未処理のメッセージIDだけを返す(純粋関数) */
function filterNewIds_(ids, processed) {
  const seen = {}; processed.forEach(function (id) { seen[id] = true; });
  return ids.filter(function (id) { if (seen[id]) return false; seen[id] = true; return true; });
}

/** Gmailを検索して一覧を更新する。トリガー(例: 1時間ごと)から呼ぶ。 */
function processMails() {
  return runSafely_('メール取り込み', function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(MAIL_CFG.SHEET) || ss.insertSheet(MAIL_CFG.SHEET);
    const pr = ss.getSheetByName(MAIL_CFG.PROCESSED_SHEET) || ss.insertSheet(MAIL_CFG.PROCESSED_SHEET);
    const processed = pr.getLastRow() ? pr.getRange(1, 1, pr.getLastRow(), 1).getValues().map(function (r) { return String(r[0]); }) : [];
    const cur = rowsToObjects_(sh.getDataRange().getValues());
    let table = cur.items;
    const headers = (cur.headers.length ? cur.headers : MAIL_CFG.FIELDS.map(function (f) { return f.name; }))
      .concat(['状態', '備考'].filter(function (h) { return cur.headers.indexOf(h) < 0; }));   // 状態・備考列が無ければ追加

    const counts = { added: 0, updated: 0, canceled: 0, canceled_orphan: 0, error: 0 };
    const msgs = [];
    GmailApp.search(MAIL_CFG.QUERY, 0, 50).forEach(function (th) {
      th.getMessages().forEach(function (m) { msgs.push(m); });
    });
    const fresh = filterNewIds_(msgs.map(function (m) { return m.getId(); }), processed);
    msgs.filter(function (m) { return fresh.indexOf(m.getId()) >= 0; })
      .sort(function (a, b) { return a.getDate() - b.getDate(); })   // 古い順に処理(予約→キャンセルの順を守る)
      .forEach(function (m) {
        const r = applyMailEvent_(table, parseMail_(m.getSubject(), m.getPlainBody(), MAIL_CFG), MAIL_CFG.KEY_FIELD);
        table = r.table; counts[r.action]++;
        if (r.action === 'error') logEvent_('メール取り込み', '警告', '読み取り不可: ' + m.getSubject());
        pr.appendRow([m.getId()]);
      });

    sh.clearContents();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (table.length) sh.getRange(2, 1, table.length, headers.length).setValues(objectsToRows_(table, headers));
    return counts;
  });
}
