/**
 * CSV取り込み → 予定との突き合わせ(出席/欠席/予定外)。
 * 痒い所: ①CSVの列名が変わっても、別名(エイリアス)で見つけて動く
 *         ②見つからないときは「どの列が足りないか・何があったか」を日本語で知らせる
 *         ③BOM・引用符・改行入りのセルを正しく読む
 */

/** CSVを読む(純粋関数)。引用符("..."、"" のエスケープ、セル内改行)・BOM・CRLFに対応。 */
function parseCsv_(text) {
  let s = String(text == null ? '' : text);
  if (s.charCodeAt(0) === 0xFEFF) s = s.slice(1);
  const rows = []; let row = [], cell = '', q = false;
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (q) {
      if (c === '"') { if (s[i + 1] === '"') { cell += '"'; i++; } else q = false; }
      else cell += c;
    } else if (c === '"') q = true;
    else if (c === ',') { row.push(cell); cell = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && s[i + 1] === '\n') i++;
      row.push(cell); cell = ''; rows.push(row); row = [];
    } else cell += c;
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
  return rows.filter(function (r) { return !(r.length === 1 && r[0] === ''); });
}

/** 見出し行から、必要な列の位置を別名で探す(純粋関数)。見つからない列は missing に入れる。 */
function mapHeaders_(headers, aliases) {
  const norm = headers.map(normalizeKey_);
  const index = {}, missing = [];
  Object.keys(aliases).forEach(function (key) {
    const found = aliases[key].map(normalizeKey_).map(function (a) { return norm.indexOf(a); })
      .filter(function (i) { return i >= 0; });
    if (found.length) index[key] = found[0]; else missing.push(key);
  });
  return { index: index, missing: missing };
}

/** 予定(planned)と実績(actual)を突き合わせる(純粋関数)。
 *  どちらも [{name, date}]。date は同じ形式の文字列。戻り値: { attended, absent, unplanned } */
function reconcile_(planned, actual) {
  const key = function (x) { return normalizeKey_(x.name) + '|' + normalizeKey_(x.date); };
  const act = {}; actual.forEach(function (a) { act[key(a)] = a; });
  const plan = {}; planned.forEach(function (p) { plan[key(p)] = p; });
  const attended = [], absent = [], unplanned = [];
  planned.forEach(function (p) { (act[key(p)] ? attended : absent).push(p); });
  actual.forEach(function (a) { if (!plan[key(a)]) unplanned.push(a); });
  return { attended: attended, absent: absent, unplanned: unplanned };
}

/** CSVの1行目から必要な列を探し、なければ分かりやすいエラーを出す。 */
function csvToRecords_(text, aliases) {
  const rows = parseCsv_(text);
  if (!rows.length) throw new Error('CSVが空です');
  const m = mapHeaders_(rows[0], aliases);
  if (m.missing.length) {
    throw new Error('CSVに次の列が見つかりません: ' + m.missing.join('、') +
      '。CSVにある列: ' + rows[0].join('、') + '(列名が変わった場合は、設定の別名に追加してください)');
  }
  return rows.slice(1).map(function (r) {
    const o = {}; Object.keys(m.index).forEach(function (k) { o[k] = (r[m.index[k]] || '').trim(); }); return o;
  });
}
