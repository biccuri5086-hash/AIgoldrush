// toolkit/*.gs の純粋関数をNodeで検証する。GAS固有のAPI(SpreadsheetApp等)を使う関数は対象外。
const fs = require('fs'), vm = require('vm'), path = require('path'), assert = require('assert');
const dir = path.join(__dirname, '..', 'gas', 'toolkit');
const src = ['Core', 'MailToSheet', 'FormRouter', 'CsvReconcile', 'Diagnose']
  .map(n => fs.readFileSync(path.join(dir, n + '.gs'), 'utf8')).join('\n');
const names = ['formatLogRow_', 'shouldNotify_', 'rowsToObjects_', 'objectsToRows_', 'parseMail_', 'applyMailEvent_',
  'filterNewIds_', 'normalizeKey_', 'dateSheetName_', 'splitDates_', 'routeByDate_', 'markSurveyDone_', 'parseCsv_',
  'mapHeaders_', 'reconcile_', 'csvToRecords_', 'evaluateDiagnosis_', 'MAIL_CFG', 'DIAG_CFG', 'TK'];
const ctx = vm.createContext({});
vm.runInContext(src + '\n' + names.map(n => `this.${n}=${n};`).join(''), ctx);
// vmの別コンテキストのオブジェクトを、通常のオブジェクトと比較できる形にする
const J = x => JSON.parse(JSON.stringify(x));
const F = new Proxy({}, { get: (_, k) => ctx[k] });
let n = 0;
const t = (name, fn) => { try { fn(); n++; } catch (e) { console.log('FAIL:', name, '\n ', e.message); process.exitCode = 1; } };
const eq = (a, b) => assert.deepStrictEqual(J(a), J(b));

t('formatLogRow_ 500字で切る/null', () => {
  eq(F.formatLogRow_('t', 'n', '成功', 'x'.repeat(600))[3].length, 500);
  eq(F.formatLogRow_('t', 'n', '成功', null)[3], '');
});
t('shouldNotify_', () => {
  assert.strictEqual(F.shouldNotify_(null, 1000, 60), true);
  assert.strictEqual(F.shouldNotify_(NaN, 1000, 60), true);
  assert.strictEqual(F.shouldNotify_(0, 59 * 60000, 60), false);
  assert.strictEqual(F.shouldNotify_(0, 60 * 60000, 60), true);
});
t('rowsToObjects_/objectsToRows_', () => {
  const r = F.rowsToObjects_([[' 氏名 ', '日時'], ['A', '1'], ['', ''], ['B']]);
  eq(r.headers, ['氏名', '日時']); eq(r.items, [{ 氏名: 'A', 日時: '1' }, { 氏名: 'B', 日時: '' }]);
  eq(F.objectsToRows_(r.items, ['日時', '氏名', '備考']), [['1', 'A', ''], ['', 'B', '']]);
  eq(F.rowsToObjects_([]), { headers: [], items: [] });
});
const mailBody = '予約番号: AB-001\r\nお名前: 山田 太郎\r\n日時: 6/30 19:00\r\n';
t('parseMail_ 新規/キャンセル/欠落', () => {
  const a = F.parseMail_('【予約】ご予約ありがとう', mailBody, F.MAIL_CFG);
  eq(a.record, { 予約番号: 'AB-001', 氏名: '山田 太郎', 日時: '6/30 19:00' }); eq(a.missing, []); assert.strictEqual(a.type, 'new');
  assert.strictEqual(F.parseMail_('予約キャンセルのお知らせ', mailBody, F.MAIL_CFG).type, 'cancel');
  eq(F.parseMail_('予約', '予約番号: X1', F.MAIL_CFG).missing, ['氏名', '日時']);
  eq(F.parseMail_(null, null, F.MAIL_CFG).missing, ['予約番号', '氏名', '日時']);
});
t('applyMailEvent_ 追加→更新→キャンセル→孤立キャンセル→エラー、入力は不変', () => {
  const K = '予約番号';
  const p1 = F.parseMail_('予約', mailBody, F.MAIL_CFG);
  const original = [];
  let r = F.applyMailEvent_(original, p1, K);
  assert.strictEqual(r.action, 'added'); assert.strictEqual(original.length, 0); assert.strictEqual(r.table[0]['状態'], '有効');
  const p2 = F.parseMail_('予約変更', mailBody.replace('19:00', '20:00'), F.MAIL_CFG);
  r = F.applyMailEvent_(r.table, p2, K); assert.strictEqual(r.action, 'updated'); assert.strictEqual(r.table.length, 1); assert.strictEqual(r.table[0]['日時'], '6/30 20:00');
  const keep = J(r.table);
  r = F.applyMailEvent_(r.table, F.parseMail_('キャンセル', mailBody, F.MAIL_CFG), K);
  assert.strictEqual(r.action, 'canceled'); assert.strictEqual(r.table[0]['状態'], 'キャンセル'); eq(keep[0]['状態'], '有効');
  r = F.applyMailEvent_(r.table, F.parseMail_('キャンセル', mailBody.replace('AB-001', 'ZZ-9'), F.MAIL_CFG), K);
  assert.strictEqual(r.action, 'canceled_orphan'); assert.strictEqual(r.table.length, 2);
  r = F.applyMailEvent_(r.table, F.parseMail_('予約', 'お名前: だれか', F.MAIL_CFG), K);
  assert.strictEqual(r.action, 'error'); assert.strictEqual(r.table.length, 2);
  const miss = F.applyMailEvent_([], F.parseMail_('予約', '予約番号: Q1', F.MAIL_CFG), K);
  assert.ok(miss.table[0]['備考'].includes('要確認'));
});
t('filterNewIds_', () => { eq(F.filterNewIds_(['a', 'b', 'c', 'b'], ['a']), ['b', 'c']); eq(F.filterNewIds_([], []), []); });
t('normalizeKey_', () => {
  assert.strictEqual(F.normalizeKey_('ＡＢＣ　ｄｅｆ '), 'abcdef'); assert.strictEqual(F.normalizeKey_(null), '');
  assert.strictEqual(F.normalizeKey_('Taro@Example.COM'), 'taro@example.com'); assert.strictEqual(F.normalizeKey_('山田　 太郎'), '山田太郎');
});
t('dateSheetName_ 表記ゆれ', () => {
  for (const s of ['6/30', '6月30日(月)', '2026/6/30', '２０２６年６月３０日', '2026-06-30', '6.30', ' 6 / 30 ', '令和8年6月30日'])
    assert.strictEqual(F.dateSheetName_(s), '6月30日', s);
  assert.strictEqual(F.dateSheetName_('10月5日'), '10月5日');
  for (const s of ['13/1', '6/32', '0/5', 'abc', '', null, '未定']) assert.strictEqual(F.dateSheetName_(s), null, String(s));
});
t('splitDates_', () => {
  eq(F.splitDates_('6/30, 7/1、7月3日\n6/30'), { names: ['6月30日', '7月1日', '7月3日'], unknown: [] });
  eq(F.splitDates_('6/30, 未定'), { names: ['6月30日'], unknown: ['未定'] }); eq(F.splitDates_(''), { names: [], unknown: [] });
});
t('routeByDate_', () => {
  const rows = [['a', 'x', '6/30'], ['b', 'x', '6/30, 7/1'], ['c', 'x', '未定'], ['d', 'x', '7/1, 未定'], ['e', 'x', '']];
  const r = F.routeByDate_(rows, 2);
  eq(Object.keys(r.sheets).sort(), ['6月30日', '7月1日']); eq(r.sheets['6月30日'].map(x => x[0]), ['a', 'b']); eq(r.sheets['7月1日'].map(x => x[0]), ['b', 'd']);
  eq(r.review.map(x => x[0]), ['c', 'd', 'e']);
});
t('markSurveyDone_', () => {
  const ps = [{ name: '山田 太郎', email: 'Taro@Example.com' }, { name: '佐藤花子', email: '' }, { name: '鈴木', email: 's@x.jp' }];
  const sv = [{ name: '別名', email: 'taro@example.com' }, { name: '佐藤　花子', email: '' }, { name: '不明', email: 'no@x.jp' }, { name: '', email: '' }];
  const r = F.markSurveyDone_(ps, sv);
  eq(r.participants.map(p => p['済'] || ''), ['済', '済', '']); eq(r.unmatched.map(x => x.name), ['不明', '']); assert.strictEqual(ps[0]['済'], undefined);
});
t('parseCsv_', () => {
  eq(F.parseCsv_('﻿a,b\r\n"x,1","y""z"\r\n"l1\nl2",3\r\n\r\n'), [['a', 'b'], ['x,1', 'y"z'], ['l1\nl2', '3']]);
  eq(F.parseCsv_('a,b'), [['a', 'b']]); eq(F.parseCsv_(''), []); eq(F.parseCsv_('a,\n,b'), [['a', ''], ['', 'b']]);
});
t('mapHeaders_', () => {
  const al = { name: ['利用者名', '氏名'], date: ['利用日', '日付'] };
  eq(F.mapHeaders_(['No', ' 氏　名'.replace('　', ''), '日付'], al), { index: { name: 1, date: 2 }, missing: [] });
  eq(F.mapHeaders_(['Ｎｏ', '日付'], al), { index: { date: 1 }, missing: ['name'] });
});
t('reconcile_', () => {
  const planned = [{ name: '山田 太郎', date: '6/3' }, { name: '佐藤', date: '6/3' }, { name: '鈴木', date: '6/5' }];
  const actual = [{ name: '山田太郎', date: '6/3' }, { name: '田中', date: '6/5' }];
  const r = F.reconcile_(planned, actual);
  eq(r.attended.map(x => x.name), ['山田 太郎']); eq(r.absent.map(x => x.name), ['佐藤', '鈴木']); eq(r.unplanned.map(x => x.name), ['田中']);
});
t('csvToRecords_ 列名変更の吸収とエラー文', () => {
  const al = { name: ['利用者名', '氏名'], date: ['利用日', '日付'] };
  eq(F.csvToRecords_('氏名,日付\n山田,6/3\n', al), [{ name: '山田', date: '6/3' }]);
  assert.throws(() => F.csvToRecords_('名前,No\n山田,1\n', al), e => e.message.includes('name、date') && e.message.includes('名前、No'));
  assert.throws(() => F.csvToRecords_('', al), /空/);
});
t('evaluateDiagnosis_', () => {
  const ok = { sheets: ['フォーム回答', '実行ログ'], handlers: ['routeFormResponses'], lastSuccessMs: 1000000, nowMs: 1000000 + 3600000, failures24h: 0, mailQuota: 100 };
  eq(F.evaluateDiagnosis_(ok, F.DIAG_CFG).map(x => x.level), ['OK']);
  const bad = Object.assign({}, ok, { sheets: ['実行ログ'], handlers: [], lastSuccessMs: 0, nowMs: 49 * 3600000, failures24h: 2, mailQuota: 5 });
  eq(F.evaluateDiagnosis_(bad, F.DIAG_CFG).map(x => x.level), ['要対応', '要対応', '注意', '要対応', '注意']);
  eq(F.evaluateDiagnosis_(Object.assign({}, ok, { lastSuccessMs: null }), F.DIAG_CFG).map(x => x.level), ['注意']);
});
console.log(process.exitCode ? 'FAILED' : 'PASS', n + ' tests');
