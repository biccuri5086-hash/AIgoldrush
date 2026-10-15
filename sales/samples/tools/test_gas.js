// Code.gs の aggregateOrders_ を取り出し、Python側の期待値(fixture.json)と照合する。
const fs = require('fs'), vm = require('vm'), path = require('path');
const dir = path.join(__dirname, '..', 'gas');
const ctx = {}; vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(dir, 'Code.gs'), 'utf8') + '\nthis.agg = aggregateOrders_;', ctx);
const fx = JSON.parse(fs.readFileSync(path.join(dir, 'fixture.json'), 'utf8'));
const D = vm.runInContext('Date', ctx);  // vmの別コンテキストのDateを使う(instanceof対策)
const rows = fx.rows.map(r => [r[0] ? new D(r[0]) : null, r[1], r[2], r[3], r[4]]);
const res = ctx.agg(rows);
const a = JSON.stringify(res.list), b = JSON.stringify(fx.expected);
let ok = a === b && res.skipped === fx.skipped;
// 追加ケース: 空・全件不正
const e = ctx.agg([]); ok = ok && e.list.length === 0 && e.skipped === 0;
const z = ctx.agg([[new D('2026-09-01T10:00:00'), 'x', 'A', 0, 100], [new D('2026-09-01T10:00:00'), 'x', 'A', 1, -5]]);
ok = ok && z.list.length === 0 && z.skipped === 2;
// 月またぎ・金額の整数性
const m = ctx.agg([[new D('2026-08-31T23:59:00'), 'x', 'A', 3, 333], [new D('2026-09-01T00:00:00'), 'x', 'A', 3, 333]]);
ok = ok && m.list.length === 2 && m.list[0].amount === 999 && m.list[1].amount === 999;
console.log(ok ? 'PASS' : 'FAIL', 'groups=' + res.list.length, 'skipped=' + res.skipped);
if (!ok) { console.log(a); console.log(b); process.exit(1); }
