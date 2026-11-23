// Code.gs の判定部分を Node.js で確かめるテスト（node test/test.js で実行）
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const src = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
const ctx = { module: { exports: {} }, console };
vm.createContext(ctx);
vm.runInContext(src, ctx);
const { findDueSoon, buildMail, toDate, CONFIG } = ctx.module.exports;

const now = new Date(2026, 10, 20, 7, 30); // 2026/11/20 7:30 を「今」とする
const d = (m, day) => new Date(2026, m - 1, day);
const values = [
  ['タスク名', '担当', '期限', '状態'],
  ['見積もりを送る', 'Aさん', d(11, 20), '未着手'],      // 今日 → 通知
  ['チラシの入稿', 'Bさん', d(11, 23), '作業中'],         // 3日後 → 通知
  ['発注の確認', 'Cさん', d(11, 24), ''],                // 4日後 → 通知しない
  ['請求書を作る', 'Aさん', d(11, 21), '完了'],           // 完了 → 通知しない
  ['昨日までの作業', 'Bさん', d(11, 19), '作業中'],       // 期限切れ → 通知しない
  ['期限なしの作業', 'Cさん', '', '未着手'],               // 期限なし → 飛ばす
  ['文字の日付', '', '2026/11/22', '作業中'],             // 文字の日付 → 通知（担当なし）
  ['完了済みの書き方', 'Dさん', d(11, 21), '対応済み'],    // 「済み」→ 通知しない
];

const items = JSON.parse(JSON.stringify(findDueSoon(values, now, CONFIG)));
assert.deepStrictEqual(items.map(i => i.task), ['見積もりを送る', '文字の日付', 'チラシの入稿']);
assert.deepStrictEqual(items.map(i => i.daysLeft), [0, 2, 3]);
assert.deepStrictEqual(items.map(i => i.row), [2, 8, 3]);

// ※ vm の中と外では配列の作りが違うため、JSONを通して比べています
// 見出しの列がないときは、わかりやすいエラーにする
assert.throws(() => findDueSoon([['タスク', '担当', '期限', '状態'], ['x', 'y', d(11, 20), '']], now, CONFIG), /タスク名/);
// データが見出しだけなら0件
assert.strictEqual(findDueSoon([values[0]], now, CONFIG).length, 0);
// 日付の読み取り
const t = toDate('2026年11月22日');
assert.ok(t && t.getFullYear() === 2026 && t.getMonth() === 10 && t.getDate() === 22);
assert.strictEqual(toDate('あした'), null);

// メールの中身
const mail = buildMail(findDueSoon(values, now, CONFIG), 'https://example.com/sheet', 'Asia/Tokyo');
assert.ok(mail.subject.includes('3件') && mail.subject.includes('今日まで'));
assert.ok(mail.text.includes('担当：未定'));
assert.ok(mail.html.includes('<table') && !mail.html.includes('<script'));
// HTMLに危ない文字が入っても、そのまま表示される
const x = buildMail([{ task: '<b>A&B</b>', owner: '', due: d(11, 20), status: '', daysLeft: 0 }], 'u', 'Asia/Tokyo');
assert.ok(x.html.includes('&lt;b&gt;A&amp;B&lt;/b&gt;'));

console.log('すべてのテストに通りました（' + items.length + '件が通知対象）');
console.log('--- メールの件名 ---\n' + mail.subject);
console.log('--- メールの本文（テキスト版）---\n' + mail.text);
