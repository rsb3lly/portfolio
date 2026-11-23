/**
 * 期限が近いタスクを毎朝メールで知らせる Google Apps Script
 *
 * シートの1行目を見出しにして、次の4つの列を用意してください（順番は自由）。
 *   タスク名 / 担当 / 期限 / 状態
 * 期限が「今日〜3日後」で、状態が「完了」でない行を、1通のメールにまとめて送ります。
 *
 * 送り先のメールアドレスはコードに書かず、
 * 「プロジェクトの設定」→「スクリプト プロパティ」に NOTIFY_EMAIL として登録します。
 * MIT License (c) 2026 Berry
 */

// ===== 設定（必要に応じて変えてください）=====
var CONFIG = {
  SHEET_NAME: 'タスク',   // 読み込むシートの名前
  DAYS_AHEAD: 3,          // 何日後までを「期限が近い」とするか
  DONE_WORDS: ['完了', '済', '済み', 'done'], // この言葉が状態に入っていたら通知しない
  NOTIFY_HOUR: 7,         // 毎朝何時台に送るか（0〜23）
  COLUMNS: { task: 'タスク名', owner: '担当', due: '期限', status: '状態' }
};

/** 毎朝トリガーから呼ばれる関数 */
function notifyDeadlines() {
  var to = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL');
  if (!to) {
    throw new Error('スクリプト プロパティに NOTIFY_EMAIL が登録されていません（README の手順4を見てください）');
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    throw new Error('シート「' + CONFIG.SHEET_NAME + '」が見つかりません。シート名か CONFIG.SHEET_NAME を確認してください');
  }
  var values = sheet.getDataRange().getValues();
  var items = findDueSoon(values, new Date(), CONFIG);
  if (items.length === 0) {
    console.log('期限が近いタスクはありませんでした。メールは送りません。');
    return;
  }
  var tz = Session.getScriptTimeZone();
  var mail = buildMail(items, sheet.getParent().getUrl(), tz);
  MailApp.sendEmail({ to: to, subject: mail.subject, body: mail.text, htmlBody: mail.html });
  console.log(items.length + '件を ' + to + ' に送りました。');
}

/**
 * 期限が近い行を取り出す（シートやメールに触らない純粋な関数なので、テストしやすい）
 * @param {Array<Array<*>>} values シートの全データ（1行目が見出し）
 * @param {Date} now 今の日時
 * @param {Object} config 設定
 * @return {Array<Object>} {row, task, owner, due, status, daysLeft} の配列（期限の近い順）
 */
function findDueSoon(values, now, config) {
  if (!values || values.length < 2) return [];
  var head = values[0].map(function (h) { return String(h).trim(); });
  var col = {};
  Object.keys(config.COLUMNS).forEach(function (key) {
    var idx = head.indexOf(config.COLUMNS[key]);
    if (idx < 0) throw new Error('見出しに「' + config.COLUMNS[key] + '」の列がありません');
    col[key] = idx;
  });
  var today = startOfDay(now);
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var due = toDate(row[col.due]);
    if (!due) continue; // 期限が空・読めない行は飛ばす
    var status = String(row[col.status] || '').trim();
    var done = config.DONE_WORDS.some(function (w) { return status.toLowerCase().indexOf(w.toLowerCase()) >= 0; });
    if (done) continue;
    var daysLeft = Math.round((startOfDay(due) - today) / 86400000);
    if (daysLeft >= 0 && daysLeft <= config.DAYS_AHEAD) {
      out.push({ row: r + 1, task: String(row[col.task]), owner: String(row[col.owner] || ''),
                 due: startOfDay(due), status: status, daysLeft: daysLeft });
    }
  }
  out.sort(function (a, b) { return a.daysLeft - b.daysLeft || a.row - b.row; });
  return out;
}

/** メールの件名と本文（テキスト版・HTML版）を作る */
function buildMail(items, sheetUrl, tz) {
  var fmt = function (d) {
    return typeof Utilities !== 'undefined' ? Utilities.formatDate(d, tz, 'M/d（E）')
      : (d.getMonth() + 1) + '/' + d.getDate();
  };
  var label = function (n) { return n === 0 ? '今日まで' : 'あと' + n + '日'; };
  var subject = '【期限が近いタスク】' + items.length + '件（' + label(items[0].daysLeft) + 'のものがあります）';
  var lines = items.map(function (it) {
    return '・' + label(it.daysLeft) + '｜' + fmt(it.due) + '｜' + it.task + '（担当：' + (it.owner || '未定') + '）';
  });
  var text = '期限が近いタスクが' + items.length + '件あります。\n\n' + lines.join('\n') + '\n\nシート：' + sheetUrl;
  var esc = function (s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var rows = items.map(function (it) {
    var color = it.daysLeft === 0 ? '#b3261e' : (it.daysLeft === 1 ? '#a15c00' : '#2f6b4f');
    return '<tr><td style="padding:6px 10px;border:1px solid #d9e1d6;color:' + color + ';font-weight:bold">' + label(it.daysLeft) +
      '</td><td style="padding:6px 10px;border:1px solid #d9e1d6">' + fmt(it.due) +
      '</td><td style="padding:6px 10px;border:1px solid #d9e1d6">' + esc(it.task) +
      '</td><td style="padding:6px 10px;border:1px solid #d9e1d6">' + esc(it.owner || '未定') +
      '</td><td style="padding:6px 10px;border:1px solid #d9e1d6">' + esc(it.status || '—') + '</td></tr>';
  }).join('');
  var html = '<p>期限が近いタスクが <b>' + items.length + '件</b> あります。</p>' +
    '<table style="border-collapse:collapse;font-size:14px"><tr style="background:#2f6b4f;color:#fff">' +
    '<th style="padding:6px 10px">残り</th><th style="padding:6px 10px">期限</th><th style="padding:6px 10px">タスク</th>' +
    '<th style="padding:6px 10px">担当</th><th style="padding:6px 10px">状態</th></tr>' + rows + '</table>' +
    '<p><a href="' + esc(sheetUrl) + '">シートを開く</a></p>';
  return { subject: subject, text: text, html: html };
}

/** 毎朝のトリガーを1つだけ作る（2回実行しても重複しない） */
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'notifyDeadlines') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('notifyDeadlines').timeBased().everyDays(1).atHour(CONFIG.NOTIFY_HOUR).create();
  console.log('毎朝' + CONFIG.NOTIFY_HOUR + '時台に notifyDeadlines を実行するトリガーを作りました。');
}

// ===== 小さな道具 =====
function startOfDay(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }

function toDate(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) return v;
  if (typeof v === 'string' && v.trim()) {
    var m = v.trim().match(/^(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})日?$/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }
  return null;
}

// Node.js でテストするときだけ使う（GAS では無視されます）
if (typeof module !== 'undefined') {
  module.exports = { findDueSoon: findDueSoon, buildMail: buildMail, toDate: toDate, CONFIG: CONFIG };
}
