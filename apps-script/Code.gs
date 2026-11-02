/**
 * iPad端末管理簿 データ保存用スクリプト
 * Googleスプレッドシートの「拡張機能 > Apps Script」に貼り付けて使います。
 * 手順は README.md を見てください。
 */

const SHEETS = {
  devices: {
    name: '端末',
    cols: [
      ['id', 'ID'], ['no', '管理番号'], ['status', '状態'], ['holder', '使用者'], ['holderClass', 'クラス'],
      ['outAt', '渡した日'], ['due', '返却予定日'], ['model', '機種'], ['serial', 'シリアル番号'],
      ['bought', '購入日'], ['account', 'アカウント'], ['place', '保管場所'], ['memo', 'メモ'],
      ['loanId', '貸出ID'], ['prevStatus', '修理前の状態'], ['updatedAt', '更新日時'], ['updatedBy', '更新した人']
    ]
  },
  loans: {
    name: '配布貸出履歴',
    cols: [
      ['id', 'ID'], ['deviceId', '端末ID'], ['no', '管理番号'], ['type', '種別'], ['name', '氏名'], ['cls', 'クラス'],
      ['outAt', '渡した日'], ['due', '返却予定日'], ['inAt', '返却日'], ['memo', 'メモ'], ['inMemo', '返却時メモ'],
      ['outBy', '渡した人'], ['inBy', '受け取った人'], ['createdAt', '作成日時'], ['updatedAt', '更新日時'], ['updatedBy', '更新した人']
    ]
  },
  accounts: {
    name: 'アカウント',
    cols: [
      ['id', 'ID'], ['accountId', 'アカウントID'], ['kind', '種類'], ['state', '状態'], ['name', '氏名'], ['cls', 'クラス'],
      ['deviceNo', '管理番号'], ['memo', 'メモ'], ['updatedAt', '更新日時'], ['updatedBy', '更新した人']
    ]
  },
  repairs: {
    name: '修理',
    cols: [
      ['id', 'ID'], ['deviceId', '端末ID'], ['no', '管理番号'], ['step', '進み具合'], ['reportedAt', '発生日'],
      ['cause', '区分'], ['symptom', '症状'], ['holder', '使用者'], ['holderClass', 'クラス'], ['sentAt', '修理依頼日'],
      ['doneAt', '完了日'], ['vendor', '依頼先'], ['cost', '費用'], ['memo', 'メモ'], ['prevStatus', '修理前の状態'],
      ['createdAt', '作成日時'], ['updatedAt', '更新日時'], ['updatedBy', '更新した人']
    ]
  }
};

/* ---------- スプレッドシートのメニュー ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('iPad管理簿')
    .addItem('① シートを準備する', 'setupSheets')
    .addItem('② 合言葉を設定する', 'setPassphrase')
    .addToUi();
}

function setupSheets() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(SHEETS).forEach(function (key) {
    const def = SHEETS[key];
    let sh = ss.getSheetByName(def.name);
    if (!sh) sh = ss.insertSheet(def.name);
    const head = def.cols.map(function (c) { return c[1]; });
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e1efea');
    sh.setFrozenRows(1);
    sh.getRange('A:Z').setNumberFormat('@'); // 日付や番号が勝手に変換されないよう文字列扱い
  });
  const first = ss.getSheetByName('シート1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  SpreadsheetApp.getUi().alert('シートを準備しました。次に「② 合言葉を設定する」を選んでください。');
}

function setPassphrase() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('合言葉を設定', '先生方がアプリに接続するときに入れる合言葉を決めてください(8文字以上がおすすめ)。', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const p = r.getResponseText().trim();
  if (p.length < 4) { ui.alert('4文字以上にしてください。'); return; }
  PropertiesService.getScriptProperties().setProperty('PASSPHRASE', p);
  ui.alert('合言葉を設定しました。');
}

/* ---------- アプリからの受付 ---------- */
function doGet() {
  return json({ ok: true, message: 'iPad端末管理簿のデータ受付は動いています。' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'bad_request' }); }
  const pass = PropertiesService.getScriptProperties().getProperty('PASSPHRASE');
  if (!pass) return json({ ok: false, error: 'no_passphrase' });
  if (req.key !== pass) return json({ ok: false, error: 'wrong_key' });

  try {
    if (req.action === 'load') return json({ ok: true, data: loadAll() });
    if (req.action === 'batch') {
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        applyOps(req.ops || [], String(req.by || '').slice(0, 40));
        SpreadsheetApp.flush();
      } finally { lock.releaseLock(); }
      return json({ ok: true, data: loadAll() });
    }
    return json({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json({ ok: false, error: 'server_error', message: String(err && err.message || err) });
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheetOf(key) {
  const def = SHEETS[key];
  if (!def) throw new Error('unknown sheet: ' + key);
  const sh = SpreadsheetApp.getActive().getSheetByName(def.name);
  if (!sh) throw new Error('シート「' + def.name + '」がありません。メニューの「① シートを準備する」を実行してください。');
  return { def: def, sh: sh };
}

function loadAll() {
  const out = {};
  Object.keys(SHEETS).forEach(function (key) {
    const s = sheetOf(key);
    const keys = s.def.cols.map(function (c) { return c[0]; });
    const n = s.sh.getLastRow() - 1;
    if (n <= 0) { out[key] = []; return; }
    const vals = s.sh.getRange(2, 1, n, keys.length).getDisplayValues();
    out[key] = vals.filter(function (r) { return r[0]; }).map(function (r) {
      const o = {};
      keys.forEach(function (k, i) { if (r[i] !== '') o[k] = r[i]; });
      return o;
    });
  });
  return out;
}

function applyOps(ops, by) {
  const cache = {};
  function get(key) {
    if (cache[key]) return cache[key];
    const s = sheetOf(key);
    const keys = s.def.cols.map(function (c) { return c[0]; });
    const n = s.sh.getLastRow() - 1;
    const ids = n > 0 ? s.sh.getRange(2, 1, n, 1).getDisplayValues().map(function (r) { return r[0]; }) : [];
    cache[key] = { sh: s.sh, keys: keys, ids: ids, pending: {}, order: [] };
    return cache[key];
  }
  const now = new Date().toISOString();
  ops.forEach(function (op) {
    const t = get(op.sheet);
    const id = String(op.id || '');
    if (!id) return;
    if (op.op === 'delete') {
      if (t.pending[id]) { delete t.pending[id]; t.order = t.order.filter(function (x) { return x !== id; }); return; }
      const i = t.ids.indexOf(id);
      if (i >= 0) { t.sh.deleteRow(i + 2); t.ids.splice(i, 1); }
      return;
    }
    if (op.op !== 'upsert') return;
    const rec = op.data || {};
    rec.updatedAt = now;
    if (by) rec.updatedBy = by;
    if (t.pending[id]) {
      t.keys.forEach(function (k, j) { if (k !== 'id' && Object.prototype.hasOwnProperty.call(rec, k)) t.pending[id][j] = clean(rec[k]); });
      return;
    }
    const i = t.ids.indexOf(id);
    if (i >= 0) {
      const range = t.sh.getRange(i + 2, 1, 1, t.keys.length);
      const row = range.getDisplayValues()[0];
      t.keys.forEach(function (k, j) { if (k !== 'id' && Object.prototype.hasOwnProperty.call(rec, k)) row[j] = clean(rec[k]); });
      range.setValues([row]);
    } else {
      t.pending[id] = t.keys.map(function (k) { return k === 'id' ? id : clean(rec[k]); });
      t.order.push(id);
    }
  });
  // 新しい行はまとめて書き込む
  Object.keys(cache).forEach(function (key) {
    const t = cache[key];
    if (!t.order.length) return;
    const rows = t.order.map(function (id) { return t.pending[id]; });
    t.sh.getRange(t.sh.getLastRow() + 1, 1, rows.length, t.keys.length).setValues(rows);
  });
}

function clean(v) {
  if (v === undefined || v === null) return '';
  let s = String(v).slice(0, 2000);
  if (/^[=+\-@]/.test(s)) s = "'" + s; // 数式として実行されないように
  return s;
}
