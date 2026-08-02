/**
 * Web Dashboard 書き込みAPI — Phase W2
 * 設計書: WEB_SPEC_v1.0.md §2 / §3 / §9
 *
 * updateStatus(orderId, to): Configの遷移表(区分=遷移)による検証つき状態遷移。
 * updateOrderFields(orderId, fields): 担当/メモ/期限/セレクト予定日の更新。
 *
 * 全ての書き込みはLockServiceで直列化し、変更前→変更後をEvent_Logへ
 * 追記してからOrdersを更新する。不正な遷移はAPIで拒否し、拒否もEvent_Logに残す。
 * 読み取り用の getAllowedNextStatuses_ / readTransitionTable_ は 10_web_api.js にある
 * (画面のボタン表示用。ここでは検証にも同じ関数を再利用する)。
 */

// ===== 状態遷移 =====

function updateStatus(orderId, to) {
  if (!orderId || !to) throw new Error('orderIdとtoは必須です。');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    throw new Error('他の操作が実行中のため待機できませんでした。しばらくして再実行してください。');
  }
  try {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const found = findOrderRow_(ss, orderId);
    if (!found) throw new Error('注文が見つかりません: ' + orderId);
    const from = found.order.status;

    const allowed = getAllowedNextStatuses_(ss, from, orderId);
    if (allowed.indexOf(to) < 0) {
      appendEvent_(ss, orderId, 'status', from, to, 'Web(拒否)');
      throw new Error('この状態遷移は許可されていません: 「' + from + '」→「' + to + '」');
    }

    appendEvent_(ss, orderId, 'status', from, to, 'Web');
    found.sh.getRange(found.rowIndex, found.colOf['status']).setValue(to);
    if (to === '完了' && found.colOf['完了日']) {
      found.sh.getRange(found.rowIndex, found.colOf['完了日']).setValue(new Date());
    }
    return sanitizeForClient_({ orderId: orderId, status: to });
  } finally {
    lock.releaseLock();
  }
}

// ===== 属性更新 =====

const EDITABLE_ORDER_FIELDS_ = ['担当', 'メモ', '期限', 'セレクト予定日'];
const DATE_ORDER_FIELDS_ = ['期限', 'セレクト予定日'];

function updateOrderFields(orderId, fields) {
  if (!orderId || !fields) throw new Error('orderIdとfieldsは必須です。');
  const keys = Object.keys(fields);
  keys.forEach(function (k) {
    if (EDITABLE_ORDER_FIELDS_.indexOf(k) < 0) throw new Error('更新できない項目です: ' + k);
  });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    throw new Error('他の操作が実行中のため待機できませんでした。しばらくして再実行してください。');
  }
  try {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const found = findOrderRow_(ss, orderId);
    if (!found) throw new Error('注文が見つかりません: ' + orderId);

    keys.forEach(function (key) {
      const isDateField = DATE_ORDER_FIELDS_.indexOf(key) >= 0;
      const oldRaw = found.order[key];
      const newRaw = fields[key];
      const oldStr = isDateField ? isoDate_(oldRaw) : String(oldRaw || '');
      const newStr = isDateField ? String(newRaw || '') : String(newRaw || '');
      if (oldStr === newStr) return; // 変更なしはログしない

      appendEvent_(ss, orderId, key, oldStr, newStr, 'Web');
      const col = found.colOf[key];
      const setVal = isDateField ? (newRaw ? new Date(newRaw) : '') : newRaw;
      found.sh.getRange(found.rowIndex, col).setValue(setVal);
    });

    return sanitizeForClient_({ orderId: orderId });
  } finally {
    lock.releaseLock();
  }
}

// ===== 遷移表の初期設定(1回だけ手動実行) =====
// Configシートに「区分=遷移」の行がなければ、設計書の状態遷移(§3)を書き込む。
// 何度実行しても安全(既にある組み合わせは追加しない)。

const ORDER_TRANSITIONS_SEED_ = [
  ['データ納品待ち', '完了'],
  ['店頭セレクト待ち', '発注待ち'],
  ['発注待ち', '仕上がり待ち'],
  ['仕上がり待ち', '納品連絡待ち'],
  ['納品連絡待ち', '引渡し待ち'],
  ['引渡し待ち', '完了']
];

function setupOrderTransitions() {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const sh = ss.getSheetByName('Config');
  const values = sh.getDataRange().getValues();
  const existing = {};
  for (var r = 1; r < values.length; r++) {
    if (values[r][0] === '遷移') existing[values[r][1] + '>' + values[r][2]] = true;
  }
  const toAppend = [];
  ORDER_TRANSITIONS_SEED_.forEach(function (pair) {
    const key = pair[0] + '>' + pair[1];
    if (!existing[key]) toAppend.push(['遷移', pair[0], pair[1], '', '']);
  });
  if (!toAppend.length) {
    Logger.log('遷移表は設定済みです。追加なし。');
    return;
  }
  sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, 5).setValues(toAppend);
  Logger.log('遷移表を' + toAppend.length + '件追加しました。');
}

// ===== 共通ヘルパー =====

// Ordersシートから対象行を探し、行番号・列番号マップ・現在値をまとめて返す
function findOrderRow_(ss, orderId) {
  const sh = ss.getSheetByName('Orders');
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const values = sh.getRange(1, 1, lastRow, sh.getLastColumn()).getValues();
  const headers = values[0].map(function (h) { return String(h).trim(); });
  const colOf = {};
  headers.forEach(function (h, i) { colOf[h] = i + 1; });

  for (var r = 1; r < values.length; r++) {
    if (values[r][colOf['orderId'] - 1] === orderId) {
      const obj = {};
      headers.forEach(function (h, i) { obj[h] = values[r][i]; });
      return { sh: sh, rowIndex: r + 1, colOf: colOf, order: obj };
    }
  }
  return null;
}

// Event_Logへ1行追記する(経路='Web' または 'Web(拒否)')
function appendEvent_(ss, target, item, before, after, route) {
  const sh = ss.getSheetByName('Event_Log');
  const lastRow = sh.getLastRow();
  const nextId = lastRow >= 2 ? Number(sh.getRange(lastRow, 1).getValue()) + 1 : 1;
  var operator = 'Web';
  try {
    const email = Session.getActiveUser().getEmail();
    if (email) operator = email;
  } catch (e) {
    // アクセス権限次第でメールが取得できないことがある。既定値のまま続行する。
  }
  sh.appendRow([nextId, new Date(), operator, target, item, before, after, route]);
}
