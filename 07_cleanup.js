/**
 * 過去データの一括整理(一度きり)
 * 設計書: SYNC_SPEC_v1.0.md §6(業務ジャンル)
 *
 * 業務ジャンル対応(05_sync.js)より前に同期・移行されてしまった「その他撮影」等の
 * 業務ジャンル分について、誤って起票されたOrdersと顧客の要確認フラグを整理する。
 *
 * 使い方: cleanupBusinessJobsDryRun() で対象を確認 → 問題なければ cleanupBusinessJobs()
 *
 * 既存データの削除は一切しない。Orders.statusとCustomers.要確認/備考の変更のみ。
 * Shootsには触れない(SYNC_SPEC §6(c): 将来の売上統合のため記録を残す)。
 *
 * readBusinessGenreNames_ / readSheetObjects_ は 05_sync.js、
 * findRowByKey_ は 12_web_review_api.js、appendEvent_ は 11_web_write_api.js の
 * ものをそれぞれ再利用する(再定義しない)。
 */

function cleanupBusinessJobsDryRun() {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const plan = buildCleanupPlan_(ss);
  logCleanupPlan_(plan);
  Logger.log('dryRun完了。書き込みは行っていません。');
}

function cleanupBusinessJobs() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    throw new Error('他の操作が実行中のため待機できませんでした。しばらくして再実行してください。');
  }
  try {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const plan = buildCleanupPlan_(ss);
    logCleanupPlan_(plan);

    const ordersSh = ss.getSheetByName('Orders');
    plan.orders.forEach(function (o) {
      appendEvent_(ss, o.orderId, 'status', o.status, '不要', '整理');
      appendEvent_(ss, o.orderId, 'メモ', o.memo, '広告案件のため対象外', '整理');
      ordersSh.getRange(o.rowIndex, o.colOf['status']).setValue('不要');
      ordersSh.getRange(o.rowIndex, o.colOf['メモ']).setValue('広告案件のため対象外');
    });

    const custSh = ss.getSheetByName('Customers');
    plan.customers.forEach(function (c) {
      const newRemarks = c.remarks ? (c.remarks + ' / 広告案件') : '広告案件';
      if (c.needsReview) {
        appendEvent_(ss, c.customerId, '要確認', 'TRUE', 'FALSE(整理)', '整理');
        custSh.getRange(c.rowIndex, c.colOf['要確認']).setValue(false);
      }
      if (newRemarks !== c.remarks) {
        appendEvent_(ss, c.customerId, '備考', c.remarks, newRemarks, '整理');
        custSh.getRange(c.rowIndex, c.colOf['備考']).setValue(newRemarks);
      }
    });

    Logger.log('整理完了: 注文' + plan.orders.length + '件をstatus=不要へ、顧客' +
      plan.customers.length + '件の要確認解除・備考記録を行いました。');
  } finally {
    lock.releaseLock();
  }
}

// ===== 対象の洗い出し =====

function buildCleanupPlan_(ss) {
  const businessNames = readBusinessGenreNames_(ss); // 05_sync.js

  const custSh = ss.getSheetByName('Customers');
  const custRows = readSheetObjects_(ss, 'Customers'); // 05_sync.js
  const targetCustomerIds = {};
  const customers = [];
  custRows.forEach(function (c) {
    if (businessNames.indexOf(c.顧客名) < 0) return;
    targetCustomerIds[c.customerId] = true;
    const found = findRowByKey_(custSh, 'customerId', c.customerId); // 12_web_review_api.js
    customers.push({
      customerId: c.customerId, rowIndex: found.rowIndex, colOf: found.colOf,
      remarks: c.備考 || '', needsReview: !!c.要確認
    });
  });

  const shootRows = readSheetObjects_(ss, 'Shoots');
  const targetShootIds = {};
  shootRows.forEach(function (s) {
    if (targetCustomerIds[s.customerId]) targetShootIds[s.shootId] = true;
  });

  const orderSh = ss.getSheetByName('Orders');
  const orderRows = readSheetObjects_(ss, 'Orders');
  const orders = [];
  orderRows.forEach(function (o) {
    if (!targetShootIds[o.shootId]) return;
    if (o.status === '不要') return; // 既に整理済み(冪等)
    const found = findRowByKey_(orderSh, 'orderId', o.orderId);
    orders.push({
      orderId: o.orderId, shootId: o.shootId, orderType: o.注文種別, status: o.status,
      memo: o.メモ || '', rowIndex: found.rowIndex, colOf: found.colOf
    });
  });

  return { customers: customers, orders: orders };
}

function logCleanupPlan_(plan) {
  Logger.log('=== 業務ジャンル整理: 対象一覧 ===');
  Logger.log('対象顧客: ' + plan.customers.length + '件');
  plan.customers.forEach(function (c) {
    Logger.log('  顧客 ' + c.customerId + '(要確認=' + c.needsReview + ', 備考=「' + c.remarks + '」)');
  });
  Logger.log('対象注文: ' + plan.orders.length + '件');
  plan.orders.forEach(function (o) {
    Logger.log('  注文 ' + o.orderId + '(' + o.orderType + ', 現status=' + o.status + ') shoot=' + o.shootId);
  });
}
