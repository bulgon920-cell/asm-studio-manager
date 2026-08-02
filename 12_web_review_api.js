/**
 * Web Dashboard 書き込みAPI — Phase W3(確定する)
 * 設計書: WEB_SPEC_v1.0.md §1.2 / §3
 *
 * resolveReview(kind, targetId, payload): 要確認の解消(注文/顧客/撮影)
 * confirmAdultSelect(orderId, 商品, オプション[]): 成人セレクト確定の複合操作
 * addOrder(shootId, 種別): 注文追加
 *
 * findOrderRow_ / appendEvent_ は 11_web_write_api.js、
 * readOrderTypeCategory_ は 10_web_api.js、
 * pad_ / writeSheet_ は 02_migrate.js のものをそれぞれ再利用する(再定義しない)。
 * 11_web_write_api.js(updateStatus/updateOrderFields)には一切手を加えていない。
 */

// ===== 要確認の解消 =====
// 「実状態を確定」は通常の遷移制限(Configの遷移表)とは別扱い。
// その注文の系統(データ系/商品系/作業系)で使える全statusへ直接ジャンプできる。

const STATUS_BY_CATEGORY_ = {
  'データ系': ['データ納品待ち', '完了'],
  '作業系': ['データ納品待ち', '完了'],
  '商品系': ['店頭セレクト待ち', '発注待ち', '仕上がり待ち', '納品連絡待ち', '引渡し待ち', '完了']
};

// 要確認注文の「実状態を確定」で選べる全status一覧(読み取りのみ)。
// 10_web_api.js の getShootDetail から呼ばれ、画面のボタン生成に使う。
function getReviewStatusOptions_(ss, orderType) {
  const category = readOrderTypeCategory_(ss)[orderType] || 'データ系';
  const base = (STATUS_BY_CATEGORY_[category] || STATUS_BY_CATEGORY_['データ系']).slice();
  ['保留', '不要'].forEach(function (s) { if (base.indexOf(s) < 0) base.push(s); });
  return base;
}

function resolveReview(kind, targetId, payload) {
  if (!kind || !targetId) throw new Error('kindとtargetIdは必須です。');
  payload = payload || {};

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    throw new Error('他の操作が実行中のため待機できませんでした。しばらくして再実行してください。');
  }
  try {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    if (kind === 'order') return resolveOrderReview_(ss, targetId, payload);
    if (kind === 'customer') return resolveCustomerReview_(ss, targetId);
    if (kind === 'shoot') return resolveShootReview_(ss, targetId);
    throw new Error('不明な種別です: ' + kind);
  } finally {
    lock.releaseLock();
  }
}

function resolveOrderReview_(ss, orderId, payload) {
  const found = findOrderRow_(ss, orderId);
  if (!found) throw new Error('注文が見つかりません: ' + orderId);
  if (found.order.status !== '要確認') {
    throw new Error('この注文は要確認状態ではありません(現在: ' + found.order.status + ')。');
  }

  var to, route;
  if (payload.invalid) {
    // 移行データ不備など、案件自体が実在しない場合の無効化
    to = '不要';
    route = 'Web(要確認解消:無効)';
  } else {
    to = payload.status;
    if (!to) throw new Error('statusが指定されていません。');
    const allowed = getReviewStatusOptions_(ss, found.order.注文種別);
    if (allowed.indexOf(to) < 0) {
      appendEvent_(ss, orderId, 'status', '要確認', to, 'Web(拒否)');
      throw new Error('この注文種別では選べない状態です: ' + to);
    }
    route = 'Web(要確認解消)';
  }

  appendEvent_(ss, orderId, 'status', '要確認', to, route);
  found.sh.getRange(found.rowIndex, found.colOf['status']).setValue(to);
  if (to === '完了' && found.colOf['完了日']) {
    found.sh.getRange(found.rowIndex, found.colOf['完了日']).setValue(new Date());
  }
  return { orderId: orderId, status: to };
}

function resolveCustomerReview_(ss, customerId) {
  const sh = ss.getSheetByName('Customers');
  const found = findRowByKey_(sh, 'customerId', customerId);
  if (!found) throw new Error('顧客が見つかりません: ' + customerId);
  // 統合はv1に含めない(WEB_SPEC §3)。「別家族として確認済み」のマークのみ。
  appendEvent_(ss, customerId, '要確認', 'TRUE', 'FALSE(別家族として確認済み)', 'Web(要確認解消)');
  sh.getRange(found.rowIndex, found.colOf['要確認']).setValue(false);
  return { customerId: customerId };
}

function resolveShootReview_(ss, shootId) {
  const sh = ss.getSheetByName('Shoots');
  const found = findRowByKey_(sh, 'shootId', shootId);
  if (!found) throw new Error('撮影が見つかりません: ' + shootId);
  appendEvent_(ss, shootId, '要確認', 'TRUE', 'FALSE(確認済み)', 'Web(要確認解消)');
  sh.getRange(found.rowIndex, found.colOf['要確認']).setValue(false);
  return { shootId: shootId };
}

// ===== 成人セレクト確定 =====

const ADULT_PRODUCTS_ = ['白台紙', '雅', 'ローズ'];
const ADULT_OPTION_TYPES_ = ['全カット美肌補正', '全データ'];

function confirmAdultSelect(orderId, product, options) {
  if (!orderId || !product) throw new Error('orderIdと商品は必須です。');
  if (ADULT_PRODUCTS_.indexOf(product) < 0) throw new Error('選べない商品です: ' + product);
  const opts = (options || []).filter(function (o) { return ADULT_OPTION_TYPES_.indexOf(o) >= 0; });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    throw new Error('他の操作が実行中のため待機できませんでした。しばらくして再実行してください。');
  }
  try {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const found = findOrderRow_(ss, orderId);
    if (!found) throw new Error('注文が見つかりません: ' + orderId);
    if (found.order.注文種別 !== '成人商品(セレクト前)') {
      throw new Error('この注文はセレクト確定の対象ではありません: ' + found.order.注文種別);
    }

    // 1. 種別変更(白台紙/雅/ローズへ)
    appendEvent_(ss, orderId, '注文種別', found.order.注文種別, product, 'Web(セレクト確定)');
    found.sh.getRange(found.rowIndex, found.colOf['注文種別']).setValue(product);

    // 2. status=発注待ちへ遷移
    appendEvent_(ss, orderId, 'status', found.order.status, '発注待ち', 'Web(セレクト確定)');
    found.sh.getRange(found.rowIndex, found.colOf['status']).setValue('発注待ち');

    // 3. オプション注文行を追加(全カット美肌補正・全データはデータ系: データ納品待ちで起票)
    var seq = nextOrderSeq_(ss);
    const newRows = [];
    opts.forEach(function (opt) {
      const newId = 'O-' + pad_(seq, 5);
      seq++;
      newRows.push([newId, found.order.shootId, opt, 'データ納品待ち', '', '', '', '', 'セレクト確定時に自動追加', new Date(), '']);
      appendEvent_(ss, newId, '生成', '', opt + '(セレクト確定時に自動追加)', 'Web(セレクト確定)');
    });
    if (newRows.length) writeSheet_(ss, 'Orders', newRows);

    return { orderId: orderId, product: product, addedOptions: opts };
  } finally {
    lock.releaseLock();
  }
}

// ===== 注文追加 =====

const ADD_ORDER_INITIAL_STATUS_ = {
  '撮影データ': 'データ納品待ち',
  '全データ': 'データ納品待ち',
  '全カット美肌補正': 'データ納品待ち',
  '成人商品(セレクト前)': '店頭セレクト待ち',
  '白台紙': '発注待ち',
  '雅': '発注待ち',
  'ローズ': '発注待ち',
  '台紙': '発注待ち',
  '飾れる商品': '発注待ち'
};

function addOrder(shootId, orderType) {
  if (!shootId || !orderType) throw new Error('shootIdと種別は必須です。');
  const initialStatus = ADD_ORDER_INITIAL_STATUS_[orderType];
  if (!initialStatus) throw new Error('選べない注文種別です: ' + orderType);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    throw new Error('他の操作が実行中のため待機できませんでした。しばらくして再実行してください。');
  }
  try {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const shootFound = findRowByKey_(ss.getSheetByName('Shoots'), 'shootId', shootId);
    if (!shootFound) throw new Error('撮影が見つかりません: ' + shootId);

    const seq = nextOrderSeq_(ss);
    const newId = 'O-' + pad_(seq, 5);
    writeSheet_(ss, 'Orders', [[newId, shootId, orderType, initialStatus, '', '', '', '', '', new Date(), '']]);
    appendEvent_(ss, newId, '生成', '', orderType + '(status=' + initialStatus + ')', 'Web(注文追加)');

    return { orderId: newId, status: initialStatus };
  } finally {
    lock.releaseLock();
  }
}

// ===== 共通ヘルパー =====

// 任意シートから主キー列で1行を探す(Customers/Shootsで共用)
function findRowByKey_(sh, keyHeader, keyVal) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const values = sh.getRange(1, 1, lastRow, sh.getLastColumn()).getValues();
  const headers = values[0].map(function (h) { return String(h).trim(); });
  const colOf = {};
  headers.forEach(function (h, i) { colOf[h] = i + 1; });
  const keyIdx = colOf[keyHeader] - 1;
  for (var r = 1; r < values.length; r++) {
    if (values[r][keyIdx] === keyVal) {
      const obj = {};
      headers.forEach(function (h, i) { obj[h] = values[r][i]; });
      return { rowIndex: r + 1, colOf: colOf, obj: obj };
    }
  }
  return null;
}

function nextOrderSeq_(ss) {
  const sh = ss.getSheetByName('Orders');
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 1;
  const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  var max = 0;
  ids.forEach(function (row) {
    const v = String(row[0] || '');
    if (v.indexOf('O-') === 0) {
      const n = Number(v.slice(2));
      if (!isNaN(n) && n > max) max = n;
    }
  });
  return max + 1;
}
