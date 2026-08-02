/**
 * 差分同期(S1) — 大本Master_Logの新規行だけを新環境へ追記する常設機構。
 * 設計書: SYNC_SPEC_v1.0.md
 *
 * 使い方(必ずこの順で):
 *   1. syncDryRun() を実行 → 「同期レポート」シートで内容確認(書き込みなし)
 *   2. 問題なければ syncRun() を実行 → 本実行(追記のみ)
 *
 * 大本(SOURCE_MASTER_LOG_ID)へは一切書き込まない。新環境の既存行(Customers/
 * Shoots/Ordersの既存レコード)は変更・削除しない。特にOrders.statusには触れない。
 * 想定と違う状態を検出したら推測で続行せず、エラーを報告して停止する。
 *
 * customerRow_/memberRow_/shootRow_/orderRow_/writeSheet_/norm_/childKey_/
 * toDate_/isoDate_/pad_/hasLine_ は 02_migrate.js のものをそのまま再利用する
 * (このファイルでは再定義しない)。
 */

// ===== 入口 =====

function syncDryRun() { return syncCore_(true); }
function syncRun() { return syncCore_(false); }

function syncCore_(dryRun) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('他の処理が実行中のため待機できませんでした。しばらくして再実行してください。');
  }
  try {
    const targetSs = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    ensureSyncInitialized_(targetSs);

    const sourceSh = getSourceMasterLogSheet_();
    const allRows = readMasterLogRows_(sourceSh);

    const lastRow = Number(getConfigValue_(targetSs, 'SYNC_LAST_ROW') || 0);
    const lastHash = String(getConfigValue_(targetSs, 'SYNC_LAST_HASH') || '');

    const checkRow = allRows.filter(function (r) { return r.srcRow === lastRow; })[0];
    if (!checkRow || hashRow_(checkRow) !== lastHash) {
      const msg = '同期位置(大本の行' + lastRow + ')の整合性チェックに失敗しました。' +
        '行の追加・削除・並べ替え・編集が大本側で起きた可能性があります。処理を中断しました。' +
        '状況を確認し、必要ならConfigシートの sync/SYNC_LAST_ROW・sync/SYNC_LAST_HASH を' +
        '人が再設定してください。';
      Logger.log(msg);
      writeSyncReport_(targetSs, [], 0, msg, dryRun);
      throw new Error(msg);
    }

    const newRows = allRows.filter(function (r) { return r.srcRow > lastRow; });

    if (!newRows.length) {
      Logger.log('新規行はありません。同期の必要なし。');
      writeSyncReport_(targetSs, [], 0, '', dryRun);
      return;
    }

    const idx = buildSyncIndex_(targetSs);
    const plan = { customers: [], members: [], shoots: [], orders: [], events: [] };
    const results = newRows.map(function (row) { return processSyncRow_(row, idx, plan); });

    writeSyncReport_(targetSs, results, newRows.length, '', dryRun);

    if (dryRun) {
      Logger.log('syncDryRun完了。差分' + newRows.length + '件。「同期レポート」シートを確認してください。');
      return;
    }

    if (plan.customers.length) writeSheet_(targetSs, 'Customers', plan.customers.map(customerRow_));
    if (plan.members.length) writeSheet_(targetSs, 'Family_Members', plan.members.map(memberRow_));
    if (plan.shoots.length) writeSheet_(targetSs, 'Shoots', plan.shoots.map(shootRow_));
    if (plan.orders.length) writeSheet_(targetSs, 'Orders', plan.orders.map(orderRow_));
    if (plan.events.length) writeSheet_(targetSs, 'Event_Log', plan.events);

    const last = newRows[newRows.length - 1];
    setConfigValue_(targetSs, 'SYNC_LAST_ROW', String(last.srcRow));
    setConfigValue_(targetSs, 'SYNC_LAST_HASH', hashRow_(last));

    Logger.log('syncRun完了: 顧客+' + plan.customers.length + ' 家族+' + plan.members.length +
      ' 撮影+' + plan.shoots.length + ' 注文+' + plan.orders.length +
      '(大本行' + last.srcRow + 'まで処理済み)');
  } finally {
    lock.releaseLock();
  }
}

// 初回のみ: 移行時に使ったスナップショットの最終行を同期の開始位置として設定する。
// (0から始めると移行済みの全件を新規行として重複作成してしまうため)
function ensureSyncInitialized_(targetSs) {
  if (getConfigValue_(targetSs, 'SYNC_LAST_ROW')) return; // 初期化済み

  const snapSh = SNAPSHOT_SHEET_NAME
    ? SpreadsheetApp.openById(SNAPSHOT_SPREADSHEET_ID).getSheetByName(SNAPSHOT_SHEET_NAME)
    : SpreadsheetApp.openById(SNAPSHOT_SPREADSHEET_ID).getSheets()[0];
  if (!snapSh) throw new Error('同期位置の初期化に失敗: スナップショットシートが見つかりません。');

  const snapRows = readMasterLogRows_(snapSh);
  if (!snapRows.length) throw new Error('同期位置の初期化に失敗: スナップショットにデータ行がありません。');

  const lastSnapRow = snapRows[snapRows.length - 1];
  setConfigValue_(targetSs, 'SYNC_LAST_ROW', String(lastSnapRow.srcRow));
  setConfigValue_(targetSs, 'SYNC_LAST_HASH', hashRow_(lastSnapRow));
  Logger.log('同期位置を移行スナップショット基準で初期化しました: 大本行' + lastSnapRow.srcRow +
    '(この行までは移行済みとして扱います)');
}

// ===== 大本Master_Logの読み取り(構造は02_migrate.jsのreadSnapshot_と同じ) =====

function getSourceMasterLogSheet_() {
  const ss = SpreadsheetApp.openById(SOURCE_MASTER_LOG_ID);
  const sh = SOURCE_MASTER_LOG_SHEET_NAME
    ? ss.getSheetByName(SOURCE_MASTER_LOG_SHEET_NAME)
    : ss.getSheets()[0];
  if (!sh) throw new Error('大本Master_Logのシートが見つかりません: ' + SOURCE_MASTER_LOG_SHEET_NAME);
  return sh;
}

function readMasterLogRows_(sh) {
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(function (h) { return String(h).trim(); });

  const idx = {};
  headers.forEach(function (h, i) { if (!(h in idx)) idx[h] = i; });
  const genderCols = [];
  headers.forEach(function (h, i) { if (h === '性別') genderCols.push(i); });
  const remarksCol = headers.findIndex(function (h) { return h.indexOf('備考') === 0; });

  function col(name) { return (name in idx) ? idx[name] : -1; }

  const rows = [];
  for (var r = 1; r < values.length; r++) {
    const v = values[r];
    const get = function (name) {
      const c = col(name);
      return c >= 0 ? v[c] : '';
    };
    if (!String(get('お名前')).trim() && !get('日付')) continue; // 完全な空行はスキップ

    const children = [];
    for (var n = 1; n <= 3; n++) {
      const cname = String(get('子どもの名前' + n)).trim();
      if (!cname) continue;
      children.push({
        name: cname,
        kana: String(get('ふりがな' + n)).trim(),
        gender: genderCols[n - 1] != null ? String(v[genderCols[n - 1]]).trim() : '',
        birth: toDate_(get('誕生日' + n))
      });
    }
    rows.push({
      srcRow: r + 1,
      date: toDate_(get('日付')),
      genre: String(get('撮影ジャンル')).trim(),
      name: String(get('お名前')).trim(),
      zip: String(get('郵便番号')).trim(),
      addr: String(get('住所')).trim(),
      contact: String(get('連絡先')).trim(),
      hp: String(get('HP掲載について')).trim(),
      finish: toDate_(get('仕上がり予定日')),
      amount: get('合計金額'),
      payment: String(get('決済方法')).trim(),
      remarks: remarksCol >= 0 ? String(v[remarksCol]).trim() : '',
      children: children
    });
  }
  return rows;
}

function hashRow_(row) {
  const s = [
    row.date ? isoDate_(row.date) : '',
    row.name || '',
    row.genre || '',
    row.amount != null ? String(row.amount) : ''
  ].join('|');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8);
  return digest.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

// ===== 新環境の既存データを読み、同定・採番用インデックスを作る =====

function buildSyncIndex_(targetSs) {
  const customers = readSheetObjects_(targetSs, 'Customers');
  const members = readSheetObjects_(targetSs, 'Family_Members');
  const shoots = readSheetObjects_(targetSs, 'Shoots');
  const orders = readSheetObjects_(targetSs, 'Orders');
  const events = readSheetObjects_(targetSs, 'Event_Log');

  const byChildKey = {}, byNamePhone = {}, byNameAddr = {}, nameSeen = {};
  const memberKeysByCustomer = {};

  customers.forEach(function (c) {
    nameSeen[norm_(c.顧客名)] = true;
    if (c.連絡先) byNamePhone[norm_(c.顧客名) + '|' + norm_(c.連絡先)] = c.customerId;
    if (c.住所) byNameAddr[norm_(c.顧客名) + '|' + norm_(c.住所)] = c.customerId;
  });
  members.forEach(function (m) {
    if (!memberKeysByCustomer[m.customerId]) memberKeysByCustomer[m.customerId] = {};
    const birthIso = m.誕生日 instanceof Date ? isoDate_(m.誕生日) : '';
    memberKeysByCustomer[m.customerId][norm_(m.名前) + '|' + birthIso] = m.memberId;
    if (birthIso) {
      const full = norm_(m.名前) + '|' + birthIso;
      byChildKey[full] = m.customerId;
    }
  });

  function maxSeq_(rows, key, prefix) {
    var max = 0;
    rows.forEach(function (r) {
      const v = String(r[key] || '');
      if (v.indexOf(prefix) === 0) {
        const n = Number(v.slice(prefix.length));
        if (!isNaN(n) && n > max) max = n;
      }
    });
    return max;
  }

  const shootSeqByDate = {};
  shoots.forEach(function (s) {
    if (s.撮影日 instanceof Date) {
      const dstr = Utilities.formatDate(s.撮影日, 'Asia/Tokyo', 'yyyyMMdd');
      const m = String(s.shootId).match(/-(\d+)$/);
      const seq = m ? Number(m[1]) : 0;
      shootSeqByDate[dstr] = Math.max(shootSeqByDate[dstr] || 0, seq);
    }
  });

  var maxEventSeq = 0;
  events.forEach(function (e) {
    const n = Number(e.eventId);
    if (!isNaN(n) && n > maxEventSeq) maxEventSeq = n;
  });

  return {
    customerSeq: maxSeq_(customers, 'customerId', 'C-'),
    memberSeq: maxSeq_(members, 'memberId', 'M-'),
    orderSeq: maxSeq_(orders, 'orderId', 'O-'),
    eventSeq: maxEventSeq,
    shootSeqByDate: shootSeqByDate,
    byChildKey: byChildKey,
    byNamePhone: byNamePhone,
    byNameAddr: byNameAddr,
    nameSeen: nameSeen,
    memberKeysByCustomer: memberKeysByCustomer
  };
}

function readSheetObjects_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const values = sh.getRange(1, 1, lastRow, sh.getLastColumn()).getValues();
  const headers = values[0].map(function (h) { return String(h).trim(); });
  const rows = [];
  for (var r = 1; r < values.length; r++) {
    const obj = {};
    headers.forEach(function (h, i) { obj[h] = values[r][i]; });
    if (headers[0] && !String(obj[headers[0]]).trim()) continue;
    rows.push(obj);
  }
  return rows;
}

// ===== Configの同期位置(区分=sync)の読み書き =====

function getConfigValue_(ss, key) {
  const sh = ss.getSheetByName('Config');
  const values = sh.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    if (values[r][0] === 'sync' && values[r][1] === key) return values[r][2];
  }
  return '';
}
function setConfigValue_(ss, key, value) {
  const sh = ss.getSheetByName('Config');
  const values = sh.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    if (values[r][0] === 'sync' && values[r][1] === key) {
      sh.getRange(r + 1, 3).setValue(value);
      return;
    }
  }
  sh.appendRow(['sync', key, value, '', '']);
}

// ===== 差分1行の処理(新規案件テンプレートで起票。SYNC_SPEC §3) =====

function processSyncRow_(row, idx, plan) {
  var custId = null, isNewCustomer = false, customerNeedsReview = false;

  row.children.forEach(function (ch) {
    if (custId) return;
    const k = childKey_(ch);
    if (k && idx.byChildKey[k]) custId = idx.byChildKey[k];
  });
  if (!custId && row.contact) {
    const k = norm_(row.name) + '|' + norm_(row.contact);
    if (idx.byNamePhone[k]) custId = idx.byNamePhone[k];
  }
  if (!custId && row.addr) {
    const k = norm_(row.name) + '|' + norm_(row.addr);
    if (idx.byNameAddr[k]) custId = idx.byNameAddr[k];
  }

  if (!custId) {
    isNewCustomer = true;
    idx.customerSeq++;
    custId = 'C-' + pad_(idx.customerSeq, 5);
    customerNeedsReview = !!idx.nameSeen[norm_(row.name)];
    plan.customers.push({
      id: custId, name: row.name, kana: '', contact: row.contact, zip: row.zip,
      addr: row.addr, line: hasLine_(row.remarks), remarks: '', needsReview: customerNeedsReview
    });
    idx.nameSeen[norm_(row.name)] = true;
    idx.memberKeysByCustomer[custId] = {};
  }
  if (row.contact) idx.byNamePhone[norm_(row.name) + '|' + norm_(row.contact)] = custId;
  if (row.addr) idx.byNameAddr[norm_(row.name) + '|' + norm_(row.addr)] = custId;

  // 家族構成員(未知の子どものみ追加)
  const memberIds = [];
  const newMembersThisRow = [];
  row.children.forEach(function (ch) {
    if (!ch.name) return;
    const birthIso = ch.birth ? isoDate_(ch.birth) : '';
    const mk = norm_(ch.name) + '|' + birthIso;
    if (!idx.memberKeysByCustomer[custId]) idx.memberKeysByCustomer[custId] = {};
    if (idx.memberKeysByCustomer[custId][mk]) {
      memberIds.push(idx.memberKeysByCustomer[custId][mk]);
      return;
    }
    idx.memberSeq++;
    const mid = 'M-' + pad_(idx.memberSeq, 5);
    plan.members.push({ id: mid, customerId: custId, name: ch.name, kana: ch.kana, gender: ch.gender, birth: ch.birth });
    idx.memberKeysByCustomer[custId][mk] = mid;
    if (birthIso) idx.byChildKey[norm_(ch.name) + '|' + birthIso] = custId;
    memberIds.push(mid);
    newMembersThisRow.push(mid);
  });

  // 撮影(新規カルテはアクティブな状態で起票。移行時の「過去案件ルール」は使わない)
  const dstr = row.date ? Utilities.formatDate(row.date, 'Asia/Tokyo', 'yyyyMMdd') : '00000000';
  idx.shootSeqByDate[dstr] = (idx.shootSeqByDate[dstr] || 0) + 1;
  const shootId = 'S-' + dstr + '-' + pad_(idx.shootSeqByDate[dstr], 2);

  var shootNeedsReview = !row.date;
  var orderType = '', orderStatus = '';
  if (!row.genre) {
    shootNeedsReview = true; // ジャンル不明: 注文を作らずShoots.要確認=TRUEのみ
  } else {
    const isAdult = ADULT_GENRE_KEYWORDS.some(function (kw) { return row.genre.indexOf(kw) >= 0; });
    if (isAdult) { orderType = '成人商品(セレクト前)'; orderStatus = '店頭セレクト待ち'; }
    else { orderType = '撮影データ'; orderStatus = 'データ納品待ち'; }
  }

  plan.shoots.push({
    id: shootId, customerId: custId, date: row.date, genre: row.genre,
    hp: row.hp || '未確認', amount: row.amount, payment: row.payment,
    memberIds: memberIds.join(','), needsReview: shootNeedsReview, remarks: row.remarks
  });

  var newOrderId = '';
  if (orderType) {
    idx.orderSeq++;
    newOrderId = 'O-' + pad_(idx.orderSeq, 5);
    plan.orders.push({
      id: newOrderId, shootId: shootId, type: orderType, status: orderStatus,
      finish: row.finish, memo: '同期: 新規カルテとして起票(大本行' + row.srcRow + ')'
    });
  }

  // Event_Log(経路=同期)
  const now = new Date();
  function addEvent(target, item) {
    idx.eventSeq++;
    plan.events.push([idx.eventSeq, now, 'システム', target, item, '', '生成', '同期']);
  }
  if (isNewCustomer) addEvent(custId, '同期生成' + (customerNeedsReview ? '(同名要確認)' : ''));
  newMembersThisRow.forEach(function (mid) { addEvent(mid, '同期生成'); });
  addEvent(shootId, '同期生成(大本行' + row.srcRow + ')' + (shootNeedsReview ? '(要確認)' : ''));
  if (newOrderId) addEvent(newOrderId, '同期生成(status=' + orderStatus + ')');

  return {
    srcRow: row.srcRow, name: row.name,
    date: row.date ? isoDate_(row.date) : '(日付不明)', genre: row.genre || '(不明)',
    customerId: custId, isNewCustomer: isNewCustomer, customerNeedsReview: customerNeedsReview,
    shootId: shootId, shootNeedsReview: shootNeedsReview,
    orderId: newOrderId, orderStatus: orderStatus
  };
}

// ===== 同期レポート =====

function writeSyncReport_(targetSs, results, totalNew, errorMsg, dryRun) {
  var sh = targetSs.getSheetByName('同期レポート');
  if (sh) targetSs.deleteSheet(sh);
  sh = targetSs.insertSheet('同期レポート');

  const rows = [
    ['同期レポート' + (dryRun ? '(dryRun・書き込みなし)' : '(本実行)'), ''],
    ['実行日時', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm')],
    ['差分件数', totalNew],
    ['', '']
  ];

  if (errorMsg) {
    rows.push(['■ 停止', errorMsg]);
  } else if (!results.length) {
    rows.push(['■ 結果', '新規行はありませんでした。']);
  } else {
    const newCustCount = results.filter(function (r) { return r.isNewCustomer; }).length;
    const reviewCustCount = results.filter(function (r) { return r.customerNeedsReview; }).length;
    const reviewShootCount = results.filter(function (r) { return r.shootNeedsReview; }).length;
    rows.push(['■ 内訳', '']);
    rows.push(['新規顧客', newCustCount]);
    rows.push(['既存顧客への追加撮影', results.length - newCustCount]);
    rows.push(['要確認(顧客・同名)', reviewCustCount]);
    rows.push(['要確認(撮影・日付/ジャンル不明)', reviewShootCount]);
    rows.push(['', '']);
    rows.push(['■ 明細', '']);
    results.forEach(function (r) {
      rows.push(['大本行' + r.srcRow,
        r.date + ' ' + r.name + ' ' + r.genre + ' → ' +
        r.customerId + (r.isNewCustomer ? '(新規)' : '') + ' / ' +
        r.shootId + (r.shootNeedsReview ? '(要確認)' : '') + ' / ' +
        (r.orderId ? r.orderId + '(' + r.orderStatus + ')' : '(注文なし)')
      ]);
    });
  }

  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1).setFontWeight('bold');
  sh.autoResizeColumns(1, 2);
}
