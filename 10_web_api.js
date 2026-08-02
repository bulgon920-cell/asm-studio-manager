/**
 * Web Dashboard API — 読み取り専用
 * 設計書: WEB_SPEC_v1.0.md §1.1 / §1.2 / §3
 *
 * このファイルにあるのは読み取りAPIのみ(副作用なし)。
 * 書き込みAPI(updateStatus/updateOrderFields等)は 11_web_write_api.js。
 */

function doGet() {
  return HtmlService.createTemplateFromFile('WebApp')
    .evaluate()
    .setTitle('AI Studio Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// google.script.runは戻り値にDate型が混じっていると応答そのものをnullにすることがある。
// 返す直前にJSON往復させ、Dateを文字列化(等)して安全なプレーンオブジェクトにする。
// 全API(読み取り・書き込み)共通で使う。
function sanitizeForClient_(payload) {
  return JSON.parse(JSON.stringify(payload));
}

// スクリプトエディタから直接実行して getToday() の中身を確認するための手動テスト用関数。
// Webアプリを経由せずに、今日の予定・要確認・やること・待ちの件数と内容をログへ出す。
function testGetToday() {
  const result = getToday();
  Logger.log('=== getToday() 結果 ===');
  Logger.log('todayShoots: %s件', result.todayShoots.length);
  Logger.log(JSON.stringify(result.todayShoots, null, 2));
  Logger.log('needsReview: %s件 / todo: %s件 / waiting: %s件 / hiddenEvents: %s件',
    result.needsReview.length, result.todo.length, result.waiting.length,
    (result.hiddenEvents || []).length);
  if (result.morningError) Logger.log('morningError: ' + result.morningError);
}

// ===== 今日の画面 =====

const ACTIVE_ORDER_STATUS_ = [
  'データ納品待ち', '店頭セレクト待ち', '発注待ち', '納品連絡待ち', '引渡し待ち'
];

function getToday() {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const customers = readObjects_(ss, 'Customers');
  const shoots = readObjects_(ss, 'Shoots');
  const orders = readObjects_(ss, 'Orders');
  const orderCategory = readOrderTypeCategory_(ss);

  const customerById = indexBy_(customers, 'customerId');
  const shootById = indexBy_(shoots, 'shootId');
  const today = new Date();
  const todayStr = isoDate_(today);

  function customerNameOfShoot_(shootId) {
    const shoot = shootById[shootId];
    const cust = shoot ? customerById[shoot.customerId] : null;
    return cust ? cust.顧客名 : '(不明な顧客)';
  }
  // 商品系・作業系は「プリント・商品(外注)」ブロック、データ系は「データ」ブロック
  function categoryOf_(orderType) {
    return orderCategory[orderType] === 'データ系' ? 'data' : 'goods';
  }
  function elapsedDaysOf_(shoot) {
    if (!shoot || !(shoot.撮影日 instanceof Date)) return null;
    return daysBetween_(shoot.撮影日, today) + 1; // 撮影当日を1日目とする
  }

  const todo = [];
  const waiting = [];
  const needsReview = [];

  orders.forEach(function (o) {
    const shoot = shootById[o.shootId] || null;
    const custName = customerNameOfShoot_(o.shootId);
    const elapsedDays = elapsedDaysOf_(shoot);
    const item = {
      orderId: o.orderId,
      shootId: o.shootId,
      customerName: custName,
      orderType: o.注文種別,
      status: o.status,
      dueDate: isoDate_(o.期限),
      finishDate: isoDate_(o.仕上がり予定日),
      shootDate: shoot ? isoDate_(shoot.撮影日) : '',
      category: categoryOf_(o.注文種別),
      elapsedDays: elapsedDays
    };

    if (o.status === '要確認') {
      needsReview.push({
        kind: 'order', shootId: o.shootId, category: item.category, elapsedDays: elapsedDays,
        shootDate: item.shootDate,
        label: custName + '様 実状態を確定する'
      });
    } else if (ACTIVE_ORDER_STATUS_.indexOf(o.status) >= 0) {
      todo.push(item);
    } else if (o.status === '仕上がり待ち') {
      const overdue = item.finishDate && item.finishDate < todayStr;
      (overdue ? todo : waiting).push(item);
    } else if (o.status === '保留') {
      waiting.push(item);
    }
    // 完了・完了(移行時推定)・不要は今日の画面に出さない
  });

  shoots.forEach(function (s) {
    if (s.要確認) {
      needsReview.push({
        kind: 'shoot', shootId: s.shootId, category: 'other', elapsedDays: elapsedDaysOf_(s),
        shootDate: isoDate_(s.撮影日),
        label: customerById[s.customerId] ? customerById[s.customerId].顧客名 + '様 撮影情報の確認(' + (s.ジャンル || '') + ')'
          : '撮影情報の確認(' + (s.ジャンル || '') + ')'
      });
    }
  });

  customers.forEach(function (c) {
    if (c.要確認) {
      needsReview.push({
        kind: 'customer', shootId: '', customerId: c.customerId, category: 'other', elapsedDays: null,
        label: c.顧客名 + '様 同名顧客の確認'
      });
    }
  });

  sortByElapsedDesc_(todo);
  sortByElapsedDesc_(needsReview);
  sortByDue_(waiting);

  return sanitizeForClient_({
    generatedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'),
    todayShoots: readTodayShoots_(ss),
    // 朝の更新(13_morning.js)が失敗した場合のエラーを画面上部に表示するため(W4)
    morningError: getMorningConfigValue_(ss, 'LAST_ERROR') || '',
    // 「非表示にした予定」の折りたたみ表示用(13_morning.js)
    hiddenEvents: readHiddenEvents_(ss),
    needsReview: needsReview,
    todo: todo,
    waiting: waiting
  });
}

// Configの注文種別マスタ(区分=注文種別)から 種別名→系統 を引く
function readOrderTypeCategory_(ss) {
  const sh = ss.getSheetByName('Config');
  const values = sh.getDataRange().getValues();
  const map = {};
  for (var r = 1; r < values.length; r++) {
    if (values[r][0] === '注文種別') map[values[r][1]] = values[r][2];
  }
  return map;
}

function sortByDue_(list) {
  list.sort(function (a, b) {
    if (a.dueDate && b.dueDate) return a.dueDate < b.dueDate ? -1 : (a.dueDate > b.dueDate ? 1 : 0);
    if (a.dueDate && !b.dueDate) return -1;
    if (!a.dueDate && b.dueDate) return 1;
    const sa = a.shootDate || '9999-99-99';
    const sb = b.shootDate || '9999-99-99';
    return sa < sb ? -1 : (sa > sb ? 1 : 0);
  });
}

// 経過日数の降順(古い撮影ほど上)。経過日数不明は最後尾。
function sortByElapsedDesc_(list) {
  list.sort(function (a, b) {
    const ea = a.elapsedDays == null ? -1 : a.elapsedDays;
    const eb = b.elapsedDays == null ? -1 : b.elapsedDays;
    return eb - ea;
  });
}

// Today_Shoots/Today_Shoot_Matchesは13_morning.jsのmorningUpdate()が生成するビュー。
// 未実行時は存在しない/空のことがある。
// 非表示リスト(Config 区分=非表示予定)は再生成時に13_morning.js側で除外済みだが、
// 念のためここでも二重にフィルタする(読み取り時点の安全網)。
function readTodayShoots_(ss) {
  const sh = ss.getSheetByName('Today_Shoots');
  if (!sh) return [];
  const matchesById = {};
  readObjects_(ss, 'Today_Shoot_Matches').forEach(function (m) {
    if (!matchesById[m.todayShootId]) matchesById[m.todayShootId] = [];
    matchesById[m.todayShootId].push(m);
  });
  const hidden = readHiddenEvents_(ss); // 13_morning.js
  // Sheetsが「2026-08-02」「15:00」等の文字列を日付/時刻値へ自動変換してしまうことがある
  // (書き込み側でテキスト書式は指定済みだが、念のため読み取り側でも二重に安全化する)。
  function safeDateStr_(v) { return (v instanceof Date) ? isoDate_(v) : (v || ''); }
  function safeTimeStr_(v) { return (v instanceof Date) ? Utilities.formatDate(v, 'Asia/Tokyo', 'HH:mm') : (v || ''); }

  return readObjects_(ss, 'Today_Shoots')
    .map(function (r) {
      const cand = (matchesById[r.id] || [])[0]; // 最有力候補のみ表示(確定は人)
      return {
        eventId: r.eventId || '',
        date: safeDateStr_(r.日付),
        time: safeTimeStr_(r.時刻),
        allDay: !!r.allDay,
        category: r.category === 'shoot' ? 'shoot' : 'other',
        genre: r.ジャンル || '',
        lastName: r.姓 || '',
        people: r.人数 || '',
        ageGender: r.年齢性別 || '',
        memo: r.自由メモ || '',
        title: r.タイトル || '',
        matchCustomerName: cand ? cand.顧客名 : '',
        matchShootId: cand ? cand.直近shootId : ''
      };
    })
    .filter(function (r) { return !isEventHidden_(hidden, r.eventId, r.date); });
}

// ===== 案件詳細 =====

function getShootDetail(shootId) {
  if (!shootId) throw new Error('shootIdが指定されていません。');
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);

  const shoot = readObjects_(ss, 'Shoots')
    .filter(function (s) { return s.shootId === shootId; })[0];
  if (!shoot) throw new Error('撮影が見つかりません: ' + shootId);

  const customer = readObjects_(ss, 'Customers')
    .filter(function (c) { return c.customerId === shoot.customerId; })[0] || null;
  // 業務ジャンル(広告撮影等)の判定。readBusinessGenreNames_は05_sync.jsのものを再利用(SYNC_SPEC §6)
  const isBusinessShoot = !!(customer && readBusinessGenreNames_(ss).indexOf(customer.顧客名) >= 0);

  const members = readObjects_(ss, 'Family_Members')
    .filter(function (m) { return m.customerId === shoot.customerId; });

  const sameCustomerShoots = readObjects_(ss, 'Shoots')
    .filter(function (s) { return s.customerId === shoot.customerId; })
    .sort(function (a, b) {
      const da = a.撮影日 instanceof Date ? a.撮影日.getTime() : 0;
      const db = b.撮影日 instanceof Date ? b.撮影日.getTime() : 0;
      return db - da;
    });
  const pastShoots = sameCustomerShoots
    .filter(function (s) { return s.shootId !== shootId; })
    .map(function (s) {
      return { shootId: s.shootId, date: isoDate_(s.撮影日), genre: s.ジャンル };
    });

  const orders = readObjects_(ss, 'Orders')
    .filter(function (o) { return o.shootId === shootId; })
    .map(function (o) {
      return {
        orderId: o.orderId, orderType: o.注文種別, status: o.status,
        finishDate: isoDate_(o.仕上がり予定日), selectDate: isoDate_(o.セレクト予定日),
        dueDate: isoDate_(o.期限), owner: o.担当 || '', memo: o.メモ || '',
        nextStatuses: getAllowedNextStatuses_(ss, o.status, o.orderId),
        // 要確認の「実状態を確定」用。通常の遷移制限とは別に、系統内の全statusを出す(W3)。
        reviewOptions: o.status === '要確認' ? getReviewStatusOptions_(ss, o.注文種別) : []
      };
    });

  const targetMemberIds = String(shoot.対象memberId || '')
    .split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  const targetMembers = members
    .filter(function (m) { return targetMemberIds.indexOf(m.memberId) >= 0; })
    .map(function (m) { return { name: m.名前, birth: isoDate_(m.誕生日) }; });

  return sanitizeForClient_({
    shootId: shoot.shootId,
    date: isoDate_(shoot.撮影日),
    genre: shoot.ジャンル,
    hp: shoot.HP掲載について,
    needsReview: !!shoot.要確認,
    isBusiness: isBusinessShoot,
    remarks: shoot.備考 || '',
    driveUrl: shoot.DriveフォルダURL || '',
    albumUrl: shoot.オンラインアルバムURL || '',
    customer: customer ? {
      customerId: customer.customerId, name: customer.顧客名,
      kana: customer.ふりがな, contact: customer.連絡先, addr: customer.住所,
      needsReview: !!customer.要確認
    } : null,
    targetMembers: targetMembers,
    pastShoots: pastShoots,
    orders: orders,
    // 「+ 注文を追加」のドロップダウン用(W3)
    orderTypeOptions: Object.keys(readOrderTypeCategory_(ss))
  });
}

// ===== 状態遷移(表示用・読み取りのみ) =====
// 実際の書き込み(updateStatus)は 11_web_write_api.js。ここでは
// 「画面にどのボタンを出せるか」を判定するためだけに使う(副作用なし)。

const TERMINAL_ORDER_STATUS_ = ['完了', '完了(移行時推定)', '不要'];

function getAllowedNextStatuses_(ss, fromStatus, orderId) {
  if (TERMINAL_ORDER_STATUS_.indexOf(fromStatus) >= 0) return [];
  if (fromStatus === '要確認') return []; // 実状態の確定(resolveReview)はW3
  if (fromStatus === '保留') {
    const prev = findPreviousStatusBeforeHold_(ss, orderId);
    return prev ? [prev] : [];
  }
  const forward = (readTransitionTable_(ss)[fromStatus] || []).slice();
  ['保留', '不要'].forEach(function (s) { if (forward.indexOf(s) < 0) forward.push(s); });
  return forward;
}

// Config(区分=遷移)から 現在status→遷移可能なstatus[] を読む。
// 未設定(setupOrderTransitions()未実行)の場合は空マップを返す(エラーにしない)。
function readTransitionTable_(ss) {
  const sh = ss.getSheetByName('Config');
  const values = sh.getDataRange().getValues();
  const map = {};
  for (var r = 1; r < values.length; r++) {
    if (values[r][0] === '遷移') {
      const from = values[r][1], to = values[r][2];
      if (from && to) {
        if (!map[from]) map[from] = [];
        map[from].push(to);
      }
    }
  }
  return map;
}

// 「保留」に入る直前の状態をEvent_Logから逆引きする(復帰先の判定用)
function findPreviousStatusBeforeHold_(ss, orderId) {
  const sh = ss.getSheetByName('Event_Log');
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return '';
  const values = sh.getRange(2, 1, lastRow - 1, 8).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    const row = values[i]; // [eventId,日時,操作者,対象,項目,変更前,変更後,経路]
    if (row[3] === orderId && row[4] === 'status' && row[6] === '保留' && row[7] === 'Web') {
      return row[5];
    }
  }
  return '';
}

// ===== 共通ユーティリティ =====
// isoDate_ / norm_ / daysBetween_ 等の日付・文字列ユーティリティは 02_migrate.js のものを共用する。

function readObjects_(ss, sheetName) {
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
    if (headers[0] && !String(obj[headers[0]]).trim()) continue; // 主キー空の行はスキップ
    rows.push(obj);
  }
  return rows;
}

function indexBy_(rows, key) {
  const idx = {};
  rows.forEach(function (r) { idx[r[key]] = r; });
  return idx;
}
