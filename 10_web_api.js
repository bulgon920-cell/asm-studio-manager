/**
 * Web Dashboard API — Phase W1(読み取りのみ)
 * 設計書: WEB_SPEC_v1.0.md §1.1 / §1.2 / §3
 *
 * このファイルにあるのは読み取りAPIのみ。書き込み(updateStatus等)はW2以降。
 */

function doGet() {
  return HtmlService.createTemplateFromFile('WebApp')
    .evaluate()
    .setTitle('AI Studio Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
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

  const customerById = indexBy_(customers, 'customerId');
  const shootById = indexBy_(shoots, 'shootId');
  const todayStr = isoDate_(new Date());

  function customerNameOfShoot_(shootId) {
    const shoot = shootById[shootId];
    const cust = shoot ? customerById[shoot.customerId] : null;
    return cust ? cust.顧客名 : '(不明な顧客)';
  }

  const todo = [];
  const waiting = [];
  const needsReview = [];

  orders.forEach(function (o) {
    const shoot = shootById[o.shootId] || null;
    const custName = customerNameOfShoot_(o.shootId);
    const item = {
      orderId: o.orderId,
      shootId: o.shootId,
      customerName: custName,
      orderType: o.注文種別,
      status: o.status,
      dueDate: isoDate_(o.期限),
      finishDate: isoDate_(o.仕上がり予定日),
      shootDate: shoot ? isoDate_(shoot.撮影日) : ''
    };

    if (o.status === '要確認') {
      needsReview.push({
        kind: 'order', shootId: o.shootId,
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
        kind: 'shoot', shootId: s.shootId,
        label: customerById[s.customerId] ? customerById[s.customerId].顧客名 + '様 撮影情報の確認(' + (s.ジャンル || '') + ')'
          : '撮影情報の確認(' + (s.ジャンル || '') + ')'
      });
    }
  });

  customers.forEach(function (c) {
    if (c.要確認) {
      needsReview.push({
        kind: 'customer', shootId: '',
        label: c.顧客名 + '様 同名顧客の確認'
      });
    }
  });

  sortByDue_(todo);
  sortByDue_(waiting);

  return {
    generatedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'),
    todayShoots: readTodayShoots_(ss, customerById),
    needsReview: needsReview,
    todo: todo,
    waiting: waiting
  };
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

// Today_ShootsはmorningUpdate(W4)が生成するビュー。W1時点では存在しない/空のことがある。
function readTodayShoots_(ss, customerById) {
  const sh = ss.getSheetByName('Today_Shoots');
  if (!sh) return [];
  return readObjects_(ss, 'Today_Shoots').map(function (r) {
    const custName = r.顧客名 || (r.customerId && customerById[r.customerId] ? customerById[r.customerId].顧客名 : '');
    return {
      time: r.時刻 || '',
      genre: r.ジャンル || '',
      customerName: custName || '',
      shootId: r.shootId || ''
    };
  });
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
        dueDate: isoDate_(o.期限), owner: o.担当 || '', memo: o.メモ || ''
      };
    });

  const targetMemberIds = String(shoot.対象memberId || '')
    .split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  const targetMembers = members
    .filter(function (m) { return targetMemberIds.indexOf(m.memberId) >= 0; })
    .map(function (m) { return { name: m.名前, birth: isoDate_(m.誕生日) }; });

  return {
    shootId: shoot.shootId,
    date: isoDate_(shoot.撮影日),
    genre: shoot.ジャンル,
    hp: shoot.HP掲載について,
    needsReview: !!shoot.要確認,
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
    orders: orders
  };
}

// ===== 共通ユーティリティ =====
// isoDate_ / norm_ 等の日付・文字列ユーティリティは 02_migrate.js のものを共用する。

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
