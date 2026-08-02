/**
 * 朝の更新(W4) — Google Calendarから今日・明日の予定を読み、
 * Today_Shoots / Today_Shoot_Matches / Today_Task_Board を再生成する。
 * 設計書: WEB_SPEC_v1.0.md §4 / §8
 *
 * 使い方:
 *   installMorningTrigger() を一度実行 → 毎朝6:30に syncRun→morningUpdate の順で自動実行
 *   画面の[朝の更新]ボタンからも morningUpdate() を手動実行できる
 *
 * 全再生成・冪等。Calendar取得等で失敗した場合は既存のTodayシートを書き換えず、
 * エラー内容をConfig(区分=morning)に記録して画面上部に表示する(getToday経由)。
 *
 * readSheetObjects_ は 05_sync.js、appendEvent_ は 11_web_write_api.js、
 * isoDate_ / pad_ は 02_migrate.js のものをそれぞれ再利用する(再定義しない)。
 */

// ===== 入口 =====

function morningUpdate() {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  try {
    const events = fetchCalendarEvents_();
    const seqByDate = {};
    const todayRows = events.map(function (ev) {
      const start = ev.getStartTime();
      const dstr = Utilities.formatDate(start, 'Asia/Tokyo', 'yyyyMMdd');
      seqByDate[dstr] = (seqByDate[dstr] || 0) + 1;
      return buildTodayShootRow_(ev, dstr, seqByDate[dstr]);
    });
    const matches = buildMatches_(ss, todayRows);

    regenerateTodayShootsSheet_(ss, todayRows);
    regenerateTodayShootMatchesSheet_(ss, matches);
    regenerateTodayTaskBoard_(ss);

    setMorningConfigValue_(ss, 'LAST_ERROR', '');
    setMorningConfigValue_(ss, 'LAST_RUN', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'));
    Logger.log('朝の更新完了: カレンダー' + todayRows.length + '件、照合候補' + matches.length + '件');
    return { ok: true, count: todayRows.length, matchCount: matches.length };
  } catch (e) {
    // 失敗時は既存のTodayシートを書き換えていないため、前回の表示がそのまま残る。
    setMorningConfigValue_(ss, 'LAST_ERROR', String((e && e.message) || e));
    Logger.log('朝の更新エラー: ' + e + '(前回の表示を維持します)');
    throw e;
  }
}

// ===== トリガー =====

function installMorningTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runMorningRoutine') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runMorningRoutine')
    .timeBased().atHour(6).nearMinute(30).everyDays(1).create();
  Logger.log('朝の自動更新トリガーを設置しました(毎日6:30、sync→morningUpdateの順)。');
}

function runMorningRoutine() {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  var syncOk = true, syncMsg = '';
  try {
    syncRun();
  } catch (e) {
    syncOk = false;
    syncMsg = String((e && e.message) || e);
  }

  var morningOk = true, morningMsg = '', morningCount = 0;
  try {
    const res = morningUpdate();
    morningCount = res.count;
  } catch (e) {
    morningOk = false;
    morningMsg = String((e && e.message) || e);
  }

  const summary = 'sync=' + (syncOk ? 'OK' : ('失敗: ' + syncMsg)) +
    ' / morningUpdate=' + (morningOk ? ('OK(' + morningCount + '件)') : ('失敗: ' + morningMsg));
  appendEvent_(ss, 'runMorningRoutine', '朝ルーチン', '', summary, 'トリガー');
  Logger.log(summary);
}

// ===== Calendar取得・解析 =====

function fetchCalendarEvents_() {
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!cal) throw new Error('カレンダーが見つかりません: ' + CALENDAR_ID);
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const end = new Date(start.getTime() + 2 * 86400000); // 当日+翌日
  return cal.getEvents(start, end);
}

function buildTodayShootRow_(ev, dstr, seq) {
  const title = ev.getTitle();
  const parsedTitle = parseEventTitle_(title);
  const desc = parseEventDescription_(ev.getDescription());
  const start = ev.getStartTime();
  return {
    id: 'T-' + dstr + '-' + pad_(seq, 2),
    time: Utilities.formatDate(start, 'Asia/Tokyo', 'HH:mm'),
    date: start,
    genre: parsedTitle ? parsedTitle.genre : '',
    lastName: parsedTitle ? parsedTitle.lastName : '',
    people: desc.people,
    ageGender: desc.ageGender,
    memo: desc.memo,
    phone: desc.phone,
    title: title
  };
}

// オンライン予約書式: 「●ジャンル プラン for 姓」。解析できなければnull(タイトルのまま表示)。
function parseEventTitle_(title) {
  const m = String(title || '').match(/^[●○]?\s*(\S+)[\s\S]*?\bfor\s+(\S+)\s*$/);
  if (!m) return null;
  return { genre: m[1], lastName: m[2] };
}

// 説明欄: ラベル付き行(phone number/ご来店人数/年齢・性別)+自由メモ行
function parseEventDescription_(desc) {
  const text = String(desc || '');
  if (!text) return { phone: '', people: '', ageGender: '', memo: '' };
  const lines = text.split(/\r?\n/);
  var phone = '', people = '', ageGender = '';
  const memoLines = [];
  const labelPatterns = [
    { key: 'phone', re: /Client'?s?\s*phone\s*number\s*[:：]\s*(.+)/i },
    { key: 'people', re: /ご来店人数\s*[:：]\s*(.+)/ },
    { key: 'ageGender', re: /主役のお子様の年齢[・:：]?\s*性別\s*[:：]\s*(.+)/ },
    { key: 'skip', re: /Client'?s?\s*email\s*[:：]\s*(.+)/i } // メールは今のところ使わないが行としては消費する
  ];
  lines.forEach(function (line) {
    var matched = false;
    labelPatterns.forEach(function (p) {
      const m = line.match(p.re);
      if (!m) return;
      matched = true;
      if (p.key === 'phone') phone = normalizePhone_(m[1].trim());
      if (p.key === 'people') people = m[1].trim();
      if (p.key === 'ageGender') ageGender = m[1].trim();
    });
    if (!matched && line.trim()) memoLines.push(line.trim());
  });
  return { phone: phone, people: people, ageGender: ageGender, memo: memoLines.join(' / ') };
}

// +81→0への正規化
function normalizePhone_(s) {
  var v = String(s || '').trim();
  v = v.replace(/^\+81[-\s]?/, '0');
  v = v.replace(/[^\d\-]/g, '');
  return v;
}

// ===== 顧客照合(姓+電話番号。候補提示のみ、自動紐付けしない) =====

function buildMatches_(ss, todayRows) {
  const customers = readSheetObjects_(ss, 'Customers'); // 05_sync.js
  const shoots = readSheetObjects_(ss, 'Shoots');

  const lastShootByCustomer = {};
  shoots.forEach(function (s) {
    if (!(s.撮影日 instanceof Date)) return;
    const cur = lastShootByCustomer[s.customerId];
    if (!cur || s.撮影日.getTime() > cur.撮影日.getTime()) lastShootByCustomer[s.customerId] = s;
  });

  const matches = [];
  todayRows.forEach(function (row) {
    if (!row.lastName && !row.phone) return;
    customers.forEach(function (c) {
      var reason = '';
      if (row.phone && c.連絡先 && normalizePhone_(c.連絡先) === row.phone) reason = '電話番号一致';
      else if (row.lastName && c.顧客名 && String(c.顧客名).indexOf(row.lastName) === 0) reason = '姓一致';
      if (!reason) return;
      const lastShoot = lastShootByCustomer[c.customerId];
      matches.push({
        todayShootId: row.id, customerId: c.customerId, customerName: c.顧客名,
        reason: reason, lastShootId: lastShoot ? lastShoot.shootId : ''
      });
    });
  });
  return matches;
}

// ===== Today_Shoots / Today_Shoot_Matches / Today_Task_Board の再生成 =====

const TODAY_SHOOTS_HEADERS_ = ['id', '時刻', 'ジャンル', '姓', '人数', '年齢性別', '自由メモ', '電話番号', 'タイトル', '日付'];

function getOrCreateTodayShootsSheet_(ss) {
  var sh = ss.getSheetByName('Today_Shoots');
  if (sh) return sh;
  sh = ss.insertSheet('Today_Shoots');
  sh.getRange(1, 1, 1, TODAY_SHOOTS_HEADERS_.length).setValues([TODAY_SHOOTS_HEADERS_]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

function regenerateTodayShootsSheet_(ss, rows) {
  const sh = getOrCreateTodayShootsSheet_(ss);
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
  if (!rows.length) return;
  const values = rows.map(function (r) {
    return [r.id, r.time, r.genre, r.lastName, r.people, r.ageGender, r.memo, r.phone, r.title, r.date];
  });
  sh.getRange(2, 1, values.length, values[0].length).setValues(values);
}

const TODAY_MATCHES_HEADERS_ = ['todayShootId', 'customerId', '顧客名', '一致理由', '直近shootId'];

function getOrCreateTodayMatchesSheet_(ss) {
  var sh = ss.getSheetByName('Today_Shoot_Matches');
  if (sh) return sh;
  sh = ss.insertSheet('Today_Shoot_Matches');
  sh.getRange(1, 1, 1, TODAY_MATCHES_HEADERS_.length).setValues([TODAY_MATCHES_HEADERS_]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

function regenerateTodayShootMatchesSheet_(ss, matches) {
  const sh = getOrCreateTodayMatchesSheet_(ss);
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
  if (!matches.length) return;
  const values = matches.map(function (m) {
    return [m.todayShootId, m.customerId, m.customerName, m.reason, m.lastShootId || ''];
  });
  sh.getRange(2, 1, values.length, values[0].length).setValues(values);
}

// Today_Task_Board再生成(設計書_v1.0.md §2.7)
const TASK_BOARD_ACTIVE_STATUS_ = ['データ納品待ち', '店頭セレクト待ち', '発注待ち', '納品連絡待ち', '引渡し待ち'];

function regenerateTodayTaskBoard_(ss) {
  const sh = ss.getSheetByName('Today_Task_Board');
  if (!sh) return; // ビューが無くても致命的にはしない

  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();

  const orders = readSheetObjects_(ss, 'Orders');
  const shoots = readSheetObjects_(ss, 'Shoots');
  const customers = readSheetObjects_(ss, 'Customers');
  const shootById = {}; shoots.forEach(function (s) { shootById[s.shootId] = s; });
  const custById = {}; customers.forEach(function (c) { custById[c.customerId] = c; });
  const actionText = readNextActionMap_(ss);
  const todayStr = isoDate_(new Date());

  const rows = [];
  orders.forEach(function (o) {
    if (TASK_BOARD_ACTIVE_STATUS_.indexOf(o.status) < 0) return;
    const shoot = shootById[o.shootId];
    const cust = shoot ? custById[shoot.customerId] : null;
    rows.push(['注文', o.orderId, cust ? cust.顧客名 : '', o.注文種別,
      actionText[o.status] || '', o.期限 ? isoDate_(o.期限) : '', o.担当 || '', o.status]);
  });
  shoots.forEach(function (s) {
    if (!(s.撮影日 instanceof Date) || isoDate_(s.撮影日) !== todayStr) return;
    const cust = custById[s.customerId];
    rows.push(['本日撮影', s.shootId, cust ? cust.顧客名 : '', s.ジャンル || '', '', '', '', '']);
  });
  shoots.forEach(function (s) {
    if (!s.要確認) return;
    const cust = custById[s.customerId];
    rows.push(['要確認(撮影)', s.shootId, cust ? cust.顧客名 : '', s.ジャンル || '', '内容を確認する', '', '', '要確認']);
  });
  customers.forEach(function (c) {
    if (!c.要確認) return;
    rows.push(['要確認(顧客)', c.customerId, c.顧客名, '', '内容を確認する', '', '', '要確認']);
  });

  if (rows.length) sh.getRange(2, 1, rows.length, 8).setValues(rows);
}

function readNextActionMap_(ss) {
  const sh = ss.getSheetByName('Config');
  const values = sh.getDataRange().getValues();
  const map = {};
  for (var r = 1; r < values.length; r++) {
    if (values[r][0] === '次の行動') map[values[r][1]] = values[r][2];
  }
  return map;
}

// ===== Configの朝の更新状態(区分=morning)の読み書き =====

function getMorningConfigValue_(ss, key) {
  const sh = ss.getSheetByName('Config');
  const values = sh.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    if (values[r][0] === 'morning' && values[r][1] === key) return values[r][2];
  }
  return '';
}
function setMorningConfigValue_(ss, key, value) {
  const sh = ss.getSheetByName('Config');
  const values = sh.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    if (values[r][0] === 'morning' && values[r][1] === key) {
      sh.getRange(r + 1, 3).setValue(value);
      return;
    }
  }
  sh.appendRow(['morning', key, value, '', '']);
}
