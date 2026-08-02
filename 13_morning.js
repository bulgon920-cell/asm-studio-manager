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
    ensureGenreListConfig_(ss);
    pruneOldHiddenEntries_(ss);
    const hidden = readHiddenEvents_(ss);
    const genreList = readGenreList_(ss);

    const events = fetchCalendarEvents_();
    const seqByDate = {};
    const allRows = events.map(function (ev) {
      const start = ev.getStartTime();
      const dstr = Utilities.formatDate(start, 'Asia/Tokyo', 'yyyyMMdd');
      seqByDate[dstr] = (seqByDate[dstr] || 0) + 1;
      return buildTodayShootRow_(ev, dstr, seqByDate[dstr], genreList);
    });
    const todayRows = allRows.filter(function (r) { return !isEventHidden_(hidden, r.eventId, r.dateIso); });
    const matches = buildMatches_(ss, todayRows);

    regenerateTodayShootsSheet_(ss, todayRows);
    regenerateTodayShootMatchesSheet_(ss, matches);
    regenerateTodayTaskBoard_(ss);

    setMorningConfigValue_(ss, 'LAST_ERROR', '');
    setMorningConfigValue_(ss, 'LAST_RUN', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'));
    Logger.log('朝の更新完了: カレンダー' + todayRows.length + '件、照合候補' + matches.length + '件');
    return sanitizeForClient_({ ok: true, count: todayRows.length, matchCount: matches.length });
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

function buildTodayShootRow_(ev, dstr, seq, genreList) {
  const title = ev.getTitle();
  const parsedTitle = parseEventTitle_(title);
  const desc = parseEventDescription_(ev.getDescription());
  const allDay = ev.isAllDayEvent();
  const start = ev.getStartTime();
  return {
    id: 'T-' + dstr + '-' + pad_(seq, 2),
    eventId: ev.getId(),
    // 終日予定は時刻を持たない(all-day予定のgetStartTime()はタイムゾーンの都合で
    // ずれた時刻を返すため、そもそも文字列化しない)
    time: allDay ? '' : Utilities.formatDate(start, 'Asia/Tokyo', 'HH:mm'),
    allDay: allDay,
    dateIso: Utilities.formatDate(start, 'Asia/Tokyo', 'yyyy-MM-dd'),
    category: isShootEvent_(title, genreList) ? 'shoot' : 'other',
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

// 「撮影」判定: タイトルが●/○で始まる、またはConfig(区分=ジャンル)の語をタイトルに含む
function isShootEvent_(title, genreList) {
  const t = String(title || '');
  if (/^[●○]/.test(t)) return true;
  return genreList.some(function (g) { return g && t.indexOf(g) >= 0; });
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

// 列定義はこの1箇所のみ。書き込み側(regenerateTodayShootsSheet_)はこの配列の順序で
// 行を組み立て、読み込み側(10_web_api.jsのreadTodayShoots_)はヘッダー名から列を引く
// (readObjects_)。ヘッダー行と中身が別々に列位置を知っている状態を作らないための対策。
const TODAY_SHOOTS_HEADERS_ = [
  'id', 'eventId', '時刻', 'allDay', 'category', 'ジャンル', '姓', '人数',
  '年齢性別', '自由メモ', '電話番号', 'タイトル', '日付'
];

// row(buildTodayShootRow_が返すオブジェクト)→ TODAY_SHOOTS_HEADERS_の順に並んだ配列
function todayShootRowToArray_(r) {
  const fieldByHeader = {
    id: r.id, eventId: r.eventId, 時刻: r.time, allDay: r.allDay, category: r.category,
    ジャンル: r.genre, 姓: r.lastName, 人数: r.people, 年齢性別: r.ageGender,
    自由メモ: r.memo, 電話番号: r.phone, タイトル: r.title, 日付: r.dateIso
  };
  return TODAY_SHOOTS_HEADERS_.map(function (h) { return fieldByHeader[h]; });
}

function getOrCreateTodayShootsSheet_(ss) {
  var sh = ss.getSheetByName('Today_Shoots');
  if (!sh) {
    sh = ss.insertSheet('Today_Shoots');
    sh.setFrozenRows(1);
  }
  // 再生成のたびに見出し行も書き直す(派生ビューなので安全。列を増やしても
  // 見出しと中身が絶対にずれない)。
  sh.getRange(1, 1, 1, TODAY_SHOOTS_HEADERS_.length).setValues([TODAY_SHOOTS_HEADERS_]).setFontWeight('bold');
  const maxCols = sh.getMaxColumns();
  if (maxCols > TODAY_SHOOTS_HEADERS_.length) {
    sh.getRange(1, TODAY_SHOOTS_HEADERS_.length + 1, 1, maxCols - TODAY_SHOOTS_HEADERS_.length).clearContent();
  }
  return sh;
}

function regenerateTodayShootsSheet_(ss, rows) {
  const sh = getOrCreateTodayShootsSheet_(ss);
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
  if (!rows.length) return;
  // 時刻列は「15:00」のような文字列をSheetsが時刻値へ自動変換してしまうことがあるため、
  // 書き込み前にテキスト書式を明示する(読み出し時にDate化されるのを防ぐ)。
  const timeCol = TODAY_SHOOTS_HEADERS_.indexOf('時刻') + 1;
  sh.getRange(2, timeCol, rows.length, 1).setNumberFormat('@');
  const values = rows.map(todayShootRowToArray_);
  sh.getRange(2, 1, values.length, values[0].length).setValues(values);
}

const TODAY_MATCHES_HEADERS_ = ['todayShootId', 'customerId', '顧客名', '一致理由', '直近shootId'];

function todayMatchRowToArray_(m) {
  const fieldByHeader = {
    todayShootId: m.todayShootId, customerId: m.customerId, 顧客名: m.customerName,
    一致理由: m.reason, 直近shootId: m.lastShootId || ''
  };
  return TODAY_MATCHES_HEADERS_.map(function (h) { return fieldByHeader[h]; });
}

function getOrCreateTodayMatchesSheet_(ss) {
  var sh = ss.getSheetByName('Today_Shoot_Matches');
  if (!sh) {
    sh = ss.insertSheet('Today_Shoot_Matches');
    sh.setFrozenRows(1);
  }
  sh.getRange(1, 1, 1, TODAY_MATCHES_HEADERS_.length).setValues([TODAY_MATCHES_HEADERS_]).setFontWeight('bold');
  const maxCols = sh.getMaxColumns();
  if (maxCols > TODAY_MATCHES_HEADERS_.length) {
    sh.getRange(1, TODAY_MATCHES_HEADERS_.length + 1, 1, maxCols - TODAY_MATCHES_HEADERS_.length).clearContent();
  }
  return sh;
}

function regenerateTodayShootMatchesSheet_(ss, matches) {
  const sh = getOrCreateTodayMatchesSheet_(ss);
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
  if (!matches.length) return;
  const values = matches.map(todayMatchRowToArray_);
  sh.getRange(2, 1, values.length, values[0].length).setValues(values);
}

// Today_Task_Board再生成(設計書_v1.0.md §2.7)
const TASK_BOARD_ACTIVE_STATUS_ = ['データ納品待ち', '店頭セレクト待ち', '発注待ち', '納品連絡待ち', '引渡し待ち'];
const TASK_BOARD_HEADERS_ = ['種別', '対象ID', '顧客名', '内容', '次の行動', '期限', '担当', 'status'];

function taskBoardRowToArray_(r) {
  const fieldByHeader = {
    種別: r.kind, 対象ID: r.targetId, 顧客名: r.customerName, 内容: r.content,
    次の行動: r.nextAction, 期限: r.due, 担当: r.owner, status: r.status
  };
  return TASK_BOARD_HEADERS_.map(function (h) { return fieldByHeader[h]; });
}

function regenerateTodayTaskBoard_(ss) {
  const sh = ss.getSheetByName('Today_Task_Board');
  if (!sh) return; // ビューが無くても致命的にはしない

  // 再生成のたびに見出し行も書き直す(他の派生ビューと同じ一元化)。
  sh.getRange(1, 1, 1, TASK_BOARD_HEADERS_.length).setValues([TASK_BOARD_HEADERS_]).setFontWeight('bold');
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
    rows.push({
      kind: '注文', targetId: o.orderId, customerName: cust ? cust.顧客名 : '', content: o.注文種別,
      nextAction: actionText[o.status] || '', due: o.期限 ? isoDate_(o.期限) : '', owner: o.担当 || '', status: o.status
    });
  });
  shoots.forEach(function (s) {
    if (!(s.撮影日 instanceof Date) || isoDate_(s.撮影日) !== todayStr) return;
    const cust = custById[s.customerId];
    rows.push({
      kind: '本日撮影', targetId: s.shootId, customerName: cust ? cust.顧客名 : '', content: s.ジャンル || '',
      nextAction: '', due: '', owner: '', status: ''
    });
  });
  shoots.forEach(function (s) {
    if (!s.要確認) return;
    const cust = custById[s.customerId];
    rows.push({
      kind: '要確認(撮影)', targetId: s.shootId, customerName: cust ? cust.顧客名 : '', content: s.ジャンル || '',
      nextAction: '内容を確認する', due: '', owner: '', status: '要確認'
    });
  });
  customers.forEach(function (c) {
    if (!c.要確認) return;
    rows.push({
      kind: '要確認(顧客)', targetId: c.customerId, customerName: c.顧客名, content: '',
      nextAction: '内容を確認する', due: '', owner: '', status: '要確認'
    });
  });

  if (rows.length) {
    sh.getRange(2, 1, rows.length, TASK_BOARD_HEADERS_.length).setValues(rows.map(taskBoardRowToArray_));
  }
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

// ===== ジャンル一覧(区分=ジャンル)。「撮影」判定に使う(§1.1) =====

const GENRE_LIST_SEED_ = [
  '七五三', '成人式', '成人記念', '振袖', '二十歳', 'はたち', 'お宮参り',
  'ハーフバースデー', 'バースデー', '卒業', '入学', '家族写真', '記念撮影'
];

function ensureGenreListConfig_(ss) {
  const sh = ss.getSheetByName('Config');
  const values = sh.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    if (values[r][0] === 'ジャンル') return; // 既に1件でもあれば初期化済み
  }
  const rows = GENRE_LIST_SEED_.map(function (g) { return ['ジャンル', g, '', '', '']; });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
  Logger.log('Configにジャンル一覧の初期値を' + rows.length + '件登録しました。');
}

function readGenreList_(ss) {
  const sh = ss.getSheetByName('Config');
  const values = sh.getDataRange().getValues();
  const list = [];
  for (var r = 1; r < values.length; r++) {
    if (values[r][0] === 'ジャンル' && values[r][1]) list.push(values[r][1]);
  }
  return list;
}

// ===== 非表示にした予定(区分=非表示予定)。今日の予定カードの×/戻す(§1.1) =====

const HIDDEN_EVENT_EXPIRE_DAYS_ = 15;

function hideEvent(eventId, date) {
  if (!eventId || !date) throw new Error('eventIdとdateは必須です。');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    throw new Error('他の操作が実行中のため待機できませんでした。しばらくして再実行してください。');
  }
  try {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    addHiddenEvent_(ss, eventId, date);
    appendEvent_(ss, eventId, '予定を非表示', '', date, 'Web(非表示)');
    return sanitizeForClient_({ eventId: eventId, date: date });
  } finally {
    lock.releaseLock();
  }
}

function unhideEvent(eventId, date) {
  if (!eventId || !date) throw new Error('eventIdとdateは必須です。');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    throw new Error('他の操作が実行中のため待機できませんでした。しばらくして再実行してください。');
  }
  try {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    removeHiddenEvent_(ss, eventId, date);
    appendEvent_(ss, eventId, '予定の非表示を解除', date, '', 'Web(非表示)');
    return sanitizeForClient_({ eventId: eventId, date: date });
  } finally {
    lock.releaseLock();
  }
}

function hiddenEventKey_(eventId, date) { return eventId + '|' + date; }

function addHiddenEvent_(ss, eventId, date) {
  const sh = ss.getSheetByName('Config');
  const key = hiddenEventKey_(eventId, date);
  const values = sh.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    if (values[r][0] === '非表示予定' && values[r][1] === key) return; // 既に登録済み(冪等)
  }
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  sh.appendRow(['非表示予定', key, now, date, eventId]);
}

function removeHiddenEvent_(ss, eventId, date) {
  const sh = ss.getSheetByName('Config');
  const key = hiddenEventKey_(eventId, date);
  const values = sh.getDataRange().getValues();
  for (var r = values.length - 1; r >= 1; r--) {
    if (values[r][0] === '非表示予定' && values[r][1] === key) {
      sh.deleteRow(r + 1);
      return;
    }
  }
}

// 非表示リストを読む。{eventId, date, registeredAt} の配列。
function readHiddenEvents_(ss) {
  const sh = ss.getSheetByName('Config');
  const values = sh.getDataRange().getValues();
  const list = [];
  for (var r = 1; r < values.length; r++) {
    if (values[r][0] === '非表示予定') {
      list.push({ registeredAt: values[r][2], date: values[r][3], eventId: values[r][4] });
    }
  }
  return list;
}

function isEventHidden_(hiddenList, eventId, date) {
  return hiddenList.some(function (h) { return h.eventId === eventId && h.date === date; });
}

// 登録から15日を超えた非表示エントリを削除する(Configの肥大防止。morningUpdateから呼ぶ)
function pruneOldHiddenEntries_(ss) {
  const sh = ss.getSheetByName('Config');
  const values = sh.getDataRange().getValues();
  const today = new Date();
  const rowsToDelete = [];
  for (var r = 1; r < values.length; r++) {
    if (values[r][0] !== '非表示予定') continue;
    const registeredAt = toDate_(values[r][2]); // 02_migrate.js
    if (!registeredAt) continue;
    if (daysBetween_(registeredAt, today) > HIDDEN_EVENT_EXPIRE_DAYS_) rowsToDelete.push(r + 1);
  }
  rowsToDelete.sort(function (a, b) { return b - a; }); // 下から消す
  rowsToDelete.forEach(function (rowIndex) { sh.deleteRow(rowIndex); });
  if (rowsToDelete.length) Logger.log('期限切れの非表示予定を' + rowsToDelete.length + '件削除しました。');
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
