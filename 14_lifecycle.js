/**
 * Lifecycle 節目一覧 — 読み取り中心(書き込みは「案内済み」記録の1つだけ)
 * 設計書: LIFECYCLE_SPEC_v1.0.md
 *
 * getLifecycle(): 節目対象を判定して返す読み取りAPI(表示期間はConfigの表示開始値で決まる)
 * markContacted(memberId, 節目名, 対象年): 「案内済み」をLifecycle_Contactsへ記録
 * testLifecycle(): 判定ロジックのダミーデータ検証(スクリプトエディタから手動実行)
 *
 * 節目マスタ(Config 区分=節目)は「提案メモ」が主役(業務側が直接編集する)。
 * 表示開始(値3)は「目安日の何ヶ月前から表示するか」の実効値(数値)としてコードが読む。
 * 年齢・対象年の判定ロジック本体はこのファイルのLIFECYCLE_RULES_で保持し、
 * Config側の型/条件は人が読むための参考表示として扱う(§0-2)。
 *
 * readObjects_ / sanitizeForClient_ は 10_web_api.js、appendEvent_ は
 * 11_web_write_api.js、readBusinessGenreNames_ は 05_sync.js、
 * isoDate_ は 02_migrate.js のものをそれぞれ再利用する(再定義しない)。
 */

// ===== 節目マスタ(Config 区分=節目) =====

const LIFECYCLE_DEFAULT_LEAD_MONTHS_ = 2; // 表示開始の既定値(目安日の何ヶ月前から)

// [節目名, 型, 条件(参考表示), 提案メモ]。表示開始(数値)は別途LIFECYCLE_DEFAULT_LEAD_MONTHS_を使う。
const LIFECYCLE_MILESTONES_SEED_ = [
  ['お宮参り', '誕生日型', '生後0〜3ヶ月', '生後1〜3ヶ月が目安'],
  ['ハーフバースデー', '誕生日型', '生後6ヶ月', '寝返り・ハイハイができると撮影バリエーションが増える'],
  ['1歳バースデー', '誕生日型', '満1歳', 'スマッシュケーキやファーストシューズ提案'],
  ['2歳記念', '誕生日型', '満2歳', '女の子なら次の3歳は七五三へ接続'],
  ['七五三(3歳・女)', '年次型', 'その年に満3歳(女)。目安日11/15', '満年齢推奨。数えなら早生まれは満年齢が1つ上・着物サイズ選びに影響'],
  ['七五三(5歳・男)', '年次型', 'その年に満5歳(男)。目安日11/15', ''],
  ['七五三(7歳・女)', '年次型', 'その年に満7歳(女)。目安日11/15', 'ランドセルと同時撮影プラン'],
  ['ランドセル(入学)', '年次型', '学年境界(4/2〜翌4/1)で満6歳。目安日4/1', '7歳女は七五三と同時プラン'],
  ['ハーフ成人式', '誕生日型', '満10歳', ''],
  ['中学入学', '年次型', '学年境界(4/2〜翌4/1)で満12歳。目安日4/1', ''],
  ['成人式', '誕生日型', '満20歳', '振袖層への提案に特化。振袖在庫4本・自社対応。大量案件は狙わない'],
  ['背の順家族写真', '経過型', '末っ子が満7歳以上かつ最終来店から11ヶ月以上', '背の順に並んで、毎年一枚の家族写真']
];

// 初回・更新時: Configに節目マスタを登録/移行する(冪等)。
// 既存行があれば表示開始(値3)だけを数値へ移行し、提案メモ(値4)は上書きしない。
function ensureLifecycleConfig_(ss) {
  const sh = ss.getSheetByName('Config');
  if (sh.getLastColumn() < 6 || !sh.getRange(1, 6).getValue()) {
    sh.getRange(1, 6).setValue('値4');
  }
  const values = sh.getDataRange().getValues();
  const existingRowIndexByName = {};
  for (var r = 1; r < values.length; r++) {
    if (values[r][0] === '節目') existingRowIndexByName[values[r][1]] = r + 1; // シート上の行番号(1始まり)
  }

  const toAppend = [];
  LIFECYCLE_MILESTONES_SEED_.forEach(function (seed) {
    const name = seed[0], type = seed[1], condition = seed[2], memo = seed[3];
    const rowIndex = existingRowIndexByName[name];
    if (rowIndex) {
      const currentLead = sh.getRange(rowIndex, 5).getValue(); // 値3列(表示開始)
      if (typeof currentLead !== 'number') {
        sh.getRange(rowIndex, 5).setValue(LIFECYCLE_DEFAULT_LEAD_MONTHS_);
      }
      // 型・条件(参考表示)は最新化してよいが、提案メモ(値4)には一切触れない。
      sh.getRange(rowIndex, 3, 1, 2).setValues([[type, condition]]);
    } else {
      toAppend.push(['節目', name, type, condition, LIFECYCLE_DEFAULT_LEAD_MONTHS_, memo]);
    }
  });
  if (toAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, 6).setValues(toAppend);
    Logger.log('Configに節目マスタを' + toAppend.length + '件登録しました。');
  }
}

function readLifecycleMilestones_(ss) {
  const sh = ss.getSheetByName('Config');
  const values = sh.getDataRange().getValues();
  const map = {};
  for (var r = 1; r < values.length; r++) {
    if (values[r][0] === '節目') {
      const leadRaw = values[r][4];
      map[values[r][1]] = {
        type: values[r][2], condition: values[r][3],
        leadMonths: (typeof leadRaw === 'number' && leadRaw >= 0) ? leadRaw : LIFECYCLE_DEFAULT_LEAD_MONTHS_,
        memo: values[r][5] || ''
      };
    }
  }
  return map;
}

// ===== 判定ロジック(型ごと。名称はConfigの節目名と対応させる) =====

const LIFECYCLE_RULES_ = [
  { name: 'お宮参り', kind: 'range', minMonths: 0, maxMonths: 3, genre: 'お宮参り' },
  { name: 'ハーフバースデー', kind: 'point', months: 6, genre: 'ハーフバースデー' },
  { name: '1歳バースデー', kind: 'point', years: 1 },
  { name: '2歳記念', kind: 'point', years: 2 },
  { name: '七五三(3歳・女)', kind: 'yearly', targetAge: 3, gender: '女', genre: '七五三' },
  { name: '七五三(5歳・男)', kind: 'yearly', targetAge: 5, gender: '男', genre: '七五三' },
  { name: '七五三(7歳・女)', kind: 'yearly', targetAge: 7, gender: '女', genre: '七五三' },
  { name: 'ランドセル(入学)', kind: 'school', targetAge: 6 },
  { name: 'ハーフ成人式', kind: 'point', years: 10, genre: 'ハーフ成人式' },
  { name: '中学入学', kind: 'school', targetAge: 12 },
  { name: '成人式', kind: 'point', years: 20, genre: '成人' },
  { name: '背の順家族写真', kind: 'elapsed', minAge: 7, minMonthsSinceVisit: 11 }
];

// 節目の「対象年+目安日」を求める(型ごとに1通りに定まる。表示可否はここでは判定しない)。
function findMilestoneOccurrence_(rule, birth) {
  if (rule.kind === 'range') {
    const start = addMonths_(birth, rule.minMonths);
    const end = addMonths_(birth, rule.maxMonths);
    return { year: end.getFullYear(), dueDate: end, rangeStart: start };
  }
  if (rule.kind === 'point') {
    const target = rule.years != null ? addYears_(birth, rule.years) : addMonths_(birth, rule.months);
    return { year: target.getFullYear(), dueDate: target };
  }
  if (rule.kind === 'yearly') {
    // 「その年に満N歳になる子」は 誕生年 = 対象年 − N で判定する(現在の年齢では判定しない)
    const year = birth.getFullYear() + rule.targetAge;
    return { year: year, dueDate: new Date(year, 10, 15) }; // 目安日: 対象年11/15
  }
  if (rule.kind === 'school') {
    // 学年境界(4/2〜翌4/1)で満targetAge歳になる学年の入学年
    const year = schoolEntranceYear_(birth, rule.targetAge);
    return { year: year, dueDate: new Date(year, 3, 1) }; // 目安日: 入学年4/1
  }
  return null;
}

// 学年区分(コホート年): 4/2〜翌年4/1生まれが同学年。4/2以降生まれはその年+1年の区分に属する。
function schoolCohortYear_(birth) {
  const m = birth.getMonth() + 1, d = birth.getDate();
  const isOnOrAfterCutoff = (m > 4) || (m === 4 && d >= 2);
  return isOnOrAfterCutoff ? birth.getFullYear() + 1 : birth.getFullYear();
}

// 満ageAtEntry歳になる学年の入学年(西暦)
function schoolEntranceYear_(birth, ageAtEntry) {
  return schoolCohortYear_(birth) + ageAtEntry;
}

// 表示可否の判定。range型(お宮参り)は誕生からの自然な期間そのものを表示期間とする。
// それ以外は「目安日のleadMonths ヶ月前」〜「目安日の翌月末」を表示期間とする(Config実効値)。
function isWithinDisplayWindow_(occurrence, leadMonths, today) {
  if (occurrence.rangeStart) {
    return today >= occurrence.rangeStart && today <= occurrence.dueDate;
  }
  const displayStart = addMonths_(occurrence.dueDate, -leadMonths);
  const displayEnd = endOfNextMonth_(occurrence.dueDate);
  return today >= displayStart && today <= displayEnd;
}

// dateが属する月の「翌月末」の日付を返す
function endOfNextMonth_(date) {
  return new Date(date.getFullYear(), date.getMonth() + 2, 0);
}

// ===== 読み取りAPI =====

function getLifecycle() {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const today = new Date();
  return sanitizeForClient_(computeLifecycle_(ss, today));
}

// getLifecycle()とtestLifecycle()から共用する本体。todayを外から注入できるようにして検証可能にする。
function computeLifecycle_(ss, today) {
  ensureLifecycleConfig_(ss);

  const memoByName = readLifecycleMilestones_(ss);
  const members = readObjects_(ss, 'Family_Members');
  const customers = readObjects_(ss, 'Customers');
  const shoots = readObjects_(ss, 'Shoots');
  const custById = {};
  customers.forEach(function (c) { custById[c.customerId] = c; });
  const businessNames = readBusinessGenreNames_(ss); // 05_sync.js
  const contacted = readLifecycleContacts_(ss);

  const shootsByCustomer = {};
  shoots.forEach(function (s) {
    if (!shootsByCustomer[s.customerId]) shootsByCustomer[s.customerId] = [];
    shootsByCustomer[s.customerId].push(s);
  });

  const results = [];
  var noBirthdayCount = 0;

  members.forEach(function (m) {
    const cust = custById[m.customerId];
    if (!cust) return;
    if (businessNames.indexOf(cust.顧客名) >= 0) return; // 広告案件は対象外
    if (!(m.誕生日 instanceof Date)) { noBirthdayCount++; return; }

    const custShoots = shootsByCustomer[m.customerId] || [];
    const lastShoot = latestShoot_(custShoots);

    LIFECYCLE_RULES_.forEach(function (rule) {
      if (rule.kind === 'elapsed') return; // 家族単位。下の別ループで処理
      if (rule.gender && String(m.性別 || '') !== rule.gender) return;

      const occurrence = findMilestoneOccurrence_(rule, m.誕生日);
      if (!occurrence) return;

      const meta = memoByName[rule.name] || {};
      const leadMonths = meta.leadMonths != null ? meta.leadMonths : LIFECYCLE_DEFAULT_LEAD_MONTHS_;
      if (!isWithinDisplayWindow_(occurrence, leadMonths, today)) return;

      // 済みフィルタ: 七五三・お宮参り・ハーフBD・ハーフ成人式・成人式等の一度きりの節目は、
      // 対象年に関わらず該当ジャンルの撮影歴があれば除外する(§2)。
      if (rule.genre && hasGenreShotEver_(custShoots, m.memberId, rule.genre)) return;

      const key = rule.name + '|' + occurrence.year + '|' + m.memberId;
      results.push(buildLifecycleItem_(m.memberId, m.customerId, m.名前, m.誕生日, rule.name, occurrence.year,
        occurrence.dueDate, cust, lastShoot, memoByName, contacted, key));
    });
  });

  // 経過型(背の順家族写真)は家族(顧客)単位・対象年(=今年)で判定する(唯一の毎年型)
  const elapsedRule = LIFECYCLE_RULES_.filter(function (r) { return r.kind === 'elapsed'; })[0];
  if (elapsedRule) {
    customers.forEach(function (cust) {
      if (businessNames.indexOf(cust.顧客名) >= 0) return;
      const famMembers = members.filter(function (m) {
        return m.customerId === cust.customerId && m.誕生日 instanceof Date;
      });
      if (!famMembers.length) return;
      const youngest = famMembers.reduce(function (a, b) {
        return a.誕生日.getTime() > b.誕生日.getTime() ? a : b; // 誕生日が新しい方が年少
      });
      if (ageInYears_(youngest.誕生日, today) < elapsedRule.minAge) return;

      const custShoots = shootsByCustomer[cust.customerId] || [];
      const lastShoot = latestShoot_(custShoots);
      if (!lastShoot || !(lastShoot.撮影日 instanceof Date)) return;
      if (monthsBetween_(lastShoot.撮影日, today) < elapsedRule.minMonthsSinceVisit) return;

      const key = elapsedRule.name + '|' + today.getFullYear() + '|' + youngest.memberId;
      results.push(buildLifecycleItem_(youngest.memberId, cust.customerId, youngest.名前 + '(末っ子)',
        youngest.誕生日, elapsedRule.name, today.getFullYear(), null, cust, lastShoot, memoByName, contacted, key));
    });
  }

  results.sort(function (a, b) {
    const da = a.dueDate || '9999-99-99';
    const db = b.dueDate || '9999-99-99';
    return da < db ? -1 : (da > db ? 1 : 0);
  });

  return { items: results, noBirthdayCount: noBirthdayCount };
}

function buildLifecycleItem_(memberId, customerId, childName, birthDate, milestoneName, targetYear,
  targetDate, cust, lastShoot, memoByName, contacted, key) {
  return {
    memberId: memberId, customerId: customerId, childName: childName,
    birthDate: (birthDate instanceof Date) ? isoDate_(birthDate) : '',
    milestoneName: milestoneName, targetYear: targetYear,
    dueDate: targetDate ? isoDate_(targetDate) : '',
    parentName: cust.顧客名, contact: cust.連絡先 || '', line: !!cust.LINE有無,
    lastVisitDate: (lastShoot && lastShoot.撮影日 instanceof Date) ? isoDate_(lastShoot.撮影日) : '',
    lastVisitGenre: lastShoot ? (lastShoot.ジャンル || '') : '',
    lastShootId: lastShoot ? lastShoot.shootId : '',
    memo: (memoByName[milestoneName] || {}).memo || '',
    contacted: !!contacted[key]
  };
}

// ===== 案内済みの記録(唯一の書き込み) =====

function markContacted(memberId, milestoneName, targetYear) {
  if (!memberId || !milestoneName || !targetYear) {
    throw new Error('memberId・節目名・対象年は必須です。');
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    throw new Error('他の操作が実行中のため待機できませんでした。しばらくして再実行してください。');
  }
  try {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const sh = getOrCreateLifecycleContactsSheet_(ss);
    var operator = 'Web';
    try {
      const email = Session.getActiveUser().getEmail();
      if (email) operator = email;
    } catch (e) { /* 権限次第で取得できないことがある */ }
    sh.appendRow([memberId, milestoneName, String(targetYear),
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'), operator]);
    appendEvent_(ss, memberId, '節目案内', '', milestoneName + '(' + targetYear + ')', 'Web(節目)');
    return sanitizeForClient_({ memberId: memberId, milestoneName: milestoneName, targetYear: targetYear });
  } finally {
    lock.releaseLock();
  }
}

const LIFECYCLE_CONTACTS_HEADERS_ = ['memberId', '節目名', '対象年', '案内日', '操作者'];

function getOrCreateLifecycleContactsSheet_(ss) {
  var sh = ss.getSheetByName('Lifecycle_Contacts');
  if (!sh) {
    sh = ss.insertSheet('Lifecycle_Contacts');
    sh.getRange(1, 1, 1, LIFECYCLE_CONTACTS_HEADERS_.length).setValues([LIFECYCLE_CONTACTS_HEADERS_]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readLifecycleContacts_(ss) {
  const sh = ss.getSheetByName('Lifecycle_Contacts');
  if (!sh) return {};
  const map = {};
  readObjects_(ss, 'Lifecycle_Contacts').forEach(function (r) {
    map[r.節目名 + '|' + r.対象年 + '|' + r.memberId] = true;
  });
  return map;
}

// ===== 共通ヘルパー =====

function latestShoot_(shootsArr) {
  var best = null;
  shootsArr.forEach(function (s) {
    if (!(s.撮影日 instanceof Date)) return;
    if (!best || s.撮影日.getTime() > best.撮影日.getTime()) best = s;
  });
  return best;
}

// 対象年を問わず、過去いつでも該当ジャンルの撮影歴があればtrue(一度きりの節目の済みフィルタ用)
function hasGenreShotEver_(shootsArr, memberId, genreKeyword) {
  if (!genreKeyword) return false;
  return shootsArr.some(function (s) {
    if (String(s.ジャンル || '').indexOf(genreKeyword) < 0) return false;
    const targetIds = String(s.対象memberId || '').split(',').map(function (x) { return x.trim(); });
    return targetIds.indexOf(memberId) >= 0;
  });
}

function ageInYears_(birth, asOf) {
  var age = asOf.getFullYear() - birth.getFullYear();
  if (asOf.getMonth() < birth.getMonth() ||
    (asOf.getMonth() === birth.getMonth() && asOf.getDate() < birth.getDate())) age--;
  return age;
}

function monthsBetween_(from, to) {
  return (to.getFullYear() - from.getFullYear()) * 12 + (to.getMonth() - from.getMonth());
}

function addMonths_(date, months) {
  return new Date(date.getFullYear(), date.getMonth() + months, date.getDate());
}

function addYears_(date, years) {
  return new Date(date.getFullYear() + years, date.getMonth(), date.getDate());
}

// ===== 検証用(スクリプトエディタから手動実行。実データには触れないダミーデータのみ) =====

function testLifecycle() {
  const testToday = new Date(2026, 7, 2); // 2026-08-02(この検証の基準日)
  var pass = 0, fail = 0;

  function check(label, actual, expected) {
    const ok = (actual === expected);
    if (ok) { pass++; Logger.log('PASS: ' + label); }
    else { fail++; Logger.log('FAIL: ' + label + ' — 期待=' + expected + ' 実際=' + actual); }
  }

  const rule3 = LIFECYCLE_RULES_.filter(function (r) { return r.name === '七五三(3歳・女)'; })[0];
  const ruleJunior = LIFECYCLE_RULES_.filter(function (r) { return r.name === '中学入学'; })[0];
  const ruleBd1 = LIFECYCLE_RULES_.filter(function (r) { return r.name === '1歳バースデー'; })[0];

  Logger.log('--- a. 2022-11-12生・女、2025年10月に七五三撮影済み → 2026年の3歳七五三に出ない ---');
  const birthA = new Date(2022, 10, 12);
  const occA = findMilestoneOccurrence_(rule3, birthA);
  check('a-1: 対象年は2025年(2026年の枠には入らない)', occA.year, 2025);
  const shootsA = [{ ジャンル: '七五三', 対象memberId: 'M-TEST-A' }];
  check('a-2: 済みフィルタが検知する', hasGenreShotEver_(shootsA, 'M-TEST-A', rule3.genre), true);

  Logger.log('--- b. 2023年生まれの女児 → 2026年の3歳七五三対象。表示は2026-09-15以降 ---');
  const birthB = new Date(2023, 3, 10);
  const occB = findMilestoneOccurrence_(rule3, birthB);
  check('b-1: 対象年は2026年', occB.year, 2026);
  check('b-2: 目安日は2026-11-15', isoDate_(occB.dueDate), '2026-11-15');
  const leadDefault = LIFECYCLE_DEFAULT_LEAD_MONTHS_;
  check('b-3: 表示開始は2026-09-15', isoDate_(addMonths_(occB.dueDate, -leadDefault)), '2026-09-15');
  check('b-4: 2026-09-14はまだ非表示', isWithinDisplayWindow_(occB, leadDefault, new Date(2026, 8, 14)), false);
  check('b-5: 2026-09-15から表示', isWithinDisplayWindow_(occB, leadDefault, new Date(2026, 8, 15)), true);

  Logger.log('--- c. 2013-06-28生 → 中学入学は2026年4月で通過済み。2027年春入学は2014-04-02〜2015-04-01生 ---');
  const birthC = new Date(2013, 5, 28);
  const occC = findMilestoneOccurrence_(ruleJunior, birthC);
  check('c-1: 入学年は2026年', occC.year, 2026);
  check('c-2: 2026-08-02(基準日)時点で表示ウィンドウ外', isWithinDisplayWindow_(occC, leadDefault, testToday), false);
  check('c-3: 2014-04-02生は2027年入学', schoolEntranceYear_(new Date(2014, 3, 2), ruleJunior.targetAge), 2027);
  check('c-4: 2014-04-01生は2027年ではなく2026年入学', schoolEntranceYear_(new Date(2014, 3, 1), ruleJunior.targetAge), 2026);
  check('c-5: 2015-04-01生は2027年入学', schoolEntranceYear_(new Date(2015, 3, 1), ruleJunior.targetAge), 2027);
  check('c-6: 2015-04-02生は2027年ではなく2028年入学', schoolEntranceYear_(new Date(2015, 3, 2), ruleJunior.targetAge), 2028);

  Logger.log('--- d. 誕生日型(1歳バースデー)は誕生日の2ヶ月前から表示され、誕生日翌月末に消える ---');
  const birthD = new Date(2025, 5, 15); // 誕生日(1歳)=2026-06-15
  const occD = findMilestoneOccurrence_(ruleBd1, birthD);
  check('d-1: 目安日は2026-06-15', isoDate_(occD.dueDate), '2026-06-15');
  check('d-2: 2ヶ月前の前日(2026-04-14)はまだ非表示',
    isWithinDisplayWindow_(occD, leadDefault, new Date(2026, 3, 14)), false);
  check('d-3: 2ヶ月前ちょうど(2026-04-15)から表示',
    isWithinDisplayWindow_(occD, leadDefault, new Date(2026, 3, 15)), true);
  check('d-4: 翌月末(2026-07-31)はまだ表示',
    isWithinDisplayWindow_(occD, leadDefault, new Date(2026, 6, 31)), true);
  check('d-5: 翌月末の翌日(2026-08-01)は非表示',
    isWithinDisplayWindow_(occD, leadDefault, new Date(2026, 7, 1)), false);

  Logger.log('=== testLifecycle結果: PASS=' + pass + ' / FAIL=' + fail + ' / 合計=' + (pass + fail) + ' ===');
}
