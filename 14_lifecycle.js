/**
 * Lifecycle 節目一覧 — 読み取り中心(書き込みは「案内済み」記録の1つだけ)
 * 設計書: LIFECYCLE_SPEC_v1.0.md
 *
 * getLifecycle(): 今月〜3ヶ月先の節目対象を判定して返す読み取りAPI
 * markContacted(memberId, 節目名, 対象年): 「案内済み」をLifecycle_Contactsへ記録
 *
 * 節目マスタ(Config 区分=節目)は「提案メモ」が主役(業務側が直接編集する)。
 * 年齢・月数などの実際の判定ロジックはこのファイルのLIFECYCLE_RULES_で保持し、
 * Config側の型/条件/表示開始は人が読むための参考表示として扱う(§0-2)。
 *
 * readObjects_ / sanitizeForClient_ は 10_web_api.js、appendEvent_ は
 * 11_web_write_api.js、readBusinessGenreNames_ は 05_sync.js、
 * isoDate_ は 02_migrate.js のものをそれぞれ再利用する(再定義しない)。
 */

// ===== 節目マスタ(Config 区分=節目) =====

const LIFECYCLE_MILESTONES_SEED_ = [
  ['お宮参り', '誕生日型', '生後0〜3ヶ月', '誕生直後', '生後1〜3ヶ月が目安'],
  ['ハーフバースデー', '誕生日型', '生後6ヶ月', '2ヶ月前', '寝返り・ハイハイができると撮影バリエーションが増える'],
  ['1歳バースデー', '誕生日型', '満1歳', '2ヶ月前', 'スマッシュケーキやファーストシューズ提案'],
  ['2歳記念', '誕生日型', '満2歳', '2ヶ月前', '女の子なら次の3歳は七五三へ接続'],
  ['七五三(3歳・女)', '年次型', 'その年に満3歳(女)', '3月から', '満年齢推奨。数えなら早生まれは満年齢が1つ上・着物サイズ選びに影響'],
  ['七五三(5歳・男)', '年次型', 'その年に満5歳(男)', '3月から', ''],
  ['七五三(7歳・女)', '年次型', 'その年に満7歳(女)', '3月から', 'ランドセルと同時撮影プラン'],
  ['ランドセル(入学)', '年次型', '翌年に小学校入学', '前年10月から', '7歳女は七五三と同時プラン'],
  ['ハーフ成人式', '誕生日型', '満10歳', '3ヶ月前', ''],
  ['中学入学', '年次型', '翌年に中学入学', '前年12月から', ''],
  ['成人式', '年次型', '二十歳の前年から', '前年4月から', '振袖層への提案に特化。振袖在庫4本・自社対応。大量案件は狙わない'],
  ['背の順家族写真', '経過型', '末っ子が満7歳以上かつ最終来店から11ヶ月以上', 'ー', '背の順に並んで、毎年一枚の家族写真']
];

// 初回のみ: Configに節目マスタを登録する(冪等)。値4(提案メモ)列が無ければヘッダーも足す。
function ensureLifecycleConfig_(ss) {
  const sh = ss.getSheetByName('Config');
  if (sh.getLastColumn() < 6 || !sh.getRange(1, 6).getValue()) {
    sh.getRange(1, 6).setValue('値4');
  }
  const values = sh.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    if (values[r][0] === '節目') return; // 既に1件でもあれば初期化済み
  }
  const rows = LIFECYCLE_MILESTONES_SEED_.map(function (m) {
    return ['節目', m[0], m[1], m[2], m[3], m[4]];
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  Logger.log('Configに節目マスタを' + rows.length + '件登録しました。');
}

function readLifecycleMilestones_(ss) {
  const sh = ss.getSheetByName('Config');
  const values = sh.getDataRange().getValues();
  const map = {};
  for (var r = 1; r < values.length; r++) {
    if (values[r][0] === '節目') {
      map[values[r][1]] = {
        type: values[r][2], condition: values[r][3],
        displayStart: values[r][4], memo: values[r][5] || ''
      };
    }
  }
  return map;
}

// ===== 判定ロジック(型ごと。名称はConfigの節目名と対応させる) =====

const LIFECYCLE_RULES_ = [
  { name: 'お宮参り', kind: 'range', minMonths: 0, maxMonths: 3 },
  { name: 'ハーフバースデー', kind: 'point', months: 6, leadMonths: 2 },
  { name: '1歳バースデー', kind: 'point', years: 1, leadMonths: 2 },
  { name: '2歳記念', kind: 'point', years: 2, leadMonths: 2 },
  { name: '七五三(3歳・女)', kind: 'yearly', targetAge: 3, gender: '女', seasonStartMonth: 3, genre: '七五三' },
  { name: '七五三(5歳・男)', kind: 'yearly', targetAge: 5, gender: '男', seasonStartMonth: 3, genre: '七五三' },
  { name: '七五三(7歳・女)', kind: 'yearly', targetAge: 7, gender: '女', seasonStartMonth: 3, genre: '七五三' },
  { name: 'ランドセル(入学)', kind: 'school', targetAge: 6, seasonStartMonth: 10 },
  { name: 'ハーフ成人式', kind: 'point', years: 10, leadMonths: 3 },
  { name: '中学入学', kind: 'school', targetAge: 12, seasonStartMonth: 12 },
  { name: '成人式', kind: 'comingofage', targetAge: 20, seasonStartMonth: 4, genre: '成人' },
  { name: '背の順家族写真', kind: 'elapsed', minAge: 7, minMonthsSinceVisit: 11 }
];

function evalRule_(rule, member, today, windowEnd) {
  const birth = member.誕生日;
  if (rule.kind === 'range') {
    const start = addMonths_(birth, rule.minMonths);
    const end = addMonths_(birth, rule.maxMonths);
    if (today < start || today > end) return null;
    return { year: today.getFullYear(), targetDate: end };
  }
  if (rule.kind === 'point') {
    const target = rule.years != null ? addYears_(birth, rule.years) : addMonths_(birth, rule.months);
    const displayStart = addMonths_(target, -(rule.leadMonths || 0));
    if (today < displayStart || today > windowEnd || today > target) return null;
    return { year: target.getFullYear(), targetDate: target };
  }
  if (rule.kind === 'yearly' || rule.kind === 'school') {
    var found = null;
    [today.getFullYear(), today.getFullYear() + 1].forEach(function (year) {
      if (found) return;
      const ageAtApril1 = ageInYears_(birth, new Date(year, 3, 1));
      if (ageAtApril1 !== rule.targetAge) return;
      const displayStart = (rule.kind === 'yearly')
        ? new Date(year, rule.seasonStartMonth - 1, 1)
        : new Date(year - 1, rule.seasonStartMonth - 1, 1);
      const seasonEnd = new Date(year, 11, 31);
      if (today >= displayStart && today <= seasonEnd) {
        found = { year: year, targetDate: seasonEnd };
      }
    });
    return found;
  }
  if (rule.kind === 'comingofage') {
    const birthday20 = addYears_(birth, rule.targetAge);
    const displayStart = new Date(birthday20.getFullYear() - 1, rule.seasonStartMonth - 1, 1);
    const seasonEnd = new Date(birthday20.getFullYear(), 11, 31);
    if (today < displayStart || today > seasonEnd) return null;
    return { year: birthday20.getFullYear(), targetDate: birthday20 };
  }
  return null;
}

// ===== 読み取りAPI =====

function getLifecycle() {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
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

  const today = new Date();
  const windowEnd = addMonths_(today, 3);
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

      const hit = evalRule_(rule, m, today, windowEnd);
      if (!hit) return;
      // 済みフィルタ: 該当ジャンルを対象年に撮影済みの子は除外(§2)
      if (rule.genre && hasGenreShotThisYear_(custShoots, m.memberId, rule.genre, hit.year)) return;

      const key = rule.name + '|' + hit.year + '|' + m.memberId;
      results.push(buildLifecycleItem_(m.memberId, m.customerId, m.名前, m.誕生日, rule.name, hit.year,
        hit.targetDate, cust, lastShoot, memoByName, contacted, key));
    });
  });

  // 経過型(背の順家族写真)は家族(顧客)単位で判定
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

  return sanitizeForClient_({ items: results, noBirthdayCount: noBirthdayCount });
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

function hasGenreShotThisYear_(shootsArr, memberId, genreKeyword, year) {
  return shootsArr.some(function (s) {
    if (!(s.撮影日 instanceof Date) || s.撮影日.getFullYear() !== year) return false;
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
