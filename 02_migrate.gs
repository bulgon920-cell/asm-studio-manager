/**
 * 移行スクリプト本体。
 *
 * 使い方(必ずこの順で):
 *   1. setupNewEnvironment() を実行 → IDを00_config.gsへ貼る
 *   2. dryRun() を実行 → 新環境の「移行レポート」シートを確認
 *   3. 問題なければ migrate() を実行
 *
 * dryRun() は移行元・移行先のデータを一切変更しない(レポートを書くだけ)。
 */

// ===== 入口 =====

function dryRun() {
  const model = buildModel_();
  writeReport_(model, true);
  Logger.log('dryRun完了。新環境の「移行レポート」シートを確認してください。');
}

function migrate() {
  const model = buildModel_();
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);

  guardOrClear_(ss);
  writeSheet_(ss, 'Customers', model.customers.map(customerRow_));
  writeSheet_(ss, 'Family_Members', model.members.map(memberRow_));
  writeSheet_(ss, 'Shoots', model.shoots.map(shootRow_));
  writeSheet_(ss, 'Orders', model.orders.map(orderRow_));
  writeEventLog_(ss, model);
  writeReport_(model, false);

  Logger.log('移行完了: 顧客' + model.customers.length +
    ' / 家族構成員' + model.members.length +
    ' / 撮影' + model.shoots.length +
    ' / 注文' + model.orders.length);
}

// ===== 移行元の読み取り =====

function readSnapshot_() {
  const ss = SpreadsheetApp.openById(SNAPSHOT_SPREADSHEET_ID);
  const sh = SNAPSHOT_SHEET_NAME
    ? ss.getSheetByName(SNAPSHOT_SHEET_NAME)
    : ss.getSheets()[0];
  if (!sh) throw new Error('移行元シートが見つかりません: ' + SNAPSHOT_SHEET_NAME);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error('移行元にデータ行がありません。');
  const headers = values[0].map(function (h) { return String(h).trim(); });

  // 列名→列番号。「性別」は3回現れるので子どもの名前nに続く順で対応付ける
  const idx = {};
  headers.forEach(function (h, i) { if (!(h in idx)) idx[h] = i; });
  const genderCols = [];
  headers.forEach(function (h, i) { if (h === '性別') genderCols.push(i); });
  const remarksCol = headers.findIndex(function (h) {
    return h.indexOf('備考') === 0;
  });

  function col(name) { return (name in idx) ? idx[name] : -1; }

  const rows = [];
  for (var r = 1; r < values.length; r++) {
    const v = values[r];
    const get = function (name) {
      const c = col(name);
      return c >= 0 ? v[c] : '';
    };
    // 完全な空行はスキップ
    if (!String(get('お名前')).trim() && !get('日付')) continue;

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

// ===== モデル構築(顧客同定・ID採番) =====

function buildModel_() {
  const src = readSnapshot_();
  // 古い行から処理する(顧客の初出を古い順に採番するため)
  src.sort(function (a, b) {
    return (a.date ? a.date.getTime() : 0) - (b.date ? b.date.getTime() : 0);
  });

  const customers = [];
  const members = [];
  const shoots = [];
  const orders = [];
  const issues = [];

  const byChildKey = {};   // 子どもの名前+誕生日 → customer
  const byNamePhone = {};  // 顧客名+連絡先 → customer
  const byNameAddr = {};   // 顧客名+住所 → customer
  const nameSeen = {};     // 顧客名(正規化) → true

  const today = new Date();
  const shootSeqByDate = {};
  var eventSeq = 0;

  src.forEach(function (row) {
    // --- 顧客同定(設計書 §4-1 の優先順) ---
    var cust = null;
    var matchedBy = '';
    row.children.forEach(function (ch) {
      if (cust) return;
      const k = childKey_(ch);
      if (k && byChildKey[k]) { cust = byChildKey[k]; matchedBy = '子ども'; }
    });
    if (!cust && row.contact) {
      const k = norm_(row.name) + '|' + norm_(row.contact);
      if (byNamePhone[k]) { cust = byNamePhone[k]; matchedBy = '名前+連絡先'; }
    }
    if (!cust && row.addr) {
      const k = norm_(row.name) + '|' + norm_(row.addr);
      if (byNameAddr[k]) { cust = byNameAddr[k]; matchedBy = '名前+住所'; }
    }

    if (!cust) {
      const dup = !!nameSeen[norm_(row.name)];
      cust = {
        id: 'C-' + pad_(customers.length + 1, 5),
        name: row.name, kana: '', contact: row.contact, zip: row.zip,
        addr: row.addr, line: hasLine_(row.remarks), remarks: '',
        needsReview: dup, memberKeys: {}
      };
      customers.push(cust);
      nameSeen[norm_(row.name)] = true;
      if (dup) {
        issues.push(['要確認(顧客)', cust.id, row.srcRow,
          '同名の既存顧客がいるがキー不一致。同一家族か確認して統合を検討: ' + row.name]);
      }
    } else {
      // 新しい行の非空欄値で上書き(最新情報を優先)
      if (row.contact) cust.contact = row.contact;
      if (row.zip) cust.zip = row.zip;
      if (row.addr) cust.addr = row.addr;
      if (hasLine_(row.remarks)) cust.line = true;
    }

    // --- 家族構成員 ---
    const memberIds = [];
    row.children.forEach(function (ch) {
      const mk = ch.name ? (norm_(ch.name) + '|' + isoDate_(ch.birth)) : '';
      if (!mk) return;
      if (!cust.memberKeys[mk]) {
        const m = {
          id: 'M-' + pad_(members.length + 1, 5),
          customerId: cust.id, name: ch.name, kana: ch.kana,
          gender: ch.gender, birth: ch.birth
        };
        members.push(m);
        cust.memberKeys[mk] = m;
      }
      const kFull = childKey_(ch);
      if (kFull) byChildKey[kFull] = cust;
      memberIds.push(cust.memberKeys[mk].id);
    });
    if (row.contact) byNamePhone[norm_(row.name) + '|' + norm_(row.contact)] = cust;
    if (row.addr) byNameAddr[norm_(row.name) + '|' + norm_(row.addr)] = cust;

    // --- 撮影 ---
    const dstr = row.date
      ? Utilities.formatDate(row.date, 'Asia/Tokyo', 'yyyyMMdd') : '00000000';
    shootSeqByDate[dstr] = (shootSeqByDate[dstr] || 0) + 1;
    const shoot = {
      id: 'S-' + dstr + '-' + pad_(shootSeqByDate[dstr], 2),
      customerId: cust.id, date: row.date, genre: row.genre,
      hp: row.hp || '未確認', amount: row.amount, payment: row.payment,
      memberIds: memberIds.join(','), needsReview: false,
      remarks: row.remarks, srcRow: row.srcRow
    };
    if (!row.date) {
      shoot.needsReview = true;
      issues.push(['要確認(撮影)', shoot.id, row.srcRow, '撮影日が読み取れない']);
    }
    if (!row.genre) {
      shoot.needsReview = true;
      issues.push(['要確認(撮影)', shoot.id, row.srcRow, 'ジャンルが空欄']);
    }
    shoots.push(shoot);

    // --- 注文(設計書 §4-4) ---
    const isAdult = ADULT_GENRE_KEYWORDS.some(function (kw) {
      return row.genre.indexOf(kw) >= 0;
    });
    const o = {
      id: 'O-' + pad_(orders.length + 1, 5),
      shootId: shoot.id, finish: row.finish, memo: ''
    };
    if (row.finish && row.finish.getTime() > today.getTime()) {
      o.type = isAdult ? '成人商品(セレクト前)' : '撮影データ';
      o.status = '要確認';
      o.memo = '移行: 仕上がり予定日が未来のため進行中の可能性。実状態を確認して更新';
      issues.push(['進行中の可能性', o.id, row.srcRow,
        '仕上がり予定日 ' + isoDate_(row.finish) + ' → 要確認で起票']);
    } else if (row.finish) {
      o.type = isAdult ? '成人商品(セレクト前)' : '撮影データ';
      o.status = '完了(移行時推定)';
      o.memo = '移行: 仕上がり予定日が過去のため納品済みと推定';
    } else if (row.date && daysBetween_(row.date, today) > COMPLETION_DAYS) {
      o.type = '撮影データ';
      o.status = '完了(移行時推定)';
      o.memo = '移行: 撮影日から' + COMPLETION_DAYS + '日超のため完了と推定';
    } else {
      o.type = isAdult ? '成人商品(セレクト前)' : '撮影データ';
      o.status = '要確認';
      o.memo = '移行: 撮影日から' + COMPLETION_DAYS + '日以内。実状態を確認して更新';
      issues.push(['進行中の可能性', o.id, row.srcRow,
        '直近' + COMPLETION_DAYS + '日以内の撮影 → 要確認で起票']);
    }
    orders.push(o);
  });

  return {
    customers: customers, members: members, shoots: shoots,
    orders: orders, issues: issues, srcCount: src.length
  };
}

// ===== 書き込み =====

function guardOrClear_(ss) {
  const cust = ss.getSheetByName('Customers');
  if (!cust) throw new Error('移行先にCustomersシートがありません。setupNewEnvironment()の実行とIDの設定を確認してください。');
  if (cust.getLastRow() > 1) {
    if (!CLEAR_BEFORE_MIGRATE) {
      throw new Error('移行先に既存データがあります。流し直す場合は 00_config.gs の CLEAR_BEFORE_MIGRATE を true にしてください。');
    }
    ['Customers', 'Family_Members', 'Shoots', 'Orders', 'Event_Log',
      'Today_Task_Board'].forEach(function (name) {
        const sh = ss.getSheetByName(name);
        if (sh && sh.getLastRow() > 1) {
          sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
        }
      });
  }
}

function writeSheet_(ss, name, rows) {
  if (!rows.length) return;
  const sh = ss.getSheetByName(name);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length)
    .setValues(rows);
}

function writeEventLog_(ss, model) {
  const now = new Date();
  const rows = [];
  var seq = 1;
  function add(target, item) {
    rows.push([seq++, now, '移行スクリプト', target, item, '', '生成', '移行']);
  }
  model.customers.forEach(function (c) { add(c.id, '移行生成'); });
  model.members.forEach(function (m) { add(m.id, '移行生成'); });
  model.shoots.forEach(function (s) { add(s.id, '移行生成'); });
  model.orders.forEach(function (o) { add(o.id, '移行生成(status=' + o.status + ')'); });
  writeSheet_(ss, 'Event_Log', rows);
}

// 行変換
function customerRow_(c) {
  return [c.id, c.name, c.kana, c.contact, c.zip, c.addr,
    c.line, c.remarks, c.needsReview, new Date(), new Date()];
}
function memberRow_(m) {
  return [m.id, m.customerId, m.name, m.kana, m.gender, m.birth || '', ''];
}
function shootRow_(s) {
  return [s.id, s.customerId, s.date || '', s.genre, s.hp,
    s.amount, s.payment, '', '', s.memberIds, s.needsReview, s.remarks];
}
function orderRow_(o) {
  return [o.id, o.shootId, o.type, o.status, o.finish || '',
    '', '', '', o.memo, new Date(),
    o.status === '完了(移行時推定)' ? new Date() : ''];
}

// ===== ユーティリティ =====

function norm_(s) {
  return String(s).replace(/[\s\u3000]+/g, '')
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (ch) {
      return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
    }).toLowerCase();
}
function childKey_(ch) {
  if (!ch.name || !ch.birth) return '';
  return norm_(ch.name) + '|' + isoDate_(ch.birth);
}
function toDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  const s = String(v).trim();
  if (!s) return null;
  const m = s.match(/^(\d{4})[\/\-年\.](\d{1,2})[\/\-月\.](\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return null;
}
function isoDate_(d) {
  return d ? Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') : '';
}
function daysBetween_(a, b) {
  return Math.floor((b.getTime() - a.getTime()) / 86400000);
}
function pad_(n, len) {
  return ('00000' + n).slice(-len);
}
function hasLine_(remarks) {
  return String(remarks).indexOf('LINE有') >= 0;
}
