/**
 * 顧客同定の診断(読み取り専用。どのデータも変更しない)
 * 実行: diagnoseMatching()
 * 結果: 新環境の「同定診断」シートへ出力(顧客名や子どもの名前は出力しない)
 *
 * 目的: dryRunで顧客数が移行元行数に近すぎる(=リピーター統合が効いていない)
 *       原因を、記入率・日付形式・キー別一致数の数値で特定する。
 */

function diagnoseMatching() {
  const src = readSnapshot_();
  const raw = readRawBirthCells_();

  // --- 1. 記入率 ---
  var contactFilled = 0, addrFilled = 0;
  var childWithName = 0, childWithNameAndBirth = 0, rowsWithAnyChild = 0;
  src.forEach(function (row) {
    if (row.contact) contactFilled++;
    if (row.addr) addrFilled++;
    if (row.children.length) rowsWithAnyChild++;
    row.children.forEach(function (ch) {
      childWithName++;
      if (ch.birth) childWithNameAndBirth++;
    });
  });

  // --- 2. 誕生日セルの形式分布(値の中身は日付のみ。名前は見ない) ---
  var fmtDate = 0, fmtParsedStr = 0, fmtFailedStr = 0, fmtEmpty = 0;
  const failedSamples = {};
  raw.forEach(function (v) {
    if (v instanceof Date && !isNaN(v.getTime())) { fmtDate++; return; }
    const s = String(v).trim();
    if (!s) { fmtEmpty++; return; }
    if (toDate_(s)) { fmtParsedStr++; return; }
    fmtFailedStr++;
    const shape = s.replace(/\d/g, '9');
    if (Object.keys(failedSamples).length < 10 && !failedSamples[shape]) {
      failedSamples[shape] = s;
    }
  });

  // --- 3. キー別の一致シミュレーション ---
  const stats = { child: 0, phone: 0, addr: 0, none: 0 };
  var hypoNameChild = 0; // 仮キー「顧客名+子ども名(誕生日なし)」で追加統合できる件数
  const byChildKey = {}, byNamePhone = {}, byNameAddr = {}, byNameChildName = {};

  src.slice().sort(function (a, b) {
    return (a.date ? a.date.getTime() : 0) - (b.date ? b.date.getTime() : 0);
  }).forEach(function (row) {
    var matched = '';
    row.children.forEach(function (ch) {
      if (matched) return;
      const k = childKey_(ch);
      if (k && byChildKey[k]) matched = 'child';
    });
    if (!matched && row.contact &&
      byNamePhone[norm_(row.name) + '|' + norm_(row.contact)]) matched = 'phone';
    if (!matched && row.addr &&
      byNameAddr[norm_(row.name) + '|' + norm_(row.addr)]) matched = 'addr';

    if (!matched) {
      // 仮キーなら一致していたか
      var hypo = false;
      row.children.forEach(function (ch) {
        if (ch.name &&
          byNameChildName[norm_(row.name) + '|' + norm_(ch.name)]) hypo = true;
      });
      if (hypo) hypoNameChild++;
      stats.none++;
    } else {
      stats[matched]++;
    }
    // 登録
    row.children.forEach(function (ch) {
      const k = childKey_(ch);
      if (k) byChildKey[k] = true;
      if (ch.name) byNameChildName[norm_(row.name) + '|' + norm_(ch.name)] = true;
    });
    if (row.contact) byNamePhone[norm_(row.name) + '|' + norm_(row.contact)] = true;
    if (row.addr) byNameAddr[norm_(row.name) + '|' + norm_(row.addr)] = true;
  });

  // --- 出力 ---
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  var sh = ss.getSheetByName('同定診断');
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet('同定診断');

  const pct = function (n, d) {
    return d ? Math.round(n * 100 / d) + '%' : '-';
  };
  const rows = [
    ['同定診断(読み取り専用)', ''],
    ['実行日時', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm')],
    ['', ''],
    ['■ 記入率', ''],
    ['移行元の有効行数', src.length],
    ['連絡先あり', contactFilled + ' (' + pct(contactFilled, src.length) + ')'],
    ['住所あり', addrFilled + ' (' + pct(addrFilled, src.length) + ')'],
    ['子どもの名前がある行', rowsWithAnyChild + ' (' + pct(rowsWithAnyChild, src.length) + ')'],
    ['子ども延べ人数(名前あり)', childWithName],
    ['うち誕生日もある', childWithNameAndBirth + ' (' + pct(childWithNameAndBirth, childWithName) + ')'],
    ['', ''],
    ['■ 誕生日セルの形式(空欄以外)', ''],
    ['日付型として読める', fmtDate],
    ['文字列だが変換できた', fmtParsedStr],
    ['文字列で変換できない', fmtFailedStr],
    ['空欄', fmtEmpty],
    ['', ''],
    ['■ キー別の一致数(2回目以降の来店を統合できた数)', ''],
    ['子どもの名前+誕生日', stats.child],
    ['顧客名+連絡先', stats.phone],
    ['顧客名+住所', stats.addr],
    ['どのキーでも一致せず新規顧客扱い', stats.none],
    ['', ''],
    ['■ 仮キーの効果測定', ''],
    ['「顧客名+子ども名(誕生日不要)」なら追加で統合できる件数', hypoNameChild],
    ['', ''],
    ['↓ 変換できない誕生日の形式サンプル(9=数字)', '']
  ];
  Object.keys(failedSamples).forEach(function (shape) {
    rows.push([shape, failedSamples[shape]]);
  });
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1).setFontWeight('bold');
  sh.autoResizeColumns(1, 2);
  Logger.log('診断完了。新環境の「同定診断」シートを確認してください。');
}

/** 誕生日1〜3列の生の値を全行分集める(形式診断用) */
function readRawBirthCells_() {
  const ss = SpreadsheetApp.openById(SNAPSHOT_SPREADSHEET_ID);
  const sh = SNAPSHOT_SHEET_NAME
    ? ss.getSheetByName(SNAPSHOT_SHEET_NAME)
    : ss.getSheets()[0];
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(function (h) { return String(h).trim(); });
  const cols = [];
  ['誕生日1', '誕生日2', '誕生日3'].forEach(function (name) {
    const i = headers.indexOf(name);
    if (i >= 0) cols.push(i);
  });
  const out = [];
  for (var r = 1; r < values.length; r++) {
    // 名前がある子の誕生日セルだけを対象にする
    for (var n = 0; n < cols.length; n++) {
      const nameIdx = headers.indexOf('子どもの名前' + (n + 1));
      if (nameIdx >= 0 && String(values[r][nameIdx]).trim()) {
        out.push(values[r][cols[n]]);
      }
    }
  }
  return out;
}
