/**
 * 売上タブ — 現行シート読み取り方式(v1)
 * 設計書: SALES_SPEC_v1.0.md
 *
 * 現行の「スタジオデイズ顧客名簿&月毎売上」スプレッドシート(SALES_SOURCE_ID)を
 * 読み取るだけ。ASM側では一切再計算しない。書き込みも一切行わない。
 *
 * 【現在の状態】
 * getSales(年,月): 実装済み(2025年04月シートを実読してブロック定義を確定)。
 * getAnnual(年): 未実装。年次分析シート(例: "2026年店分析")の実際のセル配置が
 * まだ分かっていないため、先に inspectSalesSheet() でレイアウトを特定してから実装する
 * (SALES_SPEC_v1.0.md §0-3「推測しない」)。
 *
 * sanitizeForClient_ は 10_web_api.js のものを再利用する(再定義しない)。
 */

// ===== 調査用(スクリプトエディタから手動実行) =====

// 引数なしで実行するとシート一覧をログに出す。
// シート名を渡すと、そのシートの中身(セル参照=値)を行ごとにログへダンプする。
// これは調査専用の読み取りツールで、書き込みは一切行わない。
function inspectSalesSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SALES_SOURCE_ID);

  if (!sheetName) {
    Logger.log('=== シート一覧(' + ss.getName() + ') ===');
    ss.getSheets().forEach(function (sh) {
      Logger.log(sh.getName() + '(行' + sh.getLastRow() + ' x 列' + sh.getLastColumn() + ')');
    });
    Logger.log('---');
    Logger.log('inspectSalesSheet("シート名") の形で再実行すると、中身をダンプします。');
    return;
  }

  const sh = ss.getSheetByName(sheetName);
  if (!sh) {
    Logger.log('シートが見つかりません: ' + sheetName);
    return;
  }
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  Logger.log('=== ' + sheetName + ' (行' + lastRow + ' x 列' + lastCol + ') ===');
  // getDisplayValues: 画面表示通りの文字列(数式の結果・エラー表示も含む)を取得する。
  // 実データ抽出(getSales等)ではgetValues()を使うが、レイアウト調査ではこちらが読みやすい。
  const values = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  for (var r = 0; r < values.length; r++) {
    const cells = [];
    for (var c = 0; c < values[r].length; c++) {
      const v = String(values[r][c]).trim();
      if (!v) continue;
      cells.push(sh.getRange(r + 1, c + 1).getA1Notation() + '=' + v);
    }
    if (cells.length) Logger.log('行' + (r + 1) + ': ' + cells.join(' | '));
  }
}

// ===== 調査用・一時呼び出し(セル配置確認が終わったら削除する) =====
// スクリプトエディタのRunボタンは引数なし関数しか実行できないため、
// 実際のシート名を書いた呼び出し専用の関数を用意する。
function inspectSalesSheet_MonthlySample() {
  inspectSalesSheet('2025年04月');
}
function inspectSalesSheet_AnnualSample() {
  inspectSalesSheet('2026年店分析');
}

// ===== 月次シート ブロック定義(2025年04月を実読して確定。2026-08-02) =====
//
// シート名: "YYYY年MM月"(月はゼロ埋め2桁)。
//
// レイアウトは「ジャンル1件につき2行1組」の繰り返し。
//   対象行(r)  : ジャンル名(A) / 目標売上(C) / 達成率(D) / 単月売上合計(F) /
//                旬別単月売上(G,H,I=1〜10日,11〜20日,21日〜末) /
//                目標件数(M) / 単月件数合計(P) / 旬別単月件数(Q,R,S) /
//                目標単価(X) / 旬別単月単価(Z,AA,AB)
//   直後の行(r+1): 前年売上(C) / 前年比(D) / 旬別“累計”売上(G,H,I) /
//                  前年件数(M) / 旬別“累計”件数(Q,R,S) /
//                  前年単価(X) / 旬別“累計”単価(Z,AA,AB)
// (r+1行のG/H/Iは「前年の値」ではなく、当月の1〜10日→11〜20日→21日〜末の累計。
//  実読データで確認済み: 21日〜末の累計(I[r+1])が単月合計(F[r])と一致する)
//
// 行番号:
//   行2-3  = 全体合計(業務ジャンル(葬儀・学校)を含む全体)
//   行5-6  = スタジオ月合計(写真ジャンルのみの合計。Web上のメイン表示に使う)
//   行7-32 = 写真ジャンル13種、2行1組(七五三,お宮参り,バースデー,家族写真,
//            プロフィール,マタニティ,成人式,ブライダル,その他撮影,ハーフBD,
//            1/2成人,入卒業,ペット撮影の順)
//   行34-39 = 業務ジャンル3種、2行1組(全体(葬儀学校),葬儀関係,学校関係)
//
// #REF!/#DIV/0!等のエラー値は "—" に変換して返す(件数はerrorCellCountに集計)。
// 完全一致確認のため、数値は再計算・再フォーマットせず getDisplayValues() の
// 表示文字列をそのまま返す。

const SALES_ERROR_PATTERN_ = /^#(REF|DIV\/0|N\/A|VALUE|NAME\?|NULL|NUM)!?$/;

function salesCell_(values, errCounter, row, col) {
  const cellRow = values[row - 1];
  const raw = (cellRow && cellRow[col - 1] != null) ? String(cellRow[col - 1]).trim() : '';
  if (SALES_ERROR_PATTERN_.test(raw)) {
    errCounter.count++;
    return '—';
  }
  return raw;
}

// r: 対象行(ジャンル名・目標・単月がある行)。r+1が前年・累計行。
function readSalesBlock_(values, errCounter, r) {
  return {
    genre: salesCell_(values, errCounter, r, 1),
    targetAmount: salesCell_(values, errCounter, r, 3),
    lastYearAmount: salesCell_(values, errCounter, r + 1, 3),
    achieveRate: salesCell_(values, errCounter, r, 4),
    yoyRate: salesCell_(values, errCounter, r + 1, 4),
    monthAmount: salesCell_(values, errCounter, r, 6),
    monthAmountByDecade: [
      salesCell_(values, errCounter, r, 7),
      salesCell_(values, errCounter, r, 8),
      salesCell_(values, errCounter, r, 9)
    ],
    cumulativeAmountByDecade: [
      salesCell_(values, errCounter, r + 1, 7),
      salesCell_(values, errCounter, r + 1, 8),
      salesCell_(values, errCounter, r + 1, 9)
    ],
    targetCount: salesCell_(values, errCounter, r, 13),
    lastYearCount: salesCell_(values, errCounter, r + 1, 13),
    monthCount: salesCell_(values, errCounter, r, 16),
    monthCountByDecade: [
      salesCell_(values, errCounter, r, 17),
      salesCell_(values, errCounter, r, 18),
      salesCell_(values, errCounter, r, 19)
    ],
    cumulativeCountByDecade: [
      salesCell_(values, errCounter, r + 1, 17),
      salesCell_(values, errCounter, r + 1, 18),
      salesCell_(values, errCounter, r + 1, 19)
    ],
    targetPrice: salesCell_(values, errCounter, r, 24),
    lastYearPrice: salesCell_(values, errCounter, r + 1, 24),
    priceByDecade: [
      salesCell_(values, errCounter, r, 26),
      salesCell_(values, errCounter, r, 27),
      salesCell_(values, errCounter, r, 28)
    ],
    cumulativePriceByDecade: [
      salesCell_(values, errCounter, r + 1, 26),
      salesCell_(values, errCounter, r + 1, 27),
      salesCell_(values, errCounter, r + 1, 28)
    ]
  };
}

const SALES_GENRE_ROWS_ = [7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31];
const SALES_BUSINESS_ROWS_ = [34, 36, 38];

// ===== 読み取りAPI =====

function getSales(year, month) {
  const sheetName = year + '年' + String(month).padStart(2, '0') + '月';
  const ss = SpreadsheetApp.openById(SALES_SOURCE_ID);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) {
    return sanitizeForClient_({ found: false, sheetName: sheetName, year: year, month: month });
  }

  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  const values = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const errCounter = { count: 0 };

  const result = {
    found: true,
    sheetName: sheetName,
    year: year,
    month: month,
    grandTotal: readSalesBlock_(values, errCounter, 2),
    studioTotal: readSalesBlock_(values, errCounter, 5),
    genres: SALES_GENRE_ROWS_.map(function (r) { return readSalesBlock_(values, errCounter, r); }),
    business: SALES_BUSINESS_ROWS_.map(function (r) { return readSalesBlock_(values, errCounter, r); }),
    errorCellCount: errCounter.count
  };
  return sanitizeForClient_(result);
}

// function getAnnual(year) { ... } — 年次分析シートの実読後に実装する。

// ===== 調査用・一時呼び出し(getSales実装の検証用。確認後に削除する) =====
function inspectSalesSheet_GetSalesTest() {
  const data = getSales(2025, 4);
  Logger.log(JSON.stringify(data, null, 2));
}
