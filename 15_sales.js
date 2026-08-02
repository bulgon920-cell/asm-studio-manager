/**
 * 売上タブ — 現行シート読み取り方式(v1)
 * 設計書: SALES_SPEC_v1.0.md
 *
 * 現行の「スタジオデイズ顧客名簿&月毎売上」スプレッドシート(SALES_SOURCE_ID)を
 * 読み取るだけ。ASM側では一切再計算しない。書き込みも一切行わない。
 *
 * 【現在の状態】getSales()/getAnnual() は未実装。
 * 月次シート・年次分析シートの実際のセル配置がまだ分かっていないため、
 * 先にこのファイルの inspectSalesSheet() を実行してレイアウトを特定し、
 * その結果をもとに読み取りブロック定義(コメント付き)を書いてから実装する
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

// ===== 読み取りAPI(セル配置が判明してから実装する) =====
//
// function getSales(yearMonth) { ... }
// function getAnnual(year) { ... }
//
// 実装時はここに、判明した月次シートのブロック定義を記録する。例:
//   スタジオ月合計: セルXX (売上) / セルYY (件数) ...
//   ジャンル別テーブル: 行N〜M, 列P(ジャンル名)/列Q(売上)/... ...
