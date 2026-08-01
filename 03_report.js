/**
 * 移行レポートを新環境の「移行レポート」シートへ出力する。
 * dryRun() / migrate() の両方から呼ばれる。
 */

function writeReport_(model, isDryRun) {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  var sh = ss.getSheetByName('移行レポート');
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet('移行レポート');

  const activeOrders = model.orders.filter(function (o) {
    return o.status === '要確認';
  }).length;
  const doneOrders = model.orders.filter(function (o) {
    return o.status === '完了(移行時推定)';
  }).length;
  const reviewCustomers = model.customers.filter(function (c) {
    return c.needsReview;
  }).length;
  const reviewShoots = model.shoots.filter(function (s) {
    return s.needsReview;
  }).length;

  const summary = [
    ['AI Studio Manager 移行レポート', ''],
    ['実行種別', isDryRun ? 'dryRun(何も書き込んでいません)' : '本実行'],
    ['実行日時', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm')],
    ['', ''],
    ['移行元の有効行数', model.srcCount],
    ['生成: 顧客(Customers)', model.customers.length],
    ['生成: 家族構成員(Family_Members)', model.members.length],
    ['生成: 撮影(Shoots)', model.shoots.length],
    ['生成: 注文(Orders)', model.orders.length],
    ['', ''],
    ['注文: 完了(移行時推定)', doneOrders],
    ['注文: 要確認(Web初回起動時に人が確定)', activeOrders],
    ['顧客: 要確認(同名別キー。統合検討)', reviewCustomers],
    ['撮影: 要確認(日付/ジャンル不備)', reviewShoots],
    ['', ''],
    ['↓ 個別の確認事項一覧', '']
  ];
  sh.getRange(1, 1, summary.length, 2).setValues(summary);
  sh.getRange(1, 1).setFontWeight('bold');

  if (model.issues.length) {
    const start = summary.length + 1;
    sh.getRange(start, 1, 1, 4)
      .setValues([['種類', '対象ID', '移行元の行番号', '内容']])
      .setFontWeight('bold');
    sh.getRange(start + 1, 1, model.issues.length, 4)
      .setValues(model.issues);
  }
  sh.autoResizeColumns(1, 4);
}
