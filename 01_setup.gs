/**
 * 新環境スプレッドシートを設計書v1.0の7シート構成で新規作成する。
 * 実行: setupNewEnvironment()
 * 実行後、ログに表示されるIDを 00_config.gs の TARGET_SPREADSHEET_ID へ貼る。
 */

const SHEET_DEFS = {
  Customers: [
    'customerId', '顧客名', 'ふりがな', '連絡先', '郵便番号', '住所',
    'LINE有無', '備考', '要確認', '登録日', '更新日'
  ],
  Family_Members: [
    'memberId', 'customerId', '名前', 'ふりがな', '性別', '誕生日', '備考'
  ],
  Shoots: [
    'shootId', 'customerId', '撮影日', 'ジャンル', 'HP掲載について',
    '合計金額', '決済方法', 'DriveフォルダURL', 'オンラインアルバムURL',
    '対象memberId', '要確認', '備考'
  ],
  Orders: [
    'orderId', 'shootId', '注文種別', 'status', '仕上がり予定日',
    'セレクト予定日', '期限', '担当', 'メモ', '作成日', '完了日'
  ],
  Event_Log: [
    'eventId', '日時', '操作者', '対象', '項目', '変更前', '変更後', '経路'
  ],
  Config: ['区分', 'キー', '値1', '値2', '値3'],
  Today_Task_Board: [
    '種別', '対象ID', '顧客名', '内容', '次の行動', '期限', '担当', 'status'
  ]
};

const STATUS_LIST = [
  'データ納品待ち', '店頭セレクト待ち', '発注待ち', '仕上がり待ち',
  '納品連絡待ち', '引渡し待ち', '完了', '完了(移行時推定)', '要確認',
  '保留', '不要'
];

const ORDER_TYPES = [
  // [注文種別, 系統, 標準価格メモ]
  ['撮影データ', 'データ系', ''],
  ['全データ', 'データ系', '38500'],
  ['成人商品(セレクト前)', '商品系', ''],
  ['白台紙', '商品系', '撮影代7700+台紙カスタマイズ制'],
  ['雅', '商品系', '27500'],
  ['ローズ', '商品系', '33000'],
  ['台紙', '商品系', ''],
  ['飾れる商品', '商品系', ''],
  ['全カット美肌補正', '作業系', '11000']
];

function setupNewEnvironment() {
  const name = 'AI_Studio_Manager_本体_' +
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  const ss = SpreadsheetApp.create(name);

  Object.keys(SHEET_DEFS).forEach(function (sheetName, i) {
    const sh = (i === 0)
      ? ss.getSheets()[0].setName(sheetName)
      : ss.insertSheet(sheetName);
    const headers = SHEET_DEFS[sheetName];
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold');
    sh.setFrozenRows(1);
    // 999行問題の再発防止: 使わない行・列を最初から削る
    const maxRows = sh.getMaxRows();
    if (maxRows > 2) sh.deleteRows(3, maxRows - 2);
    const maxCols = sh.getMaxColumns();
    if (maxCols > headers.length) {
      sh.deleteColumns(headers.length + 1, maxCols - headers.length);
    }
  });

  writeInitialConfig_(ss.getSheetByName('Config'));

  Logger.log('新環境を作成しました。');
  Logger.log('名前: ' + name);
  Logger.log('ID: ' + ss.getId());
  Logger.log('URL: ' + ss.getUrl());
  Logger.log('→ このIDを 00_config.gs の TARGET_SPREADSHEET_ID へ貼ってください。');
}

function writeInitialConfig_(sh) {
  const rows = [];
  STATUS_LIST.forEach(function (s) { rows.push(['status', s, '', '', '']); });
  ORDER_TYPES.forEach(function (t) {
    rows.push(['注文種別', t[0], t[1], t[2], '']);
  });
  // 表示文言辞書(Task Board用)
  const actions = [
    ['データ納品待ち', 'データ納品を進める'],
    ['店頭セレクト待ち', '店頭セレクトを行う'],
    ['発注待ち', '発注作業を進める'],
    ['納品連絡待ち', '納品の連絡をする'],
    ['引渡し待ち', '引渡しを進める'],
    ['要確認', '内容を確認する']
  ];
  actions.forEach(function (a) { rows.push(['次の行動', a[0], a[1], '', '']); });
  rows.push(['設定', 'COMPLETION_DAYS', String(COMPLETION_DAYS), '', '']);
  sh.getRange(2, 1, rows.length, 5).setValues(rows);
}
