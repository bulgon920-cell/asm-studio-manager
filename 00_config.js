/**
 * ===== 設定(書き換えるのはこのファイルだけ) =====
 * AI Studio Manager 移行スクリプト v1.0
 * 設計書: AI_Studio_Manager_新スキーマ設計書_v1.0.md
 */

// 移行元: Master_Log_snapshot_YYYYMMDD のスプレッドシートID
// (URLの /d/ と /edit の間の文字列)
const SNAPSHOT_SPREADSHEET_ID = '1hmI0Tdi9XcboG8JmWoIHndFfsCTjLr7eGsEO0WE4ZtM';

// 移行元のシート名。空欄 '' なら1枚目のシートを使う
const SNAPSHOT_SHEET_NAME = 'Master_Log_snapshot_20260801';

// 移行先: setupNewEnvironment() 実行後にログへ表示されるIDを貼る
const TARGET_SPREADSHEET_ID = '1kD95cWFXY-hDQ8EBxm0nxSTijG9cvg5TVByyzQEKcx0';

// 撮影日からこの日数を超えて仕上がり予定日がない案件は「完了(移行時推定)」にする
const COMPLETION_DAYS = 90;

// ジャンル名にこの文字が含まれていたら成人記念として扱う
const ADULT_GENRE_KEYWORDS = ['成人', '振袖', '二十歳', 'はたち'];

// migrate() の再実行時に、新環境の既存データ行を消してから流し直すか
// 通常は false。本番切替時に最新snapshotで流し直すときだけ true にする
const CLEAR_BEFORE_MIGRATE = false;

// 差分同期(05_sync.js)の読み先: 大本Master_Log(稼働中のもの)のスプレッドシートID
// スナップショットではない。ここには一切書き込まない(読み取り専用)
const SOURCE_MASTER_LOG_ID = '1a4-g8kkowwKVTYO66gptatzknZhXjI-EKXSijXza_TE';

// 大本Master_Logのシート(タブ)名。空欄 '' なら1枚目のシートを使う
const SOURCE_MASTER_LOG_SHEET_NAME = 'Master_Log';

// 朝の更新(13_morning.js)が読むGoogle CalendarのID。
// カレンダーの表示名は「bulgon920@yahoo.co.jp」だが、これはbulgon920@gmail.com本人の
// カレンダーに付けた表示名で、実際のカレンダーID(カレンダー設定→カレンダーの統合で確認)は
// 本人のアドレスと同じ bulgon920@gmail.com。2026-08-02に実機確認して修正(WEB_SPEC §8も修正)。
const CALENDAR_ID = 'bulgon920@gmail.com';
