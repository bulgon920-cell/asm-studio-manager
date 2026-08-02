/**
 * 日次バックアップ(B1)
 * 設計書: ROADMAP_v2.md P1
 *
 * 使い方:
 *   installBackupTrigger() を一度実行 → 毎日深夜に runDailyBackup() が自動実行される
 *   手動でバックアップしたいときは runDailyBackup() を直接実行してもよい
 *
 * 新環境スプレッドシート(TARGET_SPREADSHEET_ID)をDriveの「ASM_backup」フォルダへ
 * 日付名でコピーし、30世代を超えた古いものを削除する。
 */

const BACKUP_FOLDER_NAME = 'ASM_backup';
const BACKUP_KEEP_GENERATIONS = 30;
const BACKUP_TRIGGER_HOUR = 3; // 深夜3時台に実行

function installBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runDailyBackup') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runDailyBackup')
    .timeBased()
    .atHour(BACKUP_TRIGGER_HOUR)
    .nearMinute(0)
    .everyDays(1)
    .create();
  Logger.log('日次バックアップトリガーを設置しました(毎日' + BACKUP_TRIGGER_HOUR + '時台に実行)。');
}

// トリガーを待たず、今すぐ1回バックアップを取りたいとき用(RECOVERY.md参照)。
// runDailyBackup()と同じ処理を分かりやすい名前で呼べるようにしただけ。
function runBackupNow() {
  runDailyBackup();
}

function runDailyBackup() {
  const folder = getOrCreateBackupFolder_();
  const name = 'ASM_backup_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const file = DriveApp.getFileById(TARGET_SPREADSHEET_ID).makeCopy(name, folder);
  Logger.log('バックアップ作成: ' + name + ' (' + file.getId() + ')');

  const removed = pruneOldBackups_(folder);
  Logger.log('バックアップ完了。保持世代(' + BACKUP_KEEP_GENERATIONS + ')を超えた' + removed + '件を削除しました。');
}

function getOrCreateBackupFolder_() {
  const it = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(BACKUP_FOLDER_NAME);
}

function pruneOldBackups_(folder) {
  const files = [];
  const it = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (it.hasNext()) {
    const f = it.next();
    files.push({ file: f, created: f.getDateCreated().getTime() });
  }
  files.sort(function (a, b) { return b.created - a.created; }); // 新しい順

  var removed = 0;
  for (var i = BACKUP_KEEP_GENERATIONS; i < files.length; i++) {
    files[i].file.setTrashed(true);
    removed++;
  }
  return removed;
}
