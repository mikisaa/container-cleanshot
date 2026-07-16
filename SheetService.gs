/**
 * SheetService.gs
 * スプレッドシートの読み書きおよび初期化処理を担当します。
 */

/**
 * スプレッドシートの全シートを初期化し、ヘッダーとマスタデータを設定します。
 * 既存のシートがある場合は、データをクリアして再設定します。
 */
function initSpreadsheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. Settings シートの作成と初期データ登録
  setupSheet(ss, SHEET_NAMES.SETTINGS, ["Key", "Value"], [
    ["ROOT_FOLDER_ID", "1NMrLsEye_L2KyLpqYd5xeMOmB9OOfUgB"],
    ["APP_NAME", "Container CleanShot"],
    ["VERSION", "0.1.0"]
  ]);

  // 2. Jobs シートの作成とヘッダー設定
  setupSheet(ss, SHEET_NAMES.JOBS, [
    "JobId", "StartedAt", "CompletedAt", "Status", "PhotoCount", "CreatedBy", "UserAgent"
  ]);

  // 3. RequiredShots シートの作成と初期データ登録
  setupSheet(ss, SHEET_NAMES.REQUIRED_SHOTS, [
    "ShotKey", "ShotNo", "ShotName", "SortOrder", "Required", "OverlayType", "Active"
  ], [
    ["BEFORE_FULL", 11, "清掃前のコンテナ全体", 1, true, "WIDE", true],
    ["BACK_WALL", 1, "コンテナ奥の壁", 2, true, "BACK", true],
    ["LEFT_BACK", 2, "左壁面・奥半分", 3, true, "LEFT_BACK", true],
    ["RIGHT_BACK", 3, "右壁面・奥半分", 4, true, "RIGHT_BACK", true],
    ["CEILING_BACK", 4, "天井・奥", 5, true, "CEILING_BACK", true],
    ["FLOOR_BACK", 5, "床・奥", 6, true, "FLOOR_BACK", true],
    ["LEFT_FRONT", 6, "左壁面・手前半分", 7, true, "LEFT_FRONT", true],
    ["RIGHT_FRONT", 7, "右壁面・手前半分", 8, true, "RIGHT_FRONT", true],
    ["CEILING_FRONT", 8, "天井・手前", 9, true, "CEILING_FRONT", true],
    ["FLOOR_FRONT", 9, "床・手前", 10, true, "FLOOR_FRONT", true],
    ["AFTER_FULL", 12, "コンテナ全体", 11, true, "WIDE", true],
    ["CONTAINER_NO", 10, "コンテナNo.", 12, true, "BOX", true]
  ]);

  // 4. ChecklistMaster シートの作成と初期データ登録
  setupSheet(ss, SHEET_NAMES.CHECKLIST_MASTER, [
    "CheckKey", "Label", "SortOrder", "Active"
  ], [
    ["FLOOR_TRASH", "床面のゴミ・木片・釘を除去した", 1, true],
    ["CONDENSATION", "天井や壁面に結露がないことを確認した", 2, true],
    ["FLOOR_WET_MUD", "床面の水濡れ・泥汚れを確認した", 3, true],
    ["FOOTPRINTS", "床面に靴跡がないことを確認した", 4, true]
  ]);

  // 5. JobChecklist シートの作成とヘッダー設定
  setupSheet(ss, SHEET_NAMES.JOB_CHECKLIST, [
    "JobId", "CheckKey", "Checked", "CheckedAt"
  ]);

  // 6. Photos シートの作成とヘッダー設定
  setupSheet(ss, SHEET_NAMES.PHOTOS, [
    "JobId", "ShotKey", "ShotNo", "ShotName", "FileId", "FileUrl", "FileName", "CapturedAt", "RetryNo", "IsLatest"
  ]);

  // 7. Logs シートの作成とヘッダー設定
  setupSheet(ss, SHEET_NAMES.LOGS, [
    "Timestamp", "Level", "JobId", "Action", "Message", "Detail"
  ]);

  console.log("スプレッドシートの初期化が正常に完了しました。");
}

/**
 * 個別のシートを作成・初期化するヘルパー関数
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} sheetName - 対象シート名
 * @param {string[]} headers - ヘッダー（1行目）のカラム名配列
 * @param {Array[]} [initialData] - 初期投入するデータの二次元配列（任意）
 */
function setupSheet(ss, sheetName, headers, initialData) {
  let sheet = ss.getSheetByName(sheetName);
  
  if (sheet) {
    // 既存シートがある場合はクリア
    sheet.clear();
  } else {
    // シートがない場合は作成
    sheet = ss.insertSheet(sheetName);
  }
  
  // ヘッダーを書き込み
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // 初期データがある場合は書き込み
  if (initialData && initialData.length > 0) {
    sheet.getRange(2, 1, initialData.length, headers.length).setValues(initialData);
  }
  
  // 1行目を太字・固定にするなどの装飾処理
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
}

/**
 * 撮影マスタから指定された ShotKey の情報を取得します。
 * @param {string} shotKey - 撮影キー
 * @return {Object|null} 撮影項目情報 { shotNo, shotName }
 */
function getRequiredShotInfo(shotKey) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.REQUIRED_SHOTS);
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  // ヘッダーを除いて検索 (0: ShotKey, 1: ShotNo, 2: ShotName)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === shotKey) {
      return {
        shotNo: data[i][1],
        shotName: data[i][2]
      };
    }
  }
  return null;
}

/**
 * Photos シートから対象の JobId ＋ ShotKey の過去撮影回数（リトライ数）を取得します。
 * @param {string} jobId - 作業ID
 * @param {string} shotKey - 撮影キー
 * @return {number} リトライ回数（既存レコード数）
 */
function getRetryCount(jobId, shotKey) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.PHOTOS);
  if (!sheet) return 0;
  
  const data = sheet.getDataRange().getValues();
  let count = 0;
  // ヘッダーを除いて検索 (0: JobId, 1: ShotKey)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === jobId && data[i][1] === shotKey) {
      count++;
    }
  }
  return count;
}

/**
 * Photos シートに撮影記録を追加します。
 * 同一ジョブ・同一撮影キーの古い写真レコードがあれば、それらの IsLatest を false に更新します。
 * @param {Object} record - 登録レコード情報
 */
function savePhotoRecord(record) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.PHOTOS);
  if (!sheet) throw new Error("Photosシートが見つかりません。");
  
  const range = sheet.getDataRange();
  const data = range.getValues();
  
  // 1. 同一の JobId かつ ShotKey を持つ既存レコードの IsLatest (列10 / インデックス9) を false に更新
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === record.jobId && data[i][1] === record.shotKey) {
      // IsLatest を false に設定 (スプレッドシート上は 1-indexed なので行は i+1, 列は 10)
      sheet.getRange(i + 1, 10).setValue(false);
    }
  }
  
  // 2. 新しいレコードを追加
  const newRow = [
    record.jobId,
    record.shotKey,
    record.shotNo,
    record.shotName,
    record.fileId,
    record.fileUrl,
    record.fileName,
    new Date(), // CapturedAt
    record.retryNo,
    record.isLatest
  ];
  sheet.appendRow(newRow);
}

/**
 * Jobs シート上の対象 JobId レコードの PhotoCount（列5）を、最新写真のユニーク数に更新します。
 * @param {string} jobId - 作業ID
 */
function updateJobPhotoCount(jobId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const photosSheet = ss.getSheetByName(SHEET_NAMES.PHOTOS);
  const jobsSheet = ss.getSheetByName(SHEET_NAMES.JOBS);
  if (!photosSheet || !jobsSheet) return;
  
  // 最新の有効写真数をカウント
  const photosData = photosSheet.getDataRange().getValues();
  let latestPhotoCount = 0;
  for (let i = 1; i < photosData.length; i++) {
    if (photosData[i][0] === jobId && photosData[i][9] === true) { // IsLatest === true
      latestPhotoCount++;
    }
  }
  
  // Jobs シートの該当行を更新 (0: JobId, 4: PhotoCount)
  const jobsData = jobsSheet.getDataRange().getValues();
  for (let j = 1; j < jobsData.length; j++) {
    if (jobsData[j][0] === jobId) {
      jobsSheet.getRange(j + 1, 5).setValue(latestPhotoCount);
      break;
    }
  }
}

/**
 * Jobs シートに新規作業レコードを作成します。
 * @param {string} jobId - 生成された作業ID
 * @param {string} userAgent - ブラウザのユーザーエージェント情報
 */
function createJob(jobId, userAgent) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.JOBS);
  if (!sheet) throw new Error("Jobsシートが見つかりません。");
  
  // カラム: JobId, StartedAt, CompletedAt, Status, PhotoCount, CreatedBy, UserAgent
  const newRow = [
    jobId,
    new Date(), // StartedAt
    "",         // CompletedAt (未完了)
    "STARTED",  // Status
    0,          // PhotoCount
    "Anonymous",// CreatedBy
    userAgent   // UserAgent
  ];
  sheet.appendRow(newRow);
}

/**
 * RequiredShots マスタから、有効な（Active=TRUE）撮影項目リストを取得し、
 * SortOrder 順にソートして返します。
 * @return {Array} 撮影項目リスト
 */
function getActiveRequiredShots() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.REQUIRED_SHOTS);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const shots = [];
  
  // 1. データをオブジェクトに変換し、有効なものだけ抽出
  for (let i = 1; i < data.length; i++) {
    const active = data[i][6];
    if (active === true) {
      shots.push({
        ShotKey: data[i][0],
        ShotNo: data[i][1],
        ShotName: data[i][2],
        SortOrder: data[i][3],
        Required: data[i][4],
        OverlayType: data[i][5]
      });
    }
  }
  
  // 2. SortOrder 順にソート
  shots.sort((a, b) => a.SortOrder - b.SortOrder);
  return shots;
}

/**
 * JobChecklist シートにチェックリスト結果を保存します。
 * また、Jobs シートのステータスを CHECKLIST_DONE に更新します。
 * @param {string} jobId - 作業ID
 * @param {Object} checklistItems - チェック状態オブジェクト { CheckKey: Checked }
 */
function saveJobChecklist(jobId, checklistItems) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const checklistSheet = ss.getSheetByName(SHEET_NAMES.JOB_CHECKLIST);
  const jobsSheet = ss.getSheetByName(SHEET_NAMES.JOBS);
  if (!checklistSheet || !jobsSheet) throw new Error("必要なシートが見つかりません。");
  
  // 1. 各項目のチェック結果を記録
  const now = new Date();
  for (const checkKey in checklistItems) {
    // カラム: JobId, CheckKey, Checked, CheckedAt
    const newRow = [
      jobId,
      checkKey,
      checklistItems[checkKey],
      now
    ];
    checklistSheet.appendRow(newRow);
  }
  
  // 2. Jobsシートのステータスを CHECKLIST_DONE に更新 (0: JobId, 3: Status)
  const jobsData = jobsSheet.getDataRange().getValues();
  for (let i = 1; i < jobsData.length; i++) {
    if (jobsData[i][0] === jobId) {
      jobsSheet.getRange(i + 1, 4).setValue("CHECKLIST_DONE");
      break;
    }
  }
}

/**
 * Jobs シート上の対象ジョブステータスを COMPLETED（完了）に更新します。
 * @param {string} jobId - 作業ID
 */
function markJobCompleted(jobId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.JOBS);
  if (!sheet) throw new Error("Jobsシートが見つかりません。");
  
  const data = sheet.getDataRange().getValues();
  // 該当するJobIdを検索して更新 (2: CompletedAt, 3: Status)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === jobId) {
      sheet.getRange(i + 1, 3).setValue(new Date()); // CompletedAt
      sheet.getRange(i + 1, 4).setValue("COMPLETED"); // Status
      break;
    }
  }
}

/**
 * 対象の JobId が完了条件（必須写真12枚、チェック4項目すべて完了）を満たしているか検証します。
 * @param {string} jobId - 作業ID
 * @return {Object} 検証結果 { success: boolean, message: string }
 */
function validateJobCompletion(jobId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. アクティブな必須撮影項目キーを取得
  const requiredShots = getActiveRequiredShots();
  const requiredKeys = requiredShots.map(s => s.ShotKey);
  
  // 2. Photos シートから、該当ジョブで IsLatest = true の写真を取得
  const photosSheet = ss.getSheetByName(SHEET_NAMES.PHOTOS);
  const photosData = photosSheet ? photosSheet.getDataRange().getValues() : [];
  const capturedKeys = [];
  for (let i = 1; i < photosData.length; i++) {
    if (photosData[i][0] === jobId && photosData[i][9] === true) { // IsLatest === true
      capturedKeys.push(photosData[i][1]); // ShotKey
    }
  }
  
  // 3. マスタのチェック項目を取得
  const checklistMasterSheet = ss.getSheetByName(SHEET_NAMES.CHECKLIST_MASTER);
  const masterCheckData = checklistMasterSheet ? checklistMasterSheet.getDataRange().getValues() : [];
  const requiredCheckKeys = [];
  for (let i = 1; i < masterCheckData.length; i++) {
    if (masterCheckData[i][3] === true) { // Active === true
      requiredCheckKeys.push(masterCheckData[i][0]); // CheckKey
    }
  }
  
  // 4. JobChecklist シートから、該当ジョブで Checked = true の項目を取得
  const checklistSheet = ss.getSheetByName(SHEET_NAMES.JOB_CHECKLIST);
  const checklistData = checklistSheet ? checklistSheet.getDataRange().getValues() : [];
  const checkedKeys = [];
  for (let i = 1; i < checklistData.length; i++) {
    if (checklistData[i][0] === jobId && checklistData[i][2] === true) { // Checked === true
      checkedKeys.push(checklistData[i][1]); // CheckKey
    }
  }
  
  // 5. 写真の未撮影項目をチェック
  const missingShots = requiredKeys.filter(k => !capturedKeys.includes(k));
  
  // 6. チェックリストの未チェック項目をチェック
  const missingChecks = requiredCheckKeys.filter(k => !checkedKeys.includes(k));
  
  if (missingShots.length > 0 || missingChecks.length > 0) {
    let msg = "";
    if (missingShots.length > 0) {
      msg += `未撮影の写真があります（${missingShots.length}件）。`;
    }
    if (missingChecks.length > 0) {
      msg += `未完了の清掃チェックがあります（${missingChecks.length}件）。`;
    }
    return {
      success: false,
      message: msg
    };
  }
  
  return {
    success: true,
    message: "完了条件を満たしています。"
  };
}


