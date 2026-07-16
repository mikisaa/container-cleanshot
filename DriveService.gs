/**
 * DriveService.gs
 * Google Drive上のフォルダ操作および画像ファイルの保存処理を担当します。
 */

/**
 * Settings シートから Google Drive の保存先ルートフォルダIDを取得します。
 * @return {string} ルートフォルダID
 */
function getRootFolderId() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === "ROOT_FOLDER_ID") {
        return data[i][1];
      }
    }
  } catch (err) {
    console.error("ROOT_FOLDER_IDの取得に失敗しました: ", err);
  }
  // 取得失敗時のフォールバック
  return "1NMrLsEye_L2KyLpqYd5xeMOmB9OOfUgB";
}

/**
 * 指定した親フォルダ配下に、同名のフォルダがあれば取得し、なければ新規作成します。
 * @param {GoogleAppsScript.Drive.Folder} parentFolder - 親フォルダオブジェクト
 * @param {string} folderName - 作成または検索したいフォルダ名
 * @return {GoogleAppsScript.Drive.Folder} 対象のフォルダオブジェクト
 */
function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(folderName);
}

/**
 * ジョブ専用の保存先フォルダ階層（年/月/日/JobId）を作成または取得します。
 * @param {string} jobId - 作業ID
 * @return {GoogleAppsScript.Drive.Folder} 作成・取得されたジョブフォルダ
 */
function getOrCreateJobFolder(jobId) {
  const rootId = getRootFolderId();
  let currentFolder = DriveApp.getFolderById(rootId);
  
  // 今日の日付情報を取得してフォルダ階層を決定 (例: 2026/07/2026-07-09)
  const now = new Date();
  const yyyy = Utilities.formatDate(now, "Asia/Tokyo", "yyyy");
  const mm = Utilities.formatDate(now, "Asia/Tokyo", "MM");
  const yyyymmdd = Utilities.formatDate(now, "Asia/Tokyo", "yyyy-MM-dd");
  
  // 階層を順番に作成/取得
  currentFolder = getOrCreateFolder(currentFolder, yyyy);
  currentFolder = getOrCreateFolder(currentFolder, mm);
  currentFolder = getOrCreateFolder(currentFolder, yyyymmdd);
  currentFolder = getOrCreateFolder(currentFolder, jobId);
  
  return currentFolder;
}

/**
 * フロントエンドから送信されたBase64画像データをデコードしてGoogle Driveへ保存し、
 * スプレッドシートへ撮影情報を記録します。
 * 
 * @param {string} jobId - 作業ID
 * @param {string} shotKey - 撮影項目の識別キー (例: CONTAINER_NO)
 * @param {string} base64Image - 画像のBase64データ
 * @return {Object} 保存されたファイルの { fileId, fileUrl }
 */
function savePhoto(jobId, shotKey, base64Image) {
  try {
    logInfo(jobId, "savePhoto", `写真保存処理を開始しました (ShotKey: ${shotKey})`);
    
    // 1. スプレッドシートから撮影項目マスタの情報を取得（追加撮影の場合は動的に生成）
    let shotInfo = null;
    if (shotKey && shotKey.indexOf("EXTRA_") === 0) {
      const extraNum = shotKey.split("_")[1] || "1";
      shotInfo = {
        shotNo: "追加",
        shotName: `追加撮影 ${extraNum}`
      };
    } else {
      shotInfo = getRequiredShotInfo(shotKey);
    }

    if (!shotInfo) {
      throw new Error(`撮影項目マスタに ShotKey: ${shotKey} が登録されていません。`);
    }
    
    // 2. 過去の撮り直し回数（RetryNo）を取得してファイル名を決定
    const retryCount = getRetryCount(jobId, shotKey);
    let fileName = "";
    if (retryCount === 0) {
      fileName = `${shotInfo.shotNo}_${shotInfo.shotName}.jpg`;
    } else {
      fileName = `${shotInfo.shotNo}_${shotInfo.shotName}_retry${retryCount}.jpg`;
    }
    
    // 3. 保存先ジョブフォルダを取得・作成
    const jobFolder = getOrCreateJobFolder(jobId);
    if (!jobFolder) {
      throw new Error("保存先ジョブフォルダの取得または作成に失敗しました。");
    }
    
    // 4. Base64イメージのデコードと保存
    // "data:image/jpeg;base64,xxxx..." の形式からヘッダー部分を切り取る
    const base64Data = base64Image.substring(base64Image.indexOf(",") + 1);
    const decodedData = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedData, "image/jpeg", fileName);
    
    const file = jobFolder.createFile(blob);
    // 誰でも閲覧（リンクを知っている全員）できるように公開設定にする場合は追加設定可能ですが、
    // ここではプライベートな保存とし、ファイルのアクセスURLのみを取得します。
    const fileId = file.getId();
    const fileUrl = file.getUrl();
    
    // 5. スプレッドシートの Photos シートへ記録
    savePhotoRecord({
      jobId: jobId,
      shotKey: shotKey,
      shotNo: shotInfo.shotNo,
      shotName: shotInfo.shotName,
      fileId: fileId,
      fileUrl: fileUrl,
      fileName: fileName,
      retryNo: retryCount,
      isLatest: true
    });
    
    // 6. Jobs シートの保存済み写真カウント（PhotoCount）を更新
    updateJobPhotoCount(jobId);
    
    logInfo(jobId, "savePhoto", `写真保存に成功しました: ${fileName} (FileId: ${fileId})`);
    
    return {
      fileId: fileId,
      fileUrl: fileUrl
    };
    
  } catch (err) {
    logError(jobId, "savePhoto", `写真保存処理中に例外エラーが発生しました: ${err.message}`, err.stack);
    throw err;
  }
}
