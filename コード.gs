function runUltimateDailyImport() {
  // =================================================================
  // 【設定エリア】ご自身の環境の Google ドライブのフォルダIDに書き換えてください
  // =================================================================
  // 📥 各拠点の【投入用】フォルダID
  var FOLDER_MATSUZAKA    = "1T5RvBsnmUZ3lo7_Uyx3KeZTcX1kqRF5M";
  var FOLDER_HON_TEN      = "1JrdUP04kqK90QgSYlS0LpNPtlAMRkgrH";
  var FOLDER_OHARAIMACHI  = "1zqRp_CB4hcbJa1KVMUFTJITEC99BhYxz";
  
  // 📦 各拠点の【処理済み保存用（アーカイブ）】フォルダID
  var ARC_MATSUZAKA       = "1f6wSdJTKEQmOyUcSQMncL8465UEvt7Co";
  var ARC_HON_TEN         = "1UfVr_yycPArMuNLBhdk7gZ2Lo6JAdEcg";
  var ARC_OHARAIMACHI     = "1ofrrfCxLvYAQ6Tvz8qeTC7X_u5KUC-jZ";
  // =================================================================

  // ① 松阪センターの取り込み（アーカイブ先：ARC_MATSUZAKA）
  Logger.log("--- 松阪センターの取り込み開始 ---");
  processFolder(FOLDER_MATSUZAKA, "データ蓄積_松阪センター", "松阪センター", "松阪", ARC_MATSUZAKA);
  
  // ② 松阪本店の取り込み（アーカイブ先：ARC_HON_TEN）
  Logger.log("--- 松阪本店の取り込み開始 ---");
  processFolder(FOLDER_HON_TEN, "データ蓄積_店舗A", "松阪本店", "店舗", ARC_HON_TEN);
  
  // ③ 伊勢おはらい町店の取り込み（アーカイブ先：ARC_OHARAIMACHI）
  Logger.log("--- 伊勢おはらい町店の取り込み開始 ---");
  processFolder(FOLDER_OHARAIMACHI, "データ蓄積_店舗A", "伊勢おはらい町店", "店舗", ARC_OHARAIMACHI);

  Logger.log("すべての取り込み・店舗別アーカイブ処理が完了しました！");
}

// 📦 フォルダ巡回・日付解析・DB書き込み・各店舗フォルダへアーカイブを行う関数
function processFolder(folderId, sheetName, hubName, dataType, archiveFolderId) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  var archiveFolder = DriveApp.getFolderById(archiveFolderId);
  
  if (!sheet) {
    Logger.log("エラー：シート「" + sheetName + "」が見つかりません。");
    return;
  }

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    Logger.log("ファイルを発見しました: " + fileName);
    
    // --- 🔍 日付解析センサー（日本語表記、または西暦8桁の連続数字に対応） ---
    var targetDate = null;
    var jpMatch = fileName.match(/(\d{4})年(\d{1,2})月(?:(\d{1,2})日)?/);
    var numMatch = fileName.match(/(\d{4})(\d{2})(\d{2})/);
    
    if (jpMatch) {
      var year = parseInt(jpMatch[1], 10);
      var month = parseInt(jpMatch[2], 10) - 1;
      if (jpMatch[3]) {
        targetDate = new Date(year, month, parseInt(jpMatch[3], 10));
      } else {
        targetDate = new Date(year, month, 1);
      }
    } else if (numMatch) {
      targetDate = new Date(
        parseInt(numMatch[1], 10),
        parseInt(numMatch[2], 10) - 1,
        parseInt(numMatch[3], 10)
      );
    } else {
      targetDate = file.getDateCreated();
    }
    
    var formattedDate = Utilities.formatDate(targetDate, "JST", "yyyy/MM/dd");
    
    // --- 📥 CSVデータのパース（Shift_JIS対応） ---
    var csvData;
    try {
      csvData = Utilities.parseCsv(file.getBlob().getDataAsString("Shift_JIS"));
    } catch(e) {
      Logger.log("CSVの解析に失敗しました: " + e.toString());
      continue;
    }
    
    var dataToAppend = [];
    
    for (var i = 1; i < csvData.length; i++) {
      var row = csvData[i];
      if (!row[0] || row[0].indexOf("合計") !== -1 || row[0].indexOf("総合計") !== -1) {
        continue;
      }
      
      if (dataType === "松阪") {
        dataToAppend.push([
          formattedDate, hubName, row[2], row[3], row[4], row[5], row[6], row[7]
        ]);
      } else if (dataType === "店舗") {
        dataToAppend.push([
          formattedDate, hubName, row[2], row[3], row[4], row[5], row[6]
        ]);
      }
    }
    
    // シートの末尾に一括追記
    if (dataToAppend.length > 0) {
      var lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, dataToAppend.length, dataToAppend[0].length).setValues(dataToAppend);
      Logger.log(hubName + " のデータを " + dataToAppend.length + " 行追記しました。");
    }
    
    // --- 📦 それぞれ指定された店舗別アーカイブフォルダへ移動 ---
    file.moveTo(archiveFolder);
    Logger.log(fileName + " を【" + hubName + "】用のアーカイブフォルダへ移動しました。");
  }
}