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
// 📧 売上データ（DB）と作業者の気づきを合体させて日報メールを送るメイン関数
function sendDailyReportMail() {
  // =================================================================
  // 【設定エリア】メールの送り先アドレスを指定してください
  // =================================================================
  var TO_EMAIL = "honoki@matsujiro.co.jp"; // 👈 メインの送信先
  var CC_EMAIL = "";     // 👈 Ccのアドレス（なければ "" でOK）
  // =================================================================
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var today = new Date();
  var todayStr = Utilities.formatDate(today, "JST", "yyyy/MM/dd");
  var currentMonthStr = Utilities.formatDate(today, "JST", "yyyy/MM"); // 当月集計用の文字列（例：2026/06）
  
  // 各データ蓄積シートの定義
  var sheetCenter = ss.getSheetByName("データ蓄積_松阪センター");
  var sheetShop = ss.getSheetByName("データ蓄積_店舗A");
  
  // 集計用変数の初期化（すべて0からスタート）
  var todaySales = 0;
  var monthSales = 0;
  var orderCount = 0;
  var shippingCount = 0;
  
  // -----------------------------------------------------------------
  // 📊 ① 松阪センターシートの集計（A列:日付, C列:金額, F列:受注数, G列:出荷数 と仮定）
  // -----------------------------------------------------------------
  if (sheetCenter) {
    var data = sheetCenter.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var rowDate = data[i][0];
      if (!rowDate) continue;
      
      // 日付オブジェクト、または文字列を判定してフォーマットを統一
      var rowDateStr = (rowDate instanceof Date) ? Utilities.formatDate(rowDate, "JST", "yyyy/MM/dd") : rowDate.toString();
      
      // 【本日】の集計
      if (rowDateStr === todayStr) {
        todaySales += Number(data[i][2] || 0);    // C列：金額
        orderCount += Number(data[i][5] || 0);    // F列：受注数
        shippingCount += Number(data[i][6] || 0); // G列：出荷数
      }
      // 【当月累計】の集計（年月が一致するもの）
      if (rowDateStr.indexOf(currentMonthStr) === 0) {
        monthSales += Number(data[i][2] || 0);
      }
    }
  }

  // -----------------------------------------------------------------
  // 📊 ② 店舗A（本店・おはらい町）シートの集計（A列:日付, C列:金額 と仮定）
  // -----------------------------------------------------------------
  if (sheetShop) {
    var data = sheetShop.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var rowDate = data[i][0];
      if (!rowDate) continue;
      
      var rowDateStr = (rowDate instanceof Date) ? Utilities.formatDate(rowDate, "JST", "yyyy/MM/dd") : rowDate.toString();
      
      // 【本日】の集計
      if (rowDateStr === todayStr) {
        todaySales += Number(data[i][2] || 0); // C列：金額
        // 店舗側に受注・出荷数がある場合は、該当する列番号を足すロジックを追加してください
      }
      // 【当月累計】の集計
      if (rowDateStr.indexOf(currentMonthStr) === 0) {
        monthSales += Number(data[i][2] || 0);
      }
    }
  }

  // 昨対比・予算比の計算（※マスタや前年データシートがあれば自動化可能。現在は仮数値を設定）
  var budgetRatio = "100.0%"; 
  var YoYRatio = "100.0%"; 

  // -----------------------------------------------------------------
  // 📝 ③ 【非定型データ】作業者の自由欄（気づき）を取得
  // -----------------------------------------------------------------
  var inputSheet = ss.getSheetByName("日報入力");
  var workerData = "";
  
  if (inputSheet) {
    var rows = inputSheet.getRange(2, 1, 3, 3).getValues(); // 2行目から3名分、3列を取得
    for (var i = 0; i < rows.length; i++) {
      var name = rows[i][0];
      var work = rows[i][1] ? rows[i][1] : "（未入力）";
      var note = rows[i][2] ? rows[i][2] : "（未入力）";
      
      if (!name) continue; // 名前が空欄ならスキップ
      
      workerData += "#### 👤 " + name + "\n" +
                    "* **本日の作業内容：**\n  " + work + "\n" +
                    "* **気づき・コメント：**\n  " + note + "\n\n";
    }
  } else {
    workerData = "※「日報入力」シートが見つからないため、自由欄をスキップしました。\n\n";
  }
  
  // -----------------------------------------------------------------
  // ✉️ ④ メール本文（テキスト形式）の組み立てと送信
  // -----------------------------------------------------------------
  var subject = "【売上・業務日報】" + todayStr + " 全社速報";
  
  var body = "関係者各位\n\nお疲れ様です。本日の売上、および業務状況を報告いたします。\n\n" +
             "### 📊 1. 本日の業績速報（定型データ）\n" +
             "・本日売上高： ¥" + todaySales.toLocaleString() + " (昨対比: " + YoYRatio + " / 予算比: " + budgetRatio + ")\n" +
             "・当月累計売上： ¥" + monthSales.toLocaleString() + "\n" +
             "・本日受注数： " + orderCount + " 件 / 本日出荷数： " + shippingCount + " 件\n\n" +
             "---\n\n" +
             "### 📝 2. 現場の作業内容・気づき（自由欄）\n\n" + workerData +
             "以上、よろしくお願いいたします。";
             
  // メールの実際の送信処理
  MailApp.sendEmail({
    to: TO_EMAIL,
    cc: CC_EMAIL,
    subject: subject,
    body: body
  });
  
  Logger.log(todayStr + " の日報メールの自動集計・送信が完了しました！");
}