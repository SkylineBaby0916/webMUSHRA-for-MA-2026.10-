// ============================================================================
// Google Apps Script (GAS) - webMUSHRA 小提琴音色問卷資料自動寫入後端
// ============================================================================
// 使用方法：
// 1. 打開你的 Google Sheet 試算表
// 2. 點擊頂部選單「擴充功能」 -> 「Apps Script」
// 3. 將本程式碼全部複製貼上，覆蓋原有的 Code.gs
// 4. 點擊右上角「部署」 -> 「新建部署」
// 5. 類型選擇「網頁應用程式 (Web app)」
// 6. 「誰可以存取 (Who has access)」選擇：【所有人 (Anyone)】  <--- 極度重要！
// 7. 點擊「部署」，複製產生的「網頁應用程式網址 (Web app URL)」
// ============================================================================

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // 防並發鎖定 10 秒
  
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // 如果是全新空白工作表，自動寫入表頭 Header (共 35 欄)
    if (sheet.getLastRow() === 0) {
      var headers = [
        "填寫時間",
        "問卷語言",
        // Chapter 1: Sarasate (流浪者之歌)
        "Sarasate_飽滿度_Reference(Original)",
        "Sarasate_飽滿度_3.5k_Anchor",
        "Sarasate_飽滿度_Inverse_Filter",
        "Sarasate_飽滿度_Titian_With_Res",
        "Sarasate_飽滿度_Titian_No_Res",
        "Sarasate_真實度_Reference(Original)",
        "Sarasate_真實度_3.5k_Anchor",
        "Sarasate_真實度_Inverse_Filter",
        "Sarasate_真實度_Titian_With_Res",
        "Sarasate_真實度_Titian_No_Res",
        // Chapter 2: Mendelssohn (孟德爾頌)
        "Mendelssohn_明亮度_Reference(Original)",
        "Mendelssohn_明亮度_3.5k_Anchor",
        "Mendelssohn_明亮度_Inverse_Filter",
        "Mendelssohn_明亮度_Titian_With_Res",
        "Mendelssohn_明亮度_Titian_No_Res",
        "Mendelssohn_真實度_Reference(Original)",
        "Mendelssohn_真實度_3.5k_Anchor",
        "Mendelssohn_真實度_Inverse_Filter",
        "Mendelssohn_真實度_Titian_With_Res",
        "Mendelssohn_真實度_Titian_No_Res",
        // Chapter 3: Kreisler (莫札特迴旋曲)
        "Kreisler_清晰度_Reference(Original)",
        "Kreisler_清晰度_3.5k_Anchor",
        "Kreisler_清晰度_Inverse_Filter",
        "Kreisler_清晰度_Titian_With_Res",
        "Kreisler_清晰度_Titian_No_Res",
        "Kreisler_真實度_Reference(Original)",
        "Kreisler_真實度_3.5k_Anchor",
        "Kreisler_真實度_Inverse_Filter",
        "Kreisler_真實度_Titian_With_Res",
        "Kreisler_真實度_Titian_No_Res",
        // 背景問卷調查
        "Q1_音樂與樂器背景",
        "Q2_演奏學習年資",
        "Q3_所屬機關團體"
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e9ecef");
    }
    
    // 解析送來的 JSON 資料
    var data = JSON.parse(e.postData.contents);
    
    var row = [
      data.timestamp || Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm:ss"),
      data.language || "zh",
      
      // Sarasate 飽滿度
      data.sarasate_warmth_reference,
      data.sarasate_warmth_low_anchor,
      data.sarasate_warmth_inv_filter,
      data.sarasate_warmth_gha_with_res,
      data.sarasate_warmth_gha_no_res,
      
      // Sarasate 真實度
      data.sarasate_realism_reference,
      data.sarasate_realism_low_anchor,
      data.sarasate_realism_inv_filter,
      data.sarasate_realism_gha_with_res,
      data.sarasate_realism_gha_no_res,
      
      // Mendelssohn 明亮度
      data.mendelssohn_brightness_reference,
      data.mendelssohn_brightness_low_anchor,
      data.mendelssohn_brightness_inv_filter,
      data.mendelssohn_brightness_gha_with_res,
      data.mendelssohn_brightness_gha_no_res,
      
      // Mendelssohn 真實度
      data.mendelssohn_realism_reference,
      data.mendelssohn_realism_low_anchor,
      data.mendelssohn_realism_inv_filter,
      data.mendelssohn_realism_gha_with_res,
      data.mendelssohn_realism_gha_no_res,
      
      // Kreisler 清晰度
      data.kreisler_clarity_reference,
      data.kreisler_clarity_low_anchor,
      data.kreisler_clarity_inv_filter,
      data.kreisler_clarity_gha_with_res,
      data.kreisler_clarity_gha_no_res,
      
      // Kreisler 真實度
      data.kreisler_realism_reference,
      data.kreisler_realism_low_anchor,
      data.kreisler_realism_inv_filter,
      data.kreisler_realism_gha_with_res,
      data.kreisler_realism_gha_no_res,
      
      // 背景調查
      data.bg_q1 || "",
      data.bg_q2 || "",
      data.bg_q3 || ""
    ];
    
    sheet.appendRow(row);
    
    return ContentService.createTextOutput(JSON.stringify({ "result": "success" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "result": "error", "error": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
