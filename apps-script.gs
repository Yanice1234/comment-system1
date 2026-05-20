/**
 * MoodFlow Pro - Google Apps Script (GAS) 資料庫同步腳本
 * 
 * 作用：將網頁的每日心情打卡紀錄，利用 Google Sheet (試算表) 作為雲端資料庫。
 * 請將此代碼複製並貼上至 Google 試算表之「擴充功能」 -> 「Apps Script」中。
 * 
 * 欄位結構：
 * - id (唯一碼)
 * - createdAt (時間)
 * - mood (心情)
 * - reason (原因)
 * - status (狀態："active" 預設，或 "done" 已完成)
 */

// 試算表設定：取得或新建名為 "records" 的工作表
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("records");
  if (!sheet) {
    sheet = ss.insertSheet("records");
    // 初始化首行標題 (Headers)
    sheet.appendRow(["id", "createdAt", "mood", "reason", "status"]);
  }
  return sheet;
}

/**
 * 處理 GET 請求：獲取所有心情打卡紀錄
 */
function doGet(e) {
  try {
    const sheet = getSheet();
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const records = [];

    // 將二維試算表轉化為 JSON 物件陣列
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const record = {};
      headers.forEach((header, index) => {
        record[header] = row[index];
      });
      records.push(record);
    }

    // 依據 status 排序：active 排上面，done 排下面；同 status 則依時間降序 (最新排前面)
    records.sort((a, b) => {
      if (a.status !== b.status) {
        return a.status === 'active' ? -1 : 1;
      }
      return new Date(b.createdAt) - new Date(a.createdAt);
    });

    return ContentService.createTextOutput(JSON.stringify({ 
      status: "success", 
      data: records 
    }))
    .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "error", 
      message: error.toString() 
    }))
    .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 處理 POST 請求：用於新增打卡、切換狀態
 */
function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action || "save"; // action 預設為 "save"，也可為 "toggle"
    const sheet = getSheet();

    if (action === "save") {
      // 1. 新增心情打卡
      const newRecord = postData.record;
      if (!newRecord || !newRecord.mood) {
        throw new Error("缺少心情或記錄資料");
      }

      // 檢查當天是否已經有同日期打卡
      const today = new Date().toLocaleDateString();
      const rows = sheet.getDataRange().getValues();
      const headers = rows[0];
      const createdAtIdx = headers.indexOf("createdAt");
      
      const alreadyCheckedIn = rows.slice(1).some(row => {
        const date = new Date(row[createdAtIdx]);
        return date.toLocaleDateString() === today;
      });

      if (alreadyCheckedIn) {
        return ContentService.createTextOutput(JSON.stringify({
          status: "error",
          message: "今天已經記錄過心情了，明天再開啟新的一天吧！"
        })).setMimeType(ContentService.MimeType.JSON);
      }

      // 寫入新資料列
      sheet.appendRow([
        newRecord.id,
        newRecord.createdAt,
        newRecord.mood,
        newRecord.reason || "",
        newRecord.status || "active"
      ]);

      return ContentService.createTextOutput(JSON.stringify({ 
        status: "success", 
        message: "資料儲存成功",
        record: newRecord
      }))
      .setMimeType(ContentService.MimeType.JSON);

    } else if (action === "toggle") {
      // 2. 切換狀態 (active 變 done，或 done 變 active)
      const targetId = postData.id;
      if (!targetId) {
        throw new Error("缺少目標 ID");
      }

      const rows = sheet.getDataRange().getValues();
      const headers = rows[0];
      const idIdx = headers.indexOf("id");
      const statusIdx = headers.indexOf("status");
      let foundRowIndex = -1;
      let currentStatus = "";

      for (let i = 1; i < rows.length; i++) {
        if (rows[i][idIdx] === targetId) {
          foundRowIndex = i + 1; // Apps Script 試算表索引從 1 開始，標題為 1，第一筆資料為 2
          currentStatus = rows[i][statusIdx];
          break;
        }
      }

      if (foundRowIndex === -1) {
        throw new Error("找不到對應的記錄 ID");
      }

      const newStatus = currentStatus === "active" ? "done" : "active";
      sheet.getRange(foundRowIndex, statusIdx + 1).setValue(newStatus); // 列和行都是 1-index

      return ContentService.createTextOutput(JSON.stringify({ 
        status: "success", 
        newStatus: newStatus 
      }))
      .setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "error", 
      message: error.toString() 
    }))
    .setMimeType(ContentService.MimeType.JSON);
  }
}
