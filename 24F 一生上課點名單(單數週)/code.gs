// 顯示 HTML 頁面
function doGet(e) {
  if(e.parameter.page){
    var pageName = e.parameter.page.trim().toLowerCase();
    if (pageName !== "home"){
      var template = HtmlService.createTemplateFromFile(pageName);
      template.url = getPageUrl();
      return template.evaluate();
    }else{
      return homePage();
    }
  }else{
    return homePage();
  }
}

function homePage(){
  var pages = ["修改出席紀錄"];
  var urls = pages.map(function(name){
    return getPageUrl(name);
  });
  var template = HtmlService.createTemplateFromFile("home");
  template.urls = urls;
  return template.evaluate();
}

function getPageUrl(name){
  if (name){
    var url = ScriptApp.getService().getUrl();
    return url + "?page=" + name;
  }else{
    return ScriptApp.getService().getUrl();
  }
}
// 獲取特定群組的點名清單
function getAttendeesByGroup(group) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // 自動取得工作簿 ID
    var sheet = spreadsheet.getSheetByName(group);
    var attendees = sheet.getRange("B4:B40").getValues(); // 假設點名清單位於B4到B40
    attendees = attendees.flat().filter(String); // 展平和過濾空值
    return attendees;
  } catch (error) {
    Logger.log("Error getting attendees by group: " + error.message);
    throw error;
  }
}

// 更新或附加多行數據
function updateOrAppendMultipleRows(group, name, date, time, attendees) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // 自動取得工作簿 ID
    var sheet = spreadsheet.getSheetByName(group);

    // 格式化時間為四位數字（例如：07:10 -> 0710）
    time = time.replace(":", "");

    // 查找下一個可用的空白列 (從C列開始)
    var col = 3; // 開始檢查的列，3 = C列
    while (sheet.getRange(3, col).getValue() !== "") {
      col++;
    }

    // 在指定儲存格中填寫點名者、日期和時間
    sheet.getRange("B2").setValue(name); // 點名者 (B2)
    sheet.getRange("B1").setValue(date); // 日期 (B1)
    sheet.getRange(3, col).setValue(time); // 時間 (C3:J3 中的下一個可用列)

    // 填寫狀態
    var sheetAttendees = sheet.getRange("B4:B40").getValues().flat(); // 假設點名清單位於B4到B40
    attendees.forEach(function(item) {
      var rowIndex = sheetAttendees.indexOf(item.attendee);
      if (rowIndex !== -1) {
        sheet.getRange(rowIndex + 4, col).setValue(item.status); // 將狀態填入相應的儲存格
      }
    });
  } catch (error) {
    Logger.log("Error updating or appending rows: " + error.message);
    throw error;
  }
}

// 獲取出席紀錄

// 獲取指定週次的出席紀錄
function getAttendanceRecords(group) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(group);  // 根據選擇的週次來獲取工作表
    if (!sheet) {
      Logger.log("找不到工作表: " + group);
      return [];
    }
    var records = [];
    var data = sheet.getRange('a3:L33').getValues();  // 調整範圍讀取名字、編號和課程狀態
    data.forEach(function(row) {
      records.push({
        studentID: row[0],   // 學號
        studentName: row[1], // 學生名字
        period1: row[2],     // 課一
        period2: row[3],     // 課二
        period3: row[4],     // 課三
        period4: row[5],     // 課四
        period5: row[6],     // 課五
        period6: row[7],     // 課六
        period7: row[8],     // 課七
        period8: row[9]      // 課八
      });
    });
    return records;
  } catch (error) {
    Logger.log("獲取出席紀錄時發生錯誤: " + error.message);
    throw error;
  }
}



// 更新指定週次的出席狀態
function updateAttendanceStatus(group, statusUpdates) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(group);
    if (!sheet) {
      Logger.log("找不到工作表: " + group);
      return false;
    }

    // 假設 statusUpdates 是一個包含 { rowIndex, columnIndex, newStatus } 的數組
    statusUpdates.forEach(function(update) {
      sheet.getRange(update.rowIndex + 3, update.columnIndex + 2).setValue(update.newStatus);  // 根據行和列索引更新狀態
    });

    return true;
  } catch (error) {
    Logger.log("更新出席狀態時發生錯誤: " + error.message);
    return false;
  }
}




// 獲取工作簿名稱
function getSpreadsheetName() {
  return SpreadsheetApp.getActiveSpreadsheet().getName();
}
