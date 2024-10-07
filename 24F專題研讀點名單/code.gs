function doGet(e) {
  if(e.parameter.page){
    var pageName = e.parameter.page.trim().toLowerCase();
    if (pageName !== "home"){
      var template = HtmlService.createTemplateFromFile(pageName);
      template.url = getPageUrl();
      return template.evaluate().setTitle('編輯記錄');
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
  return template.evaluate().setTitle('24F專題研讀點名單');
}


function renderPage(pageName) {
  try {
    var template = HtmlService.createTemplateFromFile(pageName);
    template.url = getPageUrl();
    return template.evaluate();
  } catch (e) {
    return HtmlService.createHtmlOutput("Page not found.");
  }
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
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(group);
    var attendees = sheet.getRange("B4:B20").getValues(); // 假設點名清單位於B4到B20
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
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(group);

    // 格式化時間為四位數字（例如：07:10 -> 0710）
    time = time.replace(":", "");

    // 查找下一個可用的空白列
    var col = 3; // 開始檢查的列，3 = C列
    while (sheet.getRange(2, col).getValue() !== "" || sheet.getRange(3, col).getValue() !== "") {
      col++;
    }

    // 在空白列中填寫點名者、日期和時間
    sheet.getRange(1, col).setValue(name); // 點名者
    sheet.getRange(2, col).setValue(date); // 日期
    sheet.getRange(3, col).setValue(time); // 時間

    // 填寫狀態
    var sheetAttendees = sheet.getRange("B4:B20").getValues().flat(); // 調整為正確範圍
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

function loadAttendanceRecords() {
  const group = document.getElementById('group').value;
  google.script.run.withSuccessHandler(function(records) {
    const recordsDiv = document.getElementById('records');
    recordsDiv.innerHTML = ''; // 清空之前的內容

    if (records && records.length > 0) {
      const table = document.createElement('table');
      table.className = 'table table-bordered';

      // 表格內容
      const tbody = document.createElement('tbody');
      records.forEach(row => {
        const tr = document.createElement('tr');
        row.forEach(cell => {
          const td = document.createElement('td');
          td.innerText = cell || ''; // 保持空白
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      recordsDiv.appendChild(table);
    } else {
      recordsDiv.innerText = '沒有找到紀錄。';
    }
  }).getAttendanceRecords(group);
}

function getAttendanceRecords(group) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(group);

    // Assume records are stored from row 4 and beyond.
    var dataRange = sheet.getRange(3, 1, sheet.getLastRow() - 21, sheet.getLastColumn()-4);
    var records = dataRange.getValues();  // Fetch all the records in this range

    return records;
  } catch (error) {
    Logger.log("Error getting attendance records: " + error.message);
    throw error;
  }
}




function getSpreadsheetName() {
  return SpreadsheetApp.getActiveSpreadsheet().getName();
}

// 顯示修改頁面的 HTML
function doGetEditPage() {
  return HtmlService.createHtmlOutputFromFile('EditPage');
}

// 更新紀錄
function updateAttendanceRecords(group, records) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(group);

    // 限制寫入範圍 C1:V20
    var maxRows = 20;  // 行數
    var maxCols = 22;  // 列數（C 到 V 對應的列數是 3 到 22）

    // 確保紀錄數據不超出 C1:V20 的範圍
    var limitedRecords = records.slice(0, maxRows).map(function(row) {
      return row.slice(2, 2 + maxCols);  // 只取第 3 列 (C) 到第 22 列 (V) 的資料
    });

    // 更新的範圍從 C3 開始，寬度是 C:V
    var range = sheet.getRange(3, 3, limitedRecords.length, limitedRecords[0].length);
    range.setValues(limitedRecords); // 將修改過的紀錄寫回到 C3:V20 的表中
  } catch (error) {
    Logger.log("Error updating attendance records: " + error.message);
    throw error;
  }
}
