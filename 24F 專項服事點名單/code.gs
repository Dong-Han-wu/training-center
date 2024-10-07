// 顯示 HTML 頁面
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('24F 專項服事點名單');
}

// 獲取特定群組的點名清單
function getAttendeesByGroup(group) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(group);
    var attendees = sheet.getRange("B3:B20").getValues(); // 假設點名清單位於B4到B20
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
    sheet.getRange(1, col).setValue(date); //  日期
    sheet.getRange(2, col).setValue(time); // 時間
   

    // 填寫狀態
    var sheetAttendees = sheet.getRange("B3:B20").getValues().flat(); // 調整為正確範圍
    attendees.forEach(function(item) {
      var rowIndex = sheetAttendees.indexOf(item.attendee);
      if (rowIndex !== -1) {
        sheet.getRange(rowIndex + 3, col).setValue(item.status); // 將狀態填入相應的儲存格
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
    var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 21, sheet.getLastColumn()-4);
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


