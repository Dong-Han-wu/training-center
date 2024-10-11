function doGet() {
  // 加載 HTML 文件 'Index' 並顯示網頁
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('24F二年級豫讀統計紀錄');
}

function searchRecords(sheetName, searchQuery) {
  // 自動獲取當前試算表的ID
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);

  // 獲取所有的數據（從第4行開始，假設數據從第4行開始）
  var dataRange = sheet.getRange(4, 1, sheet.getLastRow(), sheet.getLastColumn());
  var data = dataRange.getValues();
  var result = [];
  var rowIndex = -1;  // 用來記錄找到的行索引

  // 迭代數據以查找符合搜索條件的行
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == searchQuery || data[i][1] == searchQuery || data[i][2] == searchQuery || data[i][3] == searchQuery) {
      result = data[i];  // 保存匹配到的行
      rowIndex = i;  // 保存匹配到的行索引
      break;  // 找到匹配後退出循環
    }
  }

  if (rowIndex !== -1) {
    // 獲取第一行的應讀篇數
    var expectedReadings = sheet.getRange(1, 5, 1, sheet.getLastColumn() - 4).getValues()[0];  // 從第5列開始
    // 更新該行對應的數據，跳過小計和豫讀率，且不編輯隱藏的列
    updateReadings(sheetName, rowIndex, expectedReadings);
    return { result: result, expectedReadings: expectedReadings };
  } else {
    return { result: [], expectedReadings: [] };
  }
}

function updateReadings(sheetName, rowIndex, readings) {
  // 自動獲取當前試算表的ID
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);

  // 動態找到「小計」和「豫讀率」的位置
  var headerRange = sheet.getRange(3, 1, 1, sheet.getLastColumn());  // 假設標題在第3行
  var headers = headerRange.getValues()[0];  // 獲取第3行的標題
  var skipColumns = [];  // 用來儲存「小計」和「豫讀率」的列索引

  for (var i = 0; i < headers.length; i++) {
    if (headers[i] == '小計' || headers[i] == '豫讀率') {
      skipColumns.push(i + 1);  // 記錄「小計」和「豫讀率」的列（列索引從1開始）
    }
  }

  // 更新應讀篇數，只更新不在 skipColumns 且不是隱藏的列
  for (var i = 5; i <= sheet.getLastColumn(); i++) {
    if (!skipColumns.includes(i) && !sheet.isColumnHiddenByUser(i)) {  // 略過「小計」、「豫讀率」和隱藏的列
      sheet.getRange(rowIndex + 4, i).setValue(readings[i - 5]);  // 更新數據
    }
  }
}
