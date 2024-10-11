/**
 * 根據學號或姓名更新學生的出席狀態
 * @param {string} studentId 學號
 * @param {string} studentName 學生姓名
 * @param {string} week 週次 (W1 到 W27)
 * @param {string} status 出席狀態 (true 或 false)
 * @return {string} 更新結果
 */
function updateStudentAttendance(studentId, studentName, week, status) {
    // 根據選擇的週次取得對應的工作表
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(week);
    
    // 取得 A7:R142 範圍內的所有數據，包含學號、姓名及 true/false 數據
    var data = sheet.getRange('A7:R142').getValues();
    
    var rowToUpdate = null;

    // 根據學號或姓名尋找對應的行
    for (var i = 0; i < data.length; i++) {
        if (data[i][2] == studentId || data[i][3] == studentName) { // 第三列為學號，第四列為姓名
            rowToUpdate = i + 7;  // 加 7 來計算實際的行號
            break;
        }
    }

    if (rowToUpdate) {
        // 更新 E 到 R 列 (14 個 true/false 狀態)
        sheet.getRange(rowToUpdate, 5, 1, 14).setValues([Array(14).fill(status === 'true')]);
        return '成功更新讀經 ' + (studentName || studentId) + ' 的狀態。';
    } else {
        return '找不到對應的學生！';
    }
}
// 顯示 HTML 頁面
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index').setTitle('2024F二年級讀經進度表');
}
