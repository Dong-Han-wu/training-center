function generateFileLinks() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // 資料夾 ID 陣列
  var folderIds = [
    ' ',  // 一

  ];

  var lastRow = sheet.getLastRow();  // 獲取最後一行
  for (var idIndex = 0; idIndex < folderIds.length; idIndex++) {
    var folderId = folderIds[idIndex];
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    
    // 遍歷當前資料夾中的所有檔案
    while (files.hasNext()) {
      var file = files.next();
      var fileName = file.getName();
      var fileUrl = file.getUrl();
      
      // 假設檔名包含學號並與 B 欄中的學號匹配
      for (var i = 1; i <= lastRow; i++) {
        var studentId = sheet.getRange(i, 2).getValue();  // 取得 B 欄學號
        if (fileName.includes(studentId)) {
          sheet.getRange(i, 4).setValue(fileUrl);  // 將連結寫入 D 欄
        }
      }
    }
  }
}

function resizeAllImages() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // 取得所有圖片對象
  var images = sheet.getImages();
  
  // 設定統一的寬度與高度
  var newWidth = 150;  // 新的寬度
  var newHeight = 150; // 新的高度
  
  // 遍歷所有圖片並修改它們的大小
  images.forEach(function(image) {
    image.setWidth(newWidth);
    image.setHeight(newHeight);
  });
}
