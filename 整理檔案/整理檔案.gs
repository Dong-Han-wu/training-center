function categorizeFilesInAllFolders() {
  // 要分類的檔案類型和對應的資料夾名稱
  var fileTypes = {
    'application/vnd.google-apps.document': 'Google Docs',        // Google Docs 文件
    'application/vnd.google-apps.spreadsheet': 'Google Sheets',    // Google Sheets 文件
    'application/vnd.google-apps.presentation': 'Google Slides',   // Google Slides 文件
    'application/pdf': 'PDFs',                                     // PDF 文件
    'image/jpeg': 'Images',                                        // JPEG 圖片
    'image/png': 'Images',                                         // PNG 圖片
    'video/mp4': 'Videos',                                         // MP4 視頻
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'Excel Files', // Excel (.xlsx)
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'Word Files', // Word (.docx)
    'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'PowerPoint Files', // PowerPoint (.pptx)
    'application/msword': 'Word Files',                            // Word (.doc)
    'application/vnd.ms-excel': 'Excel Files',                     // Excel (.xls)
    'application/vnd.ms-powerpoint': 'PowerPoint Files',           // PowerPoint (.ppt)
    'application/vnd.google-apps.script': 'Google Apps Scripts',   // Google Apps Script JSON 文件
    'application/vnd.google-apps.form': 'Google Forms',            // Google Forms 文件
    'application/vnd.google-apps.map': 'Google My Maps',           // Google My Maps 文件
    'application/vnd.google-apps.site': 'Google Sites',            // Google Sites 文件
    'application/x-iwork-pages-sffpages': 'Apple Pages Files',     // Apple Pages 文件
    'application/x-iwork-numbers-sffnumbers': 'Apple Numbers Files', // Apple Numbers 文件
    'application/x-iwork-keynote-sffkey': 'Apple Keynote Files',   // Apple Keynote 文件
    'application/x-apple-diskimage': 'Disk Images',                // macOS 鏡像檔案 (.dmg)
    'application/zip': 'Compressed Files',                         // 壓縮檔案 (.zip)
    'video/quicktime': 'Videos',                                   // QuickTime 視頻 (.mov)
  };
  
  // 從根目錄開始掃描
  var rootFolder = DriveApp.getRootFolder();
  categorizeFilesInFolder(rootFolder, fileTypes);
}

// 遞迴掃描資料夾及其子資料夾中的檔案，略過指定的資料夾
function categorizeFilesInFolder(folder, fileTypes) {
  // 忽略的資料夾名稱
  var skipFolders = ['Videos', 'Images', 'PDFs', 'Google Slides', 'Google Sheets', 'Google Docs', 'Google Apps Scripts', 'Google Forms', 'Google My Maps', 'Google Sites', 'Apple Pages Files', 'Apple Numbers Files', 'Apple Keynote Files', 'Disk Images', 'Compressed Files', 'Text Files', 'Rich Text Files', 'Google 相簿'];
  
  // 取得當前資料夾中的檔案
  var files = folder.getFiles();
  
  while (files.hasNext()) {
    var file = files.next();
    var mimeType = file.getMimeType();
    
    if (fileTypes[mimeType]) {
      // 根據 MIME 類型獲取資料夾名稱
      var folderName = fileTypes[mimeType];
      
      // 檢查或創建目標資料夾
      var targetFolder = getOrCreateFolder(folderName);
      
      // 檢查檔案是否已經在正確的資料夾內
      if (!(file.getParents().hasNext() && file.getParents().next().getId() === targetFolder.getId())) {
        // 將檔案移動到對應的資料夾
        file.moveTo(targetFolder);
        Logger.log('Moved file: ' + file.getName() + ' to folder: ' + folderName);
      }
    }
  }
  
  // 取得當前資料夾中的子資料夾並遞迴掃描
  var subfolders = folder.getFolders();
  
  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    
    // 如果子資料夾在略過的資料夾清單中，則略過
    if (skipFolders.indexOf(subfolder.getName()) !== -1) {
      Logger.log('Skipping folder: ' + subfolder.getName());
      continue;  // 略過這個資料夾
    }
    
    // 遞迴處理子資料夾
    categorizeFilesInFolder(subfolder, fileTypes);
  }
}

// 如果資料夾不存在，則創建它
function getOrCreateFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}
