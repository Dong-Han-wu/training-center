var saveToRootFolder = true;

function onOpen() {
    // 建立自訂選單
    SpreadsheetApp.getUi()
        .createAddonMenu()
        .addItem('匯出目前工作表', 'exportCurrentSheetAsPDF') // 新增匯出目前工作表功能
        .addItem('Export selected area', 'exportPartAsPDF')
        .addToUi();
}

function exportAsPDF() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    var blob = _getAsBlob(spreadsheet.getUrl())
    _exportBlob(blob, spreadsheet.getName(), spreadsheet)
}


function exportPartAsPDF(week, number) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    var selectedRanges;
    var fileSuffix;
    var week2 = week - 1
    // Log week
    Logger.log("Week value: " + week);

    selectedRanges = extractRanges(week2, number);
    fileSuffix = '-selected';



    // 建立臨時工作表
    var tempSpreadsheet = SpreadsheetApp.create(spreadsheet.getName() + fileSuffix);
    if (!saveToRootFolder) {
        DriveApp.getFileById(tempSpreadsheet.getId()).moveTo(DriveApp.getFileById(spreadsheet.getId()).getParents().next());
    }

    var tempSheets = tempSpreadsheet.getSheets();
    var sheet1 = tempSheets.length > 0 ? tempSheets[0] : undefined;
    SpreadsheetApp.setActiveSpreadsheet(tempSpreadsheet);
    tempSpreadsheet.setSpreadsheetTimeZone(spreadsheet.getSpreadsheetTimeZone());
    tempSpreadsheet.setSpreadsheetLocale(spreadsheet.getSpreadsheetLocale());

    // 將選擇範圍複製到臨時工作表
    for (var i = 0; i < selectedRanges.length; i++) {
        var selectedRange = selectedRanges[i];
        var originalSheet = selectedRange.getSheet();
        var originalSheetName = originalSheet.getName();

        var destSheet = tempSpreadsheet.getSheetByName(originalSheetName);
        if (!destSheet) {
            destSheet = tempSpreadsheet.insertSheet(originalSheetName);
        }

        Logger.log('a1notation=' + selectedRange.getA1Notation());

        // Ensure the destination range matches the size of the original range
        var destRange = destSheet.getRange(
            selectedRange.getRow(),
            selectedRange.getColumn(),
            selectedRange.getNumRows(),
            selectedRange.getNumColumns()
        );

        // Copy values
        var values = selectedRange.getValues();
        destRange.setValues(values);

        // Only apply styles if there are values
        if (values && values.length > 0) {
            destRange.setTextStyles(selectedRange.getTextStyles());
            destRange.setBackgrounds(selectedRange.getBackgrounds());
            destRange.setFontColors(selectedRange.getFontColors());
            destRange.setFontFamilies(selectedRange.getFontFamilies());
            destRange.setFontLines(selectedRange.getFontLines());
            destRange.setFontStyles(selectedRange.getFontStyles());
            destRange.setFontWeights(selectedRange.getFontWeights());
            destRange.setHorizontalAlignments(selectedRange.getHorizontalAlignments());
            destRange.setNumberFormats(selectedRange.getNumberFormats());
            destRange.setTextDirections(selectedRange.getTextDirections());
            destRange.setTextRotations(selectedRange.getTextRotations());
            destRange.setVerticalAlignments(selectedRange.getVerticalAlignments());
            destRange.setWrapStrategies(selectedRange.getWrapStrategies());



        }
    }


    cleanAndMergeCells(tempSpreadsheet, spreadsheet.getActiveSheet(), week, number);

    // 移除空的 Sheet1
    if (sheet1) {
        Logger.log('lastcol = ' + sheet1.getLastColumn() + ', lastrow=' + sheet1.getLastRow());
        if (sheet1 && sheet1.getLastColumn() === 0 && sheet1.getLastRow() === 0) {
            tempSpreadsheet.deleteSheet(sheet1);
        }
    }

    // 匯出臨時工作表為 PDF
    exportAsPDF();

    // 回到原工作表
    SpreadsheetApp.setActiveSpreadsheet(spreadsheet);

    // 刪除臨時工作表
    DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true); // 移到垃圾桶

}


function extractRanges(week2, number) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('福音牧養名單');

    // 定義變數表示範圍
    var startRow = 1;
    var numRows = number;
    var colStart = 1; // A列起始，從1開始計算
    // Log week
    Logger.log("Week value: " + week2);
    return [
        sheet.getRange(startRow, colStart, numRows, 7),
        sheet.getRange(startRow, week2 * 7 + 1, numRows, 7),
    ];
}


function cleanAndMergeCells(tempSpreadsheet, originalSheet, week, number) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var tempSheet = tempSpreadsheet.getSheetByName('福音牧養名單');
    Logger.log("Week value: " + week);

    if (tempSheet) {
        Logger.log("Merging cells in ranges...");

        var startRowW3 = 2;
        var numRows1 = 1,
            numRows2 = 2;
        var week = week - 1
        var colStart = week * 7 + 1;
        // 合併 A1:B2
        preserveAndMerge(tempSheet, 1, 1, 2, 2, "Merged A1:B2");

        // 合併 C1:G1
        preserveAndMerge(tempSheet, 1, 3, 1, 5, "Merged C1:G1");

        // 合併 E2:F2
        preserveAndMerge(tempSheet, 2, 5, 1, 2, "Merged E2:F2");

        // 合併 E7:F7
        preserveAndMerge(tempSheet, 7, 5, 1, 2, "Merged E7:F7");

        // 合併 G7:G9
        preserveAndMerge(tempSheet, 7, 7, 3, 1, "Merged G7:G9");

        // 合併 A8:A9
        preserveAndMerge(tempSheet, 8, 1, 2, 1, "Merged A8:A9");

        // 合併 B8:B9
        preserveAndMerge(tempSheet, 8, 2, 2, 1, "Merged B8:B9");

        // 合併 C8:C9
        preserveAndMerge(tempSheet, 8, 3, 2, 1, "Merged C8:C9");

        // 合併 D8:D9
        preserveAndMerge(tempSheet, 8, 4, 2, 1, "Merged D8:D9");

        // 合併 E8:F9
        preserveAndMerge(tempSheet, 8, 5, 2, 2, "Merged E8:F9");

        // 合併W3部分
        preserveAndMerge(tempSheet, startRowW3, colStart, numRows1, 5, "本週開展方向（或詢問）格子"); // O2:S2
        preserveAndMerge(tempSheet, startRowW3 + 1, colStart, numRows2, 5, "本週開展方向（或詢問）內容"); // O3:S4
        preserveAndMerge(tempSheet, startRowW3, colStart + 5, numRows1, 2, "輔訓建議（或回答）格子"); // T2:U2
        preserveAndMerge(tempSheet, startRowW3 + 1, colStart + 5, numRows2, 2, "輔訓建議（或回答）內容"); // T3:U4
        preserveAndMerge(tempSheet, startRowW3 + 3, colStart, numRows1, 2, "主日申言 格子"); // O5:P5
        preserveAndMerge(tempSheet, startRowW3 + 4, colStart, numRows1, 2, "主日申言 內容"); // O6:P6
        preserveAndMerge(tempSheet, startRowW3 + 3, colStart + 2, numRows1, 2, "本月受浸目標 格子"); // Q5:R5
        preserveAndMerge(tempSheet, startRowW3 + 4, colStart + 2, numRows1, 2, "本月受浸目標 內容"); // Q6:R6
        preserveAndMerge(tempSheet, startRowW3 + 3, colStart + 4, numRows1, 2, "本月受浸目標 格子"); // S5:T5
        preserveAndMerge(tempSheet, startRowW3 + 4, colStart + 4, numRows1, 2, "本月受浸目標 內容"); // S6:T6
        preserveAndMerge(tempSheet, startRowW3 + 5, colStart, numRows1, 2, "項目"); // O7:P7
        preserveAndMerge(tempSheet, startRowW3 + 6, colStart, numRows1, 2, "目標"); // O8:P8
        preserveAndMerge(tempSheet, startRowW3 + 7, colStart, numRows1, 2, "結果"); // O9:P9



        for (var i = 10; i <= number; i++) {
            preserveAndMerge(tempSheet, i, colStart, numRows1, 3, "O" + i + " 到 Q" + i + " 的合併範圍"); // O10:Q10 到 O70:Q70
            preserveAndMerge(tempSheet, i, colStart + 3, numRows1, 3, "R" + i + " 到 T" + i + " 的合併範圍"); // R10:T10 到 R70:T70

        }

        if (week !== 1) {
            tempSheet.deleteColumns(8, week * 7 - 7);
        }

    }
}

// 副程式：保留資料並合併儲存格
function preserveAndMerge(sheet, row, col, numRows, numCols, description) {
    var range = sheet.getRange(row, col, numRows, numCols);
    Logger.log("Merging: " + description + " in range " + row + "," + col);

    // 取得範圍內第一個儲存格的資料作為保留資料
    var firstCellValue = range.getCell(1, 1).getValue();

    // 合併儲存格
    range.merge();

    // 將原範圍內的第一個儲存格資料放回合併後的儲存格中
    range.getCell(1, 1).setValue(firstCellValue);
}



function _exportBlob(blob, fileName, spreadsheet) {
    var user = Session.getActiveUser();
    var userEmail = user.getEmail();
    var sheetName = spreadsheet.getName();

    // 發送電子郵件
    MailApp.sendEmail({
        to: userEmail,
        subject: fileName + "開展表單 - " + sheetName,
        body: "附件为生成的 PDF 文件",
        attachments: [{
            fileName: fileName + ".pdf",
            content: blob.getBytes(),
            mimeType: "application/pdf"
        }]
    });

    // 更新執行次數
    var userProperties = PropertiesService.getUserProperties();
    var executionCount = parseInt(userProperties.getProperty('executionCount') || '0', 10);
    executionCount += 1;
    userProperties.setProperty('executionCount', executionCount);

    // 回傳email和執行次數
    return {
        email: userEmail,
        count: executionCount
    };
}

function _getAsBlob(url, sheet, range) {
    var rangeParam = '';
    var sheetParam = '';
    if (range) {
        rangeParam =
            '&r1=' + (range.getRow() - 1) +
            '&r2=' + range.getLastRow() +
            '&c1=' + (range.getColumn() - 1) +
            '&c2=' + range.getLastColumn();
    }

    if (sheet) {
        sheetParam = '&gid=' + sheet.getSheetId();
    }

    var exportUrl = url.replace(/\/edit.*$/, '') +
        '/export?exportFormat=pdf&format=pdf' +
        '&size=7' +
        '&portrait=false' +
        '&fitp=false' +
        '&fitw=true' +
        '&top_margin=0.19685' +
        '&bottom_margin=0.19685' +
        '&left_margin=0.7' +
        '&right_margin=0.7' +
        '&sheetnames=false&printtitle=false' +
        '&pagenum=UNDEFINED' +
        '&gridlines=true' +
        '&fzr=FALSE' +
        '&scale=4' +
        '&printnotes=false' +
        sheetParam +
        rangeParam;

    Logger.log('exportUrl=' + exportUrl);
    var response;
    var i = 0;
    for (; i < 5; i += 1) {
        response = UrlFetchApp.fetch(exportUrl, {
            muteHttpExceptions: true,
            headers: {
                Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
            },
        });
        if (response.getResponseCode() === 429) {
            Utilities.sleep(3000);
        } else {
            break;
        }
    }

    if (i === 5) {
        throw new Error('列印失敗。工作表太多無法列印。');
    }

    return response.getBlob();
}

function exportSheetAsPDF(sheetName) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(sheetName); // 根据传入的工作表名称获取工作表

    if (sheet) {
        var blob = _getAsBlob(spreadsheet.getUrl(), sheet); // 将指定工作表导出为Blob对象
        var result = _exportBlob(blob, sheet.getName(), spreadsheet); // 将Blob对象作为PDF文件发送

        return result;
    } else {
        SpreadsheetApp.getUi().alert('未找到名為 ' + sheetName + ' 的工作表');
    }
}
function doGet(e) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetName = spreadsheet.getName();

  if (e.parameter.page) {
    var pageName = e.parameter.page.trim().toLowerCase();
    if (pageName !== "home") {
      var template = HtmlService.createTemplateFromFile(pageName);
      template.url = getPageUrl();
      return template.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setTitle('牧養表單');
    } else {
      return homePage(spreadsheetName).setTitle('牧養表單');
    }
  } else {
    return homePage(spreadsheetName).setTitle('開展表單');
  }
}

function homePage(spreadsheetName) {
  var pages = ["開展牧養表單"];
  var urls = pages.map(function(name) {
    return getPageUrl(name);
  });

  var template = HtmlService.createTemplateFromFile("home");
  template.urls = urls;
  template.spreadsheetName = spreadsheetName;

  return template.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setTitle('開展表單');
}

function getPageUrl(name) {
  var url = ScriptApp.getService().getUrl();
  return name ? url + "?page=" + name : url;
}
