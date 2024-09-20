function doGet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
  
  var result = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    result.push({
      name: row[1],  // Column B
      marks: row[3]  // Column D
    });
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
                       .setMimeType(ContentService.MimeType.JSON);
}
