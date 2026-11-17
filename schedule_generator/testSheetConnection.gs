function testSheetConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  sheet.getRange("A1").setValue("Apps Script connected");
}
