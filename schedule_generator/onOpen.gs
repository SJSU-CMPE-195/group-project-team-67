function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Schedulr')
    .addItem('Test connection', 'testSheetConnection')
    .addItem('Generate schedule', 'generateSchedule')
    .addToUi();
}
