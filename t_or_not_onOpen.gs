function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Schedulr")
    .addItem("Generate Course Overview", "generateCourseOverview")
    .addToUi();
}
