function cleanPreferences() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("Form Responses 1");
  const cleanSheet = ss.getSheetByName("Cleaned_Preferences");

  const data = rawSheet.getDataRange().getValues();

  cleanSheet.clear();
  cleanSheet.appendRow([
    "Name",
    "Email",
    "Course",
    "Rank",
    "Preferred Time",
    "Max Courses"
  ]);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const name = row[1];
    const email = row[2];
    const pref1 = row[3];
    const pref2 = row[4];
    const pref3 = row[5];
    const time = row[6];
    const maxCourses = row[7];

    if (pref1) cleanSheet.appendRow([name, email, pref1, 1, time, maxCourses]);
    if (pref2 && pref2 !== "None") cleanSheet.appendRow([name, email, pref2, 2, time, maxCourses]);
    if (pref3 && pref3 !== "None") cleanSheet.appendRow([name, email, pref3, 3, time, maxCourses]);
  }

  const headerRange = cleanSheet.getRange(1, 1, 1, 6);
  headerRange.setFontWeight("bold");
  cleanSheet.setFrozenRows(1);
  cleanSheet.autoResizeColumns(1, 6);

  const lastRow = cleanSheet.getLastRow();
  if (lastRow > 1) {
    cleanSheet.getRange(2, 1, lastRow - 1, 6).sort([
      { column: 1, ascending: true },
      { column: 4, ascending: true }
    ]);
  }
}