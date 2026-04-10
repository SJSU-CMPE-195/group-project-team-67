function generateCourseOverview() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("Courses");

  if (!inputSheet) {
    SpreadsheetApp.getUi().alert('Input sheet "Courses" not found.');
    return;
  }

  let outputSheet = ss.getSheetByName("Course_Overview");
  if (!outputSheet) {
    outputSheet = ss.insertSheet("Course_Overview");
  }

  outputSheet.clear();

  const data = inputSheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert("No course data found.");
    return;
  }

  const headers = data[0];
  const courseCol = headers.indexOf("CourseName");
  const categoryCol = headers.indexOf("Category");
  const professorCol = headers.indexOf("Prof Assigned");

  if (courseCol === -1 || categoryCol === -1 || professorCol === -1) {
    SpreadsheetApp.getUi().alert('Columns "CourseName", "Category", or "Prof Assigned" are missing.');
    return;
  }

  const lowerDiv = [];
  const upperDiv = [];
  const prereq195 = [];
  const techElective = [];
  const profNeeded = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const course = String(row[courseCol]).trim();
    const category = String(row[categoryCol]).trim();
    const professor = String(row[professorCol]).trim();

    if (!course) {
      continue;
    }

    if (category === "Lower Div") {
      lowerDiv.push([course]);
    }

    if (category === "Upper Div") {
      upperDiv.push([course]);
    }

    if (category === "CMPE195 Prereq") {
      prereq195.push([course]);
    }

    if (category === "Tech Elective") {
      techElective.push([course]);
    }

    if (professor === "") {
      profNeeded.push([course]);
    }
  }

  outputSheet.getRange(1, 1).setValue("Lower Div Courses");
  outputSheet.getRange(1, 2).setValue("Upper Div Courses");
  outputSheet.getRange(1, 3).setValue("195A Prereq");
  outputSheet.getRange(1, 4).setValue("Tech Elective");
  outputSheet.getRange(1, 5).setValue("Prof Needed");

  if (lowerDiv.length > 0) {
    outputSheet.getRange(2, 1, lowerDiv.length, 1).setValues(lowerDiv);
  }

  if (upperDiv.length > 0) {
    outputSheet.getRange(2, 2, upperDiv.length, 1).setValues(upperDiv);
  }

  if (prereq195.length > 0) {
    outputSheet.getRange(2, 3, prereq195.length, 1).setValues(prereq195);
  }

  if (techElective.length > 0) {
    outputSheet.getRange(2, 4, techElective.length, 1).setValues(techElective);
  }

  if (profNeeded.length > 0) {
    outputSheet.getRange(2, 5, profNeeded.length, 1).setValues(profNeeded);
  }

  outputSheet.autoResizeColumns(1, 5);
}
