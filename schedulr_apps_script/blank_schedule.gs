
/**
 * Opens a selector for Pulled_ semester sheets and creates a blank schedule
 * from the selected sheet without overwriting an existing schedule.
 */
function create_blank_schedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pulledSheets = ss
    .getSheets()
    .map((sheet) => sheet.getName())
    .filter((name) => name.indexOf("Pull_") === 0)
    .sort();

  if (!pulledSheets.length) {
    SpreadsheetApp.getUi().alert(
      "No Pull_ semester sheets were found. Run pull_semester first.",
    );
    return;
  }

  const options = pulledSheets
    .map((name) => "<option value='" + cbEscapeHtml_(name) + "'>" +
      cbEscapeHtml_(name) + "</option>")
    .join("");

  const html = HtmlService.createHtmlOutput(
    "<html><head><base target='_top'><style>" +
      "body{font-family:Arial,sans-serif;padding:16px;color:#202124}" +
      "label{display:block;font-weight:bold;margin-bottom:6px}" +
      "select{width:100%;padding:8px;font-size:13px}" +
      "button{margin-top:18px;padding:8px 18px;background:#1a73e8;color:white;border:0;border-radius:4px;cursor:pointer}" +
      "#status{margin-top:12px;color:#5f6368;font-size:12px}" +
      "</style></head><body>" +
      "<h2>Create Blank Schedule</h2>" +
      "<label for='source'>Pull semester schedule</label>" +
      "<select id='source'>" + options + "</select>" +
      "<button id='createButton' onclick='createSchedule()'>Create New Schedule</button>" +
      "<div id='status'></div>" +
      "<script>" +
      "function createSchedule(){" +
      "var button=document.getElementById('createButton');button.disabled=true;" +
      "document.getElementById('status').textContent='Creating schedule...';" +
      "google.script.run.withSuccessHandler(function(result){" +
      "document.getElementById('status').textContent=result.message;button.disabled=false;" +
      "}).withFailureHandler(function(error){" +
      "document.getElementById('status').textContent=error.message||error;button.disabled=false;" +
      "}).cbCreateBlankSchedule(document.getElementById('source').value);" +
      "}" +
      "</script></body></html>",
  ).setWidth(420).setHeight(300);

  SpreadsheetApp.getUi().showModalDialog(html, "Create Blank Schedule");
}

function cbCreateBlankSchedule(sourceSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName(sourceSheetName);
  if (!source || source.getName().indexOf("Pull_") !== 0) {
    throw new Error("The selected Pull_ sheet no longer exists.");
  }

  const roomsSheet = ss.getSheetByName("Rooms");
  if (!roomsSheet) throw new Error("Missing Rooms sheet.");
  const roomCapacities = cbReadRoomCapacities_(roomsSheet);

  const sourceValues = source.getDataRange().getValues();
  if (!sourceValues.length) throw new Error("The selected sheet is empty.");
  const sourceHeaders = sourceValues[0].map((value) => String(value || "").trim());
  const columns = cbFindBlankScheduleColumns_(sourceHeaders);
  if (
    columns.course === -1 ||
    columns.section === -1 ||
    columns.location === -1 ||
    columns.openSeats === -1
  ) {
    throw new Error(
      "The selected sheet must contain Course, Section, Location, and Open Seats columns.",
    );
  }

  const sheetName = psUniqueSheetName_(ss, "New Schedule");
  const output = source.copyTo(ss);
  output.setName(sheetName);

  let headers = output
    .getRange(1, 1, 1, output.getLastColumn())
    .getValues()[0]
    .map((value) => String(value || "").trim());
  let outputColumns = cbFindBlankScheduleColumns_(headers);

  if (outputColumns.waitlisted === -1) {
    const waitlistedColumn = output.getLastColumn() + 1;
    output.getRange(1, waitlistedColumn).setValue("Waitlisted");
    headers.push("Waitlisted");
    outputColumns.waitlisted = headers.length - 1;
  }

  const lastRow = output.getLastRow();
  const rowCount = Math.max(0, lastRow - 1);
  const missingRooms = {};
  const openSeatValues = [];
  const waitlistedValues = [];

  if (rowCount) {
    const rows = output.getRange(2, 1, rowCount, headers.length).getValues();
    rows.forEach((row) => {
      const room = cbNormalizeRoom_(row[outputColumns.location]);
      const capacity = roomCapacities[room];
      if (capacity === undefined) {
        openSeatValues.push([""]);
        if (room) missingRooms[room] = true;
      } else {
        openSeatValues.push([capacity]);
      }
      waitlistedValues.push([""]);
    });

    if (outputColumns.days !== -1) {
      output.getRange(2, outputColumns.days + 1, rowCount, 1).clearContent();
    }
    if (outputColumns.times !== -1) {
      output.getRange(2, outputColumns.times + 1, rowCount, 1).clearContent();
    }
    if (outputColumns.instructor !== -1) {
      output
        .getRange(2, outputColumns.instructor + 1, rowCount, 1)
        .clearContent();
    }
    if (outputColumns.openSeats !== -1) {
      output
        .getRange(2, outputColumns.openSeats + 1, rowCount, 1)
        .setValues(openSeatValues);
    }
    output
      .getRange(2, outputColumns.waitlisted + 1, rowCount, 1)
      .setValues(waitlistedValues);
  }

  output.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  output.setFrozenRows(1);
  output.autoResizeColumns(1, headers.length);

  const missingRoomList = Object.keys(missingRooms);
  const message =
    "Created " +
    sheetName +
    " with " +
    rowCount +
    " section(s)." +
    (missingRoomList.length
      ? " Rooms not found in Rooms sheet: " + missingRoomList.join(", ") + "."
      : "");
  SpreadsheetApp.getUi().alert(message);
  return { sheetName: sheetName, message: message };
}

function cbFindBlankScheduleColumns_(headers) {
  const normalized = headers.map((header) =>
    header.toLowerCase().replace(/\s+/g, "_"),
  );
  return {
    course: cbFindColumn_(normalized, ["course", "course_code", "coursecode"]),
    section: cbFindColumn_(normalized, ["section", "section_number"]),
    days: cbFindColumn_(normalized, ["days", "day"]),
    times: cbFindColumn_(normalized, ["time", "times", "class_time", "class_times"]),
    instructor: cbFindColumn_(normalized, ["instructor", "professor", "faculty"]),
    location: cbFindColumn_(normalized, ["location", "room", "classroom"]),
    openSeats: cbFindColumn_(normalized, ["open_seats", "seats_open"]),
    waitlisted: cbFindColumn_(normalized, ["waitlisted", "waitlist"]),
  };
}

function cbFindColumn_(headers, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const index = headers.indexOf(aliases[i]);
    if (index !== -1) return index;
  }
  return -1;
}

function cbReadRoomCapacities_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error("The Rooms sheet is empty.");
  const headers = values[0].map((value) =>
    String(value || "").trim().toLowerCase().replace(/\s+/g, "_"),
  );
  const roomColumn = cbFindColumn_(headers, ["room", "room_name"]);
  const seatsColumn = cbFindColumn_(headers, [
    "student_seats",
    "student_seat",
    "seats",
    "capacity",
  ]);
  if (roomColumn === -1 || seatsColumn === -1) {
    throw new Error("Rooms must contain Room and Student Seats columns.");
  }

  const capacities = {};
  for (let i = 1; i < values.length; i++) {
    const room = cbNormalizeRoom_(values[i][roomColumn]);
    const seats = Number(values[i][seatsColumn]);
    if (room && !isNaN(seats)) capacities[room] = seats;
  }
  return capacities;
}

function cbNormalizeRoom_(value) {
  return String(value || "")
    .toUpperCase()
    .replace(/[\s-]+/g, "")
    .replace(/[^A-Z0-9]/g, "")
    .trim();
}

function cbEscapeHtml_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
