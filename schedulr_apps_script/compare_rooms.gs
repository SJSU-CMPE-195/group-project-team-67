function comparing_rooms_size() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getActiveSheet();
  const roomsSheet = ucMustGetSheet_(ss, "Rooms");
  const values = scheduleSheet.getDataRange().getValues();

  if (!values.length) throw new Error("The active sheet is empty.");

  const headers = values[0].map((value) => String(value || "").trim());
  const firstThreeHeaders = headers.slice(0, 3).map((header) => header.toLowerCase());
  if (
    firstThreeHeaders.length < 3 ||
    firstThreeHeaders[0] !== "course" ||
    firstThreeHeaders[1] !== "section" ||
    firstThreeHeaders[2] !== "days"
  ) {
    throw new Error(
      "Run comparing_rooms_size from a sheet whose first row begins: Course, Section, Days.",
    );
  }

  const columns = ucFindScheduleColumns_(headers);
  if (
    columns.course === -1 ||
    columns.section === -1 ||
    columns.days === -1 ||
    columns.times === -1 ||
    columns.location === -1 ||
    columns.openSeats === -1
  ) {
    throw new Error(
      "The active sheet must contain Course, Section, Days, Times, Location, and Open Seats columns.",
    );
  }

  const roomCapacities = crReadRoomCapacities_(roomsSheet);
  const sections = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const course = ucNormalizeCourse_(row[columns.course]);
    const section = ucNormalizeSection_(row[columns.section]);
    const room = crNormalizeRoom_(row[columns.location]);
    const openSeats = crNumber_(row[columns.openSeats]);
    const capacity = roomCapacities[room];

    if (!course || !section || !room || isNaN(openSeats)) continue;
    if (capacity === undefined) continue;

    sections.push({
      course: course,
      section: section,
      days: crNormalizeTimeText_(row[columns.days]),
      times: crNormalizeTimeText_(row[columns.times]),
      room: room,
      capacity: capacity,
      openSeats: openSeats,
    });
  }

  const swaps = [];
  sections.forEach((closedSection) => {
    if (closedSection.openSeats !== 0) return;

    sections.forEach((candidate) => {
      if (
        candidate === closedSection ||
        candidate.openSeats <= 2 ||
        candidate.days !== closedSection.days ||
        candidate.times !== closedSection.times ||
        candidate.capacity < closedSection.capacity ||
        candidate.capacity - candidate.openSeats > closedSection.capacity
      ) {
        return;
      }

      swaps.push([
        closedSection.course,
        closedSection.section,
        closedSection.days,
        closedSection.times,
        closedSection.room,
        closedSection.capacity,
        candidate.course,
        candidate.section,
        candidate.room,
        candidate.capacity,
        candidate.openSeats,
        candidate.capacity - closedSection.capacity,
      ]);
    });
  });

  swaps.sort((a, b) =>
    String(a[0] + a[1] + a[6] + a[7]).localeCompare(
      String(b[0] + b[1] + b[6] + b[7]),
      undefined,
      { numeric: true },
    ),
  );

  const output = [
    [
      "Closed Course",
      "Section",
      "Days",
      "Times",
      "Current Room",
      "Capacity",
      "Available Course",
      "Section",
      "Room",
      "Capacity",
      "Available Open Seats",
      "Capacity Gain",
    ],
    ...swaps,
  ];

  const resultSheet =
    ss.getSheetByName("Room_Swap_Options") ||
    ss.insertSheet("Room_Swap_Options");
  resultSheet.clearContents();
  resultSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  resultSheet.getRange(1, 1, 1, output[0].length).setFontWeight("bold");
  resultSheet.setFrozenRows(1);
  resultSheet.autoResizeColumns(1, output[0].length);

  crShowSwapDialog_(swaps);
}

function crReadRoomCapacities_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error("The Rooms sheet is empty.");

  const headers = values[0].map((value) =>
    String(value || "").trim().toLowerCase().replace(/\s+/g, "_"),
  );
  const roomColumn = ucFindColumn_(headers, ["room", "room_name", "building_room"]);
  const seatsColumn = ucFindColumn_(headers, [
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
    const room = crNormalizeRoom_(values[i][roomColumn]);
    const seats = crNumber_(values[i][seatsColumn]);
    if (room && !isNaN(seats)) capacities[room] = seats;
  }
  return capacities;
}

function crNormalizeRoom_(value) {
  return String(value || "")
    .toUpperCase()
    .replace(/[\s-]+/g, "")
    .replace(/[^A-Z0-9]/g, "")
    .trim();
}

function crNormalizeTimeText_(value) {
  return String(value || "")
    .toUpperCase()
    .replace(/\s+/g, "")
    .trim();
}

function crNumber_(value) {
  if (typeof value === "number") return value;
  const text = String(value || "").replace(/,/g, "").trim();
  return text === "" ? NaN : Number(text);
}

function crEscapeHtml_(value) {
  return String(value === undefined || value === null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function crShowSwapDialog_(swaps) {
  const previewLimit = 20;
  const preview = swaps.slice(0, previewLimit);
  const rows = preview
    .map(
      (row) =>
        "<tr><td>" +
        crEscapeHtml_(row[0] + " section " + row[1]) +
        "</td><td>" +
        crEscapeHtml_(row[4] + " (" + row[5] + " seats)") +
        "</td><td>&rarr;</td><td>" +
        crEscapeHtml_(row[6] + " section " + row[7]) +
        "</td><td>" +
        crEscapeHtml_(
          row[8] + " (" + row[9] + " seats, " + row[10] + " open)",
        ) +
        "</td></tr>",
    )
    .join("");

  const body = swaps.length
    ? "<p>Found " +
      swaps.length +
      " possible swap option(s). The full list is in the Room_Swap_Options sheet." +
      (swaps.length > previewLimit
        ? " Showing the first " + previewLimit + "."
        : "") +
      "</p><table border='1' cellpadding='5' cellspacing='0'>" +
      "<tr><th>Section</th><th>Current Room</th><th></th><th>Section</th><th>Room</th></tr>" +
      rows +
      "</table>"
    : "<p>No valid room swaps were found.</p><p>A valid swap requires the same day and time, more than 2 open seats, and a larger available room.</p>";

  const html = HtmlService.createHtmlOutput(
    "<html><head><style>" +
      "body{font-family:Arial,sans-serif;font-size:13px;padding:12px}" +
      "table{border-collapse:collapse;margin-top:10px}" +
      "th{background:#1f4e78;color:white}" +
      "th,td{padding:6px;text-align:left}" +
      "</style></head><body>" +
      body +
      "</body></html>",
  ).setWidth(850).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, "Available Room Swaps");
}
