function update_courses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getActiveSheet();
  const url = "https://www.sjsu.edu/classes/schedules/fall-2026.php";

  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { "User-Agent": "Mozilla/5.0" },
  });
  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error("SJSU schedule request failed with HTTP " + status + ".");
  }

  const sjsuSections = ucParseSjsuSchedule_(response.getContentText());
  if (!sjsuSections.length) {
    throw new Error("No SJSU sections were found. No sheet changes were made.");
  }

  const values = scheduleSheet.getDataRange().getValues();
  if (!values.length) throw new Error("The Schedule sheet is empty.");

  const headers = values[0].map((value) => String(value || "").trim());
  const firstThreeHeaders = headers
    .slice(0, 3)
    .map((header) => header.toLowerCase());
  if (
    firstThreeHeaders.length < 3 ||
    firstThreeHeaders[0] !== "course" ||
    firstThreeHeaders[1] !== "section" ||
    firstThreeHeaders[2] !== "days"
  ) {
    throw new Error(
      "Run update_courses from a sheet whose first row begins: Course, Section, Days.",
    );
  }

  const columns = ucFindScheduleColumns_(headers);
  if (columns.course === -1 || columns.section === -1) {
    throw new Error("Schedule must contain Course and Section columns.");
  }

  const courseCodes = {};
  const existingByKey = {};
  const rowsByCourse = {};
  for (let i = 1; i < values.length; i++) {
    const course = ucNormalizeCourse_(values[i][columns.course]);
    const section = ucNormalizeSection_(values[i][columns.section]);
    if (!course || !section) continue;
    courseCodes[course] = true;

    const rowNumber = i + 1;
    const key = ucSectionKey_(course, section);
    if (!existingByKey[key]) existingByKey[key] = [];
    existingByKey[key].push(rowNumber);
    if (!rowsByCourse[course]) rowsByCourse[course] = [];
    rowsByCourse[course].push(rowNumber);
  }
  if (!Object.keys(courseCodes).length) {
    throw new Error("No course rows were found below the header row.");
  }

  const sectionsByCourse = {};
  sjsuSections.forEach((section) => {
    if (!courseCodes[section.course]) return;
    if (!sectionsByCourse[section.course]) sectionsByCourse[section.course] = [];
    sectionsByCourse[section.course].push(section);
  });

  let updatedRows = 0;
  let insertedRows = 0;
  const missingCourses = [];
  const duplicateRows = [];
  const coursesWithRows = Object.keys(sectionsByCourse)
    .filter((course) => rowsByCourse[course] && rowsByCourse[course].length)
    .sort((a, b) => {
      const aAnchor = Math.max.apply(null, rowsByCourse[a]);
      const bAnchor = Math.max.apply(null, rowsByCourse[b]);
      return bAnchor - aAnchor;
    });

  // Update existing rows and insert missing sections. Existing groups are
  // processed from the bottom up so inserted rows cannot change later anchors.
  coursesWithRows.forEach((course) => {
    const courseRows = rowsByCourse[course];
    const missing = [];

    sectionsByCourse[course].forEach((section) => {
      const key = ucSectionKey_(section.course, section.section);
      const matchingRows = existingByKey[key] || [];
      if (!matchingRows.length) {
        missing.push(section);
        return;
      }

      matchingRows.forEach((rowNumber) => {
        ucUpdateScheduleRow_(scheduleSheet, rowNumber, columns, section);
        updatedRows++;
      });
      if (matchingRows.length > 1) {
        duplicateRows.push(key + " appears in rows " + matchingRows.join(", "));
      }
    });

    if (missing.length) {
      const anchor = Math.max.apply(null, courseRows);
      ucInsertScheduleRows_(scheduleSheet, anchor, headers.length, columns, missing);
      insertedRows += missing.length;
    }
  });

  // Courses with no existing Schedule row are appended after all existing data.
  const coursesWithoutRows = Object.keys(sectionsByCourse)
    .filter((course) => !rowsByCourse[course] || !rowsByCourse[course].length)
    .sort();
  coursesWithoutRows.forEach((course) => {
    const missing = sectionsByCourse[course];
    if (!missing.length) return;
    const anchor = scheduleSheet.getLastRow();
    ucInsertScheduleRows_(scheduleSheet, anchor, headers.length, columns, missing);
    insertedRows += missing.length;
  });

  Object.keys(courseCodes).forEach((course) => {
    if (!sectionsByCourse[course]) missingCourses.push(course);
  });

  const message =
    "Schedule update complete. Updated " +
    updatedRows +
    " existing row(s) and inserted " +
    insertedRows +
    " missing section row(s)." +
    (missingCourses.length
      ? " No SJSU sections found for: " + missingCourses.join(", ") + "."
      : "") +
    (duplicateRows.length
      ? " Duplicate Schedule rows were updated: " + duplicateRows.join("; ") + "."
      : "");
  SpreadsheetApp.getUi().alert(message);
}

function ucReadCourseCodes_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) return {};
  const headers = values[0].map((value) => String(value || "").trim().toLowerCase());
  const courseColumn = headers.findIndex(
    (header) => header === "course" || header === "coursecode" || header === "course code",
  );
  if (courseColumn === -1) throw new Error("Courses must contain a course column.");

  const codes = {};
  for (let i = 1; i < values.length; i++) {
    const code = ucNormalizeCourse_(values[i][courseColumn]);
    if (code) codes[code] = true;
  }
  return codes;
}

function ucFindScheduleColumns_(headers) {
  const lower = headers.map((header) => header.toLowerCase().replace(/\s+/g, "_"));
  return {
    course: ucFindColumn_(lower, ["course", "course_code", "coursecode"]),
    section: ucFindColumn_(lower, ["section", "section_number", "sectionnumber"]),
    days: ucFindColumn_(lower, ["days", "day"]),
    times: ucFindColumn_(lower, ["class_times", "class_time", "times", "time"]),
    instructor: ucFindColumn_(lower, ["instructor", "professor", "faculty"]),
    location: ucFindColumn_(lower, ["location", "room", "classroom"]),
    openSeats: ucFindColumn_(lower, ["open_seats", "open_seat", "seats_open"]),
    classNumber: ucFindColumn_(lower, ["class_number", "classnumber"]),
    mode: ucFindColumn_(lower, ["mode_of_instruction", "mode"]),
    title: ucFindColumn_(lower, ["course_title", "title"]),
    dates: ucFindColumn_(lower, ["dates", "date"]),
  };
}

function ucFindColumn_(headers, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const index = headers.indexOf(aliases[i]);
    if (index !== -1) return index;
  }
  return -1;
}

function ucUpdateScheduleRow_(sheet, rowNumber, columns, section) {
  ucSetIfPresent_(sheet, rowNumber, columns.course, section.course);
  ucSetIfPresent_(sheet, rowNumber, columns.section, section.section);
  ucSetIfPresent_(sheet, rowNumber, columns.days, section.days);
  ucSetIfPresent_(sheet, rowNumber, columns.times, section.times);
  ucSetIfPresent_(sheet, rowNumber, columns.instructor, section.instructor);
  ucSetIfPresent_(sheet, rowNumber, columns.location, section.location);
  ucSetIfPresent_(sheet, rowNumber, columns.openSeats, section.openSeats);
  ucSetIfPresent_(sheet, rowNumber, columns.classNumber, section.classNumber);
  ucSetIfPresent_(sheet, rowNumber, columns.mode, section.mode);
  ucSetIfPresent_(sheet, rowNumber, columns.title, section.title);
  ucSetIfPresent_(sheet, rowNumber, columns.dates, section.dates);
}

function ucInsertScheduleRows_(sheet, anchorRow, columnCount, columns, sections) {
  sheet.insertRowsAfter(anchorRow, sections.length);
  const newRows = Array.from({ length: sections.length }, () =>
    Array(columnCount).fill(""),
  );

  sections.forEach((section, index) => {
    const row = newRows[index];
    row[columns.course] = section.course;
    row[columns.section] = section.section;
    if (columns.days !== -1) row[columns.days] = section.days;
    if (columns.times !== -1) row[columns.times] = section.times;
    if (columns.instructor !== -1) row[columns.instructor] = section.instructor;
    if (columns.location !== -1) row[columns.location] = section.location;
    if (columns.openSeats !== -1) row[columns.openSeats] = section.openSeats;
    if (columns.classNumber !== -1) row[columns.classNumber] = section.classNumber;
    if (columns.mode !== -1) row[columns.mode] = section.mode;
    if (columns.title !== -1) row[columns.title] = section.title;
    if (columns.dates !== -1) row[columns.dates] = section.dates;
  });
  sheet.getRange(anchorRow + 1, 1, sections.length, columnCount).setValues(newRows);
}

function ucSetIfPresent_(sheet, rowNumber, columnIndex, value) {
  if (columnIndex !== -1) sheet.getRange(rowNumber, columnIndex + 1).setValue(value);
}

function ucParseSjsuSchedule_(html) {
  const rows = html.match(/<tr\b[\s\S]*?<\/tr>/gi) || [];
  let header = null;
  const sections = [];

  rows.forEach((rowHtml) => {
    const cells = ucExtractCells_(rowHtml).map(ucStripHtml_);
    const lower = cells.map((cell) => cell.toLowerCase());
    if (lower.includes("section") && lower.includes("open seats")) {
      header = {
        section: lower.indexOf("section"),
        days: lower.indexOf("days"),
        times: lower.indexOf("times"),
        instructor: lower.indexOf("instructor"),
        location: lower.indexOf("location"),
        openSeats: lower.indexOf("open seats"),
        classNumber: lower.indexOf("class number"),
        mode: lower.indexOf("mode of instruction"),
        title: lower.indexOf("course title"),
        dates: lower.indexOf("dates"),
      };
      return;
    }
    if (!header) return;

    const sectionText = cells[header.section] || "";
    const match = sectionText.match(/^(.+?)\s+\(Section\s+([^)]+)\)/i);
    if (!match) return;

    const openSeatsText = (cells[header.openSeats] || "").replace(/[^0-9-]/g, "");
    if (!openSeatsText) return;

    sections.push({
      course: ucNormalizeCourse_(match[1]),
      section: ucNormalizeSection_(match[2]),
      days: ucCell_(cells, header.days),
      times: ucCell_(cells, header.times),
      instructor: ucCell_(cells, header.instructor),
      location: ucCell_(cells, header.location),
      openSeats: Number(openSeatsText),
      classNumber: ucCell_(cells, header.classNumber),
      mode: ucCell_(cells, header.mode),
      title: ucCell_(cells, header.title),
      dates: ucCell_(cells, header.dates),
    });
  });
  return sections.filter((section) => section.course && section.section);
}

function ucCell_(cells, index) {
  return index === -1 || index >= cells.length ? "" : cells[index];
}

function ucExtractCells_(rowHtml) {
  const matches = rowHtml.match(/<t[dh]\b[\s\S]*?<\/t[dh]>/gi) || [];
  return matches.map((cell) =>
    cell.replace(/^<t[dh]\b[^>]*>/i, "").replace(/<\/t[dh]>$/i, ""),
  );
}

function ucStripHtml_(value) {
  return String(value || "")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&#(\d+);/g, (_, code) => String.fromCharCode(Number(code)))
    .replace(/&#x([0-9a-f]+);/gi, (_, code) =>
      String.fromCharCode(parseInt(code, 16)),
    )
    .replace(/\s+/g, " ")
    .trim();
}

function ucNormalizeCourse_(value) {
  const text = String(value || "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
  const match = text.match(/^([A-Z]{2,5})\s*(\d+[A-Z]?)/);
  return match ? match[1] + " " + match[2] : text;
}

function ucNormalizeSection_(value) {
  const text = String(value || "").trim().toUpperCase();
  const numeric = text.match(/^\d+$/);
  return numeric ? String(Number(text)) : text.replace(/\s+/g, "");
}

function ucSectionKey_(course, section) {
  return ucNormalizeCourse_(course) + "|" + ucNormalizeSection_(section);
}

function ucMustGetSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error("Missing sheet: " + name);
  return sheet;
}
