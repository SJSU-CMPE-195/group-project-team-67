function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Schedulr")
    .addItem("Clean preferences", "cleanPreferences")
    .addItem("Build faculty sheet", "buildFacultySheet")
    .addItem("Generate schedule", "generateSchedule")
    .addItem("Update open seats", "updateCourseOpenSeats")
    // .addItem("Generate course overview", "generateCourseOverview") // dont need for now
    .addItem("Generate conflict summary", "generateConflictSummary")
    .addItem("Validate schedule", "validateSchedule")
    .addSeparator()
    .addItem("View Change Log", "showChangeLog")
    .addItem("Clear Change Log", "clearChangeLog")
    .addItem("Jump to Next Change", "jumpToNextChange")
    .addItem("Accept All Changes", "acceptAllChanges")
    .addItem("Revert to Original Changes", "revertAllChanges")
    .addToUi();
}

function updateCourseOpenSeats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const coursesSheet = mustGetSheet_(ss, "Courses");
  const scheduleUrl =
    "https://www.sjsu.edu/classes/schedules/fall-2026.php";

  const response = UrlFetchApp.fetch(scheduleUrl, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { "User-Agent": "Mozilla/5.0" },
  });
  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error("SJSU schedule request failed with HTTP " + status + ".");
  }

  const scheduleByCourse = parseSjsuOpenSeats_(response.getContentText());
  const range = coursesSheet.getDataRange();
  const values = range.getValues();
  if (values.length < 2) {
    SpreadsheetApp.getUi().alert("The Courses sheet has no course rows.");
    return;
  }

  const headers = values[0].map((header) => String(header).trim().toLowerCase());
  const courseCol = findHeaderIndex_(headers, ["course", "coursecode", "course code"]);
  if (courseCol === -1) throw new Error("Courses sheet is missing a course column.");

  let seatsCol = findHeaderIndex_(headers, ["seats_open", "seats open", "open seats"]);
  if (seatsCol === -1) {
    seatsCol = values[0].length;
    coursesSheet.getRange(1, seatsCol + 1).setValue("SEATS OPEN");
  }

  const output = [];
  const missing = [];
  let updatedCount = 0;
  for (let i = 1; i < values.length; i++) {
    const courseCode = normalizeCourseCode_(values[i][courseCol]);
    if (!courseCode) {
      output.push([""]);
      continue;
    }

    if (Object.prototype.hasOwnProperty.call(scheduleByCourse, courseCode)) {
      output.push([scheduleByCourse[courseCode]]);
      updatedCount++;
    } else {
      output.push([""]);
      missing.push(courseCode);
    }
  }

  coursesSheet.getRange(2, seatsCol + 1, output.length, 1).setValues(output);
  coursesSheet.getRange(1, seatsCol + 1).setFontWeight("bold");

  const uniqueMissing = uniqueList_(missing);
  SpreadsheetApp.getUi().alert(
    "Updated open seats for " +
      updatedCount +
      " course row(s)." +
      (uniqueMissing.length
        ? " Not found on the SJSU page: " + uniqueMissing.join(", ")
        : ""),
  );
}

function parseSjsuOpenSeats_(html) {
  const courseTotals = {};
  const rowMatches = html.match(/<tr\b[\s\S]*?<\/tr>/gi) || [];
  let sectionCol = -1;
  let openSeatsCol = -1;

  rowMatches.forEach((rowHtml) => {
    const cells = extractHtmlCells_(rowHtml);
    const cellText = cells.map(stripHtml_);
    const normalizedHeaders = cellText.map((text) => text.toLowerCase());

    if (
      normalizedHeaders.indexOf("section") !== -1 &&
      normalizedHeaders.indexOf("open seats") !== -1
    ) {
      sectionCol = normalizedHeaders.indexOf("section");
      openSeatsCol = normalizedHeaders.indexOf("open seats");
      return;
    }

    if (sectionCol === -1 || openSeatsCol === -1) return;
    const courseCode = normalizeCourseCode_(cellText[sectionCol]);
    const openSeatsText = cellText[openSeatsCol].replace(/[^0-9-]/g, "");
    if (!courseCode || !openSeatsText) return;

    const openSeats = Number(openSeatsText);
    if (isNaN(openSeats)) return;
    courseTotals[courseCode] = (courseTotals[courseCode] || 0) + openSeats;
  });

  return courseTotals;
}

function extractHtmlCells_(rowHtml) {
  const cells = [];
  const cellMatches = rowHtml.match(/<t[dh]\b[\s\S]*?<\/t[dh]>/gi) || [];
  cellMatches.forEach((cellHtml) => {
    cells.push(cellHtml.replace(/^<t[dh]\b[^>]*>/i, "").replace(/<\/t[dh]>$/i, ""));
  });
  return cells;
}

function stripHtml_(value) {
  return decodeHtmlEntities_(String(value || "").replace(/<[^>]+>/g, " "))
    .replace(/\s+/g, " ")
    .trim();
}

function decodeHtmlEntities_(value) {
  return String(value || "")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&#(\d+);/g, (_, code) => String.fromCharCode(Number(code)))
    .replace(/&#x([0-9a-f]+);/gi, (_, code) => String.fromCharCode(parseInt(code, 16)));
}

function findHeaderIndex_(headers, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const index = headers.indexOf(aliases[i].toLowerCase());
    if (index !== -1) return index;
  }
  return -1;
}

function cleanPreferences() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("Form Responses 1");
  if (!rawSheet) throw new Error("Missing sheet: Form Responses 1");

  const cleanSheet = getOrCreateSheet_(ss, "Cleaned_Preferences");
  const data = rawSheet.getDataRange().getValues();

  cleanSheet.clear();
  cleanSheet.appendRow([
    "facultyName",
    "email",
    "courseCode",
    "preferenceRank",
    "preferredTime",
    "maxCourses",
  ]);

  if (data.length < 2) {
    styleOutputSheet_(cleanSheet);
    return;
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const facultyName = String(row[1] || "").trim();
    const email = String(row[2] || "")
      .trim()
      .toLowerCase();

    const pref1 = normalizeCourseCode_(row[3]);
    const pref2 = normalizeCourseCode_(row[4]);
    const pref3 = normalizeCourseCode_(row[5]);

    const availabilityList = splitMultiSelect_(row[6]).map(normalizeTimeSlot_);
    const maxCourses = Number(row[8] || 0);

    // Use the first available slot as "preferredTime" for scoring.
    const preferredTime = availabilityList.length ? availabilityList[0] : "";

    if (pref1)
      cleanSheet.appendRow([
        facultyName,
        email,
        pref1,
        1,
        preferredTime,
        maxCourses,
      ]);
    if (pref2 && pref2 !== pref1)
      cleanSheet.appendRow([
        facultyName,
        email,
        pref2,
        2,
        preferredTime,
        maxCourses,
      ]);
    if (pref3 && pref3 !== pref1 && pref3 !== pref2)
      cleanSheet.appendRow([
        facultyName,
        email,
        pref3,
        3,
        preferredTime,
        maxCourses,
      ]);
  }

  styleOutputSheet_(cleanSheet);
}

function buildFacultySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = mustGetSheet_(ss, "Form Responses 1");
  const facultySheet = getOrCreateSheet_(ss, "Faculty");
  const data = rawSheet.getDataRange().getValues();

  facultySheet.clear();
  facultySheet.appendRow([
    "facultyName",
    "email",
    "unavailableSlots",
    "maxLoad",
  ]);

  if (data.length < 2) {
    styleOutputSheet_(facultySheet);
    return;
  }

  const seen = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const facultyName = String(row[1] || "").trim();
    const email = String(row[2] || "")
      .trim()
      .toLowerCase();
    const unavailableSlots = splitMultiSelect_(row[7]).map(normalizeTimeSlot_);
    const maxLoad = Number(row[8] || 0);

    const facultyId = email || facultyName;
    if (!facultyId || seen[facultyId]) continue;
    seen[facultyId] = true;

    facultySheet.appendRow([
      facultyName,
      email,
      unavailableSlots.join(", "),
      maxLoad,
    ]);
  }

  styleOutputSheet_(facultySheet);
}

function generateSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const coursesSheet = mustGetSheet_(ss, "Courses");
  const rawSheet = mustGetSheet_(ss, "Form Responses 1");
  const prefSheet =
    ss.getSheetByName("Cleaned_Preferences") ||
    ss.getSheetByName("Preferences");

  if (!prefSheet) {
    cleanPreferences();
  }

  // Auto-build missing/empty Faculty sheet
  const facultySheet = ensureFacultySheetReady_(ss, rawSheet);
  const actualPrefSheet =
    ss.getSheetByName("Cleaned_Preferences") ||
    ss.getSheetByName("Preferences");

  const courses = getNormalizedCourses_(coursesSheet);
  const faculty = getNormalizedFaculty_(facultySheet, actualPrefSheet);
  const preferences = getNormalizedPreferences_(actualPrefSheet, faculty);
  const timeSlots = getNormalizedTimeSlots_();

  if (!courses.length) throw new Error("Courses sheet has no usable rows.");
  if (!faculty.length) throw new Error("Faculty sheet has no usable rows.");
  if (!timeSlots.length) throw new Error("No usable time slots found.");

  const scheduleRows = [];
  const conflictRows = [];
  const facultyLoad = {};
  const facultyAssignments = {};

  faculty.forEach((f) => {
    facultyLoad[f.facultyId] = 0;
    facultyAssignments[f.facultyId] = [];
  });

  const expandedCourses = expandCourses_(courses).sort((a, b) =>
    compareExpandedCourses_(a, b, preferences),
  );

  expandedCourses.forEach((course) => {
    const allowedTerms = course.termsOffered.length
      ? course.termsOffered
      : ["Spring"];
    let bestChoice = null;
    const rejectionNotes = [];

    for (const term of allowedTerms) {
      for (const slot of timeSlots) {
        if (
          course.allowedTimeSlots.length &&
          course.allowedTimeSlots.indexOf(slot) === -1
        ) {
          continue;
        }

        const candidates = faculty
          .filter((f) => !isUnavailableForSlot_(f, slot))
          .filter((f) => !hasTimeConflict_(f, term, slot, facultyAssignments))
          .filter((f) => !isOverloaded_(f, facultyLoad))
          .sort((a, b) =>
            compareCandidatesForSlot_(
              a,
              b,
              course,
              term,
              slot,
              preferences,
              facultyLoad,
            ),
          );

        if (candidates.length) {
          const picked = candidates[0];
          const score = scoreCandidateChoice_(
            picked,
            course,
            term,
            slot,
            preferences,
            facultyLoad,
          );

          if (!bestChoice || score < bestChoice.score) {
            bestChoice = {
              score: score,
              faculty: picked,
              term: term,
              timeSlot: slot,
            };
          }
        } else {
          rejectionNotes.push("No faculty fits " + term + " / " + slot);
        }
      }
    }

    if (bestChoice) {
      const picked = bestChoice.faculty;
      facultyLoad[picked.facultyId] += 1;
      facultyAssignments[picked.facultyId].push({
        term: bestChoice.term,
        timeSlot: bestChoice.timeSlot,
        courseCode: course.courseCode,
        section: course.section,
      });

      scheduleRows.push([
        bestChoice.term,
        course.courseCode,
        course.courseName,
        course.section,
        picked.facultyId,
        picked.facultyName,
        bestChoice.timeSlot,
        "ASSIGNED",
        buildAssignmentNote_(
          picked.facultyId,
          course.courseCode,
          bestChoice.timeSlot,
          preferences,
        ),
      ]);
    } else {
      const candidateFaculty = faculty.map((f) => f.facultyName).join(", ");
      const eligibleCount = faculty.length;

      const conflictType = classifyConflictType_(rejectionNotes, eligibleCount);

      conflictRows.push([
        course.courseCode,
        course.courseName,
        allowedTerms.join(", "),
        course.section,
        course.allowedTimeSlots.join(", ") || "Any standard slot",
        conflictType,
        collapseReasons_(rejectionNotes) ||
          "No valid professor + time slot combination found",
        eligibleCount,
        candidateFaculty || "None",
      ]);
    }
  });

  writeSheet_(
    ss,
    "ScheduleOutput",
    [
      "term",
      "courseCode",
      "courseName",
      "section",
      "email",
      "facultyName",
      "timeSlot",
      "status",
      "notes",
    ],
    scheduleRows,
  );

  writeSheet_(
    ss,
    "Conflicts",
    [
      "courseCode",
      "courseName",
      "termsTried",
      "section",
      "timeSlotOptions",
      "conflictType",
      "reason",
      "totalFacultyCount",
      "candidateFaculty",
    ],
    conflictRows,
  );
}

function getNormalizedCourses_(sheet) {
  const rows = sheetToObjects_(sheet);
  return rows
    .map((r) => ({
      courseCode: normalizeCourseCode_(
        pickField_(r, ["courseCode", "CourseCode", "course", "Course"]),
      ),
      courseName:
        pickField_(r, [
          "courseName",
          "CourseName",
          "course_title",
          "Course_Title",
          "title",
          "Title",
        ]) || "",
      sectionsNeeded: Number(
        pickField_(r, [
          "sectionsNeeded",
          "SectionsNeeded",
          "section_amount",
          "sectionAmount",
          "sections",
        ]) || 1,
      ),
      termsOffered: splitList_(
        pickField_(r, [
          "termsOffered",
          "TermsOffered",
          "term",
          "Term",
          "terms",
        ]),
      ),
      allowedTimeSlots: splitList_(
        pickField_(r, [
          "valid_class_times",
          "allowedTimeSlots",
          "AllowedTimeSlots",
          "allowedSlots",
          "AllowedSlots",
        ]),
      ).map(normalizeTimeSlot_),
      category: String(pickField_(r, ["category", "Category"]) || "").trim(),
    }))
    .map((r) => {
      // These courses have multiple sections at the same time but are taught at the same time by one professor
      const singleSectionCourses = ["CMPE 195A", "CMPE 195B"];
      if (singleSectionCourses.includes(r.courseCode)) {
        r.sectionsNeeded = 1;
      }
      return r;
    })
    .filter((r) => r.courseCode && Number(r.sectionsNeeded || 0) > 0);
}

function getNormalizedFaculty_(sheet, prefSheet) {
  const rows = sheetToObjects_(sheet);
  const maxFromPrefs = buildMaxLoadLookupFromPreferences_(prefSheet);

  return rows
    .map((r) => {
      const email = String(pickField_(r, ["email", "Email"]) || "")
        .trim()
        .toLowerCase();
      const facultyName = String(
        pickField_(r, ["facultyName", "FacultyName", "name", "Name"]) || "",
      ).trim();
      const facultyId = email || facultyName;
      const fallbackMax =
        email && maxFromPrefs[email] ? maxFromPrefs[email] : "";

      return {
        facultyId: facultyId,
        facultyName: facultyName || facultyId,
        email: email,
        unavailableSlots: splitList_(
          pickField_(r, [
            "unavailableSlots",
            "UnavailableSlots",
            "blockedSlots",
            "BlockedSlots",
          ]),
        ).map(normalizeTimeSlot_),
        maxLoad: Number(
          pickField_(r, ["maxLoad", "MaxLoad", "maxCourses", "Max Courses"]) ||
            fallbackMax ||
            0,
        ),
      };
    })
    .filter((r) => r.facultyId);
}

function getNormalizedPreferences_(sheet, facultyList) {
  const rows = sheetToObjects_(sheet);
  const facultyByEmail = {};
  const facultyByName = {};

  facultyList.forEach((f) => {
    if (f.email) facultyByEmail[f.email.toLowerCase()] = f.facultyId;
    if (f.facultyName) facultyByName[f.facultyName.toLowerCase()] = f.facultyId;
  });

  return rows
    .map((r) => {
      const email = String(pickField_(r, ["email", "Email"]) || "")
        .trim()
        .toLowerCase();
      const name = String(
        pickField_(r, ["facultyName", "FacultyName", "name", "Name"]) || "",
      ).trim();
      const directId = String(
        pickField_(r, ["facultyId", "FacultyId", "id", "ID"]) || "",
      ).trim();
      const facultyId =
        directId ||
        facultyByEmail[email] ||
        facultyByName[name.toLowerCase()] ||
        email ||
        name;

      return {
        facultyId: facultyId,
        courseCode: normalizeCourseCode_(
          pickField_(r, ["courseCode", "CourseCode", "course", "Course"]),
        ),
        preferenceRank: Number(
          pickField_(r, ["preferenceRank", "PreferenceRank", "Rank", "rank"]) ||
            999,
        ),
        preferredTime: normalizeTimeSlot_(
          pickField_(r, ["preferredTime", "Preferred Time"]) || "",
        ),
        maxCourses: Number(pickField_(r, ["maxCourses", "Max Courses"]) || 0),
      };
    })
    .filter((r) => r.facultyId && r.courseCode);
}

function getNormalizedTimeSlots_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const coursesSheet = ss.getSheetByName("Courses");
  if (!coursesSheet) return [];

  const rows = sheetToObjects_(coursesSheet);
  const slots = [];

  rows.forEach((r) => {
    const times = String(
      pickField_(r, ["valid_class_times", "allowedTimeSlots"]) || "",
    ).trim();
    if (!times) return;
    times
      .split(",")
      .map((t) => t.trim())
      .filter(Boolean)
      .forEach((t) => {
        if (!slots.includes(t)) slots.push(t);
      });
  });

  return slots;
}

function compareExpandedCourses_(a, b, preferences) {
  const aPref = countPreferenceMatches_(a.courseCode, preferences);
  const bPref = countPreferenceMatches_(b.courseCode, preferences);
  if (aPref !== bPref) return aPref - bPref;

  return a.courseCode.localeCompare(b.courseCode);
}

function compareCandidatesForSlot_(
  a,
  b,
  course,
  term,
  slot,
  preferences,
  facultyLoad,
) {
  return (
    scoreCandidateChoice_(a, course, term, slot, preferences, facultyLoad) -
    scoreCandidateChoice_(b, course, term, slot, preferences, facultyLoad)
  );
}

function scoreCandidateChoice_(
  faculty,
  course,
  term,
  slot,
  preferences,
  facultyLoad,
) {
  const prefScore = getPreferenceScore_(
    faculty.facultyId,
    course.courseCode,
    preferences,
  );
  const timeScore = getPreferredTimeScore_(
    faculty.facultyId,
    slot,
    preferences,
  );
  const loadScore = Number(facultyLoad[faculty.facultyId] || 0);

  return prefScore * 100 + timeScore * 10 + loadScore;
}

function buildAssignmentNote_(facultyId, courseCode, timeSlot, preferences) {
  const courseRank = getPreferenceScore_(facultyId, courseCode, preferences);
  const timeRank = getPreferredTimeScore_(facultyId, timeSlot, preferences);

  const notes = [];
  if (courseRank < 999) notes.push("course preference rank " + courseRank);
  else notes.push("no stated preference for this course");

  if (timeRank === 0) notes.push("preferred time matched");
  else notes.push("standard slot assigned");

  return notes.join("; ");
}

function countPreferenceMatches_(courseCode, preferences) {
  return preferences.filter(
    (p) =>
      normalizeCourseCode_(p.courseCode) === normalizeCourseCode_(courseCode),
  ).length;
}

function getPreferenceScore_(facultyId, courseCode, preferences) {
  const row = preferences.find(
    (p) =>
      String(p.facultyId).trim() === String(facultyId).trim() &&
      normalizeCourseCode_(p.courseCode) === normalizeCourseCode_(courseCode),
  );
  return row ? Number(row.preferenceRank || 999) : 999;
}

function getPreferredTimeScore_(facultyId, timeSlot, preferences) {
  const rows = preferences.filter(
    (p) => String(p.facultyId).trim() === String(facultyId).trim(),
  );
  if (!rows.length) return 1;
  const match = rows.find(
    (r) =>
      String(r.preferredTime).trim() &&
      normalizeTimeSlot_(r.preferredTime) === normalizeTimeSlot_(timeSlot),
  );
  return match ? 0 : 1;
}

function isUnavailableForSlot_(faculty, timeSlot) {
  if (!timeSlot) return false;
  return faculty.unavailableSlots.includes(normalizeTimeSlot_(timeSlot));
}

function isOverloaded_(faculty, facultyLoad) {
  const max = Number(faculty.maxLoad || 0);
  if (max <= 0) return false;
  return (facultyLoad[faculty.facultyId] || 0) >= max;
}

function hasTimeConflict_(faculty, term, timeSlot, facultyAssignments) {
  if (!timeSlot) return false;
  const assignments = facultyAssignments[faculty.facultyId] || [];
  return assignments.some(
    (a) =>
      String(a.term).trim() === String(term).trim() &&
      normalizeTimeSlot_(a.timeSlot) === normalizeTimeSlot_(timeSlot),
  );
}

function collapseReasons_(reasons) {
  return uniqueList_(reasons.filter(Boolean)).slice(0, 8).join("; ");
}

function classifyConflictType_(reasons, totalFacultyCount) {
  if (totalFacultyCount === 0) return "NO_FACULTY_AVAILABLE";
  const text = reasons.join(" ").toLowerCase();
  if (text.includes("overload") || text.includes("max load"))
    return "MAX_LOAD_REACHED";
  if (text.includes("time conflict") || text.includes("already uses"))
    return "TIME_CONFLICT";
  if (text.includes("no faculty fits")) return "ALL_SLOTS_BLOCKED";
  return "NO_VALID_COMBINATION";
}

function expandCourses_(courses) {
  const expanded = [];
  courses.forEach((course) => {
    const count = Number(course.sectionsNeeded || 1);
    for (let i = 1; i <= count; i++) {
      expanded.push({
        courseCode: course.courseCode,
        courseName: course.courseName,
        termsOffered: course.termsOffered,
        allowedTimeSlots: course.allowedTimeSlots,
        category: course.category,
        section: i,
      });
    }
  });
  return expanded;
}

function ensureFacultySheetReady_(ss, rawSheet) {
  let sheet = ss.getSheetByName("Faculty");
  if (!sheet || sheet.getLastRow() < 2) {
    buildFacultySheet();
    sheet = ss.getSheetByName("Faculty");
  }
  return sheet;
}

function sheetToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) return [];
  const headers = values.shift().map((h) => String(h).trim());
  return values
    .filter((row) => row.some((cell) => String(cell).trim() !== ""))
    .map((row) => {
      const obj = {};
      headers.forEach((header, index) => (obj[header] = row[index]));
      return obj;
    });
}

function writeSheet_(ss, name, headers, rows) {
  const sheet = getOrCreateSheet_(ss, name);
  clearTrackedChangesForSheet_(ss, name);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  styleOutputSheet_(sheet);
}

function writeSingleColumnBlock_(sheet, startRow, col, rows) {
  if (rows.length) {
    sheet.getRange(startRow, col, rows.length, 1).setValues(rows);
  }
}

function buildMaxLoadLookupFromPreferences_(prefSheet) {
  const map = {};
  if (!prefSheet || prefSheet.getLastRow() < 2) return map;

  const rows = sheetToObjects_(prefSheet);
  rows.forEach((r) => {
    const email = String(pickField_(r, ["email", "Email"]) || "")
      .trim()
      .toLowerCase();
    const max = Number(pickField_(r, ["maxCourses", "Max Courses"]) || 0);
    if (email && max > 0) map[email] = max;
  });
  return map;
}

function pickField_(obj, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    if (Object.prototype.hasOwnProperty.call(obj, aliases[i]))
      return obj[aliases[i]];
  }
  return "";
}

function splitList_(value) {
  return String(value || "")
    .split(",")
    .map((x) => String(x).trim())
    .filter(Boolean);
}

function splitMultiSelect_(value) {
  return String(value || "")
    .split(",")
    .map((x) => String(x).trim())
    .filter(Boolean)
    .filter(
      (x) => x.toLowerCase() !== "none" && x.toLowerCase() !== "no preference",
    );
}

function uniqueList_(arr) {
  const seen = {};
  const out = [];
  arr.forEach((item) => {
    const key = String(item);
    if (!seen[key]) {
      seen[key] = true;
      out.push(item);
    }
  });
  return out;
}

function normalizeCourseCode_(value) {
  const text = String(value || "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();

  const match = text.match(/^([A-Z]{2,4}\s*\d+[A-Z]?)/);
  return match ? match[1].replace(/\s+/g, " ").trim() : text;
}

function normalizeTimeSlot_(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeBoolean_(value) {
  const v = String(value || "")
    .trim()
    .toUpperCase();
  return v === "TRUE" || v === "YES" || v === "Y" || v === "1";
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function mustGetSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error("Missing sheet: " + name);
  return sheet;
}

function styleOutputSheet_(sheet) {
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) return;
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, sheet.getLastColumn());
}

function generateConflictSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const conflictSheet = ss.getSheetByName("Conflicts");
  if (!conflictSheet || conflictSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert(
      "No conflicts found. Run Generate Schedule first.",
    );
    return;
  }

  const rows = sheetToObjects_(conflictSheet);
  const counts = {
    NO_FACULTY_AVAILABLE: 0,
    MAX_LOAD_REACHED: 0,
    TIME_CONFLICT: 0,
    ALL_SLOTS_BLOCKED: 0,
    NO_VALID_COMBINATION: 0,
  };

  rows.forEach((r) => {
    const type = String(r["conflictType"] || "").trim();
    if (counts.hasOwnProperty(type)) counts[type]++;
    else counts["NO_VALID_COMBINATION"]++;
  });

  const summarySheet = getOrCreateSheet_(ss, "Conflict_Summary");
  summarySheet.clear();
  summarySheet.getRange(1, 1, 1, 2).setValues([["conflictType", "count"]]);

  const summaryRows = Object.entries(counts).map(([type, count]) => [
    type,
    count,
  ]);
  summarySheet.getRange(2, 1, summaryRows.length, 2).setValues(summaryRows);

  // Color-code by severity
  const colorMap = {
    NO_FACULTY_AVAILABLE: "#F4CCCC", // red - no faculty at all to draw from
    MAX_LOAD_REACHED: "#FCE8B2", // yellow - adjust maxLoad
    TIME_CONFLICT: "#FFF2CC", // light yellow - slot overlap
    ALL_SLOTS_BLOCKED: "#FCE8B2", // yellow - availability too narrow
    NO_VALID_COMBINATION: "#D9EAD3", // green - minor, mixed causes
  };

  summaryRows.forEach((row, i) => {
    const color = colorMap[row[0]] || "#FFFFFF";
    summarySheet.getRange(i + 2, 1, 1, 2).setBackground(color);
  });

  styleOutputSheet_(summarySheet);
}

function validateSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const outputSheet = ss.getSheetByName("ScheduleOutput");
  if (!outputSheet || outputSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert(
      "No schedule found. Run Generate Schedule first.",
    );
    return;
  }

  const coursesSheet = ss.getSheetByName("Courses");

  const scheduleRows = sheetToObjects_(outputSheet);
  const courseRows = coursesSheet ? getNormalizedCourses_(coursesSheet) : [];

  const logRows = [];
  const facultySlotMap = {}; // facultyId -> [term|timeSlot, ...]
  const facultyLoadMap = {}; // facultyId -> count
  const courseSectionMap = {}; // courseCode -> count assigned

  // Build expected sections needed from Courses sheet
  const sectionsNeeded = {};
  courseRows.forEach((c) => {
    sectionsNeeded[c.courseCode] = Number(c.sectionsNeeded || 1);
  });

  scheduleRows.forEach((row, i) => {
    const rowNum = i + 2; // account for header row
    const facultyId = String(row["email"] || "").trim();
    const courseCode = normalizeCourseCode_(row["courseCode"] || "");
    const timeSlot = normalizeTimeSlot_(row["timeSlot"] || "");
    const term = String(row["term"] || "").trim();

    if (!facultyId || !courseCode || !timeSlot || !term) {
      logRows.push([
        rowNum,
        courseCode,
        facultyId,
        "FAIL",
        "MISSING_DATA",
        "Row is missing one or more required fields",
      ]);
      return;
    }

    // Check 1: duplicate faculty time slot (professor double-booked)
    const slotKey = facultyId + "|" + term + "|" + timeSlot;
    if (!facultySlotMap[facultyId]) facultySlotMap[facultyId] = [];
    if (facultySlotMap[facultyId].includes(slotKey)) {
      logRows.push([
        rowNum,
        courseCode,
        facultyId,
        "FAIL",
        "DOUBLE_BOOKED",
        facultyId + " is assigned to two courses in " + term + " / " + timeSlot,
      ]);
    } else {
      facultySlotMap[facultyId].push(slotKey);
    }

    // Track load and section counts for checks 2 and 3
    facultyLoadMap[facultyId] = (facultyLoadMap[facultyId] || 0) + 1;
    courseSectionMap[courseCode] = (courseSectionMap[courseCode] || 0) + 1;
  });

  // Check 2: max load exceeded
  // Re-read faculty sheet to get maxLoad per faculty
  const facultySheet = ss.getSheetByName("Faculty");
  if (facultySheet && facultySheet.getLastRow() > 1) {
    const facultyRows = sheetToObjects_(facultySheet);
    facultyRows.forEach((f) => {
      const facultyId = String(f["email"] || "").trim();
      const maxLoad = Number(f["maxLoad"] || 0);
      const assigned = facultyLoadMap[facultyId] || 0;
      if (maxLoad > 0 && assigned > maxLoad) {
        logRows.push([
          "-",
          "-",
          facultyId,
          "FAIL",
          "OVERLOAD",
          facultyId +
            " is assigned " +
            assigned +
            " courses but maxLoad is " +
            maxLoad,
        ]);
      }
    });
  }

  // Check 3: sections assigned vs sections needed
  courseRows.forEach((c) => {
    const needed = Number(c.sectionsNeeded || 1);
    const assigned = courseSectionMap[c.courseCode] || 0;
    if (assigned < needed) {
      logRows.push([
        "-",
        c.courseCode,
        "-",
        "FAIL",
        "MISSING_SECTIONS",
        c.courseCode +
          " needed " +
          needed +
          " section(s) but only " +
          assigned +
          " were assigned",
      ]);
    }
  });

  // Write to Validation_Log sheet
  const logSheet = getOrCreateSheet_(ss, "Validation_Log");
  logSheet.clear();
  logSheet
    .getRange(1, 1, 1, 6)
    .setValues([
      [
        "scheduleRow",
        "courseCode",
        "email",
        "result",
        "checkType",
        "description",
      ],
    ]);

  if (logRows.length === 0) {
    logSheet
      .getRange(2, 1, 1, 6)
      .setValues([
        ["-", "-", "-", "PASS", "ALL_CHECKS", "No violations found"],
      ]);
    logSheet.getRange(2, 1, 1, 6).setBackground("#B7E1CD"); // green
  } else {
    logSheet.getRange(2, 1, logRows.length, 6).setValues(logRows);
    logRows.forEach((row, i) => {
      const color = row[3] === "PASS" ? "#B7E1CD" : "#F4CCCC"; // green or red
      logSheet.getRange(i + 2, 4, 1, 1).setBackground(color);
    });
  }

  styleOutputSheet_(logSheet);

  const failCount = logRows.filter((r) => r[3] === "FAIL").length;
  SpreadsheetApp.getUi().alert(
    failCount === 0
      ? "Validation passed. No violations found."
      : "Validation found " +
          failCount +
          " issue(s). See Validation_Log sheet.",
  );
}

/* =========================================================================
 * CHANGE TRACKING
 * - Every single-cell edit gets a hover note with the previous value, who
 *   changed it, and when, plus a red highlight on the cell.
 * - Multi-cell edits (paste/fill/delete) get flagged too, but Sheets doesn't
 *   expose old values for those, so the note says so instead of guessing.
 * - Three ways to review changes: (1) hover over any highlighted cell,
 *   (2) "View Change Log" for a full running audit trail, (3) "Jump to Next
 *   Change" to cycle through every flagged cell one at a time.
 * - "Accept All Changes" clears highlights/notes and keeps the current
 *   values. "Revert to Original Changes" restores each cell's prior value
 *   (where captured) and then clears highlights/notes.
 * ========================================================================= */

const CHANGE_LOG_SHEET_NAME = "Change_Log";
const CHANGE_HIGHLIGHT_COLOR = "#F4CCCC";
const CHANGE_EXCLUDED_SHEETS = [
  "Conflicts",
  "Conflict_Summary",
  "Validation_Log",
  CHANGE_LOG_SHEET_NAME,
];

function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    if (CHANGE_EXCLUDED_SHEETS.indexOf(sheetName) !== -1) return;

    // Don't track edits to the header row.
    if (e.range.getRow() === 1 && e.range.getNumRows() === 1) return;

    const numCells = e.range.getNumRows() * e.range.getNumColumns();

    if (numCells === 1) {
      const oldValue = e.oldValue !== undefined ? e.oldValue : "";
      const newValue = e.value !== undefined ? e.value : e.range.getValue();
      if (String(oldValue) === String(newValue)) return;
      recordChange_(sheet, e.range, oldValue, newValue, false);
    } else {
      // Multi-cell edit - no per-cell old values available from the event.
      const values = e.range.getValues();
      for (let r = 0; r < values.length; r++) {
        for (let c = 0; c < values[r].length; c++) {
          const cell = sheet.getRange(
            e.range.getRow() + r,
            e.range.getColumn() + c,
          );
          recordChange_(sheet, cell, null, values[r][c], true);
        }
      }
    }
  } catch (err) {
    // Never let tracking errors block the user's actual edit.
    console.error(err);
  }
}

function recordChange_(sheet, cell, oldValue, newValue, isBulk) {
  const user = getEditingUser_();
  const now = new Date();
  const timestamp = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    "MM/dd/yyyy HH:mm",
  );

  const previousText = isBulk
    ? "(bulk edit - previous value not captured)"
    : formatValueForNote_(oldValue);

  const noteText =
    "Previous: " +
    previousText +
    "\nChanged by: " +
    user +
    "\nChanged at: " +
    timestamp;

  cell.setNote(noteText);
  cell.setBackground(CHANGE_HIGHLIGHT_COLOR);

  logChange_(
    sheet.getName(),
    cell.getA1Notation(),
    previousText,
    formatValueForNote_(newValue),
    user,
    now,
  );
  addPendingChange_(sheet.getName(), cell.getA1Notation(), oldValue, isBulk);
}

function getEditingUser_() {
  try {
    const email = Session.getActiveUser().getEmail();
    return email || "Unknown user";
  } catch (err) {
    return "Unknown user";
  }
}

function formatValueForNote_(value) {
  const text = String(value === null || value === undefined ? "" : value);
  return text === "" ? "(empty)" : text;
}

function logChange_(sheetName, a1, previousValue, newValue, user, timestamp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = getOrCreateChangeLogSheet_(ss);
  logSheet.appendRow([timestamp, user, sheetName, a1, previousValue, newValue]);
}

function getOrCreateChangeLogSheet_(ss) {
  let sheet = ss.getSheetByName(CHANGE_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CHANGE_LOG_SHEET_NAME);
    sheet.appendRow([
      "timestamp",
      "user",
      "sheet",
      "cell",
      "previousValue",
      "newValue",
    ]);
    styleOutputSheet_(sheet);
    sheet.hideSheet();
  }
  return sheet;
}

function clearTrackedChangesForSheet_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() > 0 && dataRange.getNumColumns() > 0) {
      dataRange.setBackground(null);
      dataRange.clearNote();
    }
  }

  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty("pendingChanges");
  if (!raw) return;
  const map = JSON.parse(raw);
  const filtered = {};
  Object.keys(map).forEach((id) => {
    if (id.split("!")[0] !== sheetName) filtered[id] = map[id];
  });
  props.setProperty("pendingChanges", JSON.stringify(filtered));
}

function addPendingChange_(sheetName, a1, rawOldValue, isBulk) {
  const props = PropertiesService.getDocumentProperties();
  const key = "pendingChanges";
  const raw = props.getProperty(key);
  const map = raw ? JSON.parse(raw) : {};
  const id = sheetName + "!" + a1;
  map[id] = {
    previousValue: isBulk ? null : rawOldValue,
    isBulk: !!isBulk,
  };
  props.setProperty(key, JSON.stringify(map));
}

function jumpToNextChange() {
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty("pendingChanges");
  const map = raw ? JSON.parse(raw) : {};
  const ids = Object.keys(map);

  if (!ids.length) {
    SpreadsheetApp.getUi().alert("No tracked changes found.");
    return;
  }

  let idx = Number(props.getProperty("pendingChangeIndex") || 0);
  if (idx >= ids.length) idx = 0;

  const parts = ids[idx].split("!");
  const sheetName = parts[0];
  const a1 = parts[1];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.setActiveSheet(sheet);
    sheet.getRange(a1).activate();
  }

  props.setProperty("pendingChangeIndex", String(idx + 1));
}

function acceptAllChanges() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "Accept all tracked changes?",
    "This confirms the current values are correct and removes the highlight colors and notes from tracked cells. The values themselves are kept as-is — nothing is reverted.",
    ui.ButtonSet.YES_NO,
  );
  if (resp !== ui.Button.YES) return;

  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty("pendingChanges");
  const map = raw ? JSON.parse(raw) : {};
  const ids = Object.keys(map);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ids.forEach((id) => {
    const parts = id.split("!");
    const sheetName = parts[0];
    const a1 = parts[1];
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const range = sheet.getRange(a1);
      range.setBackground(null);
      range.clearNote();
    }
  });

  props.deleteProperty("pendingChanges");
  props.deleteProperty("pendingChangeIndex");
  ui.alert("Accepted " + ids.length + " change(s).");
}

function revertAllChanges() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "Revert to original values?",
    "This restores the value each tracked cell had before it was edited, then clears the highlight and note. Cells changed by a bulk edit (paste/fill) don't have a captured previous value, so those will just be cleared, not reverted.",
    ui.ButtonSet.YES_NO,
  );
  if (resp !== ui.Button.YES) return;

  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty("pendingChanges");
  const map = raw ? JSON.parse(raw) : {};
  const ids = Object.keys(map);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let reverted = 0;
  let skipped = 0;

  ids.forEach((id) => {
    const parts = id.split("!");
    const sheetName = parts[0];
    const a1 = parts[1];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const range = sheet.getRange(a1);
    const entry = map[id];

    if (entry && !entry.isBulk) {
      range.setValue(entry.previousValue);
      reverted++;
    } else {
      skipped++;
    }

    range.setBackground(null);
    range.clearNote();
  });

  props.deleteProperty("pendingChanges");
  props.deleteProperty("pendingChangeIndex");

  ui.alert(
    "Reverted " +
      reverted +
      " cell(s) to their original value." +
      (skipped > 0
        ? " " +
          skipped +
          " cell(s) from bulk edits couldn't be reverted and were just cleared."
        : ""),
  );
}

function clearChangeLog() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "Clear the Change Log?",
    "This deletes every row in the Change_Log sheet. It does not affect the highlight colors or notes on tracked cells.",
    ui.ButtonSet.YES_NO,
  );
  if (resp !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = getOrCreateChangeLogSheet_(ss);
  const lastRow = logSheet.getLastRow();
  if (lastRow > 1) {
    logSheet.deleteRows(2, lastRow - 1);
  }

  ui.alert("Change Log cleared.");
}

function showChangeLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateChangeLogSheet_(ss);
  sheet.showSheet();
  ss.setActiveSheet(sheet);
}
