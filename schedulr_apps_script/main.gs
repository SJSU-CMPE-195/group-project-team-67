function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Schedulr")
    .addItem("Clean preferences", "cleanPreferences")
    .addItem("Build faculty sheet", "buildFacultySheet")
    .addItem("Generate schedule", "generateSchedule")
    // .addItem("Generate course overview", "generateCourseOverview") // dont need for now
    .addItem("Generate conflict summary", "generateConflictSummary")
    .addItem("Validate schedule", "validateSchedule")
    .addSeparator()
    .addItem("View Change Log", "showChangeLog")
    .addItem("View Changed Rows", "showChangedRows")
    .addItem("Clear Change Log", "clearChangeLog")
    .addItem("Jump to Next Change", "jumpToNextChange")
    .addItem("Accept All Changes", "acceptAllChanges")
    .addItem("Undo Changes", "undoAllChanges")
    .addToUi();
  SpreadsheetApp.getUi()
    .createMenu("Courses")
    .addItem("Create blank schedule", "create_blank_schedule")
    .addItem("Pull Semester Schedule", "pull_semester")
    .addItem("Fill in Responses", "fill_in_responses")
    .addItem("Comparing Room Size (Current Page)", "comparing_room_size")
    //.addItem("Update Courses", "update_courses")
    .addToUi();
  try {
    seedSnapshots();
  } catch (err) {
    console.error(err);
  }
}

/**
 * Finds the form response tab by its headers instead of assuming a specific
 * tab name such as "Form Responses 1".
 */
function getFormResponsesSheet_(ss) {
  const candidates = ss.getSheets().filter((sheet) => {
    const values = sheet.getDataRange().getValues();
    if (!values.length) return false;
    const headers = values[0].map(normalizeRawFormHeader_);
    const hasEmail = findRawFormHeader_(headers, ["email", "email_address"]) !== -1;
    const hasCourseChoice = headers.some(
      (header) =>
        header.indexOf("choice_course") !== -1 ||
        header.indexOf("course_preference") !== -1,
    );
    return hasEmail && hasCourseChoice;
  });

  if (!candidates.length) {
    throw new Error(
      "No form response sheet was found. The sheet must contain Email and course-choice columns.",
    );
  }

  const namedCandidate = candidates.find((sheet) =>
    /^form responses?(?: \d+)?$/i.test(sheet.getName().trim()),
  );
  return namedCandidate || candidates[0];
}

function getRawFormColumns_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error("The form response sheet is empty.");

  const headers = values[0].map(normalizeRawFormHeader_);
  const preferences = [
    findRawFormHeader_(headers, ["first_choice_course", "course_preference_1", "course_1"]),
    findRawFormHeader_(headers, ["second_choice_course", "course_preference_2", "course_2"]),
    findRawFormHeader_(headers, ["third_choice_course", "course_preference_3", "course_3"]),
  ];

  const columns = {
    email: findRawFormHeader_(headers, ["email", "email_address", "e_mail"]),
    preferences: preferences,
    availability: findRawFormHeader_(headers, [
      "availability",
      "available_slots",
      "available_times",
    ]),
    unavailableSlots: findRawFormHeader_(headers, [
      "time_slots_you_cannot_teach",
      "slots_you_cannot_teach",
      "unavailable_slots",
      "blocked_slots",
    ]),
    maxCourses: findRawFormHeader_(headers, [
      "maximum_courses_willing_to_teach",
      "max_courses",
      "maximum_courses",
      "courses_willing_to_teach",
      "max_load",
    ]),
  };

  if (columns.email === -1) {
    throw new Error("The form response sheet is missing an Email column.");
  }
  if (columns.preferences.every((index) => index === -1)) {
    throw new Error("The form response sheet is missing course-choice columns.");
  }
  return columns;
}

function normalizeRawFormHeader_(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/&/g, "and")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function findRawFormHeader_(headers, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const alias = normalizeRawFormHeader_(aliases[i]);
    const exactIndex = headers.indexOf(alias);
    if (exactIndex !== -1) return exactIndex;
  }
  return -1;
}

function cleanPreferences() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = getFormResponsesSheet_(ss);
  const columns = getRawFormColumns_(rawSheet);

  const cleanSheet = getOrCreateSheet_(ss, "Cleaned_Preferences");
  const data = rawSheet.getDataRange().getValues();

  cleanSheet.clear();
  cleanSheet.appendRow([
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

    const email = String(row[columns.email] || "")
      .trim()
      .toLowerCase();
    if (!email) continue; // no email, no way to track this professor

    const pref1 = normalizeCourseCode_(
      columns.preferences[0] === -1 ? "" : row[columns.preferences[0]],
    );
    const pref2 = normalizeCourseCode_(
      columns.preferences[1] === -1 ? "" : row[columns.preferences[1]],
    );
    const pref3 = normalizeCourseCode_(
      columns.preferences[2] === -1 ? "" : row[columns.preferences[2]],
    );

    const availabilityList = splitMultiSelect_(
      columns.availability === -1 ? "" : row[columns.availability],
    ).map(normalizeTimeSlot_);
    const maxCourses = Number(
      columns.maxCourses === -1 ? 0 : row[columns.maxCourses] || 0,
    );

    // Use the first available slot as "preferredTime" for scoring.
    const preferredTime = availabilityList.length ? availabilityList[0] : "";

    if (pref1)
      cleanSheet.appendRow([email, pref1, 1, preferredTime, maxCourses]);
    if (pref2 && pref2 !== pref1)
      cleanSheet.appendRow([email, pref2, 2, preferredTime, maxCourses]);
    if (pref3 && pref3 !== pref1 && pref3 !== pref2)
      cleanSheet.appendRow([email, pref3, 3, preferredTime, maxCourses]);
  }

  styleOutputSheet_(cleanSheet);
}

function buildFacultySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = getFormResponsesSheet_(ss);
  const columns = getRawFormColumns_(rawSheet);
  const facultySheet = getOrCreateSheet_(ss, "Faculty");
  const data = rawSheet.getDataRange().getValues();

  facultySheet.clear();
  facultySheet.appendRow(["email", "unavailableSlots", "maxLoad"]);

  if (data.length < 2) {
    styleOutputSheet_(facultySheet);
    return;
  }

  const seen = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const email = String(row[columns.email] || "")
      .trim()
      .toLowerCase();
    const unavailableSlots = splitMultiSelect_(
      columns.unavailableSlots === -1 ? "" : row[columns.unavailableSlots],
    ).map(normalizeTimeSlot_);
    const maxLoad = Number(
      columns.maxCourses === -1 ? 0 : row[columns.maxCourses] || 0,
    );

    if (!email || seen[email]) continue; // no email, no way to track this professor
    seen[email] = true;

    facultySheet.appendRow([email, unavailableSlots.join(", "), maxLoad]);
  }

  styleOutputSheet_(facultySheet);
}

function generateSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const coursesSheet = mustGetSheet_(ss, "Courses");
  const rawSheet = getFormResponsesSheet_(ss);
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
      const candidateFaculty = faculty.map((f) => f.email).join(", ");
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
      "candidateEmails",
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
      const fallbackMax =
        email && maxFromPrefs[email] ? maxFromPrefs[email] : "";

      return {
        facultyId: email,
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

  return rows
    .map((r) => {
      const email = String(pickField_(r, ["email", "Email"]) || "")
        .trim()
        .toLowerCase();

      return {
        facultyId: email,
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
  reseedSheet_(sheet);
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

const CHANGE_LOG_SHEET_NAME = "Change_Log";
const CHANGED_ROWS_SHEET_NAME = "Changed_Rows";
const CHANGE_HIGHLIGHT_COLOR = "#F4CCCC";
const CHANGE_LOG_DATE_HEADER_COLOR = "#FFF2CC";
const CHANGE_EXCLUDED_SHEETS = [
  "Conflicts",
  "Conflict_Summary",
  "Validation_Log",
  CHANGE_LOG_SHEET_NAME,
  CHANGED_ROWS_SHEET_NAME,
];

function recordChange_(sheet, cell, oldValue, newValue, isBulk) {
  const a1 = cell.getA1Notation();
  const id = sheet.getName() + "!" + a1;
  const user = getEditingUser_();
  const now = new Date();
  const timestamp = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    "MM/dd/yyyy HH:mm",
  );

  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty("pendingChanges");
  const map = raw ? JSON.parse(raw) : {};
  const existing = map[id];
  const isRevert = !!(
    existing &&
    existing.hasOriginal &&
    normalizeValue_(existing.originalValue) === normalizeValue_(newValue)
  );

  const previousText = isBulk
    ? "(bulk edit - previous value not captured)"
    : formatValueForNote_(oldValue);
  const newText = formatValueForNote_(newValue);

  const entryText =
    previousText +
    " -> " +
    newText +
    "\nChanged by: " +
    user +
    "\nChanged at: " +
    timestamp;

  cell.setBackground(isRevert ? null : CHANGE_HIGHLIGHT_COLOR);

  logChange_(sheet.getName(), a1, previousText, newText, user, now);
  addPendingChange_(sheet.getName(), a1, oldValue, isBulk, entryText, isRevert);
  refreshCellNote_(sheet, a1);
  refreshChangedRowsSheet_();
}

function getEditingUser_() {
  try {
    const email = Session.getActiveUser().getEmail();
    return email || "Unknown user";
  } catch (err) {
    return "Unknown user";
  }
}

function normalizeValue_(value) {
  if (value === null || value === undefined) return "";
  const text = String(value).trim();
  if (/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(text)) {
    return String(Number(text));
  }
  return text;
}

function formatValueForNote_(value) {
  const text = normalizeValue_(value);
  return text === "" ? "(empty)" : text;
}

function logChange_(sheetName, a1, previousValue, newValue, user, timestamp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = getOrCreateChangeLogSheet_(ss);
  const dayStart = new Date(
    timestamp.getFullYear(),
    timestamp.getMonth(),
    timestamp.getDate(),
  );

  const hasExistingRows = logSheet.getLastRow() >= 2;
  const topRowIsToday =
    hasExistingRows && isSameDay_(logSheet.getRange(2, 1).getValue(), dayStart);

  if (!topRowIsToday) {
    // Insert a date-header row right under the column headers if new day, above whatever day was previously on top.
    logSheet.insertRowBefore(2);
    const headerRange = logSheet.getRange(2, 1, 1, 7);
    headerRange.setBackground(CHANGE_LOG_DATE_HEADER_COLOR);
    logSheet.getRange(2, 1).setValue(dayStart).setNumberFormat("M/d/yyyy");
  }

  // New entry always goes right under that day's header, so entries for the same day stack below their date row without touching earlier days.
  // insertRowBefore copies the background of the adjacent row, so the new row picks up the date header's highlight - clear it explicitly.
  // Column 7 (description) is left blank here — it's a manual field, only filled in if the user types into it directly.
  logSheet.insertRowBefore(3);
  logSheet.getRange(3, 1, 1, 7).setBackground(null);
  logSheet
    .getRange(3, 1, 1, 6)
    .setValues([[timestamp, user, sheetName, a1, previousValue, newValue]]);
}

function isSameDay_(a, b) {
  if (!(a instanceof Date) || !(b instanceof Date)) return false;
  return (
    a.getFullYear() === b.getFullYear() &&
    a.getMonth() === b.getMonth() &&
    a.getDate() === b.getDate()
  );
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
      "description",
    ]);
    styleOutputSheet_(sheet);
    sheet.hideSheet();
  }
  return sheet;
}

function getOrCreateChangedRowsSheet_(ss) {
  let sheet = ss.getSheetByName(CHANGED_ROWS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CHANGED_ROWS_SHEET_NAME);
    sheet.hideSheet();
  }
  return sheet;
}

function refreshChangedRowsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty("pendingChanges");
  const map = raw ? JSON.parse(raw) : {};

  // Group changed cells by sheet + row so each row is only pulled once.
  const rowsBySheet = {};
  Object.keys(map).forEach((id) => {
    const parts = id.split("!");
    const sheetName = parts[0];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const row = sheet.getRange(parts[1]).getRow();
    if (!rowsBySheet[sheetName]) rowsBySheet[sheetName] = [];
    if (rowsBySheet[sheetName].indexOf(row) === -1)
      rowsBySheet[sheetName].push(row);
  });

  const out = getOrCreateChangedRowsSheet_(ss);
  out.clear();

  let outRow = 1;
  Object.keys(rowsBySheet).forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    const lastCol = sheet ? sheet.getLastColumn() : 0;
    if (!sheet || !lastCol) return;

    out.getRange(outRow, 1).setValue(sheetName).setFontWeight("bold");
    outRow++;
    out
      .getRange(outRow, 1, 1, lastCol)
      .setValues(sheet.getRange(1, 1, 1, lastCol).getValues());
    outRow++;

    rowsBySheet[sheetName]
      .sort((a, b) => a - b)
      .forEach((row) => {
        out
          .getRange(outRow, 1, 1, lastCol)
          .setValues(sheet.getRange(row, 1, 1, lastCol).getValues());
        outRow++;
      });
    outRow++;
  });
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
  refreshChangedRowsSheet_();
}

function addPendingChange_(
  sheetName,
  a1,
  rawOldValue,
  isBulk,
  entryText,
  isReverted,
) {
  const props = PropertiesService.getDocumentProperties();
  const key = "pendingChanges";
  const raw = props.getProperty(key);
  const map = raw ? JSON.parse(raw) : {};
  const id = sheetName + "!" + a1;

  if (!map[id]) {
    map[id] = {
      history: [],
      originalValue: isBulk
        ? null
        : rawOldValue === undefined
          ? ""
          : rawOldValue,
      hasOriginal: !isBulk,
    };
  }

  map[id].reverted = !!isReverted;

  map[id].history.push({
    previousValue: isBulk ? null : rawOldValue,
    isBulk: !!isBulk,
    noteText: entryText,
  });

  props.setProperty(key, JSON.stringify(map));
}

function refreshCellNote_(sheet, a1) {
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty("pendingChanges");
  const map = raw ? JSON.parse(raw) : {};
  const id = sheet.getName() + "!" + a1;
  const entry = map[id];
  const range = sheet.getRange(a1);

  if (!entry || !entry.history || !entry.history.length) {
    range.clearNote();
    return;
  }

  range.setNote(entry.history.map((h) => h.noteText).join("\n\n"));
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
    "This confirms the current values are correct and removes the highlight colors and notes from tracked cells. The values themselves are kept as-is, nothing is reverted.",
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
  refreshChangedRowsSheet_();
  ui.alert("Accepted " + ids.length + " change(s).");
}

function undoAllChanges() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "Undo tracked changes?",
    "This steps each tracked cell back by one edit (its most recent change) and updates the note to show whatever history is left. The cell stays highlighted as long as history remains; once it's undone all the way back to its original value, the highlight clears automatically. Cells changed by a bulk edit (paste/fill) don't have a captured previous value, so that edit is dropped from the note without changing the cell's value.",
    ui.ButtonSet.YES_NO,
  );
  if (resp !== ui.Button.YES) return;

  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty("pendingChanges");
  const map = raw ? JSON.parse(raw) : {};
  const ids = Object.keys(map);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let undone = 0;
  let skipped = 0;

  ids.forEach((id) => {
    const parts = id.split("!");
    const sheetName = parts[0];
    const a1 = parts[1];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const entry = map[id];
    const history = (entry && entry.history) || [];
    if (!history.length) return; // already fully undone, stays highlighted

    const last = history.pop();
    const range = sheet.getRange(a1);

    if (last.isBulk) {
      skipped++;
    } else {
      range.setValue(last.previousValue);
      syncSnapshotRange_(sheet, range);
      undone++;
    }

    if (history.length) {
      range.setNote(history.map((h) => h.noteText).join("\n\n"));
      // Still has history left, stays highlighted.
    } else {
      range.clearNote();
      range.setBackground(null);
      delete map[id];
    }
  });

  props.setProperty("pendingChanges", JSON.stringify(map));
  refreshChangedRowsSheet_();

  ui.alert(
    "Undid " +
      undone +
      " cell(s)." +
      (skipped > 0
        ? " " +
          skipped +
          " bulk-edit entr" +
          (skipped === 1 ? "y" : "ies") +
          " couldn't be undone and " +
          (skipped === 1 ? "was" : "were") +
          " just dropped from the note."
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

function showChangedRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateChangedRowsSheet_(ss);
  sheet.showSheet();
  ss.setActiveSheet(sheet);
}
