function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Schedulr")
    .addItem("Test connection", "testSheetConnection")
    .addSeparator()
    .addItem("Clean preferences", "cleanPreferences")
    .addItem("Build faculty sheet", "buildFacultySheet")
    .addItem("Build eligibility sheet", "buildEligibilitySheet")
    .addItem("Generate schedule", "generateSchedule")
    .addItem("Generate course overview", "generateCourseOverview")
    .addItem("Generate conflict summary", "generateConflictSummary")
    .addItem("Validate schedule", "validateSchedule")
    .addToUi();
}

function testSheetConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  sheet.getRange("A1").setValue("Apps Script connected");
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
    "facultyId",
    "facultyName",
    "email",
    "availableTerms",
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
      facultyId,
      facultyName,
      email,
      "Spring",
      unavailableSlots.join(", "),
      maxLoad,
    ]);
  }

  styleOutputSheet_(facultySheet);
}

function buildEligibilitySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = mustGetSheet_(ss, "Form Responses 1");
  const eligibilitySheet = getOrCreateSheet_(ss, "Eligibility");
  const data = rawSheet.getDataRange().getValues();

  eligibilitySheet.clear();
  eligibilitySheet.appendRow(["facultyId", "courseCode", "eligible"]);

  if (data.length < 2) {
    styleOutputSheet_(eligibilitySheet);
    return;
  }

  const seen = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const facultyName = String(row[1] || "").trim();
    const email = String(row[2] || "")
      .trim()
      .toLowerCase();
    const facultyId = email || facultyName;

    const previouslyTaught = splitMultiSelect_(row[9]).map(
      normalizeCourseCode_,
    );

    previouslyTaught.forEach((courseCode) => {
      if (!courseCode) return;
      const key = facultyId + "|" + courseCode;
      if (seen[key]) return;
      seen[key] = true;

      eligibilitySheet.appendRow([facultyId, courseCode, true]);
    });
  }

  styleOutputSheet_(eligibilitySheet);
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

  // Auto-build missing/empty Faculty and Eligibility
  const facultySheet = ensureFacultySheetReady_(ss, rawSheet);
  const eligibilitySheet = ensureEligibilitySheetReady_(ss, rawSheet);
  const timeSlotSheet = ss.getSheetByName("TimeSlots");
  const actualPrefSheet =
    ss.getSheetByName("Cleaned_Preferences") ||
    ss.getSheetByName("Preferences");

  const courses = getNormalizedCourses_(coursesSheet);
  const faculty = getNormalizedFaculty_(facultySheet, actualPrefSheet);
  const eligibility = getNormalizedEligibility_(eligibilitySheet);
  const preferences = getNormalizedPreferences_(actualPrefSheet, faculty);
  const timeSlots = getNormalizedTimeSlots_(timeSlotSheet);

  if (!courses.length) throw new Error("Courses sheet has no usable rows.");
  if (!faculty.length) throw new Error("Faculty sheet has no usable rows.");
  if (!eligibility.length)
    throw new Error("Eligibility sheet has no usable rows.");
  if (!timeSlots.length) throw new Error("No usable time slots found.");

  const scheduleRows = [];
  const conflictRows = [];
  const facultyLoad = {};
  const facultyAssignments = {};
  const slotUsage = {};

  faculty.forEach((f) => {
    facultyLoad[f.facultyId] = 0;
    facultyAssignments[f.facultyId] = [];
  });

  const expandedCourses = expandCourses_(courses).sort((a, b) =>
    compareExpandedCourses_(a, b, eligibility, preferences),
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

        if (isCourseSlotTaken_(course.courseCode, term, slot, slotUsage)) {
          rejectionNotes.push("Course already uses " + slot + " in " + term);
          continue;
        }

        const candidates = faculty
          .filter((f) =>
            isEligible_(f.facultyId, course.courseCode, eligibility),
          )
          .filter((f) => isAvailableForTerm_(f, term))
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
      markCourseSlotUsed_(
        course.courseCode,
        bestChoice.term,
        bestChoice.timeSlot,
        slotUsage,
      );

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
      const candidateFaculty = faculty
        .filter((f) => isEligible_(f.facultyId, course.courseCode, eligibility))
        .map((f) => f.facultyName)
        .join(", ");

      const eligibleCount = faculty.filter((f) =>
        isEligible_(f.facultyId, course.courseCode, eligibility),
      ).length;

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
      "facultyId",
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
      "eligibleFacultyCount",
      "candidateFaculty",
    ],
    conflictRows,
  );

  generateCourseOverview();
}

function generateCourseOverview() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outputSheet = getOrCreateSheet_(ss, "Course_Overview");
  outputSheet.clear();

  const conflictSheet = ss.getSheetByName("Conflicts");
  const coursesSheet = ss.getSheetByName("Courses");

  const conflictSet = {};
  if (conflictSheet && conflictSheet.getLastRow() > 1) {
    const rows = sheetToObjects_(conflictSheet);
    rows.forEach((r) => {
      const code = normalizeCourseCode_(
        pickField_(r, ["courseCode", "CourseCode", "course", "Course"]),
      );
      if (code) conflictSet[code] = true;
    });
  }

  const courseRows = coursesSheet ? getNormalizedCourses_(coursesSheet) : [];
  const lowerDiv = [];
  const upperDiv = [];
  const prereq195 = [];
  const techElective = [];
  const profNeeded = [];

  courseRows.forEach((course) => {
    const row = [course.courseCode, course.courseName];
    const category = String(course.category || "").trim();

    if (category === "Lower Div") lowerDiv.push(row);
    if (category === "Upper Div") upperDiv.push(row);
    if (category === "CMPE195 Prereq" || category === "195A Prereq")
      prereq195.push(row);
    if (category === "Tech Elective") techElective.push(row);
    if (conflictSet[course.courseCode]) profNeeded.push(row);
  });

  outputSheet
    .getRange(1, 1, 1, 5)
    .setValues([
      [
        "Lower Div Courses",
        "Upper Div Courses",
        "195A Prereq",
        "Tech Elective",
        "Prof Needed",
      ],
    ]);

  writeSingleColumnBlock_(
    outputSheet,
    2,
    1,
    lowerDiv.map((r) => [r[0] + " - " + r[1]]),
  );
  writeSingleColumnBlock_(
    outputSheet,
    2,
    2,
    upperDiv.map((r) => [r[0] + " - " + r[1]]),
  );
  writeSingleColumnBlock_(
    outputSheet,
    2,
    3,
    prereq195.map((r) => [r[0] + " - " + r[1]]),
  );
  writeSingleColumnBlock_(
    outputSheet,
    2,
    4,
    techElective.map((r) => [r[0] + " - " + r[1]]),
  );
  writeSingleColumnBlock_(
    outputSheet,
    2,
    5,
    profNeeded.map((r) => [r[0] + " - " + r[1]]),
  );

  styleOutputSheet_(outputSheet);
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
          "course title",
          "Course Title",
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
          "allowedTimeSlots",
          "AllowedTimeSlots",
          "allowedSlots",
          "AllowedSlots",
        ]),
      ).map(normalizeTimeSlot_),
      category: String(pickField_(r, ["category", "Category"]) || "").trim(),
    }))
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
      const facultyId = String(
        pickField_(r, [
          "facultyId",
          "FacultyId",
          "id",
          "ID",
          "email",
          "Email",
        ]) ||
          email ||
          facultyName,
      ).trim();
      const fallbackMax =
        email && maxFromPrefs[email] ? maxFromPrefs[email] : "";

      return {
        facultyId: facultyId,
        facultyName: facultyName || facultyId,
        email: email,
        availableTerms: splitList_(
          pickField_(r, ["availableTerms", "AvailableTerms", "terms", "Terms"]),
        ),
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

function getNormalizedEligibility_(sheet) {
  const rows = sheetToObjects_(sheet);
  return rows
    .map((r) => ({
      facultyId: String(
        pickField_(r, ["facultyId", "FacultyId", "id", "ID", "email", "Email"]),
      ).trim(),
      courseCode: normalizeCourseCode_(
        pickField_(r, ["courseCode", "CourseCode", "course", "Course"]),
      ),
      eligible: normalizeBoolean_(
        pickField_(r, ["eligible", "Eligible", "canTeach", "CanTeach"]),
      ),
    }))
    .filter((r) => r.facultyId && r.courseCode);
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

function getNormalizedTimeSlots_(sheet) {
  const defaults = [
    "MW 9:00-10:15",
    "MW 10:30-11:45",
    "MW 12:00-1:15",
    "MW 1:30-2:45",
    "MW 3:00-4:15",
    "TTh 9:00-10:15",
    "TTh 10:30-11:45",
    "TTh 12:00-1:15",
    "TTh 1:30-2:45",
    "TTh 3:00-4:15",
    "M 6:00-8:45",
    "T 6:00-8:45",
    "W 6:00-8:45",
    "Th 6:00-8:45",
  ];

  if (!sheet || sheet.getLastRow() < 2) {
    return defaults;
  }

  const rows = sheetToObjects_(sheet);
  const slots = rows
    .map((r) =>
      normalizeTimeSlot_(
        pickField_(r, ["timeSlot", "TimeSlot", "slot", "Slot"]),
      ),
    )
    .filter(Boolean);

  return slots.length ? uniqueList_(slots) : defaults;
}

function compareExpandedCourses_(a, b, eligibility, preferences) {
  const aEligible = countEligibleFaculty_(a.courseCode, eligibility);
  const bEligible = countEligibleFaculty_(b.courseCode, eligibility);
  if (aEligible !== bEligible) return aEligible - bEligible;

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
  else notes.push("eligible match");

  if (timeRank === 0) notes.push("preferred time matched");
  else notes.push("standard slot assigned");

  return notes.join("; ");
}

function countEligibleFaculty_(courseCode, eligibilityRows) {
  let count = 0;
  const seen = {};
  eligibilityRows.forEach((row) => {
    if (
      normalizeCourseCode_(row.courseCode) ===
        normalizeCourseCode_(courseCode) &&
      row.eligible === true
    ) {
      if (!seen[row.facultyId]) {
        seen[row.facultyId] = true;
        count += 1;
      }
    }
  });
  return count;
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

function isEligible_(facultyId, courseCode, eligibilityRows) {
  return eligibilityRows.some(
    (row) =>
      String(row.facultyId).trim() === String(facultyId).trim() &&
      normalizeCourseCode_(row.courseCode) ===
        normalizeCourseCode_(courseCode) &&
      row.eligible === true,
  );
}

function isAvailableForTerm_(faculty, term) {
  if (!faculty.availableTerms.length) return true;
  return faculty.availableTerms.includes(String(term).trim());
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

function isCourseSlotTaken_(courseCode, term, timeSlot, slotUsage) {
  const key =
    String(term).trim() +
    "|" +
    normalizeCourseCode_(courseCode) +
    "|" +
    normalizeTimeSlot_(timeSlot);
  return slotUsage[key] === true;
}

function markCourseSlotUsed_(courseCode, term, timeSlot, slotUsage) {
  const key =
    String(term).trim() +
    "|" +
    normalizeCourseCode_(courseCode) +
    "|" +
    normalizeTimeSlot_(timeSlot);
  slotUsage[key] = true;
}

function collapseReasons_(reasons) {
  return uniqueList_(reasons.filter(Boolean)).slice(0, 8).join("; ");
}

function classifyConflictType_(reasons, eligibleCount) {
  if (eligibleCount === 0) return "NO_ELIGIBLE_FACULTY";
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

function ensureEligibilitySheetReady_(ss, rawSheet) {
  let sheet = ss.getSheetByName("Eligibility");
  if (!sheet || sheet.getLastRow() < 2) {
    buildEligibilitySheet();
    sheet = ss.getSheetByName("Eligibility");
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
    NO_ELIGIBLE_FACULTY: 0,
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
    NO_ELIGIBLE_FACULTY: "#F4CCCC", // red - needs Eligibility sheet fix
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

  const eligibilitySheet = ss.getSheetByName("Eligibility");
  const coursesSheet = ss.getSheetByName("Courses");

  const scheduleRows = sheetToObjects_(outputSheet);
  const eligibilityRows = eligibilitySheet
    ? sheetToObjects_(eligibilitySheet)
    : [];
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
    const facultyId = String(row["facultyId"] || "").trim();
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

    // Check 2: eligibility
    const eligible = eligibilityRows.some(
      (e) =>
        String(e["facultyId"]).trim() === facultyId &&
        normalizeCourseCode_(e["courseCode"]) === courseCode &&
        normalizeBoolean_(e["eligible"]) === true,
    );
    if (!eligible) {
      logRows.push([
        rowNum,
        courseCode,
        facultyId,
        "FAIL",
        "NOT_ELIGIBLE",
        facultyId + " is not listed as eligible to teach " + courseCode,
      ]);
    }

    // Track load and section counts for checks 3 and 4
    facultyLoadMap[facultyId] = (facultyLoadMap[facultyId] || 0) + 1;
    courseSectionMap[courseCode] = (courseSectionMap[courseCode] || 0) + 1;
  });

  // Check 3: max load exceeded
  // Re-read faculty sheet to get maxLoad per faculty
  const facultySheet = ss.getSheetByName("Faculty");
  if (facultySheet && facultySheet.getLastRow() > 1) {
    const facultyRows = sheetToObjects_(facultySheet);
    facultyRows.forEach((f) => {
      const facultyId = String(f["facultyId"] || "").trim();
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

  // Check 4: sections assigned vs sections needed
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
        "facultyId",
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
