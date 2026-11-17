function generateSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const courses = readSheet(ss, 'Courses');
  const faculty = readSheet(ss, 'Faculty');
  const eligibility = readSheet(ss, 'Eligibility');
  const preferences = readSheet(ss, 'Preferences');

  const scheduleRows = [];
  const conflictRows = [];
  const facultyLoad = {};
  const facultyAssignments = {};

  faculty.forEach(f => {
    facultyLoad[f.facultyId] = 0;
    facultyAssignments[f.facultyId] = [];
  });

  const expandedCourses = expandCourses(courses);

  expandedCourses.forEach(course => {
    const allowedTerms = splitList(course.termsOffered);
    let assigned = false;
    let rejectionNotes = [];

    for (const term of allowedTerms) {
      const candidates = faculty
        .filter(f => isEligible(f.facultyId, course.courseCode, eligibility))
        .filter(f => isAvailableForTerm(f, term))
        .filter(f => !isUnavailableForSlot(f, course.timeSlot))
        .filter(f => !hasTimeConflict(f, term, course.timeSlot, facultyAssignments))
        .filter(f => !isOverloaded(f, facultyLoad))
        .sort((a, b) =>
          getPreferenceScore(a.facultyId, course.courseCode, preferences) -
          getPreferenceScore(b.facultyId, course.courseCode, preferences)
        );

      if (candidates.length > 0) {
        const picked = candidates[0];

        facultyLoad[picked.facultyId] += 1;
        facultyAssignments[picked.facultyId].push({
          term: term,
          timeSlot: course.timeSlot,
          courseCode: course.courseCode,
          section: course.section
        });

        scheduleRows.push([
          term,
          course.courseCode,
          course.courseName,
          course.section,
          picked.facultyId,
          picked.facultyName,
          course.timeSlot,
          'ASSIGNED',
          ''
        ]);

        assigned = true;
        break;
      } else {
        rejectionNotes.push(`No valid faculty for ${term}`);
      }
    }

    if (!assigned) {
      const candidateFaculty = faculty
        .filter(f => isEligible(f.facultyId, course.courseCode, eligibility))
        .map(f => f.facultyName)
        .join(', ');

      conflictRows.push([
        course.courseCode,
        course.courseName,
        allowedTerms.join(', '),
        course.section,
        course.timeSlot,
        rejectionNotes.join('; ') || 'No assignment possible',
        candidateFaculty
      ]);
    }
  });

  writeSheet(ss, 'ScheduleOutput', [
    'term',
    'courseCode',
    'courseName',
    'section',
    'facultyId',
    'facultyName',
    'timeSlot',
    'status',
    'notes'
  ], scheduleRows);

  writeSheet(ss, 'Conflicts', [
    'courseCode',
    'courseName',
    'termsTried',
    'section',
    'timeSlot',
    'reason',
    'candidateFaculty'
  ], conflictRows);
}

function readSheet(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`Missing sheet: ${name}`);

  const values = sheet.getDataRange().getValues();
  const headers = values.shift();

  return values
    .filter(row => row.some(cell => String(cell).trim() !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[String(header).trim()] = row[index];
      });
      return obj;
    });
}

function writeSheet(ss, name, headers, rows) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

function expandCourses(courses) {
  const expanded = [];

  courses.forEach(course => {
    const count = Number(course.sectionsNeeded || 1);
    for (let i = 1; i <= count; i++) {
      expanded.push({
        ...course,
        section: i
      });
    }
  });

  return expanded;
}

function splitList(value) {
  return String(value || '')
    .split(',')
    .map(x => x.trim())
    .filter(Boolean);
}

function isEligible(facultyId, courseCode, eligibilityRows) {
  return eligibilityRows.some(row =>
    String(row.facultyId).trim() === String(facultyId).trim() &&
    String(row.courseCode).trim() === String(courseCode).trim() &&
    String(row.eligible).trim().toUpperCase() === 'TRUE'
  );
}

function isAvailableForTerm(faculty, term) {
  const terms = splitList(faculty.availableTerms);
  return terms.includes(term);
}

function isUnavailableForSlot(faculty, timeSlot) {
  const blocked = splitList(faculty.unavailableSlots);
  return blocked.includes(String(timeSlot).trim());
}

function isOverloaded(faculty, facultyLoad) {
  return facultyLoad[faculty.facultyId] >= Number(faculty.maxLoad || 0);
}

function hasTimeConflict(faculty, term, timeSlot, facultyAssignments) {
  const assignments = facultyAssignments[faculty.facultyId] || [];
  return assignments.some(a =>
    String(a.term).trim() === String(term).trim() &&
    String(a.timeSlot).trim() === String(timeSlot).trim()
  );
}

function getPreferenceScore(facultyId, courseCode, preferences) {
  const row = preferences.find(p =>
    String(p.facultyId).trim() === String(facultyId).trim() &&
    String(p.courseCode).trim() === String(courseCode).trim()
  );

  if (!row) return 999;
  return Number(row.preferenceRank || 999);
}
