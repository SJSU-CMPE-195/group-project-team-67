/**
 * Opens a semester/course-group selector and pulls matching SJSU sections.
 */
function pull_semester() {
  const currentYear = new Date().getFullYear();
  const years = [];
  for (let year = currentYear - 6; year <= currentYear + 2; year++) years.push(year);
  const yearOptions = years.map((year) =>
    "<option value='" + year + "'" +
    (year === currentYear ? " selected" : "") +
    ">" + year + "</option>"
  ).join("");

  const html = HtmlService.createHtmlOutput(
    "<html><head><base target='_top'><style>" +
      "body{font-family:Arial,sans-serif;padding:16px;color:#202124}" +
      "label{display:block;font-weight:bold;margin-top:12px;margin-bottom:4px}" +
      "select{width:100%;padding:7px;font-size:13px}" +
      "button{margin-top:18px;padding:8px 18px;background:#1a73e8;color:white;border:0;border-radius:4px;cursor:pointer}" +
      "button:disabled{background:#9aa0a6;cursor:default}" +
      "#status{margin-top:12px;color:#5f6368;font-size:12px}" +
      "</style></head><body>" +
      "<h2>Pull SJSU Semester Data</h2>" +
      "<label>Schedule source</label>" +
      "<select id='source' onchange='loadGroups()'>" +
      "<option value='current'>Current semester</option>" +
      "<option value='archive'>Archive</option></select>" +
      "<label>Semester</label>" +
      "<select id='semester' onchange='loadGroups()'>" +
      "<option value='fall'>Fall</option>" +
      "<option value='spring'>Spring</option>" +
      "<option value='summer'>Summer</option>" +
      "<option value='winter'>Winter</option></select>" +
      "<label>Year</label><select id='year' onchange='loadGroups()'>" +
      yearOptions + "</select>" +
      "<label>Course group</label>" +
      "<select id='group' disabled><option>Loading course groups...</option></select>" +
      "<button id='pullButton' onclick='pullData()' disabled>Pull</button>" +
      "<div id='status'>Choose a semester and year to load course groups.</div>" +
      "<script>" +
      "function setStatus(text){document.getElementById('status').textContent=text;}" +
      "function loadGroups(){" +
      "var source=document.getElementById('source').value;" +
      "var semester=document.getElementById('semester').value;" +
      "var year=document.getElementById('year').value;" +
      "var group=document.getElementById('group');" +
      "var button=document.getElementById('pullButton');" +
      "group.disabled=true;button.disabled=true;" +
      "group.innerHTML='<option>Loading course groups...</option>';" +
      "setStatus('Loading course groups from SJSU...');" +
      "google.script.run.withSuccessHandler(function(groups){" +
      "group.innerHTML='';" +
      "if(!groups.length){group.innerHTML='<option>No course groups found</option>';" +
      "setStatus('No course groups were found for that selection.');return;}" +
      "groups.forEach(function(value){var option=document.createElement('option');" +
      "option.value=value;option.textContent=value;group.appendChild(option);});" +
      "group.disabled=false;button.disabled=false;setStatus(groups.length+' course group(s) found.');" +
      "}).withFailureHandler(function(error){group.innerHTML='<option>Unable to load groups</option>';" +
      "setStatus(error.message||error);}).ps_get_course_groups(source,semester,year);}" +
      "function pullData(){" +
      "var button=document.getElementById('pullButton');button.disabled=true;" +
      "setStatus('Pulling selected sections...');" +
      "google.script.run.withSuccessHandler(function(result){setStatus(result.message);" +
      "button.disabled=false;}).withFailureHandler(function(error){setStatus(error.message||error);" +
      "button.disabled=false;}).ps_pull_semester(" +
      "document.getElementById('source').value,document.getElementById('semester').value," +
      "document.getElementById('year').value,document.getElementById('group').value);}" +
      "loadGroups();</script></body></html>",
  ).setWidth(420).setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, "Pull Semester");
}

/**
 * Run this once from the Apps Script editor before using pull_semester().
 * Google displays the authorization screen for UrlFetchApp here, where it
 * can be completed successfully.
 */
function authorizeSchedulr() {
  const response = UrlFetchApp.fetch(
    "https://www.sjsu.edu/classes/schedules/fall-2026.php",
    {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { "User-Agent": "Mozilla/5.0" },
    },
  );

  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error("Authorization test request failed with HTTP " + status + ".");
  }

  SpreadsheetApp.getUi().alert(
    "Authorization complete. You can now use Pull semester data.",
  );
}

/**
 * Pulls all CMPE sections from the hardcoded Fall 2025 archive page.
 *
 * To pull a different semester or subject later, change SCHEDULE_URL and
 * COURSE_GROUP below.
 */
function pulled_semester() {
  const SCHEDULE_URL =
    "https://www.sjsu.edu/classes/schedules/archive/fall-2025.php";
  const COURSE_GROUP = "CMPE";
  const scheduleMatch = SCHEDULE_URL.match(
    /\/(fall|spring|summer|winter)-(\d{4})\.php\/?$/i,
  );
  if (!scheduleMatch) {
    throw new Error(
      "SCHEDULE_URL must end with a semester and year, such as fall-2026.php.",
    );
  }
  const semester =
    scheduleMatch[1].charAt(0).toUpperCase() +
    scheduleMatch[1].slice(1).toLowerCase();
  const year = scheduleMatch[2];

  const response = UrlFetchApp.fetch(SCHEDULE_URL, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { "User-Agent": "Mozilla/5.0" },
  });
  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error(
      "SJSU schedule request failed with HTTP " +
        status +
        ": " +
        SCHEDULE_URL,
    );
  }

  const sections = psParseSchedule_(response.getContentText()).filter(
    (section) => section.course.indexOf(COURSE_GROUP + " ") === 0,
  );
  if (!sections.length) {
    throw new Error(
      "No " + COURSE_GROUP + " sections were found at " + SCHEDULE_URL,
    );
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = psUniqueSheetName_(
    ss,
    "Pulled_" + semester + "_" + year + "_" + COURSE_GROUP,
  );
  const sheet = ss.insertSheet(sheetName);
  const headers = [
    "Course",
    "Section",
    "Days",
    "Time",
    "Instructor",
    "Location",
    "Open Seats",
  ];
  const rows = sections.map((section) => [
    section.course,
    section.section,
    section.days,
    section.time,
    section.instructor,
    section.location,
    section.openSeats,
  ]);

  sheet
    .getRange(1, 1, rows.length + 1, headers.length)
    .setValues([headers, ...rows]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  Logger.log(
    "Created " +
      sheetName +
      " with " +
      rows.length +
      " CMPE sections from Fall 2025.",
  );
}

function ps_get_course_groups(source, semester, year) {
  const groups = {};
  psFetchSections_(source, semester, year).forEach((section) => {
    const match = section.course.match(/^([A-Z]{2,5})\s+\d+/);
    if (match) groups[match[1]] = true;
  });
  return Object.keys(groups).sort();
}

function ps_pull_semester(source, semester, year, group) {
  const normalizedGroup = String(group || "").trim().toUpperCase();
  if (!normalizedGroup) throw new Error("Choose a course group.");
  const selected = psFetchSections_(source, semester, year).filter((section) =>
    section.course.indexOf(normalizedGroup + " ") === 0,
  );
  if (!selected.length) throw new Error("No sections were found for that course group.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = psUniqueSheetName_(
    ss,
    "Pull_" + String(semester).charAt(0).toUpperCase() +
      String(semester).slice(1).toLowerCase() + "_" + year + "_" + normalizedGroup,
  );
  const sheet = ss.insertSheet(sheetName);
  const headers = ["Course", "Section", "Days", "Time", "Instructor", "Location", "Open Seats"];
  const rows = selected.map((section) => [
    section.course, section.section, section.days, section.time,
    section.instructor, section.location, section.openSeats,
  ]);
  sheet.getRange(1, 1, rows.length + 1, headers.length).setValues([headers, ...rows]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  const message = "Created " + sheetName + " with " + rows.length +
    " " + normalizedGroup + " section(s).";
  SpreadsheetApp.getUi().alert(message);
  return { sheetName: sheetName, message: message };
}

function psFetchSections_(source, semester, year) {
  const url = psBuildUrl_(source, semester, year);
  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { "User-Agent": "Mozilla/5.0" },
  });
  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error("SJSU schedule request failed with HTTP " + status + ": " + url);
  }
  const sections = psParseSchedule_(response.getContentText());
  if (!sections.length) throw new Error("No schedule sections were found at: " + url);
  return sections;
}

function psBuildUrl_(source, semester, year) {
  const sourceValue = String(source || "").toLowerCase();
  const semesterValue = String(semester || "").toLowerCase();
  const yearValue = String(year || "").trim();
  if (sourceValue !== "current" && sourceValue !== "archive") {
    throw new Error("Invalid schedule source.");
  }
  if (["fall", "spring", "summer", "winter"].indexOf(semesterValue) === -1) {
    throw new Error("Invalid semester.");
  }
  if (!/^\d{4}$/.test(yearValue)) throw new Error("Invalid year.");
  return "https://www.sjsu.edu/classes/schedules/" +
    (sourceValue === "archive" ? "archive/" : "") +
    semesterValue + "-" + yearValue + ".php";
}

function psParseSchedule_(html) {
  const rowMatches = html.match(/<tr\b[\s\S]*?<\/tr>/gi) || [];
  let header = null;
  const sections = [];
  rowMatches.forEach((rowHtml) => {
    const cells = psExtractCells_(rowHtml).map(psStripHtml_);
    const lower = cells.map((cell) => cell.toLowerCase());
    if (lower.indexOf("section") !== -1 && lower.indexOf("open seats") !== -1) {
      header = {
        section: lower.indexOf("section"),
        days: lower.indexOf("days"),
        time: lower.indexOf("times"),
        instructor: lower.indexOf("instructor"),
        location: lower.indexOf("location"),
        openSeats: lower.indexOf("open seats"),
      };
      return;
    }
    if (!header) return;

    const match = (cells[header.section] || "").match(
      /^(.+?)\s+\(Section\s+([^)]+)\)/i,
    );
    if (!match) return;
    const openSeatsText = (cells[header.openSeats] || "").replace(/[^0-9-]/g, "");
    if (!openSeatsText) return;

    sections.push({
      course: psNormalizeCourse_(match[1]),
      section: psNormalizeSection_(match[2]),
      days: psCell_(cells, header.days),
      time: psCell_(cells, header.time),
      instructor: psCell_(cells, header.instructor),
      location: psCell_(cells, header.location),
      openSeats: Number(openSeatsText),
    });
  });
  return sections.filter((section) => section.course && section.section);
}

function psExtractCells_(rowHtml) {
  const cells = rowHtml.match(/<t[dh]\b[\s\S]*?<\/t[dh]>/gi) || [];
  return cells.map((cell) =>
    cell.replace(/^<t[dh]\b[^>]*>/i, "").replace(/<\/t[dh]>$/i, ""),
  );
}

function psCell_(cells, index) {
  return index === -1 || index >= cells.length ? "" : cells[index];
}

function psStripHtml_(value) {
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

function psNormalizeCourse_(value) {
  const text = String(value || "").toUpperCase().replace(/\s+/g, " ").trim();
  const match = text.match(/^([A-Z]{2,5})\s*(\d+[A-Z]?)/);
  return match ? match[1] + " " + match[2] : text;
}

function psNormalizeSection_(value) {
  const text = String(value || "").trim().toUpperCase();
  return /^\d+$/.test(text) ? String(Number(text)) : text.replace(/\s+/g, "");
}

function psUniqueSheetName_(ss, baseName) {
  const cleanName = String(baseName).replace(/[\\/?*[\]:]/g, "_").slice(0, 90);
  let name = cleanName;
  let counter = 2;
  while (ss.getSheetByName(name)) {
    name = (cleanName.slice(0, 86) + "_" + counter).slice(0, 100);
    counter++;
  }
  return name;
}

