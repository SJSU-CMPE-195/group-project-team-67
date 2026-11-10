function showWelcomeIfNeeded() {
  const properties = PropertiesService.getUserProperties();

  if (properties.getProperty("welcomeShown") === "true") return;

  showWelcomeGuide();

  properties.setProperty("welcomeShown", "true");
}

function showWelcomeGuide() {
  const html = HtmlService.createHtmlOutput(
    `
    <html>
      <body style="font-family:Arial,sans-serif;padding:20px;color:#202124">
        <h2>How to use</h2>

        <p>This spreadsheet helps you pull and organize SJSU course sections. While also keeping a log of all changes</p>

         <h3>Getting Started Using the Schedulr Section</h3>
        <ol>
          <li>Use the <b>Schedulr</b> button to select a <b>build faculty sheet</b>. This will create a sheet with all the faculty's unwanted times.</li>
          <li>Click <b>Generate Schedule</b> to create a new sheet called ScheduleOutput.</li>
          <li>This will use the Courses sheet and create a schedule based the courses and clean preferences page.</li>
        </ol>
        <h3>Change Log</h3>
        <ol>
          <li>Any manual changes will be recorded and can viewed in change log</li>
          <li>Manual change will be marked in red with a log of recent changes</li>
          <li>Accept all changes will get rid of all mark red ups</li>
        </ol>


        <h3>Getting Started Using the Courses Section</h3>
        <ol>
          <li>Use the <b>Courses</b> button to select <b>Pull Semester Schedule</b> and select a semester and course group.</li>
          <li>Choose a course group, such as CS or CMPE.</li>
          <li>Click <b>Pull</b> to create a new sheet with information pull from SJSU course catalog.</li>
          <li>Next <b>Create Blank Schedule</b> will take a chosen schedule and remove the days, time and instructor while replaces the opens seats with the maximum open seats from the Rooms sheet</li>
          <li><b>Fill in Responses</b> will take a choosen response from and New Schedule try to feel the time and course depending on the form response, timestamp of when form was filled out</li>
          <li>Staff priority comes from the <b>Rank</b> column on the form responses sheet (1 = highest). The admin enters each staff member's rank manually; blank ranks are placed last</li>
          <li>Comparing room will look at the current page the user is on and see if it has a days, time, location and open seats. Then it will compare all courses with open seat of 0 with any courses with an open seat greater than 2. This will give a list of all course's room that can be swapped.
        </ol>

        <p>For detailed instructions, open the <b>README</b> sheet.</p>

        <button onclick="google.script.host.close()">Got it</button>
      </body>
    </html>
  `,
  )
    .setWidth(650)
    .setHeight(550);

  SpreadsheetApp.getUi().showModalDialog(html, "Welcome to Schedulr");
}

function openReadme() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName("README");

  if (!sheet) {
    sheet = spreadsheet.insertSheet("README");
    sheet.getRange("A1").setValue("Schedulr User Guide");
    sheet.getRange("A3").setValue("Pull Semester");
    sheet
      .getRange("B3")
      .setValue(
        "Opens a popup where you can select a semester, year, and course group. Type part of a course group or choose one from the suggestions.",
      );

    sheet.getRange("A4").setValue("Authorize Schedulr");
    sheet
      .getRange("B4")
      .setValue(
        "Run this once if Google asks for permission to access SJSU schedule pages.",
      );

    sheet.getRange("A5").setValue("Generated semester sheets");
    sheet
      .getRange("B5")
      .setValue(
        "Each pull creates a new sheet containing the matching course sections.",
      );

    sheet.getRange("A1:B1").setFontWeight("bold");
    sheet.getRange("A3:A5").setFontWeight("bold");
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 600);
    sheet.getRange("B:B").setWrap(true);
  }

  spreadsheet.setActiveSheet(sheet);
}

function resetWelcomeGuide() {
  PropertiesService.getUserProperties().deleteProperty("welcomeShown");
  SpreadsheetApp.getUi().alert(
    "The welcome guide will appear the next time you open the spreadsheet.",
  );
}

function setupWelcomeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();

  const alreadyExists = triggers.some(
    (trigger) => trigger.getHandlerFunction() === "showWelcomeOnOpen",
  );

  if (!alreadyExists) {
    ScriptApp.newTrigger("showWelcomeOnOpen")
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onOpen()
      .create();
  }

  SpreadsheetApp.getUi().alert("Welcome trigger has been installed.");
}

function showWelcomeOnOpen() {
  showWelcomeGuide();
}
