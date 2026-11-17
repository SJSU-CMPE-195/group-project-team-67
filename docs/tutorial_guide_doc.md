# [Tutorial/Guide on Schedulr Operation]

## **Instructor Guide on Operating a Constraint-Based Faculty Scheduling Tool**

**CMPE 195 - Senior Design Project I & II,** **Spring/Fall 2026**


**Team Members:**

| Name | SJSU Email |
|---|---|
| Nick Mamaoag | nickjason.mamaoag@sjsu.edu  |
| Vance Nguyen | vance.nguyen@sjsu.edu  |
| Yashas Satheesh | yashas.satheesh@sjsu.edu  |

**Project Advisor:** [Carlos Rojas]

**Date****:** [11/24/2026] \
 \
 \
 \
 \
 \
 \
 \
 \
 \
 \
 \
 \
 \
 \
 \
 \
 \
 \
 \
 \
 \


1. Create a Copy of the Google Sheets ?CMPE 195 - Schedulr?

Google Sheets Link:

2. Open the Google Form ?Schedulr?

Google Form Link: [Schedulr](https://docs.google.com/forms/u/1/d/e/1FAIpQLSdLRE6gahZPBa8LZ2bJOkayNBDF4Q88hslxZGCcsHvdrPdGYg/formResponse)	 \


3. Turn on Google Apps Script API \
 \
Google Apps Script Link: [Google Apps Script API](https://script.google.com/home/usersettings) \


4. Link the Google Form to the Sheet and Publish the Sheet \
 \
 \
	If the Form Responses is titled ?Form Responses X? instead of ?Form Responses 1?, the inputs from the Google Form won?t be inputted.  \
 \
 \
 \
 \
 \
 \
 \
 \


5. Getting Started on the ?Schedulr? Button in Google Sheets \
 \
	Use the ?Schedulr? button next to ?Help? on Google Sheets. Displays a list of functions.  \
<span style="text - decoration: underline;">Click on Clean Preferences</span>

 \
		Parses and normalizes responses from Google Forms. (e.g., rows 11 to 13)


<span style="text - decoration: underline;">Click on Build Faculty Sheet</span> \
 \


Extracts faculty availability and maximum load. (e.g., row 3)

<span style="text - decoration: underline;">Click on Generate Schedule</span> \
 \


The schedule is generated based on past inputs and new input. (e.g., row 7). Conflicted schedules land in ?Conflicts?

6. Other Commands/Functions Under the ?Schedulr? Button

<span style="text - decoration: underline;">Changing Value(s)</span>

Changes values on the sheet. (e.g., change the number of seats open under ?Courses?. (42 to 35). It highlights the selected region and tells which email made the change and at what timestamp.)

<span style="text - decoration: underline;">View Change Log</span>

Has the ability to view every change on Google Sheets. (e.g., changing the value of 42 to 35 tells the timestamp, user, sheet, cell, previous and new values, as well as an additional description to describe the change.)


<span style="text - decoration: underline;">View Changed Rows</span>

           \
          Has the ability to view the changed row on Google Sheets. (e.g., row 4 fully lists the changed row with all its values.) \
 \
 \
 \
 \
<span style="text - decoration: underline;">Jump to Next Change</span>

Can jump to the next changed value. (e.g., test1 was changed to Software Engineering I, and it highlights the change.)


<span style="text - decoration: underline;">Undo All Highlighted Change(s)</span>

Can undo prior changed value(s). (e.g., Software Engineering I was changed to test1, 35 was changed to 42, and it dehighlights the change.)

<span style="text - decoration: underline;">Accepts All Change(s)</span>

Can accept all changed value(s). Confirms values, dehighlights, and removes notes from the tracked cells.

<span style="text - decoration: underline;">Clear Change Log</span>

Has the ability to clear the change log on Google Sheets. (e.g., makes the ?Change_Log? blank and ready for new changes.)


<span style="text - decoration: underline;">Rooms Available for CMPE/SE Departments</span>

Shows the class building and room numbers with the amount of seats listed for the CMPE/SE departments.



7. *Future Changes*

## References

[1] V. Nguyen, N. Jason, and Y. Satheesh, ?Schedulr,? *Github*, Feb. 06, 2026. https://github.com/SJSU-CMPE-195/group-project-team-67 (accessed Nov. 01, 2026).
