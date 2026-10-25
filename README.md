# Schedulr

> A constraint-based faculty scheduling tool that automatically assigns professors to courses and generates time slots based on preferences, availability, and eligibility.

## Group 7

| Name    | GitHub                                                 | Email                      |
| ------- | ------------------------------------------------------ | -------------------------- |
| Nick    | [@nick-mama](https://github.com/nick-mama)             | nickjason.mamaoag@sjsu.edu |
| Vance   | [@vancednguyen](https://github.com/vancednguyen)       | vance.nguyen@sjsu.edu      |
| Yashas  | [@YashSat21](https://github.com/YashSat21)             | yashas.satheesh@sjsu.edu   |

**Advisor:** Professor Rojas

---

## Project Description

Schedulr is a faculty course planning system that automates the process of assigning professors to courses and generating valid schedules. It uses instructor preferences, availability, prior teaching history, and institutional constraints to produce a feasible schedule.

## Problem Statement

Departments currently rely on manual spreadsheets to plan CMPE and SE course offerings, making demand forecasting and multi-term coordination inefficient and error-prone. It’s difficult to balance prerequisites, outside course schedules, faculty preferences, and workload constraints without creating conflicts or mismatched capacity.

## Proof of Concept Scope

This PoC demonstrates:

- End-to-end scheduling pipeline:
  - Google Form → Data cleaning → Faculty & eligibility generation → Schedule output
- Automatic assignment of:
  - Professors to courses
  - Time slots to course sections
- Constraint handling:
  - Instructor availability
  - Maximum teaching load
  - Course eligibility
  - Time conflicts

**Not included yet:**

- UI/dashboard
- Optimization beyond greedy assignment
- Room assignment

## Live PoC

Google Sheet (Scheduler): https://docs.google.com/spreadsheets/d/1lojaKJ_upQNaAoNeLLrqR7sFTTfOgHD1zpiIz_9MTRs/edit?usp=sharing

Google Form (Input): https://forms.gle/BSXZEJVfRYdtAs1t9

Use the form to submit sample responses, then run the scheduler from the Sheet menu (see **Running the PoC** section).

## Prerequisites

### Apps Script/clasp

- Google account
- Google Sheets
- Node.js (for clasp)
- clasp CLI

Install clasp:

```bash
npm install -g @google/clasp
clasp login
```

### Python Scraper

- Python 3.8+
- pip

Install dependencies:

```bash
cd sjsu_data_scraper
pip install requests beautifulsoup4 pandas lxml
```

## Installation

```bash
git clone https://github.com/SJSU-CMPE-195/group-project-team-67.git
cd group-project-team-67
```

To pull the latest Apps Script code:

```bash
cd schedulr_apps_script
clasp pull
```

## Apps Script Setup

This project uses `clasp` to sync local code with Google Apps Script.

You must create a `.clasp.json` file in the **root directory** of the project.

### Quick Setup:

1. Run:

**Mac/Linux:**

```bash
cp .clasp.json.txt .clasp.json
```

**Windows (PowerShell):**

```powershell
Copy-Item .clasp.json.txt .clasp.json
```

2. Open the new `.clasp.json` file

3. Replace the placeholder with the Apps Script ID

To get the script ID:

- Open the Google Sheet
- Go to **Extensions -> Apps Script**
- Click **Project Settings**
- Copy the **Script ID**

Example `.clasp.json` file:

```json
{
  "scriptId": "YOUR_SCRIPT_ID_HERE",
  "rootDir": "schedulr_apps_script"
}
```

## Running the PoC

1. Open the connected Google Sheet
2. Submit sample responses through the Google Form
3. In the Sheet menu, click **Schedulr** and run the steps in order:

| Step | Menu Item                 | Description                                                           |
| ---- | ------------------------- | --------------------------------------------------------------------- |
| 1    | Import scraper CSV        | Load course data from the SJSU scraper CSV via Drive ID               |
| 2    | Clean preferences         | Parse and normalize raw Form responses                                |
| 3    | Build faculty sheet       | Extract faculty availability and load limits                          |
| 4    | Build eligibility sheet   | Map which faculty are eligible for which courses                      |
| 5    | Generate schedule         | Run the constraint-based scheduling algorithm                         |
| 6    | Generate conflict summary | Summarize unresolved conflicts by type                                |
| 7    | Validate schedule         | Check the output for double-bookings, overloads, and missing sections |

4. View outputs in the following sheets:

| Sheet            | Contents                                              |
| ---------------- | ----------------------------------------------------- |
| ScheduleOutput   | Assigned professors, courses, and time slots          |
| Conflicts        | Courses that could not be assigned and the reasons    |
| Conflict_Summary | Conflict counts grouped by type, color-coded          |
| Validation_Log   | Per-row validation results (eligibility, load, slots) |

## Running the Scraper

The scraper pulls current CMPE/SE course data from the SJSU catalog and class schedule.

```bash
cd sjsu_data_scraper
python sjsu_class_scraper.py
```

Output: `sjsu_major_courses_spring_2026.csv`

Upload the file to Google Drive, then use **Import scraper CSV** in the Schedulr menu and paste the Drive file ID when prompted.

## Demo

### Schedule Output

![Schedule Output](docs/screenshots/schedule.png)

Shows assigned professors and generated time slots based on preferences and constraints.

### Conflicts

![Conflicts](docs/screenshots/conflicts.png)

Shows courses that could not be assigned due to constraints such as time conflicts, limited availability, or maximum load limits. Candidate faculty are listed to show why no valid assignment could be made within the current constraints.

## Tech Stack

| Category  | Technology             |
| --------- | ---------------------- |
| Backend   | Google Apps Script     |
| Data      | Google Sheets          |
| Input     | Google Forms           |
| Scripting | JavaScript             |
| Scraping  | Python (BeautifulSoup) |
| Tools     | Git, clasp             |

## What's Next (195B)

- Build web-based schedule visualization (React or Next.js)
- Add JSON export from Apps Script to feed the frontend
- Improve scheduling algorithm (constraint optimization)
- Add UI dashboard
- Integrate real historical course data
- Add room assignment
- Improve preference weighting system

## Project Structure

```
.
├── schedulr_apps_script/ # Apps Script scheduler (main.gs + appsscript.json)
├── sjsu_data_scraper/ # Python scraper + output CSV
├── docs/
│   └── screenshots/        # Demo screenshots
├── src/                    # Frontend (195B — in development)
├── tests/                  # Tests (in development)
├── .clasp.json.txt         # Clasp config template (copy to .clasp.json)
├── .gitignore
└── README.md
```

## License

This project is licensed under the <FILL IN> License - see the [LICENSE](LICENSE) file for details.

---

_CMPE 195A/B - Senior Design Project | San Jose State University | Spring 2026_
