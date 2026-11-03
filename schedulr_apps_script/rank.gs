// Rank column on the form responses sheet.
// 1 = highest priority. Order: Seniority (High Priority > Senior > N/A/blank),
// then earliest submission time. One rank per staff member (by email); if a
// staff member submitted more than once, every one of their rows shows the same
// rank, based on their best seniority and their earliest timestamp.

const RANK_HEADER = "Rank";

// Pure logic (no Sheets calls): rows = [{ email, seniority, ts, row }]
// returns { [email]: rank }
function computeRanks_(rows) {
  const tier = (s) => {
    const v = String(s || "")
      .trim()
      .toLowerCase();
    return v === "high priority" ? 0 : v === "senior" ? 1 : 2;
  };

  const people = {};
  rows.forEach((r) => {
    const email = String(r.email || "")
      .trim()
      .toLowerCase();
    if (!email) return;
    const t = tier(r.seniority);
    const ts = r.ts instanceof Date ? r.ts.getTime() : new Date(r.ts).getTime();
    const p =
      people[email] ||
      (people[email] = { email, tier: 2, ts: Infinity, row: Infinity });
    p.tier = Math.min(p.tier, t);
    p.ts = Math.min(p.ts, isNaN(ts) ? Infinity : ts);
    p.row = Math.min(p.row, r.row);
  });

  const ranks = {};
  Object.values(people)
    .sort((a, b) => a.tier - b.tier || a.ts - b.ts || a.row - b.row)
    .forEach((p, i) => (ranks[p.email] = i + 1));
  return ranks;
}

function isFormResponsesSheet_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return false;
  const h = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map((x) => String(x).trim().toLowerCase());
  return (
    h.some((x) => /^email/.test(x)) &&
    h.some((x) => /timestamp/.test(x)) &&
    h.some((x) => /seniority/.test(x))
  );
}

function updateRanks(sheet) {
  sheet =
    sheet || getFormResponsesSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map((x) => String(x).trim().toLowerCase());

  const emailCol = headers.findIndex((x) => /^email/.test(x));
  const tsCol = headers.findIndex((x) => /timestamp/.test(x));
  const senCol = headers.findIndex((x) => /seniority/.test(x));
  if (emailCol < 0 || tsCol < 0 || senCol < 0) return;

  // Rank column: reuse if it exists, otherwise add it after the last column.
  let rankCol = headers.indexOf(RANK_HEADER.toLowerCase());
  if (rankCol < 0) {
    rankCol = lastCol;
    const hdr = sheet.getRange(1, rankCol + 1);
    sheet
      .getRange(1, senCol + 1)
      .copyTo(hdr, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    hdr.setValue(RANK_HEADER);
  }
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rows = data.map((r, i) => ({
    email: r[emailCol],
    seniority: r[senCol],
    ts: r[tsCol],
    row: i + 2,
  }));
  const ranks = computeRanks_(rows);

  const out = data.map((r) => {
    const email = String(r[emailCol] || "")
      .trim()
      .toLowerCase();
    return [email ? ranks[email] : ""];
  });
  sheet.getRange(2, rankCol + 1, out.length, 1).setValues(out);
}

// Called from onEdit (snapshotDiff.gs): re-rank when the admin changes Seniority
// or when any of the key columns is edited by hand.
function maybeUpdateRanks_(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (!isFormResponsesSheet_(sheet)) return;
    updateRanks(sheet);
  } catch (err) {
    console.error(err);
  }
}

// New form submissions don't fire onEdit; this runs from an installable
// "On form submit" trigger. Run setupRankTrigger() once to install it.
function onFormSubmitRank(e) {
  const sheet =
    (e && e.range && e.range.getSheet()) ||
    getFormResponsesSheet_(SpreadsheetApp.getActiveSpreadsheet());
  updateRanks(sheet);
}

function setupRankTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exists = ScriptApp.getProjectTriggers().some(
    (t) => t.getHandlerFunction() === "onFormSubmitRank",
  );
  if (!exists) {
    ScriptApp.newTrigger("onFormSubmitRank")
      .forSpreadsheet(ss)
      .onFormSubmit()
      .create();
  }
}
