//Recovers per-cell old values on multi-cell edits (paste / fill / delete) by diffing against hidden mirror sheets named "_snap_<SheetName>".
//Old/new values are display strings on both sides, so type coercion can't cause false diffs.

const SNAP_PREFIX = "_snap_";

function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    if (CHANGE_EXCLUDED_SHEETS.indexOf(sheetName) !== -1) return;
    if (sheetName.indexOf(SNAP_PREFIX) === 0) return;

    // Don't track edits to the header row (but keep the snapshot current).
    if (e.range.getRow() === 1 && e.range.getNumRows() === 1) {
      syncSnapshotRange_(sheet, e.range);
      return;
    }

    const numCells = e.range.getNumRows() * e.range.getNumColumns();

    if (numCells === 1) {
      const oldValue = e.oldValue !== undefined ? e.oldValue : "";
      const newValue = e.value !== undefined ? e.value : e.range.getValue();
      if (normalizeValue_(oldValue) !== normalizeValue_(newValue)) {
        recordChange_(sheet, e.range, oldValue, newValue, false);
      }
      syncSnapshotRange_(sheet, e.range);
    } else {
      // Multi-cell: diff against the snapshot. If this sheet has no snapshot
      // yet, fall back to your original "bulk edit, value not captured" path.
      if (!trackBulkEdit_(sheet, e.range)) {
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
    }
  } catch (err) {
    // Never let tracking errors block the user's actual edit.
    console.error(err);
  }
}

// snapshot helpers

function isTracked_(sheet) {
  const n = sheet.getName();
  return (
    CHANGE_EXCLUDED_SHEETS.indexOf(n) === -1 && n.indexOf(SNAP_PREFIX) !== 0
  );
}

function getSnapSheet_(sheet, create) {
  const ss = sheet.getParent();
  const name = SNAP_PREFIX + sheet.getName();
  let snap = ss.getSheetByName(name);
  if (!snap && create) {
    const active = ss.getActiveSheet(); // insertSheet steals focus
    snap = ss.insertSheet(name);
    snap.hideSheet();
    if (active) ss.setActiveSheet(active);
  }
  return snap;
}

function ensureSize_(snap, rows, cols) {
  if (snap.getMaxRows() < rows)
    snap.insertRowsAfter(snap.getMaxRows(), rows - snap.getMaxRows());
  if (snap.getMaxColumns() < cols)
    snap.insertColumnsAfter(snap.getMaxColumns(), cols - snap.getMaxColumns());
}

function colToLetter_(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// Rebuild snapshots for every tracked sheet. Safe to call repeatedly.
function seedSnapshots() {
  SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .filter(isTracked_)
    .forEach(function (s) {
      reseedSheet_(s);
    });
}

function reseedSheet_(sheet) {
  if (!isTracked_(sheet)) return;
  const snap = getSnapSheet_(sheet, true);
  snap.clear();
  const r = sheet.getLastRow(),
    c = sheet.getLastColumn();
  if (r < 1 || c < 1) return;
  ensureSize_(snap, r, c);
  snap
    .getRange(1, 1, r, c)
    .setNumberFormat("@")
    .setValues(sheet.getRange(1, 1, r, c).getDisplayValues());
}

//Copy current display values of rng into the snapshot.
//Does nothing if the sheet has no snapshot (so a half-empty snapshot is never created).
//Call after any edit or script write to a tracked sheet.
function syncSnapshotRange_(sheet, rng) {
  if (!isTracked_(sheet)) return;
  const snap = getSnapSheet_(sheet, false);
  if (!snap) return;
  ensureSize_(snap, rng.getLastRow(), rng.getLastColumn());
  snap
    .getRange(
      rng.getRow(),
      rng.getColumn(),
      rng.getNumRows(),
      rng.getNumColumns(),
    )
    .setNumberFormat("@")
    .setValues(rng.getDisplayValues());
}

//multi-cell diff

// Returns false if the sheet has no snapshot (caller falls back to the old bulk path).
function trackBulkEdit_(sheet, rng) {
  const snap = getSnapSheet_(sheet, false);
  if (!snap) return false;

  ensureSize_(snap, rng.getLastRow(), rng.getLastColumn());
  const oldVals = snap
    .getRange(
      rng.getRow(),
      rng.getColumn(),
      rng.getNumRows(),
      rng.getNumColumns(),
    )
    .getDisplayValues();
  const newVals = rng.getDisplayValues();

  recordBulkEdit_(sheet, rng, oldVals, newVals);
  syncSnapshotRange_(sheet, rng);
  return true;
}

//For a whole pasted range: one read/write of pendingChanges, one Change_Log insert, one notes write, one Changed_Rows refresh.
//Entries are stored as normal (non-bulk) history, so Undo Changes works on them.
function recordBulkEdit_(sheet, rng, oldVals, newVals) {
  const r0 = rng.getRow(),
    c0 = rng.getColumn();
  const sheetName = sheet.getName();
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
  const notes = rng.getNotes();
  const highlight = [],
    unhighlight = [],
    logEntries = [];

  for (let i = 0; i < newVals.length; i++) {
    if (r0 + i === 1) continue; // header row
    for (let j = 0; j < newVals[i].length; j++) {
      const oldV = oldVals[i][j],
        newV = newVals[i][j];
      if (normalizeValue_(oldV) === normalizeValue_(newV)) continue;

      const a1 = colToLetter_(c0 + j) + (r0 + i);
      const id = sheetName + "!" + a1;
      const existing = map[id];
      const isRevert = !!(
        existing &&
        existing.hasOriginal &&
        normalizeValue_(existing.originalValue) === normalizeValue_(newV)
      );

      const entryText =
        formatValueForNote_(oldV) +
        " -> " +
        formatValueForNote_(newV) +
        "\nChanged by: " +
        user +
        "\nChanged at: " +
        timestamp;

      if (!existing)
        map[id] = { history: [], originalValue: oldV, hasOriginal: true };
      map[id].reverted = isRevert;
      map[id].history.push({
        previousValue: oldV,
        isBulk: false,
        noteText: entryText,
      });

      notes[i][j] = map[id].history
        .map(function (h) {
          return h.noteText;
        })
        .join("\n\n");
      (isRevert ? unhighlight : highlight).push(a1);
      logEntries.push({
        a1: a1,
        prev: formatValueForNote_(oldV),
        next: formatValueForNote_(newV),
      });
    }
  }

  if (!logEntries.length) return;

  // Save state first so a storage failure doesn't leave highlights with no history.
  try {
    props.setProperty("pendingChanges", JSON.stringify(map));
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Too many tracked changes to store - this edit was not tracked.",
      "Schedulr",
      8,
    );
    return;
  }

  rng.setNotes(notes);
  if (highlight.length)
    sheet.getRangeList(highlight).setBackground(CHANGE_HIGHLIGHT_COLOR);
  if (unhighlight.length) sheet.getRangeList(unhighlight).setBackground(null);

  logChangesBatch_(sheetName, logEntries, user, now);
  refreshChangedRowsSheet_();
}

//Batched version of logChange_: one date header (if needed) + N rows under it.
function logChangesBatch_(sheetName, entries, user, timestamp) {
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
    logSheet.insertRowBefore(2);
    logSheet.getRange(2, 1, 1, 7).setBackground(CHANGE_LOG_DATE_HEADER_COLOR);
    logSheet.getRange(2, 1).setValue(dayStart).setNumberFormat("M/d/yyyy");
  }

  logSheet.insertRowsBefore(3, entries.length);
  logSheet.getRange(3, 1, entries.length, 7).setBackground(null);
  logSheet.getRange(3, 1, entries.length, 6).setValues(
    entries.map(function (en) {
      return [timestamp, user, sheetName, en.a1, en.prev, en.next];
    }),
  );
}

//row/column insert & delete

//Structural changes shift positions, so rebuild that sheet's snapshot.
function onStructureChange(e) {
  const types = ["INSERT_ROW", "REMOVE_ROW", "INSERT_COLUMN", "REMOVE_COLUMN"];
  if (types.indexOf(e.changeType) === -1) return;
  const sheet = SpreadsheetApp.getActiveSheet();
  if (isTracked_(sheet)) reseedSheet_(sheet);
}

//Run once (onChange needs an installable trigger).
function installStructureTrigger() {
  ScriptApp.newTrigger("onStructureChange")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange()
    .create();
}
