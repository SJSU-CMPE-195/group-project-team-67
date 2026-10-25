// ---------- CONFIG ----------
const ROUND_BY_CHOICE = true;      // true: everyone's 1st choice is placed (in priority order), then 2nd choices, etc.
                                   // false: each professor gets all their picks before the next professor starts
const INSTRUCTOR_AS_NAME = false;  // false: writes the email; true: "first.last@sjsu.edu" -> "First Last"


// ---------- UI: pick the two sheets ----------
function fill_in_responses() {
  const names = SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .map(s => s.getName()).filter(n => !/^_snap_/i.test(n));
  const resp = names.filter(n => /form response/i.test(n));
  const sched = names.filter(n => /schedule/i.test(n));
  if (!resp.length || !sched.length) {
    SpreadsheetApp.getUi().alert('Could not find a sheet with "Form Response" and a sheet with "Schedule" in its name.');
    return;
  }
  const j = a => JSON.stringify(a).replace(/</g, '\\u003c');
  const html = HtmlService.createHtmlOutput(
    '<style>body{font-family:Arial,sans-serif;padding:12px}label{display:block;margin-top:12px;font-weight:bold}' +
    'select{width:100%;padding:6px}button{margin-top:18px;padding:8px 16px}#msg{margin-top:14px;white-space:pre-wrap;font-size:13px}</style>' +
    '<label>Form responses sheet</label><select id="resp"></select>' +
    '<label>Schedule sheet to fill in</label><select id="sched"></select>' +
    '<button id="go" onclick="run()">Fill schedule</button>' +
    '<div id="msg"></div>' +
    '<script>' +
    'function opt(id,arr){var s=document.getElementById(id);arr.forEach(function(n){var o=document.createElement("option");o.value=n;o.textContent=n;s.appendChild(o);});}' +
    'opt("resp",' + j(resp) + ');opt("sched",' + j(sched) + ');' +
    'function run(){var b=document.getElementById("go"),m=document.getElementById("msg");b.disabled=true;m.textContent="Working...";' +
    'google.script.run.withSuccessHandler(function(r){m.textContent=r;b.disabled=false;})' +
    '.withFailureHandler(function(e){m.textContent="Error: "+e.message;b.disabled=false;})' +
    '.runFill(document.getElementById("resp").value,document.getElementById("sched").value);}' +
    '</script>'
  ).setWidth(400).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'Fill in schedule from responses');
}

// ---------- Main logic ----------
function runFill(respName, schedName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profs = readProfessors_(ss.getSheetByName(respName));
  const schedSheet = ss.getSheetByName(schedName);
  const sched = readSchedule_(schedSheet);

  // slot usage so far (spreads professors across slots)
  const usage = {};
  sched.sections.forEach(s => { if (s.preset) usage[s.preset.key] = (usage[s.preset.key] || 0) + 1; });

  // build the order in which (professor, choice) pairs are processed
  const maxPicks = Math.max(1, ...profs.map(p => p.max));
  const tasks = [];
  if (ROUND_BY_CHOICE) {
    for (let p = 0; p < maxPicks; p++) profs.forEach(prof => tasks.push([prof, p]));
  } else {
    profs.forEach(prof => { for (let p = 0; p < prof.max; p++) tasks.push([prof, p]); });
  }

  tasks.forEach(([prof, p]) => {
    if (prof.sections.length >= prof.max || !prof.choices.length) return;
    // choices beyond the 3rd pick cycle back (extra sections of their favorites)
    const code = prof.choices[p % prof.choices.length];
    assignSection_(prof, code, sched.sections, usage);
  });

  // write results
  const c = sched.cols;
  let filled = 0;
  sched.sections.forEach(s => {
    if (!s.assignedTo) return;
    filled++;
    schedSheet.getRange(s.row, c.inst + 1).setValue(s.assignedTo);
    if (s.outDays) {
      const d = schedSheet.getRange(s.row, c.days + 1); d.setNumberFormat('@'); d.setValue(s.outDays);
      const t = schedSheet.getRange(s.row, c.time + 1); t.setNumberFormat('@'); t.setValue(s.outTime);
    }
  });

  const unfilled = sched.sections.filter(s => !s.instructor && !s.assignedTo);
  const short = profs.filter(p => p.sections.length < p.max);
  return 'Filled ' + filled + ' section(s).\n\n' +
    (unfilled.length ? 'Still open (' + unfilled.length + '):\n' +
      unfilled.map(s => '  row ' + s.row + ': ' + s.code).join('\n') + '\n\n' : 'All sections filled.\n\n') +
    (short.length ? 'Professors below their max:\n' +
      short.map(p => '  ' + p.email + ' got ' + p.sections.length + ' of ' + p.max).join('\n') : '');
}

function assignSection_(prof, code, sections, usage) {
  const candidates = sections.filter(s => !s.instructor && !s.assignedTo && s.code === code);
  for (const s of candidates) {
    let slot = null;
    if (s.preset) {                                   // section already has day/time -> must match availability
      if (s.fixedUnknown) slot = 'any';
      else if (fits_(prof, s.preset) && prof.avail.some(a => a.key === s.preset.key)) slot = s.preset;
    } else {                                          // pick the least-used slot this professor can teach
      slot = prof.avail.filter(a => fits_(prof, a))
        .sort((a, b) => (usage[a.key] || 0) - (usage[b.key] || 0))[0] || null;
    }
    if (!slot) continue;

    s.assignedTo = INSTRUCTOR_AS_NAME ? nameFromEmail_(prof.email) : prof.email;
    if (slot !== 'any') {
      prof.slots.push(slot);
      usage[slot.key] = (usage[slot.key] || 0) + 1;
      if (!s.preset) { s.outDays = slot.days; s.outTime = slot.timeText; }
    }
    prof.sections.push(s);
    return true;
  }
  return false;
}

// not blocked by "cannot teach" and no overlap with their other sections
function fits_(prof, slot) {
  return !prof.cannot.some(x => overlap_(x, slot)) && !prof.slots.some(x => overlap_(x, slot));
}

// ---------- Readers ----------
function readProfessors_(sh) {
  if (!sh) throw new Error('Responses sheet not found.');
  const data = sh.getDataRange().getValues();
  const h = data[0].map(x => String(x).trim().toLowerCase());
  const col = re => h.findIndex(x => re.test(x));
  const c = { ts: col(/timestamp/), email: col(/^email/), c1: col(/first choice/), c2: col(/second choice/),
              c3: col(/third choice/), avail: col(/^availability/), cannot: col(/cannot/),
              max: col(/maximum/), sen: col(/seniority/) };
  const missing = Object.keys(c).filter(k => k !== 'sen' && c[k] < 0);
  if (missing.length) throw new Error('Missing columns in ' + sh.getName() + ': ' + missing.join(', '));

  const profs = [];
  data.slice(1).forEach(r => {
    const email = String(r[c.email]).trim();
    if (!email) return;
    let sen = c.sen >= 0 ? parseInt(r[c.sen], 10) : NaN;
    if (!(sen >= 1 && sen <= 5)) sen = 5;             // empty/invalid seniority = 5
    const t = r[c.ts] instanceof Date ? r[c.ts].getTime() : new Date(r[c.ts]).getTime();
    profs.push({
      email, seniority: sen, ts: isNaN(t) ? Infinity : t,
      choices: [r[c.c1], r[c.c2], r[c.c3]].map(courseCode_).filter(Boolean),
      avail: splitSlots_(r[c.avail]),
      cannot: splitSlots_(r[c.cannot]),
      max: parseInt(r[c.max], 10) || 1,
      sections: [], slots: []
    });
  });
  // seniority 1 first, then earliest timestamp
  profs.sort((a, b) => a.seniority - b.seniority || a.ts - b.ts);
  return profs;
}

function readSchedule_(sh) {
  if (!sh) throw new Error('Schedule sheet not found.');
  const data = sh.getDataRange().getValues();
  let hr = -1, cols = null;
  for (let i = 0; i < Math.min(data.length, 15) && hr < 0; i++) {
    const h = data[i].map(x => String(x).trim().toLowerCase());
    const c = { course: h.findIndex(x => /^course/.test(x)),
                inst: h.findIndex(x => /instructor|professor|faculty/.test(x)),
                days: h.findIndex(x => /^days?$/.test(x)),
                time: h.findIndex(x => /^time/.test(x)) };
    if (Object.values(c).every(v => v >= 0)) { hr = i; cols = c; }
  }
  if (hr < 0) throw new Error('Could not find a header row with Course, Days, Time and Instructor in ' + sh.getName());

  const sections = [];
  for (let i = hr + 1; i < data.length; i++) {
    const code = courseCode_(data[i][cols.course]);
    if (!code) continue;
    const days = String(data[i][cols.days]).trim(), time = String(data[i][cols.time]).trim();
    const sec = { row: i + 1, code, instructor: String(data[i][cols.inst]).trim() };
    if (days && time) {
      sec.preset = parseSlot_(days + ' ' + time);
      if (!sec.preset) sec.fixedUnknown = true;
    }
    sections.push(sec);
  }
  return { sections, cols };
}

// ---------- Helpers ----------
function courseCode_(v) {
  const m = String(v || '').toUpperCase().match(/([A-Z]{2,5})\s*-?\s*(\d{1,3}[A-Z]?)\b/);
  return m ? m[1] + ' ' + m[2] : '';
}

function parseSlot_(s) {
  const m = String(s).trim().match(/^([A-Za-z]+)\s+(\d{1,2}:\d{2}\s*[AP]M)\s*-\s*(\d{1,2}:\d{2}\s*[AP]M)$/i);
  if (!m) return null;
  const dayset = m[1].replace(/th/gi, 'R').toUpperCase().split('').sort();
  const start = toMin_(m[2]), end = toMin_(m[3]);
  return { days: m[1], dayset, start, end,
           timeText: m[2].replace(/\s/g, '') + '-' + m[3].replace(/\s/g, ''),
           key: dayset.join('') + ' ' + start + '-' + end };
}

function splitSlots_(v) {   // Friday slots are ignored
  return String(v || '').split(',').map(parseSlot_).filter(s => s && !s.dayset.includes('F'));
}

function toMin_(t) {
  const m = t.replace(/\s/g, '').match(/^(\d{1,2}):(\d{2})([AP])M$/i);
  let h = parseInt(m[1], 10) % 12;
  if (m[3].toUpperCase() === 'P') h += 12;
  return h * 60 + parseInt(m[2], 10);
}

function overlap_(a, b) {
  return a.dayset.some(d => b.dayset.includes(d)) && a.start < b.end && b.start < a.end;
}

function nameFromEmail_(e) {
  return e.split('@')[0].split('.').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
}