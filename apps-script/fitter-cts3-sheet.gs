// Generated from presets/fitter-cts3.json — do not edit this block; edit the preset and regenerate.
var PRESET = {
 "sheetTitle": "Fitter logbook data — Govt. ITI Nedumangad",
 "tradeName": "Fitter",
 "los": [
  {
   "number": "1",
   "title": "Plan and organize the work to make job as per specification applying different types of basic fitting operation and check for dimensional accuracy following safety precautions.",
   "description": "Basic fitting operation – marking, Hacksawing, Chiselling, Filing, Drilling, Taping and Grinding etc. Accuracy: ± 0.25mm",
   "year": 1
  },
  {
   "number": "2",
   "title": "Manufacture simple sheet metal items as per drawing and join them by soldering, brazing and riveting.",
   "description": "",
   "year": 1
  },
  {
   "number": "3",
   "title": "Join metal components by riveting observing standard procedure.",
   "description": "",
   "year": 1
  },
  {
   "number": "4",
   "title": "Join metal component by arc welding observing standard procedure.",
   "description": "",
   "year": 1
  },
  {
   "number": "5",
   "title": "Cut and join metal component by gas (oxyacetylene / LPG / oxy-hydrogen).",
   "description": "",
   "year": 1
  },
  {
   "number": "6",
   "title": "Produce components by different operations and check accuracy using appropriate measuring instruments.",
   "description": "Different Operations - Drilling, Reaming, Taping, Dieing; Appropriate Measuring Instrument – Vernier, Screw Gauge, Micrometer",
   "year": 1
  },
  {
   "number": "7",
   "title": "Make different fit of components for assembling as per required tolerance observing principle of interchangeability and check for functionality.",
   "description": "Different Fit – Sliding, Angular, Step fit, ‘T’ fit, Square fit and Profile fit; Required tolerance: ±0.04 mm, angular tolerance: 30 min.",
   "year": 1
  },
  {
   "number": "8",
   "title": "Produce components involving different operations on lathe observing standard procedure and check for accuracy.",
   "description": "Different Operations – facing, plain turning, step turning, parting, chamfering, shoulder turn, grooving, knurling, boring, taper turning, threading (external ‘V’ only)",
   "year": 1
  },
  {
   "number": "9",
   "title": "Plan & perform simple repair, overhauling of different machines and check for functionality.",
   "description": "Different Machines – Drill Machine, Power Saw, Bench Grinder and Lathe",
   "year": 1
  },
  {
   "number": "10",
   "title": "Read and apply engineering drawing for different application in the field of work.",
   "description": "",
   "year": 1
  },
  {
   "number": "11",
   "title": "Demonstrate basic mathematical concept and principles to perform practical operations. Understand and explain basic science in the field of study.",
   "description": "",
   "year": 1
  },
  {
   "number": "12",
   "title": "Perform surface finishing operations using different fastening components, tools and check functionality.",
   "description": "Different fastening components – Dowel pins, screws, bolts, keys and cotters; Different fastening tools – hand operated & power tools; Required tolerance: ±0.02mm, angular tolerance ± 10 min.",
   "year": 2
  },
  {
   "number": "13",
   "title": "Make & assemble components of different mating surfaces as per required tolerance.",
   "description": "Different Mating Surfaces – Dovetail fitting, Radius fitting, Combined fitting; Different surface finishing operations – Scraping, Lapping and Honing",
   "year": 2
  },
  {
   "number": "14",
   "title": "Make different gauges by using standard tools & equipment and checks for specified accuracy.",
   "description": "Different Gauges – Snap gauge, Gap gauge; Specified Accuracy: ±0.02mm",
   "year": 2
  },
  {
   "number": "15",
   "title": "Apply a range of skills to execute pipe joints, dismantle and assemble valves & fittings with pipes and test for leakages.",
   "description": "Range of skills – Cutting, Threading, Flaring, Bending and Joining",
   "year": 2
  },
  {
   "number": "16",
   "title": "Make drill jig & produce components on drill machine by using jigs and check for correctness.",
   "description": "",
   "year": 2
  },
  {
   "number": "17",
   "title": "Plan, dismantle, repair and assemble different damaged mechanical components used for power transmission & check functionality.",
   "description": "Different Damaged Mechanical Components – Pulley, Gear, Keys, Jibs and Shafts",
   "year": 2
  },
  {
   "number": "18",
   "title": "Identify, dismantle, replace and assemble different pneumatics and hydraulics components.",
   "description": "Different components – Compressor, Pressure Gauge, Filter Regulator Lubricator, Valves and Actuators",
   "year": 2
  },
  {
   "number": "19",
   "title": "Construct circuit of pneumatics and hydraulics observing standard operating procedure & safety aspect.",
   "description": "",
   "year": 2
  },
  {
   "number": "20",
   "title": "Plan & perform basic day to day preventive maintenance, repairing and check functionality.",
   "description": "Simple Machines – Drill Machine, Power Saw and Lathe",
   "year": 2
  },
  {
   "number": "21",
   "title": "Plan, erect simple machine and test machine tool accuracy.",
   "description": "Simple Machines – Drill Machine, Power Saw and Lathe",
   "year": 2
  },
  {
   "number": "22",
   "title": "Read and apply engineering drawing for different application in the field of work.",
   "description": "",
   "year": 2
  }
 ]
};

// ─────────────────────────────────────────────────────────────────────────────
//  LOGBOOK DATA-ENTRY SHEET BUILDER  (Google Apps Script)
//
//  HOW TO USE
//   1. Make a new, blank Google Sheet.            (sheets.new)
//   2. Extensions → Apps Script → delete what is there → paste this whole file → Save.
//   3. Pick  setupLogbookSheet  in the function list → Run. Allow the permissions (it only touches this sheet).
//   4. Go back to the sheet. It now has the tabs below and a "📘 Logbook" menu.
//
//  TABS
//   Read me · LearningOutcomes (already filled) · Jobs (practical) · Lessons (theory) · Tasks (logbook tasks)
//
//  When the sheet is filled:  File → Download → Microsoft Excel (.xlsx), then import it with
//  scripts/migrate-copa.mjs --preset presets/<trade>.json   (see FITTER-SETUP.md).
//  Photos are added later inside the app, not in this sheet.
// ─────────────────────────────────────────────────────────────────────────────

var MAX_ROWS = 500;          // rows prepared with formulas on Jobs / Lessons / Tasks
var HEAD_BG = '#1a3a5c', HEAD_FG = '#ffffff';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📘 Logbook')
    .addItem('Set up the sheets (first time only)', 'setupLogbookSheet')
    .addItem('Check my data before exporting', 'checkLogbookData')
    .addToUi();
}

function setupLogbookSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename(PRESET.sheetTitle);
  var first = ss.getSheets()[0];
  var sheets = {};
  ['Read me', 'LearningOutcomes', 'Jobs', 'Lessons', 'Tasks'].forEach(function (name, i) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = i === 0 && first.getLastRow() === 0 && !ss.getSheetByName('Read me') ? first.setName(name) : ss.insertSheet(name);
    sheets[name] = sh;
  });
  buildReadMe_(sheets['Read me']);
  buildLearningOutcomes_(sheets['LearningOutcomes']);
  buildJobs_(sheets['Jobs']);
  buildLessons_(sheets['Lessons']);
  buildTasks_(sheets['Tasks']);
  ss.setActiveSheet(sheets['Read me']);
  SpreadsheetApp.flush();
  try { SpreadsheetApp.getUi().alert('Done. Fill the Jobs, Lessons and Tasks tabs — see "Read me".'); } catch (e) {}
}

// ── helpers ──────────────────────────────────────────────────────────────────
function header_(sh, labels, widths) {
  var r = sh.getRange(1, 1, 1, labels.length);
  r.setValues([labels]).setFontWeight('bold').setBackground(HEAD_BG).setFontColor(HEAD_FG).setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
}
// IDs such as 1.10 must stay TEXT, otherwise Sheets turns them into the number 1.1
function asText_(sh, col) { sh.getRange(2, col, MAX_ROWS, 1).setNumberFormat('@'); }
function fillFormula_(sh, col, formulaForRow2) {
  var f = [];
  for (var i = 0; i < MAX_ROWS; i++) f.push([formulaForRow2.replace(/\{r\}/g, String(i + 2))]);
  sh.getRange(2, col, MAX_ROWS, 1).setFormulas(f).setFontColor('#6b7280').setBackground('#f3f4f6');
}
function dupRule_(sh, col, letter) {
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($' + letter + '2<>"",COUNTIF($' + letter + '$2:$' + letter + ',$' + letter + '2)>1)')
    .setBackground('#fecaca').setRanges([sh.getRange(2, col, MAX_ROWS, 1)]).build();
  var rules = sh.getConditionalFormatRules(); rules.push(rule); sh.setConditionalFormatRules(rules);
}
function clear_(sh) { sh.clear(); sh.clearConditionalFormatRules(); }

// ── tabs ─────────────────────────────────────────────────────────────────────
function buildReadMe_(sh) {
  clear_(sh);
  var lines = [
    [PRESET.sheetTitle],
    [''],
    ['What this sheet is for'],
    ['Type the syllabus split-up of the ' + PRESET.tradeName + ' trade here: practical JOBS, theory LESSONS and logbook TASKS. Later it is imported into the logbook app in one go.'],
    [''],
    ['Numbering (important)'],
    ['• Job and lesson numbers start with the learning-outcome number:  job 3.1, 3.2 … 3.10 belong to LO 3.  Lessons the same: 3.1, 3.2 …'],
    ['• Task numbers add one more level:  task 3.2.1 is the first logbook task of job 3.2.'],
    ['• The LO column and the Job column are filled automatically from the number. Do not type in the grey columns.'],
    ['• Number cells are TEXT, so 3.10 stays 3.10 (it is not turned into 3.1). Red number = the same number was typed twice.'],
    [''],
    ['Tabs'],
    ['LearningOutcomes — already filled from the syllabus (' + PRESET.los.length + ' LOs). You may correct a title or add a description.'],
    ['Jobs — practical jobs: Job No · title · notes (optional) · hours (optional).'],
    ['Lessons — theory lessons: Lesson No · title · hours (optional) · notes. Notes can be plain text; start a line with "- " for bullet points.'],
    ['Tasks — logbook tasks of a job: Task No · name · Aim · Procedure · Result · Conclusion. Plain text; line breaks are kept.'],
    [''],
    ['When you have finished'],
    ['1) Menu 📘 Logbook → "Check my data before exporting" and fix anything it lists.'],
    ['2) File → Download → Microsoft Excel (.xlsx).'],
    ['3) Give the .xlsx to the person who runs the import (see FITTER-SETUP.md).'],
    ['Photos and PDFs are added afterwards inside the app.'],
  ];
  sh.getRange(1, 1, lines.length, 1).setValues(lines).setWrap(true).setVerticalAlignment('top');
  sh.setColumnWidth(1, 760);
  [1, 3, 6, 12, 18].forEach(function (n) { sh.getRange(n, 1).setFontWeight('bold'); });
  sh.getRange(1, 1).setFontSize(14);
}

function buildLearningOutcomes_(sh) {
  clear_(sh);
  header_(sh, ['No', 'Learning outcome', 'Description (optional)', 'Year (info only)'], [50, 520, 420, 90]);
  var rows = PRESET.los.map(function (l) { return [String(l.number), l.title, l.description || '', l.year ? (l.year === 1 ? 'First year' : 'Second year') : '']; });
  sh.getRange(2, 1, Math.max(rows.length, 60), 1).setNumberFormat('@');
  sh.getRange(2, 1, rows.length, 4).setValues(rows);
  sh.getRange(2, 1, Math.max(rows.length, 60), 4).setWrap(true).setVerticalAlignment('top');
}

function buildJobs_(sh) {
  clear_(sh);
  header_(sh, ['Job No', 'LO No (auto)', 'Job title (practical)', 'Notes (optional)', 'Hours (optional)'], [70, 90, 420, 360, 90]);
  asText_(sh, 1);
  fillFormula_(sh, 2, '=IF(A{r}="","",LEFT(A{r},FIND(".",A{r}&".")-1))');
  sh.getRange(2, 3, MAX_ROWS, 3).setWrap(true).setVerticalAlignment('top');
  dupRule_(sh, 1, 'A');
}

function buildLessons_(sh) {
  clear_(sh);
  header_(sh, ['Lesson No', 'LO No (auto)', 'Lesson title (theory)', 'Hours', 'Notes'], [80, 90, 420, 70, 520]);
  asText_(sh, 1);
  fillFormula_(sh, 2, '=IF(A{r}="","",LEFT(A{r},FIND(".",A{r}&".")-1))');
  sh.getRange(2, 3, MAX_ROWS, 3).setWrap(true).setVerticalAlignment('top');
  dupRule_(sh, 1, 'A');
}

function buildTasks_(sh) {
  clear_(sh);
  header_(sh, ['Task No', 'Job No (auto)', 'Job title (auto)', 'Task name', 'Aim', 'Procedure', 'Result', 'Conclusion'], [80, 90, 260, 280, 260, 380, 220, 220]);
  asText_(sh, 1);
  fillFormula_(sh, 2, '=IF(A{r}="","",LEFT(A{r},FIND(".",A{r}&".",FIND(".",A{r}&".")+1)-1))');
  var t = [];
  for (var i = 0; i < MAX_ROWS; i++) t.push(['=IF(B' + (i + 2) + '="","",IFERROR(VLOOKUP(B' + (i + 2) + ',Jobs!$A$2:$C$' + (MAX_ROWS + 1) + ',3,FALSE),"job not found"))']);
  sh.getRange(2, 3, MAX_ROWS, 1).setFormulas(t).setFontColor('#6b7280').setBackground('#f3f4f6');
  sh.getRange(2, 4, MAX_ROWS, 5).setWrap(true).setVerticalAlignment('top');
  dupRule_(sh, 1, 'A');
}

// ── checker ──────────────────────────────────────────────────────────────────
function checkLogbookData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rep = checkData_(function (name) { var sh = ss.getSheetByName(name); return sh ? sh.getDataRange().getDisplayValues() : []; });
  var msg = rep.summary + (rep.problems.length ? '\n\nPlease fix:\n• ' + rep.problems.slice(0, 40).join('\n• ') + (rep.problems.length > 40 ? '\n… and ' + (rep.problems.length - 40) + ' more' : '') : '\n\nNo problems found. You can download the .xlsx now.');
  SpreadsheetApp.getUi().alert(msg);
}

// pure function (also used by the automated test): rows = function(sheetName) → array of rows incl. header
function checkData_(rows) {
  var problems = [], los = {}, jobs = {}, lessons = {}, tasks = {};
  var t = function (v) { return v === null || v === undefined ? '' : String(v).trim(); };
  rows('LearningOutcomes').slice(1).forEach(function (r) { if (t(r[0])) los[t(r[0])] = t(r[1]); });
  var loOf = function (id) { return id.split('.')[0]; };
  var scan = function (sheet, label, store, titleCol, levels) {
    rows(sheet).slice(1).forEach(function (r, i) {
      var id = t(r[0]); if (!id) return;
      var where = label + ' ' + id + ' (' + sheet + ' row ' + (i + 2) + ')';
      if (!/^\d+(\.\d+)*$/.test(id)) { problems.push(where + ': the number should look like 3.2' + (levels === 3 ? '.1' : '')); return; }
      if (id.split('.').length !== levels) problems.push(where + ': the number should have ' + levels + ' parts' + (levels === 2 ? ' (e.g. 3.2)' : ' (e.g. 3.2.1)'));
      if (store[id]) problems.push(where + ': this number is used more than once');
      store[id] = true;
      if (!los[loOf(id)]) problems.push(where + ': learning outcome ' + loOf(id) + ' does not exist in the LearningOutcomes tab');
      if (!t(r[titleCol])) problems.push(where + ': the ' + (label === 'Task' ? 'task name' : 'title') + ' is empty');
    });
  };
  scan('Jobs', 'Job', jobs, 2, 2);
  scan('Lessons', 'Lesson', lessons, 2, 2);
  scan('Tasks', 'Task', tasks, 3, 3);
  Object.keys(tasks).forEach(function (id) { var j = id.split('.').slice(0, 2).join('.'); if (!jobs[j]) problems.push('Task ' + id + ': job ' + j + ' is not in the Jobs tab'); });
  var n = function (o) { return Object.keys(o).length; };
  return { problems: problems, summary: 'Found ' + n(los) + ' learning outcomes, ' + n(jobs) + ' jobs, ' + n(lessons) + ' lessons and ' + n(tasks) + ' tasks.' };
}
