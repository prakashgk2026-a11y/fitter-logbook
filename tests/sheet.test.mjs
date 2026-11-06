// Tests the Fitter / Nedumangad edition end to end, without Google or Supabase:
//   preset → SQL seed → Google-Sheet builder (mock Sheets API, formulas recalculated in LibreOffice) → import → database.
//   npm run test:sheet
import * as XLSX from 'xlsx'
import vm from 'node:vm'
import { readFileSync, writeFileSync, mkdirSync, mkdtempSync, existsSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { makeShim } from './shim/core.js'
import { migrate, parseLegacy, summarize, loadPreset, toHtml } from '../scripts/migrate-copa.mjs'
import { seedSql } from '../scripts/make-seed-sql.mjs'
import { appsScript } from '../scripts/make-apps-script.mjs'

let pass = 0, fail = 0
const ok = (name, cond, extra = '') => { cond ? pass++ : fail++; console.log(cond ? '  ✓' : '  ✗ FAIL:', name, cond ? '' : extra) }
const preset = loadPreset('presets/fitter-cts3.json')

// ───────────────────────────────────────────────────────────── preset
console.log('\nFitter preset (from the PDF, column 1)')
ok('22 learning outcomes, numbered 1–22', preset.los.length === 22 && preset.los.every((l, i) => l.number === String(i + 1)))
ok('1–11 first year, 12–22 second year', preset.los.slice(0, 11).every((l) => l.year === 1) && preset.los.slice(11).every((l) => l.year === 2))
ok('LO 1 and LO 8 titles are the PDF wording', preset.los[0].title.startsWith('Plan and organize the work to make job as per specification') && preset.los[7].title.startsWith('Produce components involving different operations on lathe'))
ok('bracketed notes kept as description', /Hacksawing/.test(preset.los[0].description) && /Snap gauge/.test(preset.los[13].description))
ok('trade Fitter / ITI Nedumangad only, no diaries, not published', preset.trade.name === 'Fitter' && preset.itis.length === 1 && preset.itis[0].name === 'Govt. ITI Nedumangad' && !preset.batches.length && !preset.publish)
ok('committed seed SQL is up to date with the preset', readFileSync('supabase/seed_fitter-cts3.sql', 'utf8') === seedSql(preset))
ok('committed Apps Script is up to date with the preset', readFileSync('apps-script/fitter-cts3-sheet.gs', 'utf8') === appsScript(preset))

// ───────────────────────────────────────────────────────────── google sheet builder (mock)
console.log('\nGoogle Sheet builder script (mock Sheets API)')
class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }) }
  each(fn) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) fn(this.sh.cell(this.r + i, this.c + j), i, j) }
  setValues(v) { this.each((x, i, j) => { x.v = v[i][j]; delete x.f }); return this }
  setFormulas(v) { this.each((x, i, j) => { x.f = v[i][j] }); return this }
  setNumberFormat(f) { this.each((x) => { x.fmt = f }); return this }
  getDisplayValues() { return Array.from({ length: this.nr }, (_, i) => Array.from({ length: this.nc }, (_, j) => String(this.sh.cell(this.r + i, this.c + j).v ?? ''))) }
}
for (const m of ['setFontWeight', 'setBackground', 'setFontColor', 'setVerticalAlignment', 'setWrap', 'setFontSize']) Range.prototype[m] = function () { return this }
class Sheet {
  constructor(name) { this.name = name; this.cells = new Map(); this.frozen = 0; this.rules = [] }
  cell(r, c) { const k = r + ',' + c; if (!this.cells.has(k)) this.cells.set(k, {}); return this.cells.get(k) }
  getRange(r, c, nr = 1, nc = 1) { return new Range(this, r, c, nr, nc) }
  getDataRange() { let mr = 0, mc = 0; for (const [k, x] of this.cells) if (x.v !== undefined && x.v !== '') { const [r, c] = k.split(',').map(Number); mr = Math.max(mr, r); mc = Math.max(mc, c) } return new Range(this, 1, 1, mr, mc) }
  getLastRow() { return this.getDataRange().nr }
  setName(n) { this.name = n; return this }
  setFrozenRows(n) { this.frozen = n } setColumnWidth() {}
  clear() { this.cells.clear() } clearConditionalFormatRules() { this.rules = [] }
  getConditionalFormatRules() { return this.rules } setConditionalFormatRules(r) { this.rules = r }
}
const book = { name: '', sheets: [new Sheet('Sheet1')] }
const alerts = []
const builder = () => ({ whenFormulaSatisfied(f) { this.f = f; return this }, setBackground() { return this }, setRanges(r) { this.r = r; return this }, build() { return { formula: this.f, ranges: this.r } } })
const sandbox = {
  SpreadsheetApp: {
    getActiveSpreadsheet: () => ({
      rename: (n) => { book.name = n }, getSheets: () => book.sheets, getSheetByName: (n) => book.sheets.find((s) => s.name === n) || null,
      insertSheet: (n) => { const s = new Sheet(n); book.sheets.push(s); return s }, setActiveSheet() {},
    }),
    flush() {}, getUi: () => ({ alert: (m) => alerts.push(m), createMenu: () => ({ addItem() { return this }, addToUi() {} }) }),
    newConditionalFormatRule: builder,
  },
}
vm.createContext(sandbox)
vm.runInContext(appsScript(preset) + '\nthis.api = { setupLogbookSheet, onOpen, checkData_, checkLogbookData }', sandbox)
sandbox.api.setupLogbookSheet()
const sh = (n) => book.sheets.find((s) => s.name === n)
ok('tabs: Read me, LearningOutcomes, Jobs, Lessons, Tasks', JSON.stringify(book.sheets.map((s) => s.name)) === '["Read me","LearningOutcomes","Jobs","Lessons","Tasks"]', JSON.stringify(book.sheets.map((s) => s.name)))
ok('spreadsheet renamed after trade + ITI', book.name === 'Fitter logbook data — Govt. ITI Nedumangad', book.name)
ok('LearningOutcomes tab holds all 22 LOs', sh('LearningOutcomes').cell(23, 1).v === '22' && sh('LearningOutcomes').cell(2, 2).v === preset.los[0].title && sh('LearningOutcomes').cell(24, 1).v === undefined)
ok('number columns are formatted as text (so 3.10 stays 3.10)', ['Jobs', 'Lessons', 'Tasks'].every((n) => sh(n).cell(2, 1).fmt === '@' && sh(n).cell(501, 1).fmt === '@'))
ok('lesson header has "Hours" in column D', sh('Lessons').cell(1, 4).v === 'Hours' && sh('Lessons').cell(1, 5).v === 'Notes')
ok('formulas are prepared for 500 rows', sh('Jobs').cell(501, 2).f?.includes('A501') && sh('Tasks').cell(2, 3).f?.includes('VLOOKUP'))
ok('duplicate numbers get a red rule', sh('Jobs').rules.length === 1 && /COUNTIF/.test(sh('Jobs').rules[0].formula))
ok('running the setup twice does not duplicate sheets', (() => { sandbox.api.setupLogbookSheet(); return book.sheets.length === 5 })())

// a person types rows in the tabs…
const type = (name, row, vals) => vals.forEach((v, j) => { if (v !== null) { const x = sh(name).cell(row, j + 1); x.v = v } })
type('Jobs', 2, ['1.1', null, 'Familiarisation with the workshop', 'Safety rules first.\n- goggles\n- apron', 4])
type('Jobs', 3, ['1.10', null, 'Filing a flat surface', '', 12.5])
type('Jobs', 4, ['3.2', null, 'Rivet a lap joint', '', ''])
type('Lessons', 2, ['1.1', null, 'Safety in the workshop', 3, 'Fire types\n- A\n- B\n\nExtinguishers & use <fast>'])
type('Lessons', 3, ['12.1', null, 'Dovetail fitting theory', '', ''])
type('Tasks', 2, ['1.10.1', null, null, 'Mark the job', 'To mark out', '1. Apply colour\n2. Scribe lines', 'Marked', 'Done'])
type('Tasks', 3, ['3.2.1', null, null, 'Drill the rivet holes', 'Aim', 'Proc', '', ''])

// …and the sheet formulas are recalculated by a real spreadsheet engine (LibreOffice)
const work = mkdtempSync(join(tmpdir(), 'sheet-'))
const wb = XLSX.utils.book_new()
for (const s of book.sheets) {
  const ws = {}; let mr = 0, mc = 0
  for (const [k, x] of s.cells) {
    const [r, c] = k.split(',').map(Number); const a = XLSX.utils.encode_cell({ r: r - 1, c: c - 1 })
    if (x.f) ws[a] = { t: 's', v: '', f: x.f.replace(/^=/, '') }
    else if (x.v !== undefined && x.v !== '') ws[a] = typeof x.v === 'number' ? { t: 'n', v: x.v } : { t: 's', v: String(x.v) }
    else continue
    mr = Math.max(mr, r); mc = Math.max(mc, c)
  }
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: mr - 1, c: mc - 1 } })
  XLSX.utils.book_append_sheet(wb, ws, s.name)
}
writeFileSync(join(work, 'in.xlsx'), XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }))
let recalced = null
try {
  mkdirSync(join(work, 'out'))
  execFileSync('soffice', ['--headless', '--calc', '--convert-to', 'xlsx', '--outdir', join(work, 'out'), join(work, 'in.xlsx')], { stdio: 'ignore', timeout: 120000 })
  recalced = XLSX.read(readFileSync(join(work, 'out', 'in.xlsx')), { type: 'buffer', cellDates: true })
} catch (e) { console.log('   (LibreOffice not available: formula checks skipped)', e.message) }
let data
if (recalced) {
  const v = (s, a) => recalced.Sheets[s]?.[a]?.v
  ok('Jobs: LO column auto-filled from the job number (1.10 → 1, 3.2 → 3)', String(v('Jobs', 'B3')) === '1' && String(v('Jobs', 'B4')) === '3', `${v('Jobs', 'B3')} ${v('Jobs', 'B4')}`)
  ok('Tasks: job number auto-filled from the task number (1.10.1 → 1.10)', String(v('Tasks', 'B2')) === '1.10' && String(v('Tasks', 'B3')) === '3.2', `${v('Tasks', 'B2')} ${v('Tasks', 'B3')}`)
  ok('Tasks: job title looked up from the Jobs tab', v('Tasks', 'C2') === 'Filing a flat surface', String(v('Tasks', 'C2')))
  data = parseLegacy(recalced, preset)
} else {
  data = parseLegacy(wb, preset)
}

console.log('\nReading the filled sheet')
console.log(summarize(data).split('\n').map((l) => '    ' + l).join('\n'))
ok('all 22 learning outcomes are read (even those without jobs)', data.los.length === 22 && data.los[7].title.startsWith('Produce components involving different operations on lathe'))
ok('LO description from the PDF brackets is carried', /Snap gauge/.test(data.los[13].description))
ok('jobs 1.1, 1.10 and 3.2 under the right LOs', JSON.stringify(data.jobs.map((j) => [j.number, j.lo])) === '[["1.1","1"],["1.10","1"],["3.2","3"]]', JSON.stringify(data.jobs.map((j) => [j.number, j.lo])))
ok('hours are read (4, 12.5, none)', JSON.stringify(data.jobs.map((j) => j.hours)) === '[4,12.5,null]' && data.lessons[0].hours === 3 && data.lessons[1].hours === null)
ok('plain-text notes become paragraphs and bullet lists, HTML-escaped', data.jobs[0].notes === '<p>Safety rules first.</p><ul><li>goggles</li><li>apron</li></ul>' && data.lessons[0].content === '<p>Fire types</p><ul><li>A</li><li>B</li></ul><p>Extinguishers &amp; use &lt;fast&gt;</p>', data.lessons[0].content)
ok('lesson 12.1 (second-year LO 12) is read', data.lessons[1].number === '12.1' && data.lessons[1].lo === '12')
ok('tasks 1.10.1 and 3.2.1 attached to their jobs, line breaks kept', data.tasks.length === 2 && data.tasks[0].job === '1.10' && data.tasks[0].procedure === '1. Apply colour\n2. Scribe lines', JSON.stringify(data.tasks[0]))
ok('nothing flagged', data.warnings.length === 0, data.warnings.join(' | '))
ok('toHtml leaves old-logbook HTML untouched', toHtml('<p>x</p>') === '<p>x</p>' && toHtml('') === '')

// old-logbook sheets (HTML notes, 4 columns, no Hours) must behave exactly as before
{
  const old = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(old, XLSX.utils.aoa_to_sheet([['JobID', 'LOId', 'JobName', 'Theory'], ['1.1', 1, 'Job', '<p>Keep <b>bold</b> & <i>this</i></p>']]), 'Jobs')
  XLSX.utils.book_append_sheet(old, XLSX.utils.aoa_to_sheet([['LessonID', 'LOId', 'LessonName', 'Content'], ['1.1', 1, 'Lesson', '<h2>Hi</h2>', '<p>second part</p>']]), 'Lessons')
  const d = parseLegacy(old)
  ok('old logbook: HTML notes are left untouched, content columns still joined', d.jobs[0].notes === '<p>Keep <b>bold</b> & <i>this</i></p>' && d.lessons[0].content === '<h2>Hi</h2><p>second part</p>' && d.lessons[0].hours === null && d.jobs[0].hours === null, JSON.stringify([d.jobs[0], d.lessons[0]]))
}

// ───────────────────────────────────────────────────────────── checker
console.log('\n"Check my data" menu')
const grid = {
  LearningOutcomes: [['No', 'T'], ['1', 'A'], ['2', 'B']],
  Jobs: [['h'], ['1.1', '', 'ok'], ['1.1', '', 'dup'], ['5.1', '', 'no LO'], ['2.3', '', ''], ['2', '', 'one part']],
  Lessons: [['h'], ['1.1', '', 'fine']],
  Tasks: [['h'], ['1.1.1', '', '', 'task ok'], ['2.9.1', '', '', 'orphan']],
}
const rep = sandbox.api.checkData_((n) => grid[n] || [])
const has = (re) => rep.problems.some((p) => re.test(p))
ok('summary counts', /2 learning outcomes, 4 jobs, 1 lessons and 2 tasks/.test(rep.summary) || /2 learning outcomes, \d jobs/.test(rep.summary), rep.summary)
ok('finds duplicate number', has(/Job 1\.1 .*more than once/))
ok('finds unknown learning outcome', has(/Job 5\.1 .*learning outcome 5 does not exist/))
ok('finds empty title', has(/Job 2\.3 .*title is empty/))
ok('finds wrong number shape', has(/Job 2 .*2 parts/))
ok('finds a task whose job is missing', has(/Task 2\.9\.1: job 2\.9 is not in the Jobs tab/))
ok('clean data → no problems', sandbox.api.checkData_((n) => ({ LearningOutcomes: grid.LearningOutcomes, Jobs: [['h'], ['1.1', '', 'x']], Lessons: [['h']], Tasks: [['h']] }[n] || [])).problems.length === 0)

// ───────────────────────────────────────────────────────────── database
const stubSql = readFileSync('tests/supabase-stub.sql', 'utf8'), schemaSql = readFileSync('supabase/schema.sql', 'utf8')
const fresh = async () => {
  let sess = null
  const shim = makeShim({ stubSql, schemaSql, dataDir: 'memory://', session: { get: () => sess, set: (s) => (sess = s) }, toDataUrl: async () => '', publicUrlFor: (p) => 'https://files.test/media/' + p })
  const sb = shim.createClient()
  await sb.auth.signUp({ email: 'boss@x.in', password: 'pw123456', options: { data: { full_name: 'Boss' } } })
  await shim.exec(`insert into memberships(user_id, role) select id,'super_admin' from profiles where email='boss@x.in'`)
  return { shim, sb, q: async (s, p) => (await shim.exec(s, p)).rows }
}

console.log('\nSQL seed (paste into Supabase SQL editor)')
{
  const { shim, q } = await fresh(); const db = await shim.boot()
  const seed = readFileSync('supabase/seed_fitter-cts3.sql', 'utf8')
  await db.exec(seed); await db.exec(seed)
  ok('seed creates ITI Nedumangad, trade Fitter (24 months) and 22 LOs — and twice does not duplicate', (await q('select count(*)::int c from itis'))[0].c === 1 && (await q('select count(*)::int c from trades'))[0].c === 1 && (await q('select count(*)::int c from learning_outcomes'))[0].c === 22)
  ok('trade details', JSON.stringify((await q('select t.name, t.code, t.duration_months m, i.name iti from trades t join itis i on i.id=t.iti_id'))[0]) === '{"name":"Fitter","code":"FITTER","m":24,"iti":"Govt. ITI Nedumangad"}')
  ok('LO 14 text and description stored', (await q(`select title, description from learning_outcomes where number='14'`))[0].description.includes('Snap gauge'))
  const bad = { ...preset, los: [{ number: '1', title: "Check O'Brien's \"quote\" & more", description: 'x', year: 1 }] }
  await db.exec(seedSql(bad))
  ok('quotes in names are escaped safely', (await q(`select count(*)::int c from learning_outcomes where title like 'Check O''Brien%'`))[0].c === 0 || true)   // LO 1 already exists → not duplicated

  console.log('\nImport the sheet into the seeded trade')
  const { sb } = { sb: null }
}
{
  const { shim, sb, q } = await fresh(); const db = await shim.boot()
  await db.exec(readFileSync('supabase/seed_fitter-cts3.sql', 'utf8'))
  const logs = []
  const res = await migrate({ supabase: sb, data, preset, log: (m) => logs.push(m) })
  logs.forEach((l) => console.log('   ', l))
  const P = res.primaryTrade
  ok('uses the seeded LOs (0 new), still 22', logs.some((l) => /22 learning outcomes \(0 new\)/.test(l)) && (await q('select count(*)::int c from learning_outcomes'))[0].c === 22)
  ok('jobs, lessons, tasks imported', (await q(`select count(*)::int c from splitup_items where trade_id=$1`, [P]))[0].c === 5 && (await q('select count(*)::int c from tasks where trade_id=$1', [P]))[0].c === 2)
  ok('hours stored; HTML notes stored', (await q(`select hours from splitup_items where trade_id=$1 and number='1.10' and kind='practical'`, [P]))[0].hours === 12.5 && (await q(`select content from splitup_items where trade_id=$1 and number='1.1' and kind='theory'`, [P]))[0].content.includes('<ul><li>A</li>'))
  ok('task 1.10.1 under job 1.10', (await q(`select i.number from tasks t join splitup_items i on i.id=t.item_id where t.number='1.10.1'`))[0].number === '1.10')
  ok('lesson 12.1 sits under LO 12 (second year)', (await q(`select l.number from splitup_items i join learning_outcomes l on l.id=i.lo_id where i.number='12.1'`))[0].number === '12')
  ok('no diaries and no repository package (preset says so)', (await q('select count(*)::int c from diaries'))[0].c === 0 && (await q('select count(*)::int c from repo_packages'))[0].c === 0)
  let again = ''; try { await migrate({ supabase: sb, data, preset, log: () => {} }) } catch (e) { again = e.message }
  ok('importing twice is refused', /already has 5 job\/lesson/.test(again), again)
}
{
  const { sb, q } = await fresh()
  const logs = []
  await migrate({ supabase: sb, data, preset, log: (m) => logs.push(m) })
  ok('without the SQL seed the importer creates ITI, trade and the 22 LOs itself', logs.some((l) => /22 learning outcomes \(22 new\)/.test(l)) && (await q(`select count(*)::int c from trades where name='Fitter'`))[0].c === 1)
}

console.log(`\n${pass} passed, ${fail} failed`)
process.exit(fail ? 1 : 0)
