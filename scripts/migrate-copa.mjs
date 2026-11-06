#!/usr/bin/env node
// Moves the old Google-Sheets COPA logbook into the new database.
//
//   1. In the old Google Sheet:  File → Download → Microsoft Excel (.xlsx)
//   2. Run  (once, from this folder):
//        SUPABASE_URL=… SUPABASE_ANON_KEY=… ADMIN_EMAIL=… ADMIN_PASSWORD=… \
//        node scripts/migrate-copa.mjs --file old-logbook.xlsx --dry-run     ← look first
//        … same without --dry-run                                            ← then really do it
//
// The ADMIN account must already be a Super Admin (see supabase/bootstrap.sql). The script signs in
// with that e-mail/password like a normal user — no secret "service" key is needed or used.
//
// What it does: ITIs + trade → learning outcomes → jobs (practical) + tasks → lessons (theory)
// → photos/files (to Storage) → diaries (batches, per ITI, with class start dates) →
// optionally publishes the syllabus to the shared repository and gives the other ITIs their own copy.
//
// Which ITIs / trade / learning-outcome names to use comes from a PRESET file in /presets:
//   presets/copa.json (default, the original logbook)      presets/fitter-cts3.json (Fitter, ITI Nedumangad)
//   node scripts/migrate-copa.mjs --file data.xlsx --preset presets/fitter-cts3.json
// The same script reads both the old COPA logbook and the new Google-Sheet template (apps-script/).
import { readFileSync } from 'node:fs'
import { createHash, randomUUID } from 'node:crypto'
import { pathToFileURL } from 'node:url'
import * as XLSX from 'xlsx'

// ───────────────────────────────────────────────────────── presets (ITIs, trade, LO names)
export const loadPreset = (file) => JSON.parse(readFileSync(file, 'utf8'))
export const COPA = loadPreset(new URL('../presets/copa.json', import.meta.url))
const DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']

// ───────────────────────────────────────────────────────── reading the workbook
const blank = (v) => v === null || v === undefined || String(v).trim() === ''
const str = (v) => (blank(v) ? '' : v instanceof Date ? ymdLocal(v) : String(v))
const pad = (n) => String(n).padStart(2, '0')
const ymdLocal = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`
const isJunk = (v) => typeof v === 'string' && /^#(NUM|REF|N\/A|VALUE|NAME|DIV)/i.test(v.trim())

// Sheets once turned "1.10" into the number 1.1, so two different rows can look identical.
// A *number* that repeats an earlier ID is really the next-ten ("1.1" again after "1.9" → "1.10").
function makeIdResolver(label, warnings) {
  const seen = new Set()
  const aliases = {}                 // text exactly as seen → first ID it resolved to
  const resolve = (raw) => {
    if (blank(raw) || isJunk(raw)) return null
    const isNum = typeof raw === 'number'
    let id = isNum ? String(raw) : String(raw).trim()
    let note = ''
    if (seen.has(id)) {
      const alt = isNum ? (id.includes('.') ? id + '0' : id + '.0') : null
      if (alt && !seen.has(alt)) { note = `${id} → ${alt}`; id = alt }
      else return { id: alt && seen.has(alt) ? alt : id, dup: true }      // a further copy of a row already read
    }
    seen.add(id)
    aliases[isNum ? String(raw) : String(raw).trim()] ??= id
    return { id, note }
  }
  return { resolve, aliases }
}
const hoursOf = (v) => { const n = typeof v === 'number' ? v : parseFloat(String(v ?? '').replace(',', '.')); return Number.isFinite(n) && n >= 0 && n < 10000 ? n : null }
const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
// Notes typed as plain text in a Google Sheet cell → paragraphs and bullet lists (HTML from the old logbook is left alone)
export function toHtml(text, plain = false) {
  const t = String(text ?? ''); if (!t.trim() || (!plain && /<\/?[a-z][^>]*>/i.test(t))) return t
  const out = []; let list = null
  for (const raw of t.split(/\r?\n/)) {
    const line = raw.trim(), b = /^[-•*]\s+(.*)$/.exec(line)
    if (b) { (list ||= []).push(b[1]); continue }
    if (list) { out.push('<ul>' + list.map((x) => `<li>${esc(x)}</li>`).join('') + '</ul>'); list = null }
    if (line) out.push(`<p>${esc(line)}</p>`)
  }
  if (list) out.push('<ul>' + list.map((x) => `<li>${esc(x)}</li>`).join('') + '</ul>')
  return out.join('')
}
const loOf = (id) => String(id).split('.')[0]
const cleanName = (n) => String(n || '').replace(/^\s*\d+(\.\d+)*\s*[:\-–.]\s*/, '').trim()

function isoWeekMonday(weekKey) {
  const m = /^(\d{4})-W(\d{1,2})$/.exec(String(weekKey).trim()); if (!m) return null
  const y = +m[1], w = +m[2]
  const jan4 = new Date(Date.UTC(y, 0, 4))
  const mon = new Date(jan4); mon.setUTCDate(jan4.getUTCDate() - ((jan4.getUTCDay() || 7) - 1) + (w - 1) * 7)
  return `${mon.getUTCFullYear()}-${pad(mon.getUTCMonth() + 1)}-${pad(mon.getUTCDate())}`
}
const addDays = (ymd, n) => { const [y, m, d] = ymd.split('-').map(Number); const t = new Date(y, m - 1, d + n); return ymdLocal(t) }

export function parseLegacy(wb, preset = COPA) {
  const warnings = [], fixes = []
  const sheet = (name) => {
    const ws = wb.Sheets[name]
    return ws ? XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }).slice(1) : []
  }

  // The new Google-Sheet template (apps-script/) has an "Hours" column and holds PLAIN text notes; the old logbook held HTML.
  const headerOf = (n) => (wb.Sheets[n] ? XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null })[0] || [] : [])
  const lessonHours = /^hours?$/i.test(String(headerOf('Lessons')[3] ?? '').trim())
  const lessonCol = lessonHours ? 4 : 3
  const plainText = lessonHours || /^hours?/i.test(String(headerOf('Jobs')[4] ?? '').trim())

  // jobs
  const J = makeIdResolver('Job', warnings)
  const jobs = []
  const mergedCopies = []
  for (const r of sheet('Jobs')) {
    const plain = typeof r[0] === 'number' ? String(r[0]) : str(r[0]).trim()
    const twin = jobs.find((x) => x.number === plain && x.title.toLowerCase() === str(r[2]).trim().toLowerCase())   // same number AND same name = a copy
    const res = twin ? { id: plain, dup: true } : J.resolve(r[0]); if (!res) continue
    if (res.dup) {
      const ex = jobs.find((x) => x.number === res.id)
      if (ex && ex.title.toLowerCase() === str(r[2]).trim().toLowerCase()) { if (toHtml(str(r[3]), plainText).length > ex.notes.length) ex.notes = toHtml(str(r[3]), plainText); mergedCopies.push('Job ' + res.id) }
      else warnings.push(`Job ${res.id} appears twice with different names ("${ex?.title}" and "${str(r[2]).trim()}") — the second one was skipped`)
      continue
    }
    if (res.note) fixes.push('Job ' + res.note)
    jobs.push({ number: res.id, lo: loOf(res.id), title: str(r[2]).trim(), notes: toHtml(str(r[3]), plainText), hours: hoursOf(r[4]) })
    if (!blank(r[1]) && Math.round(Number(r[1])) !== Number(loOf(res.id))) warnings.push(`Job ${res.id}: LO column says ${r[1]} — placed under LO ${loOf(res.id)} by its number`)
  }
  const jobByNo = Object.fromEntries(jobs.map((j) => [j.number, j]))

  // tasks (job number comes from the task number: 1.10.2 → job 1.10)
  const tasks = []
  for (const r of sheet('Tasks')) {
    if (blank(r[0]) || isJunk(r[0])) continue
    const id = typeof r[0] === 'number' ? String(r[0]) : String(r[0]).trim()
    const parts = id.split('.')
    const jobNo = parts.length >= 3 ? parts.slice(0, 2).join('.') : J.aliases[str(r[1])] || str(r[1])
    if (!jobByNo[jobNo]) { warnings.push(`Task ${id}: job ${jobNo} not found — task skipped`); continue }
    if (tasks.some((t) => t.number === id)) { warnings.push(`Task ${id} appears twice — the second one was skipped`); continue }
    tasks.push({ number: id, job: jobNo, name: str(r[3]).trim() || `Task ${id}`, aim: str(r[4]), procedure: str(r[5]), result: str(r[6]), conclusion: str(r[7]) })
  }

  // lessons (long content was split over columns D, E, F … → join them)
  const L = makeIdResolver('Lesson', warnings)
  const lessons = []
  let strippedPrefix = 0
  for (const r of sheet('Lessons')) {
    const plain = typeof r[0] === 'number' ? String(r[0]) : str(r[0]).trim()
    const twin = lessons.find((x) => x.number === plain && x.title.toLowerCase() === (cleanName(str(r[2])) || str(r[2]).trim()).toLowerCase())
    const res = twin ? { id: plain, dup: true } : L.resolve(r[0]); if (!res) continue
    if (res.dup) {
      const ex = lessons.find((x) => x.number === res.id), content = toHtml(r.slice(lessonCol).map((c) => (c == null ? '' : String(c))).join(''), plainText)
      if (ex && ex.title.toLowerCase() === (cleanName(str(r[2])) || str(r[2]).trim()).toLowerCase()) { if (content.length > ex.content.length) ex.content = content; mergedCopies.push('Lesson ' + res.id) }
      else warnings.push(`Lesson ${res.id} appears twice with different names ("${ex?.title}" and "${str(r[2]).trim()}") — the second one was skipped`)
      continue
    }
    if (res.note) fixes.push('Lesson ' + res.note)
    const nm = str(r[2]).trim(); const cl = cleanName(nm); if (cl !== nm) strippedPrefix++
    lessons.push({ number: res.id, lo: loOf(res.id), title: cl || nm || `Lesson ${res.id}`, content: toHtml(r.slice(lessonCol).map((c) => (c == null ? '' : String(c))).join(''), plainText), hours: lessonHours ? hoursOf(r[3]) : null })
  }
  if (mergedCopies.length) {
    const c = {}; mergedCopies.forEach((k) => (c[k] = (c[k] || 1) + 1))
    fixes.push('Repeated copies of the same row were merged (the copy with the most text was kept): ' + Object.entries(c).map(([k, n]) => `${k} ×${n}`).join(', '))
  }
  if (strippedPrefix) fixes.push(`${strippedPrefix} lesson name(s) had the number typed in front (e.g. "1.9: Memory") — number removed from the name`)
  const lessonByNo = Object.fromEntries(lessons.map((l) => [l.number, l]))

  // legacy LO-level "Theory" sheet: not used any more in the old app's menus
  const oldTheory = sheet('Theory').filter((r) => !blank(r[0]) && !blank(r[1])).length
  if (oldTheory) warnings.push(`${oldTheory} row(s) in the old "Theory" sheet (per-LO notes) are not migrated — the lessons replace them`)
  const jobNotes = jobs.filter((j) => !blank(j.notes)).length
  if (jobNotes) fixes.push(`${jobNotes} job(s) had text in the old "Theory" column — kept as job notes`)

  // media
  const media = []
  for (const r of sheet('Media')) {
    const key = str(r[0]).trim(); if (!key || blank(r[4])) continue
    let kind, owner
    if (key.startsWith('job-')) { kind = 'job'; owner = J.aliases[key.slice(4)] || key.slice(4); if (!jobByNo[owner]) owner = null }
    else if (key.startsWith('lesson-')) { kind = 'lesson'; owner = L.aliases[key.slice(7)] || key.slice(7); if (!lessonByNo[owner]) owner = null }
    else if (key.startsWith('theory-')) { continue }
    else { kind = 'task'; owner = tasks.some((t) => t.number === key) ? key : null }
    if (!owner) { warnings.push(`File "${str(r[2])}" belongs to "${key}", which no longer exists — skipped`); continue }
    let b64 = String(r[4]); const m = /^data:([^;]+);base64,/.exec(b64); if (m) b64 = b64.slice(m[0].length)
    media.push({ kind, owner, index: Number(r[1]) || 0, fileName: str(r[2]) || 'file', mime: str(r[3]) || (m ? m[1] : 'application/octet-stream'), b64, type: str(r[5]) === 'file' ? 'file' : 'photo' })
  }

  // diary + class start dates
  const settings = {}
  for (const r of sheet('Settings')) if (!blank(r[0])) settings[String(r[0])] = str(r[1])
  const diary = []
  for (const r of sheet('Diary')) {
    const monday = isoWeekMonday(str(r[0])); const idx = DAYS.indexOf(str(r[1]).trim())
    if (!monday || idx < 0) { if (!blank(r[0])) warnings.push(`Diary row "${str(r[0])} ${str(r[1])}" skipped (unreadable week or day)`); continue }
    const fixNo = (raw, title, byNo) => {                  // a number cell "1.1" may really be "1.10"
      if (typeof raw !== 'number') return str(raw).trim()
      const plain = String(raw), alt = plain.includes('.') ? plain + '0' : plain
      return title && byNo[alt]?.title === title.trim() && byNo[plain]?.title !== title.trim() ? alt : plain
    }
    const iti = str(r[7]).trim().toLowerCase() || preset.defaultIti
    const batch = (str(r[8]).trim() || preset.defaultBatch || '').toUpperCase()
    const date = /^\d{4}-\d{2}-\d{2}/.test(str(r[2])) ? str(r[2]).slice(0, 10) : addDays(monday, idx)
    const e = { iti, batch, week: monday, day: idx, date, exNo: fixNo(r[3], str(r[4]), jobByNo), practicalTitle: str(r[4]), theoryNo: fixNo(r[5], str(r[6]), lessonByNo), theoryTitle: str(r[6]) }
    if (e.exNo || e.practicalTitle || e.theoryNo || e.theoryTitle) diary.push(e)
  }

  // learning outcomes: names come from the sheet's "LearningOutcomes" tab (if any), else from the preset
  const presetLo = Object.fromEntries((preset.los || []).map((l) => [String(l.number), l]))
  const sheetLo = {}
  for (const r of sheet('LearningOutcomes')) {
    const n = blank(r[0]) ? '' : String(r[0]).trim(); if (!n || !/^\d+$/.test(n) || blank(r[1])) continue
    sheetLo[n] = { number: n, title: str(r[1]).trim(), description: str(r[2]).trim() }
  }
  const used = [...new Set([...jobs, ...lessons].map((x) => x.lo))]
  const loNumbers = [...new Set([...used, ...(preset.allLos ? [...Object.keys(presetLo), ...Object.keys(sheetLo)] : [])])].sort((a, b) => a - b)
  const los = loNumbers.map((n) => {
    const l = sheetLo[n] || presetLo[n]
    return { number: n, title: l?.title || `Learning outcome ${n}`, description: sheetLo[n]?.description || presetLo[n]?.description || '' }
  })
  for (const n of loNumbers) if (!sheetLo[n] && !presetLo[n]) warnings.push(`LO ${n} has no known name — a placeholder title was used`)
  return { los, jobs, tasks, lessons, media, diary, settings, warnings, fixes }
}

export function summarize(d) {
  const perIti = {}
  d.diary.forEach((e) => { const k = `${e.iti} / batch ${e.batch}`; perIti[k] = (perIti[k] || 0) + 1 })
  const lines = [
    `Learning outcomes : ${d.los.length}`, `Jobs (practical)  : ${d.jobs.length}`, `Tasks             : ${d.tasks.length}`,
    `Lessons (theory)  : ${d.lessons.length}  (${d.lessons.filter((l) => l.content).length} with notes)`,
    `Photos / files    : ${d.media.length}`, `Diary entries     : ${d.diary.length}${Object.keys(perIti).length ? '  → ' + Object.entries(perIti).map(([k, v]) => `${k}: ${v}`).join(', ') : ''}`,
    `Class start dates : ${Object.entries(d.settings).filter(([k]) => k.startsWith('classStart:')).map(([k, v]) => `${k.slice(11)}=${v}`).join(', ') || 'none'}`,
  ]
  if (d.fixes.length) lines.push('', 'Repairs made:', ...d.fixes.map((f) => '  • ' + f))
  if (d.warnings.length) lines.push('', 'Please check:', ...d.warnings.map((w) => '  ! ' + w))
  return lines.join('\n')
}

// ───────────────────────────────────────────────────────── writing to Supabase
const must = ({ data, error }, what) => { if (error) throw new Error(`${what}: ${error.message}`); return data }
const chunks = (a, n) => Array.from({ length: Math.ceil(a.length / n) }, (_, i) => a.slice(i * n, i * n + n))
const EXT = { 'image/jpeg': 'jpg', 'image/jpg': 'jpg', 'image/png': 'png', 'image/webp': 'webp', 'image/gif': 'gif', 'application/pdf': 'pdf' }

export async function migrate({ supabase, data, log = console.log, preset = COPA, primary = preset.defaultIti, publish = preset.publish }) {
  const TRADE = preset.trade, ITIS = preset.itis, BATCHES = preset.batches || [], DEFAULT_BATCH = preset.defaultBatch
  const sb = supabase
  const upload = async (tradeId, buf, mime, name) => {
    const ext = EXT[mime] || (/\.([a-z0-9]{1,5})$/i.exec(name || '')?.[1] || 'bin').toLowerCase()
    const path = `${tradeId}/${randomUUID()}.${ext}`
    const { error } = await sb.storage.from('media').upload(path, buf, { contentType: mime, cacheControl: '31536000' })
    if (error) throw new Error(`upload ${name}: ${error.message}`)
    return path
  }
  const urlOf = (path) => sb.storage.from('media').getPublicUrl(path).data.publicUrl

  // ITIs
  const itiRow = {}
  for (const i of ITIS) {
    let row = must(await sb.from('itis').select('*').eq('name', i.name).maybeSingle(), 'find ITI')
    if (!row) row = must(await sb.from('itis').insert({ name: i.name, short_name: i.short, print_subtitle: i.subtitle || 'Thiruvananthapuram, Kerala' }).select().single(), 'create ITI')
    itiRow[i.key] = row
  }
  if (!itiRow[primary]) throw new Error(`Unknown ITI "${primary}". This preset knows: ${ITIS.map((i) => i.key).join(', ')}`)
  const order = [primary, ...ITIS.map((i) => i.key).filter((k) => k !== primary)]

  // trade in the primary ITI must be empty
  const trades = {}
  for (const k of order) {
    let t = must(await sb.from('trades').select('*').eq('iti_id', itiRow[k].id).eq('name', TRADE.name).maybeSingle(), 'find trade')
    if (!t) t = must(await sb.from('trades').insert({ iti_id: itiRow[k].id, name: TRADE.name, code: TRADE.code, duration_months: TRADE.months }).select().single(), 'create trade')
    trades[k] = t
  }
  const P = trades[primary]
  const already = must(await sb.from('splitup_items').select('id').eq('trade_id', P.id), 'check trade')
  if (already.length) throw new Error(`The ${primary} ${TRADE.name} trade already has ${already.length} job/lesson item(s). Delete that trade in Settings → Trades (or choose another ITI) and run again.`)

  // learning outcomes: reuse the ones already there (e.g. from supabase/seed_*.sql), add the missing ones
  const haveLo = must(await sb.from('learning_outcomes').select('*').eq('trade_id', P.id), 'check learning outcomes')
  const loRows = [...haveLo]
  const missing = data.los.filter((l) => !haveLo.some((h) => h.number === l.number))
  if (missing.length) loRows.push(...must(await sb.from('learning_outcomes').insert(missing.map((l) => ({ trade_id: P.id, number: l.number, title: l.title, description: l.description || '', sort_order: Number(l.number) || 0 }))).select(), 'learning outcomes'))
  const loId = Object.fromEntries(loRows.map((r) => [r.number, r.id]))
  log(`  ✓ ${loRows.length} learning outcomes (${missing.length} new)`)

  // lesson notes: pictures pasted inside the text → Storage
  const seenImg = new Map()
  async function externalize(html) {
    const re = /(<img\b[^>]*?\ssrc=)(["'])data:image\/([a-z0-9.+-]+);base64,([A-Za-z0-9+/=\s]+?)\2/gi
    const found = [...html.matchAll(re)]; if (!found.length) return html
    let out = html
    for (const m of found) {
      const buf = Buffer.from(m[4].replace(/\s+/g, ''), 'base64'); const h = createHash('sha1').update(buf).digest('hex')
      if (!seenImg.has(h)) seenImg.set(h, urlOf(await upload(P.id, buf, `image/${m[3] === 'jpeg' ? 'jpeg' : m[3]}`, 'pasted.' + m[3])))
      out = out.replace(m[0], `${m[1]}${m[2]}${seenImg.get(h)}${m[2]}`)
    }
    return out
  }

  // jobs + lessons (split-up items)
  const itemRows = []
  for (const j of data.jobs) itemRows.push({ lo_id: loId[j.lo], kind: 'practical', number: j.number, title: j.title || `Job ${j.number}`, content: j.notes || '', ...(j.hours != null && { hours: j.hours }) })
  for (const l of data.lessons) itemRows.push({ lo_id: loId[l.lo], kind: 'theory', number: l.number, title: l.title, content: await externalize(l.content), ...(l.hours != null && { hours: l.hours }) })
  const items = []
  for (const c of chunks(itemRows, 40)) items.push(...must(await sb.from('splitup_items').insert(c).select('id,lo_id,kind,number'), 'jobs & lessons'))
  const itemId = (kind, no) => items.find((i) => i.kind === kind && i.number === no)?.id
  log(`  ✓ ${data.jobs.length} jobs, ${data.lessons.length} lessons`)

  // tasks
  const taskRows = data.tasks.map((t) => ({ item_id: itemId('practical', t.job), number: t.number, name: t.name, aim: t.aim, procedure: t.procedure, result: t.result, conclusion: t.conclusion }))
  const tasks = []
  for (const c of chunks(taskRows, 50)) tasks.push(...must(await sb.from('tasks').insert(c).select('id,number'), 'tasks'))
  log(`  ✓ ${tasks.length} tasks`)

  // photos & files
  let done = 0
  for (const m of data.media) {
    const path = await upload(P.id, Buffer.from(m.b64, 'base64'), m.mime, m.fileName)
    const owner = m.kind === 'task' ? { task_id: tasks.find((t) => t.number === m.owner)?.id } : { item_id: itemId(m.kind === 'job' ? 'practical' : 'theory', m.owner) }
    must(await sb.from('media').insert({ ...owner, path, file_name: m.fileName, mime_type: m.mime, size_bytes: Buffer.byteLength(m.b64, 'base64'), kind: m.type, sort_order: m.index }).select('id').single(), 'media row')
    if (++done % 10 === 0) log(`    … ${done}/${data.media.length} files`)
  }
  log(`  ✓ ${data.media.length} photos/files`)

  // publish the finished syllabus to the repository, and give the other ITIs their own copy
  let pkgId = null
  if (publish) {
    pkgId = must(await sb.rpc('publish_package', {
      p_trade: P.id, p_lo_ids: loRows.map((r) => r.id), p_title: preset.repoTitle || `${TRADE.name} — complete syllabus`,
      p_description: preset.repoDescription || `${TRADE.name}: ${loRows.length} learning outcomes.`, p_visibility: 'all',
    }), 'publish')
    log('  ✓ published to the shared repository')
    for (const k of order.slice(1)) {
      must(await sb.rpc('import_package', { p_package: pkgId, p_trade: trades[k].id, p_lo_numbers: null, p_conflict: 'skip' }), 'copy for ' + k)
      log(`  ✓ own copy created for ${itiRow[k].name}`)
    }
  }

  // diaries (one per batch in every ITI, like the old app)
  for (const k of BATCHES.length ? order : []) {
    const tradeId = trades[k].id
    const its = must(await sb.from('splitup_items').select('id,kind,number').eq('trade_id', tradeId), 'items for diary')
    const find = (kind, no) => its.find((i) => i.kind === kind && i.number === no)?.id ?? null
    let n = 0
    for (const [bi, b] of BATCHES.entries()) {
      const key = 'classStart:' + k + (b.id === DEFAULT_BATCH ? '' : ':' + b.id)
      const start = /^\d{4}-\d{2}-\d{2}/.test(data.settings[key] || '') ? data.settings[key].slice(0, 10) : null
      const dRow = must(await sb.from('diaries').insert({ trade_id: tradeId, name: b.name, class_start_date: start, sort_order: bi + 1 }).select().single(), 'diary')
      const rows = data.diary.filter((e) => e.iti === k && e.batch === b.id).map((e) => ({
        diary_id: dRow.id, week_start: e.week, day_idx: e.day, entry_date: e.date,
        practical_no: e.exNo, practical_title: e.practicalTitle, practical_item_id: find('practical', e.exNo),
        theory_no: e.theoryNo, theory_title: e.theoryTitle, theory_item_id: find('theory', e.theoryNo),
      }))
      for (const c of chunks(rows, 100)) must(await sb.from('diary_entries').upsert(c, { onConflict: 'diary_id,week_start,day_idx' }).select('id'), 'diary entries')
      n += rows.length
    }
    log(`  ✓ diaries for ${itiRow[k].name} (${n} entries)`)
  }
  return { primaryTrade: P.id, package: pkgId }
}

// ───────────────────────────────────────────────────────── command line
const arg = (n) => { const i = process.argv.indexOf('--' + n); return i < 0 ? null : process.argv[i + 1] && !process.argv[i + 1].startsWith('--') ? process.argv[i + 1] : true }
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const file = arg('file')
  if (!file || file === true) { console.error('Usage: node scripts/migrate-copa.mjs --file old-logbook.xlsx [--preset presets/fitter-cts3.json] [--dry-run] [--primary <iti key>]'); process.exit(2) }
  const presetArg = arg('preset')
  const preset = presetArg && presetArg !== true ? loadPreset(presetArg) : COPA
  console.log(`Preset: ${preset.title || preset.id} → ${preset.trade.name}, ${preset.itis.map((i) => i.name).join(' + ')}`)
  const wb = XLSX.read(readFileSync(file), { type: 'buffer', cellDates: true })
  const data = parseLegacy(wb, preset)
  console.log('\nWhat was found in ' + file + '\n' + summarize(data) + '\n')
  if (arg('dry-run')) { console.log('Dry run — nothing was written. Run again without --dry-run to migrate.'); process.exit(0) }
  const { SUPABASE_URL, SUPABASE_ANON_KEY, ADMIN_EMAIL, ADMIN_PASSWORD } = process.env
  if (![SUPABASE_URL, SUPABASE_ANON_KEY, ADMIN_EMAIL, ADMIN_PASSWORD].every(Boolean)) { console.error('Set SUPABASE_URL, SUPABASE_ANON_KEY, ADMIN_EMAIL and ADMIN_PASSWORD first (see the README).'); process.exit(2) }
  const { createClient } = await import('@supabase/supabase-js')
  const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, { auth: { persistSession: false } })
  const { error } = await supabase.auth.signInWithPassword({ email: ADMIN_EMAIL, password: ADMIN_PASSWORD })
  if (error) { console.error('Could not sign in: ' + error.message); process.exit(1) }
  const me = await supabase.from('memberships').select('role').eq('role', 'super_admin')
  if (!me.data?.length) { console.error('This account is not a Super Admin yet — run supabase/bootstrap.sql first.'); process.exit(1) }
  try {
    console.log('Migrating…')
    await migrate({ supabase, data, preset, primary: arg('primary') || preset.defaultIti })
    console.log('\nDone. Sign in to the app and open Home to check the result.')
  } catch (e) { console.error('\nStopped: ' + e.message + '\nNothing is half-hidden: you can delete the trade in Settings → Trades and run again.'); process.exit(1) }
}
