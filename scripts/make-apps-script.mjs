#!/usr/bin/env node
// Builds the Google Apps Script that creates the data-entry Google Sheet for a trade.
//   node scripts/make-apps-script.mjs presets/fitter-cts3.json > apps-script/fitter-cts3-sheet.gs
import { readFileSync } from 'node:fs'

export function appsScript(preset) {
  const body = readFileSync(new URL('../apps-script/body.gs.txt', import.meta.url), 'utf8')
  const iti = preset.itis.find((i) => i.key === preset.defaultIti) || preset.itis[0]
  const data = {
    sheetTitle: `${preset.trade.name} logbook data — ${iti.name}`,
    tradeName: preset.trade.name,
    los: preset.los.map((l) => ({ number: String(l.number), title: l.title, description: l.description || '', year: l.year || null })),
  }
  return `// Generated from presets/${preset.id}.json — do not edit this block; edit the preset and regenerate.\nvar PRESET = ${JSON.stringify(data, null, 1)};\n\n${body}`
}

if (process.argv[1] && import.meta.url === new URL('file://' + process.argv[1]).href) {
  const f = process.argv[2]; if (!f) { console.error('Usage: node scripts/make-apps-script.mjs presets/<name>.json > apps-script/<name>-sheet.gs'); process.exit(2) }
  process.stdout.write(appsScript(JSON.parse(readFileSync(f, 'utf8'))))
}
