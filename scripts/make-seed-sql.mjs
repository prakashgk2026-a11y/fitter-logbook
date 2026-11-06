#!/usr/bin/env node
// Turns a preset (presets/*.json) into a SQL file you paste into Supabase → SQL Editor.
//   node scripts/make-seed-sql.mjs presets/fitter-cts3.json > supabase/seed_fitter-cts3.sql
// The SQL creates the ITI(s), the trade and all learning outcomes. Safe to run twice (nothing is duplicated).
import { readFileSync } from 'node:fs'

export function seedSql(preset) {
  const q = (s) => `'${String(s ?? '').replace(/'/g, "''")}'`
  const iti = preset.itis.find((i) => i.key === preset.defaultIti) || preset.itis[0]
  const t = preset.trade
  const los = preset.los.map((l, i) => `    (${q(l.number)}, ${q(l.title)}, ${q(l.description || '')}, ${Number(l.number) || i + 1})`).join(',\n')
  return `-- =====================================================================
-- SEED: ${preset.title || preset.id}
-- Creates ${iti.name} + the ${t.name} trade + ${preset.los.length} learning outcomes.
-- Run ONCE in Supabase → SQL Editor, after schema.sql. Running it again does not duplicate anything.
-- (Generated from presets/${preset.id}.json by scripts/make-seed-sql.mjs — edit the preset, not this file.)
-- =====================================================================
do $seed$
declare v_iti uuid; v_trade uuid;
begin
  select id into v_iti from public.itis where name = ${q(iti.name)};
  if v_iti is null then
    insert into public.itis (name, short_name, print_subtitle)
    values (${q(iti.name)}, ${q(iti.short)}, ${q(iti.subtitle || '')}) returning id into v_iti;
  end if;

  select id into v_trade from public.trades where iti_id = v_iti and name = ${q(t.name)};
  if v_trade is null then
    insert into public.trades (iti_id, name, code, duration_months)
    values (v_iti, ${q(t.name)}, ${q(t.code)}, ${Number(t.months) || 12}) returning id into v_trade;
  end if;

  insert into public.learning_outcomes (trade_id, number, title, description, sort_order)
  select v_trade, x.number, x.title, x.description, x.sort_order
    from (values
${los}
    ) as x(number, title, description, sort_order)
   where not exists (select 1 from public.learning_outcomes l where l.trade_id = v_trade and l.number = x.number);
end
$seed$;

-- check: should list ${preset.los.length} rows
select l.number, l.title from public.learning_outcomes l
  join public.trades t on t.id = l.trade_id where t.name = ${q(t.name)} order by l.number::int;
`
}

if (process.argv[1] && import.meta.url === new URL('file://' + process.argv[1]).href) {
  const f = process.argv[2]; if (!f) { console.error('Usage: node scripts/make-seed-sql.mjs presets/<name>.json > supabase/seed_<name>.sql'); process.exit(2) }
  process.stdout.write(seedSql(JSON.parse(readFileSync(f, 'utf8'))))
}
