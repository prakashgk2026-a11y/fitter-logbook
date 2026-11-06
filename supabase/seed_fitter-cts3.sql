-- =====================================================================
-- SEED: Fitter — CTS 3.0 (NSQF Level 4): learning outcomes
-- Creates Govt. ITI Nedumangad + the Fitter trade + 22 learning outcomes.
-- Run ONCE in Supabase → SQL Editor, after schema.sql. Running it again does not duplicate anything.
-- (Generated from presets/fitter-cts3.json by scripts/make-seed-sql.mjs — edit the preset, not this file.)
-- =====================================================================
do $seed$
declare v_iti uuid; v_trade uuid;
begin
  select id into v_iti from public.itis where name = 'Govt. ITI Nedumangad';
  if v_iti is null then
    insert into public.itis (name, short_name, print_subtitle)
    values ('Govt. ITI Nedumangad', 'ITI Nedumangad', 'Thiruvananthapuram, Kerala') returning id into v_iti;
  end if;

  select id into v_trade from public.trades where iti_id = v_iti and name = 'Fitter';
  if v_trade is null then
    insert into public.trades (iti_id, name, code, duration_months)
    values (v_iti, 'Fitter', 'FITTER', 24) returning id into v_trade;
  end if;

  insert into public.learning_outcomes (trade_id, number, title, description, sort_order)
  select v_trade, x.number, x.title, x.description, x.sort_order
    from (values
    ('1', 'Plan and organize the work to make job as per specification applying different types of basic fitting operation and check for dimensional accuracy following safety precautions.', 'Basic fitting operation – marking, Hacksawing, Chiselling, Filing, Drilling, Taping and Grinding etc. Accuracy: ± 0.25mm', 1),
    ('2', 'Manufacture simple sheet metal items as per drawing and join them by soldering, brazing and riveting.', '', 2),
    ('3', 'Join metal components by riveting observing standard procedure.', '', 3),
    ('4', 'Join metal component by arc welding observing standard procedure.', '', 4),
    ('5', 'Cut and join metal component by gas (oxyacetylene / LPG / oxy-hydrogen).', '', 5),
    ('6', 'Produce components by different operations and check accuracy using appropriate measuring instruments.', 'Different Operations - Drilling, Reaming, Taping, Dieing; Appropriate Measuring Instrument – Vernier, Screw Gauge, Micrometer', 6),
    ('7', 'Make different fit of components for assembling as per required tolerance observing principle of interchangeability and check for functionality.', 'Different Fit – Sliding, Angular, Step fit, ‘T’ fit, Square fit and Profile fit; Required tolerance: ±0.04 mm, angular tolerance: 30 min.', 7),
    ('8', 'Produce components involving different operations on lathe observing standard procedure and check for accuracy.', 'Different Operations – facing, plain turning, step turning, parting, chamfering, shoulder turn, grooving, knurling, boring, taper turning, threading (external ‘V’ only)', 8),
    ('9', 'Plan & perform simple repair, overhauling of different machines and check for functionality.', 'Different Machines – Drill Machine, Power Saw, Bench Grinder and Lathe', 9),
    ('10', 'Read and apply engineering drawing for different application in the field of work.', '', 10),
    ('11', 'Demonstrate basic mathematical concept and principles to perform practical operations. Understand and explain basic science in the field of study.', '', 11),
    ('12', 'Perform surface finishing operations using different fastening components, tools and check functionality.', 'Different fastening components – Dowel pins, screws, bolts, keys and cotters; Different fastening tools – hand operated & power tools; Required tolerance: ±0.02mm, angular tolerance ± 10 min.', 12),
    ('13', 'Make & assemble components of different mating surfaces as per required tolerance.', 'Different Mating Surfaces – Dovetail fitting, Radius fitting, Combined fitting; Different surface finishing operations – Scraping, Lapping and Honing', 13),
    ('14', 'Make different gauges by using standard tools & equipment and checks for specified accuracy.', 'Different Gauges – Snap gauge, Gap gauge; Specified Accuracy: ±0.02mm', 14),
    ('15', 'Apply a range of skills to execute pipe joints, dismantle and assemble valves & fittings with pipes and test for leakages.', 'Range of skills – Cutting, Threading, Flaring, Bending and Joining', 15),
    ('16', 'Make drill jig & produce components on drill machine by using jigs and check for correctness.', '', 16),
    ('17', 'Plan, dismantle, repair and assemble different damaged mechanical components used for power transmission & check functionality.', 'Different Damaged Mechanical Components – Pulley, Gear, Keys, Jibs and Shafts', 17),
    ('18', 'Identify, dismantle, replace and assemble different pneumatics and hydraulics components.', 'Different components – Compressor, Pressure Gauge, Filter Regulator Lubricator, Valves and Actuators', 18),
    ('19', 'Construct circuit of pneumatics and hydraulics observing standard operating procedure & safety aspect.', '', 19),
    ('20', 'Plan & perform basic day to day preventive maintenance, repairing and check functionality.', 'Simple Machines – Drill Machine, Power Saw and Lathe', 20),
    ('21', 'Plan, erect simple machine and test machine tool accuracy.', 'Simple Machines – Drill Machine, Power Saw and Lathe', 21),
    ('22', 'Read and apply engineering drawing for different application in the field of work.', '', 22)
    ) as x(number, title, description, sort_order)
   where not exists (select 1 from public.learning_outcomes l where l.trade_id = v_trade and l.number = x.number);
end
$seed$;

-- check: should list 22 rows
select l.number, l.title from public.learning_outcomes l
  join public.trades t on t.id = l.trade_id where t.name = 'Fitter' order by l.number::int;
