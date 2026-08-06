-- XPLMS migration 029
-- Increase the permanent XP Legend requirement from 1,500 to 2,000 overall XP.
-- Run after migration 019 (and after the later Battle migrations).

begin;

create or replace function public.capture_xp_badges()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  current_xp integer;
begin
  select coalesce("XP", 0) into current_xp
  from public.stud_xp
  where "NO MATRIK" = new."NO MATRIK";

  insert into public.student_badges (
    "NO MATRIK", badge_family, badge_code, badge_name,
    xp_threshold, source_xp_event_id, earned_at
  )
  select
    new."NO MATRIK", 'xp', level.badge_code, level.badge_name,
    level.threshold, new.id, coalesce(new.created_at, now())
  from (
    values
      ('rookie', 'Rookie', 200),
      ('explorer', 'Explorer', 500),
      ('expert', 'Expert', 1000),
      ('legend', 'Legend', 2000)
  ) as level(badge_code, badge_name, threshold)
  where current_xp >= level.threshold
  on conflict ("NO MATRIK", badge_family, badge_code) do nothing;
  return new;
end;
$$;

-- Remove old 1,500-XP Legend awards where the student has never reached 2,000.
with running_xp as (
  select
    "NO MATRIK",
    sum(points) over (
      partition by "NO MATRIK"
      order by created_at, id
      rows between unbounded preceding and current row
    ) as running_total
  from public.xp_events
), peak_xp as (
  select
    sx."NO MATRIK",
    greatest(coalesce(sx."XP", 0), coalesce(max(rx.running_total), 0)) as peak_total
  from public.stud_xp sx
  left join running_xp rx using ("NO MATRIK")
  group by sx."NO MATRIK", sx."XP"
)
delete from public.student_badges sb
using peak_xp px
where sb.badge_family = 'xp'
  and sb.badge_code = 'legend'
  and sb."NO MATRIK" = px."NO MATRIK"
  and px.peak_total < 2000;

update public.student_badges
set xp_threshold = 2000
where badge_family = 'xp' and badge_code = 'legend';

-- Backfill qualifying Legend badges at the exact first 2,000-XP crossing time.
with running_xp as (
  select
    id,
    "NO MATRIK",
    created_at,
    sum(points) over (
      partition by "NO MATRIK"
      order by created_at, id
      rows between unbounded preceding and current row
    ) as running_total
  from public.xp_events
), first_crossing as (
  select distinct on ("NO MATRIK")
    "NO MATRIK", id as source_event_id, created_at as earned_at
  from running_xp
  where running_total >= 2000
  order by "NO MATRIK", created_at, id
)
insert into public.student_badges (
  "NO MATRIK", badge_family, badge_code, badge_name,
  xp_threshold, source_xp_event_id, earned_at
)
select
  "NO MATRIK", 'xp', 'legend', 'Legend', 2000,
  source_event_id, earned_at
from first_crossing
on conflict ("NO MATRIK", badge_family, badge_code)
do update set
  badge_name = excluded.badge_name,
  xp_threshold = excluded.xp_threshold,
  source_xp_event_id = excluded.source_xp_event_id,
  earned_at = excluded.earned_at;

commit;
