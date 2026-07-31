-- XPLMS migration 017
-- Revised permanent XP badge thresholds:
-- Rookie 200, Explorer 500, Expert 1000, Legend 1500.

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
    xp_threshold, source_xp_event_id
  )
  select
    new."NO MATRIK", 'xp', level.badge_code, level.badge_name,
    level.threshold, new.id
  from (
    values
      ('rookie', 'Rookie', 200),
      ('explorer', 'Explorer', 500),
      ('expert', 'Expert', 1000),
      ('legend', 'Legend', 1500)
  ) as level(badge_code, badge_name, threshold)
  where current_xp >= level.threshold
  on conflict ("NO MATRIK", badge_family, badge_code) do nothing;
  return new;
end;
$$;

update public.student_badges
set xp_threshold = case badge_code
  when 'rookie' then 200
  when 'explorer' then 500
  when 'expert' then 1000
  when 'legend' then 1500
end
where badge_family = 'xp'
  and badge_code in ('rookie', 'explorer', 'expert', 'legend');

with running_xp as (
  select
    "NO MATRIK",
    sum(points) over (
      partition by "NO MATRIK"
      order by created_at, id
      rows between unbounded preceding and current row
    ) as running_total
  from public.xp_events
),
peak_xp as (
  select
    sx."NO MATRIK",
    greatest(
      coalesce(sx."XP", 0),
      coalesce(max(rx.running_total), 0)
    ) as peak_total
  from public.stud_xp sx
  left join running_xp rx using ("NO MATRIK")
  group by sx."NO MATRIK", sx."XP"
)
delete from public.student_badges sb
using peak_xp px
where sb.badge_family = 'xp'
  and sb."NO MATRIK" = px."NO MATRIK"
  and coalesce(sb.xp_threshold, 0) > px.peak_total;

with running_xp as (
  select
    "NO MATRIK",
    sum(points) over (
      partition by "NO MATRIK"
      order by created_at, id
      rows between unbounded preceding and current row
    ) as running_total
  from public.xp_events
),
peak_xp as (
  select
    sx."NO MATRIK",
    greatest(
      coalesce(sx."XP", 0),
      coalesce(max(rx.running_total), 0)
    ) as peak_total
  from public.stud_xp sx
  left join running_xp rx using ("NO MATRIK")
  group by sx."NO MATRIK", sx."XP"
),
levels (badge_code, badge_name, threshold) as (
  values
    ('rookie', 'Rookie', 200),
    ('explorer', 'Explorer', 500),
    ('expert', 'Expert', 1000),
    ('legend', 'Legend', 1500)
)
insert into public.student_badges (
  "NO MATRIK", badge_family, badge_code, badge_name, xp_threshold
)
select
  px."NO MATRIK", 'xp', level.badge_code, level.badge_name, level.threshold
from peak_xp px
cross join levels level
where px.peak_total >= level.threshold
on conflict ("NO MATRIK", badge_family, badge_code)
do update set
  badge_name = excluded.badge_name,
  xp_threshold = excluded.xp_threshold;

commit;
