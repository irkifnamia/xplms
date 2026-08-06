-- XPLMS migration 028: consultation and completed Battle qualify for streaks.
-- Run after migration 027.

begin;

alter table public.student_activity_days
  drop constraint if exists student_activity_days_activity_type_check;
alter table public.student_activity_days
  add constraint student_activity_days_activity_type_check
  check (activity_type in (
    'in_app_quiz', 'extra_practice', 'consultation', 'battle'
  ));

create or replace function public.sync_xp_event_streak_day()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  event_row public.xp_events%rowtype;
  activity_kind text;
  activity_timestamp timestamptz;
begin
  if tg_op in ('UPDATE', 'DELETE') then
    if old.rule_code in ('consultation', 'battle_result') then
      delete from public.student_activity_days
      where activity_type = case
          when old.rule_code = 'consultation' then 'consultation' else 'battle' end
        and source_id = old.id::text;
    end if;
  end if;
  if tg_op = 'DELETE' then
    return old;
  end if;
  event_row := new;
  if event_row.points <= 0
     or event_row.rule_code not in ('consultation', 'battle_result') then
    return new;
  end if;
  activity_kind := case
    when event_row.rule_code = 'consultation' then 'consultation'
    else 'battle'
  end;
  activity_timestamp := event_row.created_at;
  -- For an approved consultation request, use the original request time.
  if event_row.rule_code = 'consultation'
     and event_row.source_id like 'claim-%' then
    select claim.created_at into activity_timestamp
    from public.xp_claims claim
    where 'claim-' || claim.id::text = event_row.source_id;
    activity_timestamp := coalesce(activity_timestamp, event_row.created_at);
  end if;
  insert into public.student_activity_days(
    "NO MATRIK", activity_date, activity_type, source_id
  ) values (
    event_row."NO MATRIK",
    (activity_timestamp at time zone 'Asia/Kuala_Lumpur')::date,
    activity_kind,
    event_row.id::text
  ) on conflict ("NO MATRIK", activity_type, source_id) do nothing;
  return new;
end;
$$;

drop trigger if exists sync_xp_event_streak_day on public.xp_events;
create trigger sync_xp_event_streak_day
after insert or update or delete on public.xp_events
for each row execute function public.sync_xp_event_streak_day();

-- Backfill qualifying historical consultation and Battle XP events.
insert into public.student_activity_days(
  "NO MATRIK", activity_date, activity_type, source_id
)
select
  event."NO MATRIK",
  (
    coalesce(claim.created_at, event.created_at)
    at time zone 'Asia/Kuala_Lumpur'
  )::date,
  case when event.rule_code = 'consultation' then 'consultation' else 'battle' end,
  event.id::text
from public.xp_events event
left join public.xp_claims claim
  on event.rule_code = 'consultation'
 and event.source_id = 'claim-' || claim.id::text
where event.rule_code in ('consultation', 'battle_result')
  and event.points > 0
on conflict ("NO MATRIK", activity_type, source_id) do nothing;

-- Capture any streak badge thresholds reached through the historical backfill.
insert into public.student_badges(
  "NO MATRIK", badge_family, badge_code, badge_name, streak_threshold
)
select
  summary."NO MATRIK", 'streak', level.badge_code,
  level.badge_name, level.threshold
from public.student_streak_summary summary
cross join (
  values
    ('rookie', 'Rookie', 7),
    ('explorer', 'Explorer', 14),
    ('expert', 'Expert', 21),
    ('legend', 'Legend', 28)
) as level(badge_code, badge_name, threshold)
where summary.longest_streak >= level.threshold
on conflict ("NO MATRIK", badge_family, badge_code) do nothing;

commit;
