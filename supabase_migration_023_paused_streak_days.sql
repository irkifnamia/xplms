-- XPLMS migration 023: pause streaks for classes blocked by an XP mode.
-- Run after supabase_migration_022_daily_xp_modes.sql.

begin;

create or replace function public.is_student_xp_day_blocked(
  p_no_matrik text,
  p_activity_date date
)
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select coalesce((
    select
      schedule.mode in ('special', 'extra_special')
      and not exists (
        select 1
        from unnest(schedule.selected_classes) as allowed_class
        where upper(trim(allowed_class)) = upper(trim(background."KELAS"))
      )
    from public.xp_mode_schedule schedule
    left join public.stud_background background
      on background."NO MATRIK" = p_no_matrik
    where schedule.mode_date = p_activity_date
  ), false);
$$;

create or replace function public.current_student_streak(p_no_matrik text)
returns integer
language plpgsql
stable
security definer
set search_path = public
as $$
declare
  malaysia_today date := (now() at time zone 'Asia/Kuala_Lumpur')::date;
  target_date date;
  latest_activity date;
  streak integer := 0;
begin
  select max(activity_date)
  into latest_activity
  from public.student_activity_days
  where "NO MATRIK" = p_no_matrik
    and activity_date <= malaysia_today;

  if latest_activity is null then
    return 0;
  end if;

  -- Today is still available. Every completed date after the latest activity
  -- must be a blocked (paused) date; one missed eligible date resets the streak.
  if exists (
    select 1
    from generate_series(
      latest_activity + 1,
      malaysia_today - 1,
      interval '1 day'
    ) as calendar_day
    where not public.is_student_xp_day_blocked(
      p_no_matrik,
      calendar_day::date
    )
  ) then
    return 0;
  end if;

  target_date := latest_activity;
  loop
    -- Blocked dates are neutral: skip them without adding or breaking a day.
    while public.is_student_xp_day_blocked(p_no_matrik, target_date) loop
      target_date := target_date - 1;
    end loop;

    exit when not exists (
      select 1
      from public.student_activity_days
      where "NO MATRIK" = p_no_matrik
        and activity_date = target_date
    );

    streak := streak + 1;
    target_date := target_date - 1;
  end loop;

  return streak;
end;
$$;

create or replace function public.longest_student_streak(p_no_matrik text)
returns integer
language sql
stable
security definer
set search_path = public
as $$
  with days as (
    select distinct activity_date
    from public.student_activity_days
    where "NO MATRIK" = p_no_matrik
      and not public.is_student_xp_day_blocked(p_no_matrik, activity_date)
  ),
  effective_days as (
    select
      activity_date,
      (activity_date - date '2000-01-01')
      - (
          select count(*)::integer
          from public.xp_mode_schedule schedule
          where schedule.mode_date <= days.activity_date
            and public.is_student_xp_day_blocked(
              p_no_matrik,
              schedule.mode_date
            )
        ) as eligible_day_number
    from days
  ),
  islands as (
    select
      eligible_day_number
        - (row_number() over (order by activity_date))::integer as island
    from effective_days
  )
  select coalesce(max(day_count), 0)::integer
  from (
    select count(*) as day_count
    from islands
    group by island
  ) grouped_days;
$$;

comment on function public.is_student_xp_day_blocked(text, date) is
  'True when Special or Extra Special mode blocks the student class on the date. Blocked dates pause streaks.';

commit;
