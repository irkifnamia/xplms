-- XPLMS migration 022
-- Daily Normal, Special and Extra Special XP modes with class enforcement.
-- Run once in Supabase SQL Editor after migration 021.

begin;

create table if not exists public.xp_mode_schedule (
  mode_date date primary key,
  mode text not null default 'normal' check (
    mode in ('normal', 'special', 'extra_special')
  ),
  selected_classes text[] not null default '{}',
  configured_by uuid references public.app_users(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  check (
    mode = 'normal'
    or cardinality(selected_classes) > 0
  )
);

alter table public.xp_mode_schedule enable row level security;
revoke all on table public.xp_mode_schedule from anon, authenticated;

alter table public.xp_events
  add column if not exists base_points integer,
  add column if not exists mode_multiplier integer not null default 1
    check (mode_multiplier in (1, 2)),
  add column if not exists xp_mode text not null default 'normal'
    check (xp_mode in ('normal', 'special', 'extra_special'));

update public.xp_events
set base_points = points
where base_points is null;

alter table public.xp_events
  alter column base_points set not null;

create or replace function public.enforce_daily_xp_mode()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  activity_date date;
  active_mode text := 'normal';
  allowed_classes text[] := '{}';
  student_class text;
begin
  activity_date := (
    coalesce(new.created_at, now()) at time zone 'Asia/Kuala_Lumpur'
  )::date;

  select mode, selected_classes
  into active_mode, allowed_classes
  from public.xp_mode_schedule
  where mode_date = activity_date;

  active_mode := coalesce(active_mode, 'normal');
  allowed_classes := coalesce(allowed_classes, '{}');

  select nullif(trim("KELAS"), '')
  into student_class
  from public.stud_background
  where "NO MATRIK" = new."NO MATRIK"
  limit 1;

  if active_mode in ('special', 'extra_special')
     and not coalesce(student_class = any(allowed_classes), false) then
    raise exception 'XP activity is locked for class % on % (% mode)',
      coalesce(student_class, 'UNKNOWN'), activity_date, active_mode;
  end if;

  new.base_points := new.points;
  new.xp_mode := active_mode;
  new.mode_multiplier := case
    when active_mode = 'extra_special' then 2
    else 1
  end;
  new.points := new.base_points * new.mode_multiplier;
  return new;
end;
$$;

drop trigger if exists aa_enforce_daily_xp_mode on public.xp_events;
create trigger aa_enforce_daily_xp_mode
before insert on public.xp_events
for each row execute function public.enforce_daily_xp_mode();

create or replace function public.enforce_daily_xp_claim_mode()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  active_mode text := 'normal';
  allowed_classes text[] := '{}';
  student_class text;
  activity_date date := (now() at time zone 'Asia/Kuala_Lumpur')::date;
begin
  select mode, selected_classes
  into active_mode, allowed_classes
  from public.xp_mode_schedule
  where mode_date = activity_date;

  active_mode := coalesce(active_mode, 'normal');
  allowed_classes := coalesce(allowed_classes, '{}');

  if active_mode in ('special', 'extra_special') then
    select nullif(trim("KELAS"), '')
    into student_class
    from public.stud_background
    where "NO MATRIK" = new."NO MATRIK"
    limit 1;

    if not coalesce(student_class = any(allowed_classes), false) then
      raise exception 'XP requests are locked for class % on % (% mode)',
        coalesce(student_class, 'UNKNOWN'), activity_date, active_mode;
    end if;
  end if;
  return new;
end;
$$;

drop trigger if exists aa_enforce_daily_xp_claim_mode on public.xp_claims;
create trigger aa_enforce_daily_xp_claim_mode
before insert on public.xp_claims
for each row execute function public.enforce_daily_xp_claim_mode();

-- Replace migration 021's correction function so corrected quiz XP preserves
-- the multiplier that applied on the original quiz day.
create or replace function public.resolve_quiz_question_report(
  p_report_id bigint,
  p_status text,
  p_admin_notes text,
  p_corrected_correct_count integer,
  p_reviewed_by uuid
)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  report_row public.quiz_question_reports%rowtype;
  attempt_row public.quiz_attempts%rowtype;
  v_corrected_count integer;
  v_corrected_score numeric;
  v_multiplier integer := 1;
  v_corrected_xp integer;
begin
  if p_status not in ('accepted', 'rejected') then
    raise exception 'Resolution status must be accepted or rejected';
  end if;

  select * into report_row from public.quiz_question_reports
  where id = p_report_id for update;
  if not found then
    raise exception 'Quiz question report % was not found', p_report_id;
  end if;
  if report_row.status <> 'pending' then
    raise exception 'Quiz question report % is already resolved', p_report_id;
  end if;

  select * into attempt_row from public.quiz_attempts
  where id = report_row.attempt_id for update;
  if not found then
    raise exception 'Quiz attempt % was not found', report_row.attempt_id;
  end if;

  v_corrected_count := attempt_row.correct_count;
  if p_status = 'accepted' and p_corrected_correct_count is not null then
    v_corrected_count := greatest(
      0, least(p_corrected_correct_count, attempt_row.total_questions)
    );
  end if;
  v_corrected_score := case
    when attempt_row.total_questions > 0 then
      v_corrected_count::numeric / attempt_row.total_questions * 100
    else 0
  end;

  if attempt_row.xp_event_id is not null then
    select coalesce(mode_multiplier, 1) into v_multiplier
    from public.xp_events where id = attempt_row.xp_event_id;
  end if;
  v_corrected_xp := (5 + v_corrected_count) * coalesce(v_multiplier, 1);

  if p_status = 'accepted' then
    update public.quiz_attempts
    set correct_count = v_corrected_count,
        score = v_corrected_score,
        passed = v_corrected_score >= 60,
        xp_awarded = v_corrected_xp
    where id = attempt_row.id;

    if attempt_row.xp_event_id is not null then
      update public.xp_events
      set base_points = 5 + v_corrected_count,
          points = v_corrected_xp,
          reason = concat(
            'Corrected daily C', attempt_row.chapter, ' quiz: ',
            v_corrected_count, '/', attempt_row.total_questions,
            ' correct after question report ', p_report_id
          )
      where id = attempt_row.xp_event_id;
    end if;
  end if;

  update public.quiz_question_reports
  set status = p_status,
      admin_notes = nullif(trim(p_admin_notes), ''),
      reviewed_by = p_reviewed_by,
      reviewed_at = now(),
      original_correct_count = attempt_row.correct_count,
      corrected_correct_count = v_corrected_count,
      original_xp = attempt_row.xp_awarded,
      corrected_xp = v_corrected_xp,
      updated_at = now()
  where id = p_report_id;
end;
$$;

revoke all on function public.resolve_quiz_question_report(
  bigint, text, text, integer, uuid
) from public, anon, authenticated;
grant execute on function public.resolve_quiz_question_report(
  bigint, text, text, integer, uuid
) to service_role;

commit;
