-- XPLMS migration 001: secure student data and XP functions
-- Run this once in the Supabase SQL Editor after supabase_schema.sql.

begin;

-- Anonymous users must never execute XP award functions.
revoke execute on function public.award_manual_xp(text, text, text, integer, text) from anon;
revoke execute on function public.award_earliest_submission_xp(bigint) from anon;

-- Make the manual award role check safe when no authenticated role exists.
create or replace function public.award_manual_xp(
  p_no_matrik text,
  p_rule_code text,
  p_reason text,
  p_points_override integer default null,
  p_reference text default null
) returns bigint
language plpgsql
security definer
set search_path = public
as $$
declare
  v_event_id bigint;
  v_points integer;
begin
  if coalesce(public.current_role(), '') not in ('lecturer','admin') then
    raise exception 'Only lecturers and administrators can award XP';
  end if;

  if not exists (
    select 1 from public.profiles
    where "NO MATRIK" = p_no_matrik and role = 'student'
  ) then
    raise exception 'No student profile found for NO MATRIK %', p_no_matrik;
  end if;

  select coalesce(p_points_override, default_points)
  into v_points
  from public.xp_rules
  where code = p_rule_code and award_mode = 'manual' and active;

  if v_points is null or v_points <= 0 or v_points > 1000 then
    raise exception 'Invalid or inactive manual XP rule/points';
  end if;

  insert into public.xp_events (
    "NO MATRIK", rule_code, points, source_id, reason, award_mode, awarded_by
  )
  values (
    p_no_matrik,
    p_rule_code,
    v_points,
    coalesce(nullif(p_reference, ''), gen_random_uuid()::text),
    p_reason,
    'manual',
    auth.uid()
  )
  returning id into v_event_id;

  return v_event_id;
end;
$$;

revoke all on function public.award_manual_xp(text, text, text, integer, text) from public;
revoke execute on function public.award_manual_xp(text, text, text, integer, text) from anon;
grant execute on function public.award_manual_xp(text, text, text, integer, text) to authenticated;

-- Student datasets contain personal information and must require a login.
alter table public.stud_background enable row level security;
alter table public.stud_progress enable row level security;
alter table public.stud_xp enable row level security;

revoke all on table public.stud_background from anon;
revoke all on table public.stud_progress from anon;
revoke all on table public.stud_xp from anon;

grant select, insert, update, delete on table
  public.stud_background,
  public.stud_progress
to authenticated;
grant select on table public.stud_xp to authenticated;

drop policy if exists "stud_background_read" on public.stud_background;
create policy "stud_background_read"
on public.stud_background for select to authenticated
using (
  "NO MATRIK" = (
    select p."NO MATRIK" from public.profiles p where p.id = auth.uid()
  )
  or public.current_role() in ('lecturer','admin')
);

drop policy if exists "stud_background_admin_write" on public.stud_background;
create policy "stud_background_admin_write"
on public.stud_background for all to authenticated
using (public.current_role() = 'admin')
with check (public.current_role() = 'admin');

drop policy if exists "stud_progress_read" on public.stud_progress;
create policy "stud_progress_read"
on public.stud_progress for select to authenticated
using (
  "NO MATRIK" = (
    select p."NO MATRIK" from public.profiles p where p.id = auth.uid()
  )
  or public.current_role() in ('lecturer','admin')
);

drop policy if exists "stud_progress_admin_write" on public.stud_progress;
create policy "stud_progress_admin_write"
on public.stud_progress for all to authenticated
using (public.current_role() = 'admin')
with check (public.current_role() = 'admin');

drop policy if exists "stud_xp_read" on public.stud_xp;
create policy "stud_xp_read"
on public.stud_xp for select to authenticated
using (
  "NO MATRIK" = (
    select p."NO MATRIK" from public.profiles p where p.id = auth.uid()
  )
  or public.current_role() in ('lecturer','admin')
);

commit;
