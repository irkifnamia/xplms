-- XPLMS migration 011
-- Adds the canonical XPTEAM field and makes XP ledger edits balance-safe.
-- Review and run this once in the Supabase SQL editor.

begin;

alter table public.stud_background
  add column if not exists "XPTEAM" text;

create index if not exists stud_background_xpteam_idx
  on public.stud_background ("XPTEAM");

-- Existing approved requests and quiz attempts retain their audit history if
-- an administrator deletes the linked XP event.
alter table public.xp_claims
  drop constraint if exists xp_claims_xp_event_id_fkey;
alter table public.xp_claims
  add constraint xp_claims_xp_event_id_fkey
  foreign key (xp_event_id) references public.xp_events(id) on delete set null;

alter table public.quiz_attempts
  drop constraint if exists quiz_attempts_xp_event_id_fkey;
alter table public.quiz_attempts
  add constraint quiz_attempts_xp_event_id_fkey
  foreign key (xp_event_id) references public.xp_events(id) on delete set null;

create or replace function public.apply_xp_event()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  if tg_op = 'INSERT' then
    update public.stud_xp
    set "XP" = coalesce("XP", 0) + new.points,
        updated_at = now()
    where "NO MATRIK" = new."NO MATRIK";

    if not found then
      raise exception 'No stud_xp row found for NO MATRIK %', new."NO MATRIK";
    end if;
    return new;
  end if;

  if tg_op = 'UPDATE' then
    if old."NO MATRIK" = new."NO MATRIK" then
      update public.stud_xp
      set "XP" = coalesce("XP", 0) + new.points - old.points,
          updated_at = now()
      where "NO MATRIK" = new."NO MATRIK";
      if not found then
        raise exception 'No stud_xp row found for NO MATRIK %', new."NO MATRIK";
      end if;
    else
      update public.stud_xp
      set "XP" = coalesce("XP", 0) - old.points,
          updated_at = now()
      where "NO MATRIK" = old."NO MATRIK";
      if not found then
        raise exception 'No stud_xp row found for NO MATRIK %', old."NO MATRIK";
      end if;

      update public.stud_xp
      set "XP" = coalesce("XP", 0) + new.points,
          updated_at = now()
      where "NO MATRIK" = new."NO MATRIK";
      if not found then
        raise exception 'No stud_xp row found for NO MATRIK %', new."NO MATRIK";
      end if;
    end if;
    return new;
  end if;

  update public.stud_xp
  set "XP" = coalesce("XP", 0) - old.points,
      updated_at = now()
  where "NO MATRIK" = old."NO MATRIK";
  if not found then
    raise exception 'No stud_xp row found for NO MATRIK %', old."NO MATRIK";
  end if;
  return old;
end;
$$;

drop trigger if exists on_xp_event_created on public.xp_events;
drop trigger if exists on_xp_event_changed on public.xp_events;
create trigger on_xp_event_changed
after insert or update of points, "NO MATRIK" or delete on public.xp_events
for each row execute function public.apply_xp_event();

-- Badges remain permanent. An upward XP edit may capture a newly reached badge,
-- while a downward edit or deletion never removes an earned badge.
drop trigger if exists zz_capture_xp_badges on public.xp_events;
create trigger zz_capture_xp_badges
after insert or update of points, "NO MATRIK" on public.xp_events
for each row execute function public.capture_xp_badges();

commit;
