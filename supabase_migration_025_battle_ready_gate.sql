-- XPLMS migration 025: both players must be ready before the next round starts.
-- Run after supabase_migration_024_student_battles.sql.

begin;

create table if not exists public.battle_round_ready (
  battle_id uuid not null,
  position integer not null check (position between 1 and 9),
  student_user_id uuid not null references public.app_users(id) on delete cascade,
  ready_at timestamptz not null default clock_timestamp(),
  primary key (battle_id, position, student_user_id),
  foreign key (battle_id, position)
    references public.battle_questions(battle_id, position) on delete cascade
);

alter table public.battle_round_ready enable row level security;
revoke all on table public.battle_round_ready from anon, authenticated;

create or replace function public.hold_battle_question_advance()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare ready_count integer;
begin
  if new.current_question > old.current_question then
    select count(*) into ready_count
    from public.battle_round_ready
    where battle_id = old.id and position = old.current_question;
    if ready_count < 2 then
      new.current_question := old.current_question;
    end if;
  end if;
  return new;
end;
$$;

drop trigger if exists hold_battle_question_advance on public.battle_matches;
create trigger hold_battle_question_advance
before update of current_question on public.battle_matches
for each row execute function public.hold_battle_question_advance();

create or replace function public.guard_battle_question_open()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare active_position integer;
begin
  if old.opened_at is null and new.opened_at is not null then
    select current_question into active_position
    from public.battle_matches where id = new.battle_id;
    if active_position is distinct from new.position then
      new.opened_at := null;
    end if;
  end if;
  return new;
end;
$$;

drop trigger if exists guard_battle_question_open on public.battle_questions;
create trigger guard_battle_question_open
before update of opened_at on public.battle_questions
for each row execute function public.guard_battle_question_open();

create or replace function public.ready_for_next_battle_question(
  p_battle_id uuid,
  p_student_user_id uuid,
  p_position integer
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  battle public.battle_matches%rowtype;
  question_outcome text;
  ready_count integer;
begin
  select * into battle from public.battle_matches
  where id = p_battle_id for update;
  if battle.id is null or battle.status <> 'active'
     or p_student_user_id not in (battle.player_a_id, battle.player_b_id)
     or battle.current_question <> p_position or p_position >= 10 then
    raise exception 'This battle round cannot advance';
  end if;
  select outcome into question_outcome
  from public.battle_questions
  where battle_id = p_battle_id and position = p_position;
  if question_outcome is null then
    raise exception 'Both answers must be resolved first';
  end if;
  insert into public.battle_round_ready(battle_id, position, student_user_id)
  values (p_battle_id, p_position, p_student_user_id)
  on conflict do nothing;
  select count(*) into ready_count
  from public.battle_round_ready
  where battle_id = p_battle_id and position = p_position;
  if ready_count = 2 then
    update public.battle_matches
    set current_question = p_position + 1
    where id = p_battle_id;
    update public.battle_questions
    set opened_at = clock_timestamp()
    where battle_id = p_battle_id and position = p_position + 1;
    return jsonb_build_object('status', 'started', 'position', p_position + 1);
  end if;
  return jsonb_build_object('status', 'waiting', 'ready_count', ready_count);
end;
$$;

commit;
