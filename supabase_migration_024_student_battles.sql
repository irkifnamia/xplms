-- XPLMS migration 024: live two-player quiz battles.
-- Run after migrations 015, 022 and 023.

begin;

insert into public.xp_rules (
  code, name, description, award_mode, default_points, active
)
values (
  'battle_result', 'Quiz battle',
  'Two-player battle: 5 XP completed loss, 8 XP draw, 10 XP win.',
  'automatic', 5, true
)
on conflict (code) do update
set name = excluded.name,
    description = excluded.description,
    award_mode = excluded.award_mode,
    default_points = excluded.default_points,
    active = excluded.active,
    updated_at = now();

create table if not exists public.battle_presence (
  student_user_id uuid primary key references public.app_users(id) on delete cascade,
  "NO MATRIK" text not null,
  last_seen_at timestamptz not null default now()
);

create table if not exists public.battle_challenges (
  id uuid primary key default gen_random_uuid(),
  challenger_id uuid not null references public.app_users(id) on delete cascade,
  opponent_id uuid not null references public.app_users(id) on delete cascade,
  challenge_date date not null default ((now() at time zone 'Asia/Kuala_Lumpur')::date),
  pair_key text not null,
  status text not null default 'pending'
    check (status in ('pending', 'accepted', 'rejected', 'expired')),
  battle_id uuid,
  created_at timestamptz not null default now(),
  responded_at timestamptz,
  check (challenger_id <> opponent_id)
);

create table if not exists public.battle_matches (
  id uuid primary key default gen_random_uuid(),
  challenge_id uuid not null unique references public.battle_challenges(id) on delete cascade,
  player_a_id uuid not null references public.app_users(id) on delete cascade,
  player_b_id uuid not null references public.app_users(id) on delete cascade,
  player_a_matric text not null,
  player_b_matric text not null,
  status text not null default 'active'
    check (status in ('active', 'completed', 'cancelled')),
  current_question integer not null default 1 check (current_question between 1 and 10),
  player_a_wins integer not null default 0,
  player_b_wins integer not null default 0,
  draws integer not null default 0,
  winner_id uuid references public.app_users(id) on delete set null,
  player_a_xp integer,
  player_b_xp integer,
  started_at timestamptz not null default now(),
  completed_at timestamptz,
  check (player_a_id <> player_b_id)
);

alter table public.battle_challenges
  drop constraint if exists battle_challenges_battle_id_fkey;
alter table public.battle_challenges
  add constraint battle_challenges_battle_id_fkey
  foreign key (battle_id) references public.battle_matches(id) on delete set null;

create table if not exists public.battle_questions (
  battle_id uuid not null references public.battle_matches(id) on delete cascade,
  position integer not null check (position between 1 and 10),
  question_id bigint not null references public.quiz_questions(id) on delete restrict,
  difficulty text not null check (difficulty in ('easy', 'medium')),
  question text not null,
  options jsonb not null,
  correct_index integer not null check (correct_index between 0 and 3),
  explanation text not null,
  opened_at timestamptz,
  winner_id uuid references public.app_users(id) on delete set null,
  outcome text check (outcome in ('player_a', 'player_b', 'draw')),
  primary key (battle_id, position),
  unique (battle_id, question_id)
);

create table if not exists public.battle_answers (
  battle_id uuid not null,
  position integer not null,
  student_user_id uuid not null references public.app_users(id) on delete cascade,
  selected_index integer check (selected_index between 0 and 3),
  is_correct boolean not null default false,
  submitted_at timestamptz not null default clock_timestamp(),
  time_taken_ms integer not null default 60000 check (time_taken_ms between 0 and 60000),
  timed_out boolean not null default false,
  primary key (battle_id, position, student_user_id),
  foreign key (battle_id, position)
    references public.battle_questions(battle_id, position) on delete cascade
);

create index if not exists battle_presence_last_seen_idx
  on public.battle_presence(last_seen_at desc);
create index if not exists battle_challenges_opponent_status_idx
  on public.battle_challenges(opponent_id, status, created_at desc);
create unique index if not exists battle_challenges_active_pair_day_uidx
  on public.battle_challenges(pair_key, challenge_date)
  where status in ('pending', 'accepted');
create index if not exists battle_matches_players_idx
  on public.battle_matches(player_a_id, player_b_id, started_at desc);

alter table public.battle_presence enable row level security;
alter table public.battle_challenges enable row level security;
alter table public.battle_matches enable row level security;
alter table public.battle_questions enable row level security;
alter table public.battle_answers enable row level security;
revoke all on table public.battle_presence, public.battle_challenges,
  public.battle_matches, public.battle_questions, public.battle_answers
from anon, authenticated;

create or replace function public.battle_heartbeat(p_student_user_id uuid)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare student_matric text;
begin
  update public.battle_challenges
  set status = 'expired', responded_at = now()
  where status = 'pending'
    and challenge_date < (now() at time zone 'Asia/Kuala_Lumpur')::date;
  select "NO MATRIK" into student_matric
  from public.app_users
  where id = p_student_user_id and role = 'student' and active = true;
  if student_matric is null then
    raise exception 'Active student account not found';
  end if;
  insert into public.battle_presence(student_user_id, "NO MATRIK", last_seen_at)
  values (p_student_user_id, student_matric, now())
  on conflict (student_user_id) do update set last_seen_at = excluded.last_seen_at;
end;
$$;

create or replace function public.send_battle_challenge(
  p_challenger_id uuid,
  p_opponent_id uuid
)
returns uuid
language plpgsql
security definer
set search_path = public
as $$
declare
  new_id uuid;
  challenger_matric text;
  opponent_matric text;
  key_value text;
  malaysia_today date := (now() at time zone 'Asia/Kuala_Lumpur')::date;
begin
  if p_challenger_id = p_opponent_id then
    raise exception 'You cannot challenge yourself';
  end if;
  select "NO MATRIK" into challenger_matric from public.app_users
  where id = p_challenger_id and role = 'student' and active = true;
  select "NO MATRIK" into opponent_matric from public.app_users
  where id = p_opponent_id and role = 'student' and active = true;
  if challenger_matric is null or opponent_matric is null then
    raise exception 'Both players must be active students';
  end if;
  if public.is_student_xp_day_blocked(challenger_matric, malaysia_today)
     or public.is_student_xp_day_blocked(opponent_matric, malaysia_today) then
    raise exception 'Both classes must be eligible for XP activity today';
  end if;
  key_value := least(p_challenger_id::text, p_opponent_id::text)
    || ':' || greatest(p_challenger_id::text, p_opponent_id::text);
  insert into public.battle_challenges(
    challenger_id, opponent_id, challenge_date, pair_key
  ) values (p_challenger_id, p_opponent_id, malaysia_today, key_value)
  returning id into new_id;
  return new_id;
exception when unique_violation then
  raise exception 'You can battle the same opponent only once per day';
end;
$$;

create or replace function public.respond_battle_challenge(
  p_challenge_id uuid,
  p_opponent_id uuid,
  p_accept boolean
)
returns uuid
language plpgsql
security definer
set search_path = public
as $$
declare
  challenge_row public.battle_challenges%rowtype;
  new_battle_id uuid;
  player_a_matric text;
  player_b_matric text;
  selected_count integer;
begin
  select * into challenge_row from public.battle_challenges
  where id = p_challenge_id for update;
  if challenge_row.id is null or challenge_row.opponent_id <> p_opponent_id
     or challenge_row.status <> 'pending' then
    raise exception 'This challenge is no longer available';
  end if;
  if not p_accept then
    update public.battle_challenges
    set status = 'rejected', responded_at = now()
    where id = p_challenge_id;
    return null;
  end if;
  select "NO MATRIK" into player_a_matric from public.app_users
  where id = challenge_row.challenger_id;
  select "NO MATRIK" into player_b_matric from public.app_users
  where id = challenge_row.opponent_id;
  if public.is_student_xp_day_blocked(player_a_matric, challenge_row.challenge_date)
     or public.is_student_xp_day_blocked(player_b_matric, challenge_row.challenge_date) then
    raise exception 'Both classes must be eligible for XP activity today';
  end if;
  insert into public.battle_matches(
    challenge_id, player_a_id, player_b_id, player_a_matric, player_b_matric
  ) values (
    challenge_row.id, challenge_row.challenger_id, challenge_row.opponent_id,
    player_a_matric, player_b_matric
  ) returning id into new_battle_id;

  insert into public.battle_questions(
    battle_id, position, question_id, difficulty, question, options,
    correct_index, explanation, opened_at
  )
  select new_battle_id, row_number() over (order by random()), id,
    difficulty, question, options, correct_index, explanation,
    case when row_number() over (order by random()) = 1 then now() else null end
  from (
    (select qq.* from public.quiz_questions qq
      join public.quizzes q on q.id = qq.quiz_id
      where q.status = 'published' and qq.review_status = 'approved'
        and qq.difficulty = 'easy' order by random() limit 7)
    union all
    (select qq.* from public.quiz_questions qq
      join public.quizzes q on q.id = qq.quiz_id
      where q.status = 'published' and qq.review_status = 'approved'
        and qq.difficulty = 'medium' order by random() limit 3)
  ) selected_questions;
  get diagnostics selected_count = row_count;
  if selected_count <> 10 then
    raise exception 'Published banks require at least 7 easy and 3 medium questions';
  end if;
  -- Ensure exactly position 1 owns the common one-minute start time.
  update public.battle_questions set opened_at = null
  where battle_id = new_battle_id and position > 1;
  update public.battle_questions set opened_at = now()
  where battle_id = new_battle_id and position = 1;
  update public.battle_challenges
  set status = 'accepted', responded_at = now(), battle_id = new_battle_id
  where id = challenge_row.id;
  return new_battle_id;
end;
$$;

create or replace function public.submit_battle_answer(
  p_battle_id uuid,
  p_student_user_id uuid,
  p_position integer,
  p_selected_index integer
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  battle public.battle_matches%rowtype;
  question_row public.battle_questions%rowtype;
  answer_a public.battle_answers%rowtype;
  answer_b public.battle_answers%rowtype;
  round_outcome text;
  round_winner uuid;
  final_winner uuid;
  points_a integer;
  points_b integer;
  completed_answers_a integer;
  completed_answers_b integer;
  event_a public.xp_events%rowtype;
  event_b public.xp_events%rowtype;
  elapsed_ms integer;
begin
  select * into battle from public.battle_matches
  where id = p_battle_id for update;
  if battle.id is null or battle.status <> 'active'
     or p_student_user_id not in (battle.player_a_id, battle.player_b_id)
     or battle.current_question <> p_position then
    raise exception 'This battle question is not active';
  end if;
  select * into question_row from public.battle_questions
  where battle_id = p_battle_id and position = p_position for update;
  elapsed_ms := least(60000, greatest(0,
    floor(extract(epoch from (clock_timestamp() - question_row.opened_at)) * 1000)::integer
  ));
  insert into public.battle_answers(
    battle_id, position, student_user_id, selected_index, is_correct,
    submitted_at, time_taken_ms, timed_out
  ) values (
    p_battle_id, p_position, p_student_user_id,
    case when elapsed_ms >= 60000 then null else p_selected_index end,
    elapsed_ms < 60000 and p_selected_index = question_row.correct_index,
    clock_timestamp(), elapsed_ms, elapsed_ms >= 60000
  ) on conflict (battle_id, position, student_user_id) do nothing;

  -- Once the minute ends, missing answers become timeouts so play can continue.
  if clock_timestamp() >= question_row.opened_at + interval '1 minute' then
    insert into public.battle_answers(
      battle_id, position, student_user_id, selected_index, is_correct,
      submitted_at, time_taken_ms, timed_out
    ) values
      (p_battle_id, p_position, battle.player_a_id, null, false,
       question_row.opened_at + interval '1 minute', 60000, true),
      (p_battle_id, p_position, battle.player_b_id, null, false,
       question_row.opened_at + interval '1 minute', 60000, true)
    on conflict (battle_id, position, student_user_id) do nothing;
  end if;

  select * into answer_a from public.battle_answers
  where battle_id = p_battle_id and position = p_position
    and student_user_id = battle.player_a_id;
  select * into answer_b from public.battle_answers
  where battle_id = p_battle_id and position = p_position
    and student_user_id = battle.player_b_id;
  if answer_a.student_user_id is null or answer_b.student_user_id is null then
    return jsonb_build_object('status', 'waiting');
  end if;

  if answer_a.is_correct and not answer_b.is_correct then
    round_outcome := 'player_a'; round_winner := battle.player_a_id;
  elsif answer_b.is_correct and not answer_a.is_correct then
    round_outcome := 'player_b'; round_winner := battle.player_b_id;
  elsif answer_a.is_correct and answer_b.is_correct
        and answer_a.submitted_at < answer_b.submitted_at then
    round_outcome := 'player_a'; round_winner := battle.player_a_id;
  elsif answer_a.is_correct and answer_b.is_correct
        and answer_b.submitted_at < answer_a.submitted_at then
    round_outcome := 'player_b'; round_winner := battle.player_b_id;
  else
    round_outcome := 'draw'; round_winner := null;
  end if;
  update public.battle_questions
  set outcome = round_outcome, winner_id = round_winner
  where battle_id = p_battle_id and position = p_position and outcome is null;
  if round_outcome = 'player_a' then
    update public.battle_matches set player_a_wins = player_a_wins + 1
    where id = p_battle_id;
  elsif round_outcome = 'player_b' then
    update public.battle_matches set player_b_wins = player_b_wins + 1
    where id = p_battle_id;
  else
    update public.battle_matches set draws = draws + 1 where id = p_battle_id;
  end if;

  if p_position < 10 then
    update public.battle_matches set current_question = p_position + 1
    where id = p_battle_id;
    update public.battle_questions set opened_at = clock_timestamp()
    where battle_id = p_battle_id and position = p_position + 1;
  else
    select * into battle from public.battle_matches where id = p_battle_id;
    if battle.player_a_wins > battle.player_b_wins then
      final_winner := battle.player_a_id; points_a := 10; points_b := 5;
    elsif battle.player_b_wins > battle.player_a_wins then
      final_winner := battle.player_b_id; points_a := 5; points_b := 10;
    else
      final_winner := null; points_a := 8; points_b := 8;
    end if;
    select count(*) filter (where selected_index is not null)
    into completed_answers_a
    from public.battle_answers
    where battle_id = p_battle_id and student_user_id = battle.player_a_id;
    select count(*) filter (where selected_index is not null)
    into completed_answers_b
    from public.battle_answers
    where battle_id = p_battle_id and student_user_id = battle.player_b_id;
    if completed_answers_a = 10 then
      insert into public.xp_events(
        "NO MATRIK", rule_code, points, source_id, reason, award_mode, created_at
      ) values (
        battle.player_a_matric, 'battle_result', points_a,
        'battle:' || p_battle_id::text, 'Completed two-player quiz battle',
        'automatic', clock_timestamp()
      ) returning * into event_a;
    else
      points_a := 0;
    end if;
    if completed_answers_b = 10 then
      insert into public.xp_events(
        "NO MATRIK", rule_code, points, source_id, reason, award_mode, created_at
      ) values (
        battle.player_b_matric, 'battle_result', points_b,
        'battle:' || p_battle_id::text, 'Completed two-player quiz battle',
        'automatic', clock_timestamp()
      ) returning * into event_b;
    else
      points_b := 0;
    end if;
    update public.battle_matches
    set status = 'completed', winner_id = final_winner,
        player_a_xp = coalesce(event_a.points, 0),
        player_b_xp = coalesce(event_b.points, 0),
        completed_at = clock_timestamp()
    where id = p_battle_id;
  end if;
  return jsonb_build_object('status', 'resolved', 'outcome', round_outcome);
end;
$$;

commit;
