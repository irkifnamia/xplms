-- XPLMS migration 027: one challenge at a time and controlled battle leaving.
-- Run after migration 026.

begin;

-- Pending challenges are transient. Clear old multi-challenge state before
-- enforcing the new one-challenge-per-student rule.
update public.battle_challenges
set status = 'expired', responded_at = now()
where status = 'pending';

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
  -- Consistent locks prevent two simultaneous requests bypassing the checks.
  perform pg_advisory_xact_lock(
    hashtextextended(least(p_challenger_id::text, p_opponent_id::text), 0)
  );
  perform pg_advisory_xact_lock(
    hashtextextended(greatest(p_challenger_id::text, p_opponent_id::text), 0)
  );
  select "NO MATRIK" into challenger_matric from public.app_users
  where id = p_challenger_id and role = 'student' and active = true;
  select "NO MATRIK" into opponent_matric from public.app_users
  where id = p_opponent_id and role = 'student' and active = true;
  if challenger_matric is null or opponent_matric is null then
    raise exception 'Both players must be active students';
  end if;
  if exists (
    select 1 from public.battle_challenges
    where status = 'pending'
      and (
        challenger_id in (p_challenger_id, p_opponent_id)
        or opponent_id in (p_challenger_id, p_opponent_id)
      )
  ) then
    raise exception 'A player already has a pending challenge';
  end if;
  if exists (
    select 1 from public.battle_matches
    where status = 'active'
      and (
        player_a_id in (p_challenger_id, p_opponent_id)
        or player_b_id in (p_challenger_id, p_opponent_id)
      )
  ) then
    raise exception 'A player is already in an active battle';
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

create or replace function public.leave_battle(
  p_battle_id uuid,
  p_student_user_id uuid
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  battle public.battle_matches%rowtype;
  opponent_id uuid;
  opponent_matric text;
  opponent_wins integer;
  winner_event public.xp_events%rowtype;
begin
  select * into battle from public.battle_matches
  where id = p_battle_id for update;
  if battle.id is null or battle.status <> 'active'
     or p_student_user_id not in (battle.player_a_id, battle.player_b_id) then
    raise exception 'This active battle cannot be left';
  end if;
  if p_student_user_id = battle.player_a_id then
    opponent_id := battle.player_b_id;
    opponent_matric := battle.player_b_matric;
    opponent_wins := battle.player_b_wins;
  else
    opponent_id := battle.player_a_id;
    opponent_matric := battle.player_a_matric;
    opponent_wins := battle.player_a_wins;
  end if;

  if opponent_wins >= 6 then
    insert into public.xp_events(
      "NO MATRIK", rule_code, points, source_id, reason, award_mode, created_at
    ) values (
      opponent_matric, 'battle_result', 10,
      'battle:' || p_battle_id::text,
      'Battle won after opponent left with 6 or more question wins',
      'automatic', clock_timestamp()
    ) returning * into winner_event;
    update public.battle_matches
    set status = 'completed', winner_id = opponent_id,
        player_a_xp = case
          when opponent_id = player_a_id then winner_event.points else 0 end,
        player_b_xp = case
          when opponent_id = player_b_id then winner_event.points else 0 end,
        completed_at = clock_timestamp()
    where id = p_battle_id;
    return jsonb_build_object(
      'status', 'completed', 'winner_id', opponent_id,
      'winner_xp', winner_event.points
    );
  end if;

  update public.battle_matches
  set status = 'cancelled', winner_id = null,
      player_a_xp = 0, player_b_xp = 0,
      completed_at = clock_timestamp()
  where id = p_battle_id;
  return jsonb_build_object('status', 'cancelled', 'winner_xp', 0);
end;
$$;

commit;
