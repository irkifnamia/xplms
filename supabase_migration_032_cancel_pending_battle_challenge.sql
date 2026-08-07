begin;

create or replace function public.cancel_battle_challenge(
  p_challenge_id uuid,
  p_challenger_id uuid
)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  affected_rows integer;
begin
  update public.battle_challenges
  set status = 'expired', responded_at = clock_timestamp()
  where id = p_challenge_id
    and challenger_id = p_challenger_id
    and status = 'pending';

  get diagnostics affected_rows = row_count;
  if affected_rows <> 1 then
    raise exception 'This pending challenge cannot be cancelled';
  end if;
end;
$$;

revoke all on function public.cancel_battle_challenge(uuid, uuid)
from public;
grant execute on function public.cancel_battle_challenge(uuid, uuid)
to anon, authenticated, service_role;

commit;
