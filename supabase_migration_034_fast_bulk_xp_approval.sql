begin;

-- Process a selected group of XP requests inside one database round trip.
-- Each item remains isolated: a bad request is counted as failed without
-- rolling back valid requests from the same Admin action.
create or replace function public.approve_xp_claims_bulk(
  p_items jsonb,
  p_admin_id uuid
)
returns table (approved_count integer, failed_count integer)
language plpgsql
security definer
set search_path = public
as $$
declare
  item jsonb;
  claim_row public.xp_claims%rowtype;
  event_id bigint;
  requested_points integer;
  requested_note text;
begin
  approved_count := 0;
  failed_count := 0;

  for item in select value from jsonb_array_elements(coalesce(p_items, '[]'::jsonb))
  loop
    begin
      requested_points := (item ->> 'points')::integer;
      requested_note := coalesce(item ->> 'admin_note', 'Batch approved by Admin');
      if requested_points = 0 then
        raise exception 'XP award cannot be zero';
      end if;

      select * into strict claim_row
      from public.xp_claims
      where id = (item ->> 'claim_id')::bigint
        and status = 'pending'
      for update;

      insert into public.xp_events (
        "NO MATRIK", rule_code, points, source_id, reason,
        award_mode, created_at, awarded_by
      ) values (
        claim_row."NO MATRIK", claim_row.claim_type, requested_points,
        'claim-' || claim_row.id::text, claim_row.title,
        'manual', claim_row.created_at, p_admin_id
      ) returning id into event_id;

      update public.xp_claims
      set status = 'approved',
          admin_note = requested_note,
          reviewed_by = p_admin_id,
          reviewed_at = now(),
          xp_event_id = event_id
      where id = claim_row.id;

      approved_count := approved_count + 1;
    exception when others then
      failed_count := failed_count + 1;
    end;
  end loop;

  return next;
end;
$$;

revoke all on function public.approve_xp_claims_bulk(jsonb, uuid) from public;
grant execute on function public.approve_xp_claims_bulk(jsonb, uuid) to service_role;

commit;
