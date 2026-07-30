-- XPLMS migration 012
-- Adds Extra practice and enforces daily request quotas for request-based XP.
-- Review and run once in the Supabase SQL editor after migrations 001-011.

begin;

insert into public.xp_rules (
  code, name, description, award_mode, default_points, active
)
values
  (
    'study_group',
    'Study group',
    'Verified participation in a productive study group.',
    'manual',
    5,
    true
  ),
  (
    'extra_practice',
    'Extra practice',
    'Verified additional practice such as past-year questions or extra exercises.',
    'manual',
    15,
    true
  )
on conflict (code) do update
set name = excluded.name,
    description = excluded.description,
    award_mode = excluded.award_mode,
    default_points = excluded.default_points,
    active = excluded.active,
    updated_at = now();

alter table public.xp_claims
  drop constraint if exists xp_claims_claim_type_check;

alter table public.xp_claims
  add constraint xp_claims_claim_type_check
  check (
    claim_type in (
      'consultation',
      'class_participation',
      'commitment',
      'study_group',
      'extra_practice'
    )
  );

create index if not exists xp_claims_daily_quota_idx
  on public.xp_claims ("NO MATRIK", claim_type, created_at);

create or replace function public.enforce_xp_claim_daily_quota()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  malaysia_date date;
  request_count integer;
  quota_key text;
begin
  if new.claim_type not in ('study_group', 'extra_practice') then
    return new;
  end if;

  malaysia_date :=
    (coalesce(new.created_at, now()) at time zone 'Asia/Kuala_Lumpur')::date;
  quota_key := new."NO MATRIK" || '|' || new.claim_type || '|' || malaysia_date;

  -- Serialises simultaneous submissions for the same student/type/day.
  perform pg_advisory_xact_lock(hashtextextended(quota_key, 0));

  select count(*)
  into request_count
  from public.xp_claims
  where "NO MATRIK" = new."NO MATRIK"
    and claim_type = new.claim_type
    and (created_at at time zone 'Asia/Kuala_Lumpur')::date = malaysia_date
    and (tg_op = 'INSERT' or id <> new.id);

  if request_count >= 2 then
    raise exception
      'Daily quota reached: maximum 2 % requests per day.',
      replace(new.claim_type, '_', ' ')
      using errcode = 'P0001';
  end if;

  return new;
end;
$$;

drop trigger if exists enforce_xp_claim_daily_quota
  on public.xp_claims;
create trigger enforce_xp_claim_daily_quota
before insert or update of "NO MATRIK", claim_type, created_at
on public.xp_claims
for each row execute function public.enforce_xp_claim_daily_quota();

commit;
