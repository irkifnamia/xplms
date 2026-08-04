-- XPLMS migration 020
-- Adds Quiz correction XP with a student-proposed amount and Admin override.
-- Run once in Supabase SQL Editor after migration 019.

begin;

alter table public.xp_rules
  alter column default_points drop not null;

alter table public.xp_rules
  drop constraint if exists xp_rules_default_points_check;

alter table public.xp_rules
  add constraint xp_rules_default_points_check
  check (
    (code = 'quiz_correction' and default_points is null)
    or (code <> 'quiz_correction' and default_points > 0)
  );

insert into public.xp_rules (
  code, name, description, award_mode, default_points, active
)
values (
  'quiz_correction',
  'Quiz correction',
  'Correction work completed after reviewing an assessment or quiz.',
  'manual',
  null,
  true
)
on conflict (code) do update
set name = excluded.name,
    description = excluded.description,
    award_mode = excluded.award_mode,
    default_points = null,
    active = true,
    updated_at = now();

alter table public.xp_claims
  add column if not exists requested_points integer;

alter table public.xp_claims
  drop constraint if exists xp_claims_requested_points_check;

alter table public.xp_claims
  add constraint xp_claims_requested_points_check
  check (
    (claim_type = 'quiz_correction' and requested_points between 1 and 1000)
    or (claim_type <> 'quiz_correction' and requested_points is null)
  );

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
      'extra_practice',
      'quiz_correction'
    )
  );

commit;
