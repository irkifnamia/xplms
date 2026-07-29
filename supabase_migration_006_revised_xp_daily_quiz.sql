-- XPLMS migration 006: revised XP catalogue and daily chapter quizzes

begin;

-- Keep only the five current XP types active.
update public.xp_rules
set active = false, updated_at = now()
where code not in (
  'consultation',
  'class_participation',
  'commitment',
  'quiz_completion',
  'study_group'
);

insert into public.xp_rules (
  code, name, description, award_mode, default_points, active
)
values
  (
    'consultation',
    'Consultation',
    'Constructive academic consultation.',
    'manual',
    20,
    true
  ),
  (
    'class_participation',
    'Class participation',
    'Meaningful contribution during class.',
    'manual',
    10,
    true
  ),
  (
    'commitment',
    'Commitment',
    'Consistent commitment to learning and improvement.',
    'manual',
    10,
    true
  ),
  (
    'study_group',
    'Study group',
    'Verified participation in a productive study group.',
    'manual',
    15,
    true
  ),
  (
    'quiz_completion',
    'In-app quiz',
    'Daily chapter quiz: 1 XP per attempt or 2 XP for a correct answer.',
    'automatic',
    1,
    true
  )
on conflict (code) do update
set name = excluded.name,
    description = excluded.description,
    award_mode = excluded.award_mode,
    default_points = excluded.default_points,
    active = excluded.active,
    updated_at = now();

-- Convert legacy mathematics-discussion requests to Study group requests.
update public.xp_claims
set claim_type = 'study_group'
where claim_type = 'math_discussion';

alter table public.xp_claims
  drop constraint if exists xp_claims_claim_type_check;

alter table public.xp_claims
  add constraint xp_claims_claim_type_check
  check (
    claim_type in (
      'consultation',
      'class_participation',
      'commitment',
      'study_group'
    )
  );

-- Existing attempts remain as legacy records. New attempts are uniquely
-- identified by student, chapter and Malaysia calendar date.
alter table public.quiz_attempts
  drop constraint if exists quiz_attempts_quiz_id_student_user_id_key;

alter table public.quiz_attempts
  add column if not exists chapter integer,
  add column if not exists attempt_date date not null default current_date,
  add column if not exists xp_awarded integer not null default 0;

alter table public.quiz_attempts
  drop constraint if exists quiz_attempts_chapter_check,
  add constraint quiz_attempts_chapter_check
    check (chapter is null or chapter in (1, 2, 5, 8, 9, 10)),
  drop constraint if exists quiz_attempts_xp_awarded_check,
  add constraint quiz_attempts_xp_awarded_check
    check (xp_awarded >= 0);

create unique index if not exists quiz_attempts_student_chapter_day_uidx
  on public.quiz_attempts (student_user_id, chapter, attempt_date)
  where chapter is not null;

commit;
