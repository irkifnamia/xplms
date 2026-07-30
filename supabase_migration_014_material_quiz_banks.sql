-- XPLMS migration 014
-- Material usage controls and balanced, non-repeating daily chapter quizzes.

begin;

alter table public.materials
  add column if not exists student_visible boolean not null default true,
  add column if not exists quiz_source boolean not null default false;

comment on column public.materials.student_visible is
  'When true, students can view and download this material.';
comment on column public.materials.quiz_source is
  'When true, an admin may use this material as an AI quiz source.';

alter table public.quizzes
  add column if not exists chapter integer,
  add column if not exists source_material_ids bigint[] not null default '{}',
  add column if not exists target_question_count integer not null default 200;

update public.quizzes q
set chapter = m.chapter
from public.materials m
where q.material_id = m.id
  and q.chapter is null;

alter table public.quizzes
  drop constraint if exists quizzes_chapter_check,
  add constraint quizzes_chapter_check
    check (chapter is null or chapter in (1, 2, 5, 8, 9, 10)),
  drop constraint if exists quizzes_target_question_count_check,
  add constraint quizzes_target_question_count_check
    check (target_question_count > 0);

alter table public.quiz_questions
  add column if not exists difficulty text;

update public.quiz_questions
set difficulty = 'medium'
where difficulty is null;

alter table public.quiz_questions
  alter column difficulty set not null,
  drop constraint if exists quiz_questions_difficulty_check,
  add constraint quiz_questions_difficulty_check
    check (difficulty in ('easy', 'medium', 'hard'));

create index if not exists materials_usage_chapter_idx
  on public.materials (quiz_source, student_visible, chapter, uploaded_at desc);

create index if not exists quizzes_chapter_status_idx
  on public.quizzes (chapter, status, created_at desc);

create index if not exists quiz_questions_bank_difficulty_idx
  on public.quiz_questions (quiz_id, difficulty, id);

create index if not exists quiz_attempts_student_chapter_history_idx
  on public.quiz_attempts (student_user_id, chapter, attempt_date desc);

update public.xp_rules
set default_points = 1,
    description =
      'Daily chapter quiz: 1 XP per answered question plus 1 additional XP per correct answer.',
    updated_at = now()
where code = 'quiz_completion';

commit;
