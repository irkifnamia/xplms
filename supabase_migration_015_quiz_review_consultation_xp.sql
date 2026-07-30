-- XPLMS migration 015
-- Consultation XP update and mandatory review before quiz publication.

begin;

update public.xp_rules
set default_points = 15,
    updated_at = now()
where code = 'consultation';

alter table public.quizzes
  add column if not exists reviewed_at timestamptz,
  add column if not exists reviewed_by uuid
    references public.app_users(id) on delete set null;

alter table public.quiz_questions
  add column if not exists review_status text not null default 'pending',
  add column if not exists reviewed_at timestamptz,
  add column if not exists reviewed_by uuid
    references public.app_users(id) on delete set null;

alter table public.quiz_questions
  drop constraint if exists quiz_questions_review_status_check,
  add constraint quiz_questions_review_status_check
    check (review_status in ('pending', 'approved'));

-- Preserve already-published legacy banks.
update public.quiz_questions qq
set review_status = 'approved',
    reviewed_at = coalesce(qq.reviewed_at, now())
from public.quizzes q
where qq.quiz_id = q.id
  and q.status = 'published';

update public.quizzes
set reviewed_at = coalesce(reviewed_at, now())
where status = 'published';

create or replace function public.validate_quiz_bank_publication()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  easy_count integer;
  medium_count integer;
  hard_count integer;
  approved_count integer;
begin
  if new.status = 'published'
     and (
       tg_op = 'INSERT'
       or old.status is distinct from 'published'
     ) then
    if new.reviewed_at is null then
      raise exception 'Question bank must be reviewed before publication';
    end if;

    select
      count(*) filter (
        where review_status = 'approved' and difficulty = 'easy'
      ),
      count(*) filter (
        where review_status = 'approved' and difficulty = 'medium'
      ),
      count(*) filter (
        where review_status = 'approved' and difficulty = 'hard'
      ),
      count(*) filter (where review_status = 'approved')
    into easy_count, medium_count, hard_count, approved_count
    from public.quiz_questions
    where quiz_id = new.id;

    if approved_count <> 200
       or easy_count <> 40
       or medium_count <> 100
       or hard_count <> 60 then
      raise exception
        'Published bank requires 200 approved questions: 40 easy, 100 medium and 60 hard';
    end if;
  end if;
  return new;
end;
$$;

drop trigger if exists validate_quiz_bank_before_publish
  on public.quizzes;

create trigger validate_quiz_bank_before_publish
before insert or update of status on public.quizzes
for each row
execute function public.validate_quiz_bank_publication();

commit;
