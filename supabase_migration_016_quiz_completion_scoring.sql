-- XPLMS migration 016
-- Daily quiz scoring: 5 XP for completing the full 10-question set,
-- plus 1 XP for each correct answer (maximum 15 XP per set).

begin;

update public.xp_rules
set default_points = 5,
    description =
      'Daily chapter quiz: 5 XP for completing all 10 questions plus 1 XP per correct answer (maximum 15 XP per set).',
    updated_at = now()
where code = 'quiz_completion';

commit;
