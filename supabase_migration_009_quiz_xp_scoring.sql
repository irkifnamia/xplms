-- XPLMS migration 009
-- Daily quiz scoring: 5 XP per attempted answer plus 5 XP per correct answer.

update public.xp_rules
set default_points = 5,
    description = 'Daily chapter quiz: 5 XP per attempted answer plus 5 XP per correct answer.',
    updated_at = now()
where code = 'quiz_completion';
