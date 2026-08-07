-- XPLMS migration 031
-- Repair optional quiz-report attachment metadata and reload PostgREST schema.
-- Safe to run if migration 021 already added these columns.

begin;

alter table public.quiz_question_reports
  add column if not exists supporting_file_path text,
  add column if not exists supporting_file_name text,
  add column if not exists supporting_file_type text;

commit;

-- Make the new columns immediately visible to the Supabase Data API.
notify pgrst, 'reload schema';
