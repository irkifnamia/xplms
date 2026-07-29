-- XPLMS migration 003: merge lecturer access into administrator
-- After this migration, XPLMS has only two roles: admin and student.

begin;

alter table public.app_users
  drop constraint if exists app_users_role_check;
alter table public.app_users
  drop constraint if exists app_users_student_matric_check;

update public.app_users
set role = 'admin',
    updated_at = now()
where role = 'lecturer';

alter table public.app_users
  add constraint app_users_role_check
  check (role in ('student','admin'));

alter table public.app_users
  add constraint app_users_student_matric_check
  check (
    (role = 'student' and "NO MATRIK" is not null)
    or (role = 'admin' and "NO MATRIK" is null)
  );

commit;
