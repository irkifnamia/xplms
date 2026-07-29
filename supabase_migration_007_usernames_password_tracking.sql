-- XPLMS migration 007
-- Adds app-managed usernames and a safe audit timestamp for password changes.
-- Admin usernames remain their email address. Student usernames default to NO MATRIK.

create extension if not exists pgcrypto;

alter table public.app_users
  add column if not exists username text,
  add column if not exists password_changed_at timestamptz;

update public.app_users
set username = case
  when lower(role) = 'student' then "NO MATRIK"
  else lower(email)
end
where username is null or btrim(username) = '';

alter table public.app_users
  alter column username set not null;

create unique index if not exists app_users_username_ci_unique
  on public.app_users (lower(username));

-- Existing student accounts receive the requested initial credential:
-- username = NO MATRIK and password = NO MATRIK.
-- pgcrypto is already used by earlier XPLMS migrations.
update public.app_users
set password_hash = crypt("NO MATRIK", gen_salt('bf', 12)),
    must_change_password = false,
    password_changed_at = null
where lower(role) = 'student'
  and "NO MATRIK" is not null;

comment on column public.app_users.username is
  'Application login name: admin email or student NO MATRIK by default.';
comment on column public.app_users.password_changed_at is
  'Audit timestamp only. Plain-text passwords are never stored.';
