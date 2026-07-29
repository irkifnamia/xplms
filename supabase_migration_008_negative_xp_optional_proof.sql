-- XPLMS migration 008
-- Allows deductions in the XP ledger and makes student proof images optional.

alter table public.xp_claims
  alter column proof_path drop not null;

-- xp_events already permits any non-zero points. Remove any legacy balance
-- constraint that prevents stud_xp.XP from becoming negative.
do $$
declare
  constraint_name text;
begin
  for constraint_name in
    select con.conname
    from pg_constraint con
    join pg_class rel on rel.oid = con.conrelid
    join pg_namespace nsp on nsp.oid = rel.relnamespace
    where nsp.nspname = 'public'
      and rel.relname = 'stud_xp'
      and con.contype = 'c'
      and pg_get_constraintdef(con.oid) ilike '%"XP"%'
  loop
    execute format(
      'alter table public.stud_xp drop constraint %I',
      constraint_name
    );
  end loop;
end
$$;
