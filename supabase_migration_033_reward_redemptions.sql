begin;

create table if not exists public.reward_redemptions (
  reward_key text primary key,
  reward_section text not null,
  reward_name text not null,
  achiever text,
  redeemed_at timestamptz not null default clock_timestamp(),
  redeemed_by uuid references public.app_users(id) on delete set null
);

create index if not exists reward_redemptions_redeemed_at_idx
  on public.reward_redemptions (redeemed_at desc);

alter table public.reward_redemptions enable row level security;
revoke all on table public.reward_redemptions from anon, authenticated;
grant all on table public.reward_redemptions to service_role;

commit;

notify pgrst, 'reload schema';
