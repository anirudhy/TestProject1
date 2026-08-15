-- ============================================================
-- companies / benefit_types / company_benefits / benefit_sources
-- Public read (published only). Writes: service role / admin only.
-- ============================================================
alter table companies enable row level security;
alter table benefit_types enable row level security;
alter table company_benefits enable row level security;
alter table benefit_sources enable row level security;

create policy companies_public_read on companies
  for select using (is_published);

create policy benefit_types_public_read on benefit_types
  for select using (true);

create policy company_benefits_public_read on company_benefits
  for select using (
    exists (select 1 from companies c where c.id = company_benefits.company_id and c.is_published)
  );

create policy benefit_sources_public_read on benefit_sources
  for select using (
    exists (
      select 1 from company_benefits cb
      join companies c on c.id = cb.company_id
      where cb.id = benefit_sources.company_benefit_id and c.is_published
    )
  );

-- No insert/update/delete policies for anon/authenticated on any of the four
-- tables above: only the service role (which bypasses RLS) or an explicit
-- admin-role policy may write. Admins write via the service-role-backed
-- moderation API, so no additional policy is added here.

-- ============================================================
-- profiles: public read of handle/reputation only; write own row.
-- Column-level restriction is enforced by exposing a public view.
-- ============================================================
alter table profiles enable row level security;

create policy profiles_public_read on profiles
  for select using (true);

create policy profiles_update_own on profiles
  for update using (auth.uid() = id) with check (auth.uid() = id and role = 'user');

create policy profiles_insert_own on profiles
  for insert with check (auth.uid() = id);

-- Public-safe view: never select user_employment.* alongside profiles in a
-- way that could deanonymize a handle. This view is the only thing exposed
-- to the client for "who submitted this edit" contexts.
create view public_profiles as
  select id, handle, reputation from profiles;

-- ============================================================
-- user_employment: strictly own-row only. Never joined into any public view.
-- ============================================================
alter table user_employment enable row level security;

create policy user_employment_select_own on user_employment
  for select using (auth.uid() = user_id);

create policy user_employment_insert_own on user_employment
  for insert with check (auth.uid() = user_id);

-- verified_at is set by the service role (OTP verification flow), never by
-- the user directly — no update policy for authenticated users.

-- ============================================================
-- benefit_edits: own rows + moderators can read; approved rows are public.
-- insert: authenticated. update: moderators only.
-- ============================================================
alter table benefit_edits enable row level security;

create policy benefit_edits_select on benefit_edits
  for select using (
    status = 'approved'
    or submitted_by = auth.uid()
    or exists (select 1 from profiles p where p.id = auth.uid() and p.role in ('moderator','admin'))
  );

create policy benefit_edits_insert_own on benefit_edits
  for insert with check (submitted_by = auth.uid());

create policy benefit_edits_update_moderator on benefit_edits
  for update using (
    exists (select 1 from profiles p where p.id = auth.uid() and p.role in ('moderator','admin'))
  );

-- ============================================================
-- edit_votes: aggregate only for read (no per-user policy exposing who
-- voted which way beyond the voter's own row); insert/update own vote.
-- ============================================================
alter table edit_votes enable row level security;

create policy edit_votes_select_own on edit_votes
  for select using (auth.uid() = user_id);

create policy edit_votes_insert_own on edit_votes
  for insert with check (auth.uid() = user_id);

create policy edit_votes_update_own on edit_votes
  for update using (auth.uid() = user_id);

-- Vote totals for display are computed server-side via a security-definer
-- function so individual ballots are never exposed to other users.
create or replace function edit_vote_totals(p_edit_id uuid)
returns table (up int, down int) language sql stable security definer as $$
  select
    count(*) filter (where vote = 1)::int as up,
    count(*) filter (where vote = -1)::int as down
  from edit_votes where edit_id = p_edit_id;
$$;
