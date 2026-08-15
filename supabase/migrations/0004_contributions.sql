create table profiles (
  id            uuid primary key references auth.users(id) on delete cascade,
  handle        text unique not null,            -- pseudonymous, auto-generated
  role          text not null default 'user' check (role in ('user','trusted','moderator','admin')),
  reputation    int not null default 0,
  created_at    timestamptz default now()
);

-- Employment verification. NEVER exposed publicly with identity.
-- Per §6 privacy: only a hash of the domain match + timestamp is retained,
-- never the work email address itself.
create table user_employment (
  id              uuid primary key default gen_random_uuid(),
  user_id         uuid not null references profiles(id) on delete cascade,
  company_id      uuid not null references companies(id),
  verified_at     timestamptz,
  verification_method text default 'work_email_otp',
  is_public_badge boolean not null default false,
  unique (user_id, company_id)
);

create table benefit_edits (
  id              uuid primary key default gen_random_uuid(),
  company_id      uuid not null references companies(id),
  benefit_key     text not null references benefit_types(key),
  plan_year       int not null,
  proposed_value  jsonb not null,
  current_value   jsonb,                          -- snapshot at submission time
  source_type     text not null check (source_type in
                    ('company_public_page','sec_filing','form_5500','press','news','user_report')),
  source_url      text,
  rationale       text,
  submitted_by    uuid references profiles(id),
  submitter_is_verified_employee boolean default false,
  status          text not null default 'pending'
                  check (status in ('pending','approved','rejected','superseded')),
  reviewed_by     uuid references profiles(id),
  review_note     text,
  created_at      timestamptz default now()
);
create index benefit_edits_status_idx on benefit_edits (status, created_at);
create index benefit_edits_company_idx on benefit_edits (company_id);

create table edit_votes (
  edit_id  uuid references benefit_edits(id) on delete cascade,
  user_id  uuid references profiles(id) on delete cascade,
  vote     smallint not null check (vote in (-1, 1)),
  primary key (edit_id, user_id)
);
