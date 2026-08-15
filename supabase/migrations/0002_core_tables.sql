-- ============ Companies ============
create table companies (
  id              uuid primary key default gen_random_uuid(),
  slug            text unique not null,
  name            text not null,
  legal_name      text,
  logo_url        text,
  ticker          text,
  hq_country      text default 'US',
  employee_band   text check (employee_band in ('1-50','51-500','501-5000','5001-50000','50000+')),
  industry        text,
  careers_url     text,
  benefits_url    text,
  levels_fyi_url  text,
  email_domains   text[] not null default '{}',   -- for employee verification
  is_published    boolean not null default false,
  created_at      timestamptz default now(),
  updated_at      timestamptz default now()
);
create index companies_name_trgm_idx on companies using gin (name gin_trgm_ops);
create index companies_slug_idx on companies (slug);

-- ============ Benefit type registry ============
-- Seeded by migration/seed script, not user-editable. This IS the product's schema.
create table benefit_types (
  key             text primary key,               -- '401k_match', 'mega_backdoor_roth'
  category        text not null,
  label           text not null,
  description     text,
  value_schema    jsonb not null,                 -- JSON Schema for the value blob
  is_comparable   boolean not null default true,  -- appears in compare table?
  is_countable    boolean not null default false, -- feeds the $ calculator?
  sort_order      int not null default 100
);

-- ============ The actual data ============
create table company_benefits (
  id              uuid primary key default gen_random_uuid(),
  company_id      uuid not null references companies(id) on delete cascade,
  benefit_key     text not null references benefit_types(key),
  plan_year       int not null,
  value           jsonb not null,                 -- validated against value_schema
  notes           text check (char_length(notes) <= 280),
  confidence      text not null default 'unverified'
                  check (confidence in ('official','corroborated','single_report','unverified')),
  last_verified_at timestamptz,
  created_at      timestamptz default now(),
  updated_at      timestamptz default now(),
  unique (company_id, benefit_key, plan_year)
);
create index company_benefits_company_year_idx on company_benefits (company_id, plan_year);
create index company_benefits_value_gin_idx on company_benefits using gin (value);

-- ============ Sourcing — non-negotiable ============
create table benefit_sources (
  id                  uuid primary key default gen_random_uuid(),
  company_benefit_id  uuid not null references company_benefits(id) on delete cascade,
  source_type         text not null check (source_type in
                        ('company_public_page','sec_filing','form_5500','press','news','user_report')),
  url                 text,
  title               text,
  retrieved_at        timestamptz not null default now(),
  excerpt             text check (char_length(excerpt) <= 200)
);
create index benefit_sources_company_benefit_idx on benefit_sources (company_benefit_id);

-- updated_at maintenance
create or replace function set_updated_at()
returns trigger language plpgsql as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

create trigger companies_set_updated_at before update on companies
  for each row execute function set_updated_at();

create trigger company_benefits_set_updated_at before update on company_benefits
  for each row execute function set_updated_at();
