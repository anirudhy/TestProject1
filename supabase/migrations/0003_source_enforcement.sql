-- A company_benefits row is worthless without a citation. Enforce it at the
-- database level (deferred to end-of-transaction, so callers can insert the
-- benefit row and its source row together, which is how both the seed
-- pipeline and edit-approval flow work) rather than trusting application code.
create or replace function require_benefit_source()
returns trigger language plpgsql as $$
declare
  source_count int;
begin
  select count(*) into source_count
  from benefit_sources
  where company_benefit_id = new.id;

  if source_count = 0 then
    raise exception 'company_benefits row % has no benefit_sources row — every benefit must be sourced', new.id
      using errcode = 'integrity_constraint_violation';
  end if;

  return null;
end;
$$;

create constraint trigger company_benefits_require_source
  after insert on company_benefits
  deferrable initially deferred
  for each row execute function require_benefit_source();

-- A company can only go live once every one of its benefit rows is sourced
-- (the trigger above already guarantees that per-row, so this mirrors the
-- same invariant at publish time in case rows were inserted before sourcing
-- was wired up, e.g. during a data migration).
create or replace function guard_company_publish()
returns trigger language plpgsql as $$
declare
  unsourced_count int;
begin
  if new.is_published and (tg_op = 'INSERT' or not old.is_published) then
    select count(*) into unsourced_count
    from company_benefits cb
    where cb.company_id = new.id
      and not exists (select 1 from benefit_sources bs where bs.company_benefit_id = cb.id);

    if unsourced_count > 0 then
      raise exception 'cannot publish company %: % benefit row(s) without a source', new.slug, unsourced_count
        using errcode = 'integrity_constraint_violation';
    end if;
  end if;
  return new;
end;
$$;

create trigger companies_guard_publish before insert or update on companies
  for each row execute function guard_company_publish();
