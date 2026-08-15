-- §9 open question #2: does company_benefits need a country dimension for
-- international plan variants? Yes in principle (a company's 401(k)-equivalent
-- differs by country), but nothing in the MVP populates non-US plans yet.
-- Add the column now so the schema doesn't need a breaking migration later,
-- default 'US', unused beyond that until international data exists.
alter table company_benefits add column country text not null default 'US';
