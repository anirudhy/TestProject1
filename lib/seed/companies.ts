import type { ConfidenceLevel, SourceType } from "@/types";

/**
 * DEMO / DEVELOPMENT SEED DATA ONLY.
 *
 * These are fictional companies (note the `.example` TLD — reserved by
 * RFC 2606 for exactly this purpose: it never resolves and can't be
 * confused with a real domain). They exist to exercise the full pipeline —
 * schema, the source-required trigger, compare, leaderboards, the
 * calculator — end to end in dev without asserting unverified facts about
 * real companies.
 *
 * Populating this file with real companies is Phase 1's actual seeding
 * work per §5 of the build spec: every row needs a genuine
 * company-public-page / Form 5500 / SEC-filing source, researched and
 * cited individually. Do not swap in a real company's slug/name here
 * without doing that research — the whole product's credibility rests on
 * every number being real and sourced.
 */

export interface SeedSource {
  source_type: SourceType;
  url: string;
  title: string;
  retrieved_at: string;
  excerpt?: string;
}

export interface SeedBenefit {
  benefit_key: string;
  plan_year: number;
  value: Record<string, unknown>;
  notes?: string;
  confidence: ConfidenceLevel;
  last_verified_at: string;
  sources: SeedSource[];
}

export interface SeedCompany {
  slug: string;
  name: string;
  legal_name: string;
  logo_url: string | null;
  ticker: string | null;
  hq_country: string;
  employee_band: "1-50" | "51-500" | "501-5000" | "5001-50000" | "50000+";
  industry: string;
  careers_url: string;
  benefits_url: string;
  levels_fyi_url: string | null;
  email_domains: string[];
  is_published: boolean;
  benefits: SeedBenefit[];
}

function source(overrides: Partial<SeedSource> & Pick<SeedSource, "url" | "title">): SeedSource {
  return {
    source_type: "company_public_page",
    retrieved_at: "2026-07-01T00:00:00.000Z",
    ...overrides,
  };
}

export const SEED_COMPANIES: SeedCompany[] = [
  {
    slug: "northwind-robotics",
    name: "Northwind Robotics",
    legal_name: "Northwind Robotics, Inc.",
    logo_url: null,
    ticker: "NWRO",
    hq_country: "US",
    employee_band: "50000+",
    industry: "Technology",
    careers_url: "https://careers.northwindrobotics.example",
    benefits_url: "https://careers.northwindrobotics.example/benefits",
    levels_fyi_url: null,
    email_domains: ["northwindrobotics.example"],
    is_published: true,
    benefits: [
      {
        benefit_key: "401k_match",
        plan_year: 2026,
        value: {
          formula_type: "tiered",
          cap_type: "percent_of_irs_limit",
          cap_value: 50,
          tiers: [{ up_to_percent_of_salary: 6, match_percent: 100 }],
          vesting_type: "immediate",
          vesting_years: 0,
          true_up: true,
          eligibility_days: 0,
        },
        notes: "100% match up to 6% of pay, capped at 50% of the IRS elective deferral limit. True-up paid each January.",
        confidence: "official",
        last_verified_at: "2026-07-01T00:00:00.000Z",
        sources: [
          source({
            url: "https://careers.northwindrobotics.example/benefits/401k",
            title: "Northwind Robotics — 401(k) plan summary",
            excerpt: "Northwind matches 100% of the first 6% you contribute, with an annual true-up.",
          }),
        ],
      },
      {
        benefit_key: "mega_backdoor_roth",
        plan_year: 2026,
        value: {
          supported: true,
          after_tax_contributions_allowed: true,
          after_tax_cap_percent_of_salary: 12,
          conversion_method: "automatic_daily",
          conversion_fee_usd: 0,
          recordkeeper: "Fidelity",
          headroom_note: "Total 415(c) limit minus employee deferral minus employer match.",
        },
        notes: "In-plan Roth conversions run automatically overnight via Fidelity NetBenefits.",
        confidence: "official",
        last_verified_at: "2026-07-01T00:00:00.000Z",
        sources: [
          source({
            source_type: "form_5500",
            url: "https://www.efast.dol.gov/5500search/",
            title: "Northwind Robotics 401(k) Plan — Form 5500 (plan year 2025)",
            excerpt: "Plan permits after-tax employee contributions with in-plan Roth conversion feature.",
          }),
        ],
      },
      {
        benefit_key: "espp",
        plan_year: 2026,
        value: {
          offered: true,
          discount_percent: 15,
          lookback: true,
          offering_period_months: 6,
          contribution_cap_percent_of_salary: 10,
          holding_period_months: 0,
        },
        confidence: "official",
        last_verified_at: "2026-07-01T00:00:00.000Z",
        sources: [
          source({
            source_type: "sec_filing",
            url: "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany",
            title: "Northwind Robotics S-8 — Employee Stock Purchase Plan",
            excerpt: "15% discount with a 6-month lookback offering period.",
          }),
        ],
      },
      {
        benefit_key: "hsa_contribution",
        plan_year: 2026,
        value: { offered: true, annual_employee_only_usd: 750, annual_family_usd: 1500, requires_hdhp_enrollment: true },
        confidence: "official",
        last_verified_at: "2026-07-01T00:00:00.000Z",
        sources: [source({ url: "https://careers.northwindrobotics.example/benefits/health", title: "Northwind Robotics — Health plans" })],
      },
      {
        benefit_key: "parental_leave",
        plan_year: 2026,
        value: { birthing_parent_weeks_paid: 20, non_birthing_parent_weeks_paid: 16, pay_percent: 100, covers_adoption_foster: true, tenure_required_months: 0 },
        confidence: "official",
        last_verified_at: "2026-07-01T00:00:00.000Z",
        sources: [source({ url: "https://careers.northwindrobotics.example/benefits/family", title: "Northwind Robotics — Family leave policy" })],
      },
      {
        benefit_key: "fertility_benefit",
        plan_year: 2026,
        value: { offered: true, lifetime_max_usd: 40000, covers_ivf: true, covers_egg_freezing: true, covers_surrogacy: true, adoption_assistance_usd: 10000, provider: "Carrot" },
        confidence: "official",
        last_verified_at: "2026-07-01T00:00:00.000Z",
        sources: [source({ url: "https://careers.northwindrobotics.example/benefits/family", title: "Northwind Robotics — Family-forming benefits" })],
      },
      {
        benefit_key: "pto_policy",
        plan_year: 2026,
        value: { policy_type: "unlimited", rollover_allowed: false, paid_holidays: 11 },
        confidence: "official",
        last_verified_at: "2026-07-01T00:00:00.000Z",
        sources: [source({ url: "https://careers.northwindrobotics.example/benefits/time-off", title: "Northwind Robotics — Time off" })],
      },
      {
        benefit_key: "education_reimbursement",
        plan_year: 2026,
        value: { offered: true, annual_cap_usd: 5250, covers_degree_programs: true, requires_manager_approval: true },
        confidence: "official",
        last_verified_at: "2026-07-01T00:00:00.000Z",
        sources: [source({ url: "https://careers.northwindrobotics.example/benefits/learning", title: "Northwind Robotics — Learning & development" })],
      },
      {
        benefit_key: "lifestyle_stipend",
        plan_year: 2026,
        value: { offered: true, annual_usd: 1200, cadence: "monthly", eligible_spend: "Gym, wellness apps, home office" },
        confidence: "official",
        last_verified_at: "2026-07-01T00:00:00.000Z",
        sources: [source({ url: "https://careers.northwindrobotics.example/benefits/wellness", title: "Northwind Robotics — Wellness stipend" })],
      },
      {
        benefit_key: "perk_misc",
        plan_year: 2026,
        value: { name: "On-site robotics lab access", description: "After-hours access to the campus fabrication lab for personal projects.", est_annual_value_usd: 0, conditions: "HQ campus only" },
        confidence: "corroborated",
        last_verified_at: "2026-06-15T00:00:00.000Z",
        sources: [source({ source_type: "user_report", url: "", title: "Employee report, corroborated by two submissions", excerpt: "Confirmed by two independent verified-employee reports." })],
      },
    ],
  },
  {
    slug: "vantage-cloud",
    name: "Vantage Cloud",
    legal_name: "Vantage Cloud Corporation",
    logo_url: null,
    ticker: "VNTG",
    hq_country: "US",
    employee_band: "5001-50000",
    industry: "Technology",
    careers_url: "https://jobs.vantagecloud.example",
    benefits_url: "https://jobs.vantagecloud.example/benefits",
    levels_fyi_url: null,
    email_domains: ["vantagecloud.example"],
    is_published: true,
    benefits: [
      {
        benefit_key: "401k_match",
        plan_year: 2026,
        value: {
          formula_type: "percent_of_employee_contribution",
          match_percent: 50,
          cap_type: "percent_of_salary",
          cap_value: 6,
          vesting_type: "graded",
          vesting_years: 4,
          true_up: false,
          eligibility_days: 90,
        },
        notes: "50% match on the first 6% of pay. Vests 25%/year over 4 years, no true-up.",
        confidence: "official",
        last_verified_at: "2026-05-20T00:00:00.000Z",
        sources: [source({ url: "https://jobs.vantagecloud.example/benefits/retirement", title: "Vantage Cloud — Retirement benefits" })],
      },
      {
        benefit_key: "mega_backdoor_roth",
        plan_year: 2026,
        value: {
          supported: true,
          after_tax_contributions_allowed: true,
          after_tax_cap_percent_of_salary: 8,
          conversion_method: "manual_request",
          conversion_fee_usd: 25,
          recordkeeper: "Empower",
          headroom_note: "Total 415(c) limit minus employee deferral minus employer match.",
        },
        notes: "Conversions require a manual request to HR each quarter; not automatic.",
        confidence: "corroborated",
        last_verified_at: "2026-04-10T00:00:00.000Z",
        sources: [
          source({
            source_type: "user_report",
            url: "",
            title: "Two corroborating verified-employee reports",
            excerpt: "Both reports describe a quarterly manual conversion request process with a $25 fee.",
          }),
        ],
      },
      {
        benefit_key: "espp",
        plan_year: 2026,
        value: { offered: true, discount_percent: 10, lookback: false, offering_period_months: 3, contribution_cap_percent_of_salary: 15, holding_period_months: 0 },
        confidence: "official",
        last_verified_at: "2026-05-20T00:00:00.000Z",
        sources: [source({ source_type: "sec_filing", url: "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany", title: "Vantage Cloud S-8 — ESPP" })],
      },
      {
        benefit_key: "hsa_contribution",
        plan_year: 2026,
        value: { offered: true, annual_employee_only_usd: 500, annual_family_usd: 1000, requires_hdhp_enrollment: true },
        confidence: "official",
        last_verified_at: "2026-05-20T00:00:00.000Z",
        sources: [source({ url: "https://jobs.vantagecloud.example/benefits/health", title: "Vantage Cloud — Health benefits" })],
      },
      {
        benefit_key: "parental_leave",
        plan_year: 2026,
        value: { birthing_parent_weeks_paid: 16, non_birthing_parent_weeks_paid: 8, pay_percent: 100, covers_adoption_foster: true, tenure_required_months: 0 },
        confidence: "official",
        last_verified_at: "2026-05-20T00:00:00.000Z",
        sources: [source({ url: "https://jobs.vantagecloud.example/benefits/family", title: "Vantage Cloud — Parental leave" })],
      },
      {
        benefit_key: "pto_policy",
        plan_year: 2026,
        value: { policy_type: "accrued", days_per_year: 20, rollover_allowed: true, rollover_cap_days: 5, paid_holidays: 10 },
        confidence: "official",
        last_verified_at: "2026-05-20T00:00:00.000Z",
        sources: [source({ url: "https://jobs.vantagecloud.example/benefits/time-off", title: "Vantage Cloud — Time off" })],
      },
      {
        benefit_key: "commuter_benefit",
        plan_year: 2026,
        value: { offered: true, monthly_subsidy_usd: 150, pretax_account_available: true },
        confidence: "official",
        last_verified_at: "2026-05-20T00:00:00.000Z",
        sources: [source({ url: "https://jobs.vantagecloud.example/benefits/commute", title: "Vantage Cloud — Commuter benefits" })],
      },
    ],
  },
  {
    slug: "solstice-health-systems",
    name: "Solstice Health Systems",
    legal_name: "Solstice Health Systems, Inc.",
    logo_url: null,
    ticker: null,
    hq_country: "US",
    employee_band: "5001-50000",
    industry: "Healthcare",
    careers_url: "https://careers.solsticehealth.example",
    benefits_url: "https://careers.solsticehealth.example/benefits",
    levels_fyi_url: null,
    email_domains: ["solsticehealth.example"],
    is_published: true,
    benefits: [
      {
        benefit_key: "401k_match",
        plan_year: 2026,
        value: {
          formula_type: "tiered",
          cap_type: "percent_of_salary",
          cap_value: 5,
          tiers: [
            { up_to_percent_of_salary: 3, match_percent: 100 },
            { up_to_percent_of_salary: 5, match_percent: 50 },
          ],
          vesting_type: "cliff",
          vesting_years: 3,
          true_up: false,
          eligibility_days: 180,
        },
        notes: "100% on the first 3%, 50% on the next 2%. 3-year cliff vesting.",
        confidence: "official",
        last_verified_at: "2026-03-01T00:00:00.000Z",
        sources: [source({ source_type: "form_5500", url: "https://www.efast.dol.gov/5500search/", title: "Solstice Health Systems 401(k) Plan — Form 5500" })],
      },
      {
        benefit_key: "health_premium",
        plan_year: 2026,
        value: {
          employer_covers_percent_employee_only: 90,
          employer_covers_percent_family: 70,
          monthly_employee_cost_usd_employee_only: 45,
          monthly_employee_cost_usd_family: 310,
          deductible_usd_individual: 500,
          out_of_pocket_max_usd_individual: 3000,
        },
        confidence: "official",
        last_verified_at: "2026-03-01T00:00:00.000Z",
        sources: [source({ url: "https://careers.solsticehealth.example/benefits/medical", title: "Solstice Health Systems — Medical plan rates" })],
      },
      {
        benefit_key: "hsa_contribution",
        plan_year: 2026,
        value: { offered: true, annual_employee_only_usd: 1000, annual_family_usd: 2000, requires_hdhp_enrollment: true },
        confidence: "official",
        last_verified_at: "2026-03-01T00:00:00.000Z",
        sources: [source({ url: "https://careers.solsticehealth.example/benefits/medical", title: "Solstice Health Systems — HSA" })],
      },
      {
        benefit_key: "parental_leave",
        plan_year: 2026,
        value: { birthing_parent_weeks_paid: 12, non_birthing_parent_weeks_paid: 4, pay_percent: 100, covers_adoption_foster: false, tenure_required_months: 12 },
        confidence: "official",
        last_verified_at: "2026-03-01T00:00:00.000Z",
        sources: [source({ url: "https://careers.solsticehealth.example/benefits/family", title: "Solstice Health Systems — Parental leave" })],
      },
      {
        benefit_key: "pto_policy",
        plan_year: 2026,
        value: { policy_type: "accrued", days_per_year: 15, rollover_allowed: true, rollover_cap_days: 10, paid_holidays: 8 },
        confidence: "official",
        last_verified_at: "2026-03-01T00:00:00.000Z",
        sources: [source({ url: "https://careers.solsticehealth.example/benefits/time-off", title: "Solstice Health Systems — Time off" })],
      },
      {
        benefit_key: "life_insurance",
        plan_year: 2026,
        value: { basic_life_multiple_of_salary: 1, supplemental_life_available: true, short_term_disability_covered: true, long_term_disability_covered: true },
        confidence: "official",
        last_verified_at: "2026-03-01T00:00:00.000Z",
        sources: [source({ url: "https://careers.solsticehealth.example/benefits/insurance", title: "Solstice Health Systems — Life & disability" })],
      },
    ],
  },
  {
    slug: "meridian-capital-partners",
    name: "Meridian Capital Partners",
    legal_name: "Meridian Capital Partners LLC",
    logo_url: null,
    ticker: null,
    hq_country: "US",
    employee_band: "501-5000",
    industry: "Finance",
    careers_url: "https://careers.meridiancapital.example",
    benefits_url: "https://careers.meridiancapital.example/benefits",
    levels_fyi_url: null,
    email_domains: ["meridiancapital.example"],
    is_published: true,
    benefits: [
      {
        benefit_key: "401k_match",
        plan_year: 2026,
        value: {
          formula_type: "percent_of_salary",
          match_percent: 100,
          cap_type: "percent_of_salary",
          cap_value: 4,
          vesting_type: "immediate",
          vesting_years: 0,
          true_up: true,
          eligibility_days: 0,
        },
        confidence: "official",
        last_verified_at: "2026-06-01T00:00:00.000Z",
        sources: [source({ url: "https://careers.meridiancapital.example/benefits/retirement", title: "Meridian Capital Partners — 401(k)" })],
      },
      {
        benefit_key: "mega_backdoor_roth",
        plan_year: 2026,
        value: { supported: false, after_tax_contributions_allowed: false, conversion_method: "none" },
        notes: "Plan does not currently permit after-tax contributions.",
        confidence: "official",
        last_verified_at: "2026-06-01T00:00:00.000Z",
        sources: [source({ source_type: "form_5500", url: "https://www.efast.dol.gov/5500search/", title: "Meridian Capital Partners 401(k) Plan — Form 5500" })],
      },
      {
        benefit_key: "education_reimbursement",
        plan_year: 2026,
        value: { offered: true, annual_cap_usd: 10000, covers_degree_programs: true, requires_manager_approval: true },
        notes: "Includes CFA/CPA exam fee reimbursement.",
        confidence: "official",
        last_verified_at: "2026-06-01T00:00:00.000Z",
        sources: [source({ url: "https://careers.meridiancapital.example/benefits/learning", title: "Meridian Capital Partners — Professional development" })],
      },
      {
        benefit_key: "pto_policy",
        plan_year: 2026,
        value: { policy_type: "fixed_days", days_per_year: 25, rollover_allowed: true, rollover_cap_days: 5, paid_holidays: 10 },
        confidence: "official",
        last_verified_at: "2026-06-01T00:00:00.000Z",
        sources: [source({ url: "https://careers.meridiancapital.example/benefits/time-off", title: "Meridian Capital Partners — Time off" })],
      },
      {
        benefit_key: "sabbatical",
        plan_year: 2026,
        value: { offered: true, tenure_required_years: 7, weeks_paid: 6 },
        confidence: "single_report",
        last_verified_at: "2026-02-01T00:00:00.000Z",
        sources: [source({ source_type: "user_report", url: "", title: "Single verified-employee report, not yet corroborated" })],
      },
    ],
  },
  {
    slug: "fernbank-software",
    name: "Fernbank Software",
    legal_name: "Fernbank Software, Inc.",
    logo_url: null,
    ticker: null,
    hq_country: "US",
    employee_band: "51-500",
    industry: "Technology",
    careers_url: "https://careers.fernbanksoftware.example",
    benefits_url: "https://careers.fernbanksoftware.example/benefits",
    levels_fyi_url: null,
    email_domains: ["fernbanksoftware.example"],
    is_published: true,
    benefits: [
      {
        benefit_key: "401k_match",
        plan_year: 2026,
        value: {
          formula_type: "percent_of_employee_contribution",
          match_percent: 100,
          cap_type: "percent_of_salary",
          cap_value: 3,
          vesting_type: "immediate",
          vesting_years: 0,
          true_up: false,
          eligibility_days: 0,
        },
        confidence: "official",
        last_verified_at: "2026-07-10T00:00:00.000Z",
        sources: [source({ url: "https://careers.fernbanksoftware.example/benefits", title: "Fernbank Software — Benefits overview" })],
      },
      {
        benefit_key: "pto_policy",
        plan_year: 2026,
        value: { policy_type: "unlimited", rollover_allowed: false, paid_holidays: 9 },
        confidence: "official",
        last_verified_at: "2026-07-10T00:00:00.000Z",
        sources: [source({ url: "https://careers.fernbanksoftware.example/benefits", title: "Fernbank Software — Benefits overview" })],
      },
      {
        benefit_key: "lifestyle_stipend",
        plan_year: 2026,
        value: { offered: true, annual_usd: 600, cadence: "monthly", eligible_spend: "Wellness, home office" },
        confidence: "official",
        last_verified_at: "2026-07-10T00:00:00.000Z",
        sources: [source({ url: "https://careers.fernbanksoftware.example/benefits", title: "Fernbank Software — Benefits overview" })],
      },
      {
        benefit_key: "perk_misc",
        plan_year: 2026,
        value: { name: "Annual team offsite", description: "Company-wide offsite, location varies.", est_annual_value_usd: 1500, conditions: "Full-time employees" },
        confidence: "official",
        last_verified_at: "2026-07-10T00:00:00.000Z",
        sources: [source({ url: "https://careers.fernbanksoftware.example/benefits", title: "Fernbank Software — Benefits overview" })],
      },
    ],
  },
];
