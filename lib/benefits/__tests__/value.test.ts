import { describe, expect, it } from "vitest";
import {
  calc401kMatch,
  simulateMatchWithFrontloading,
  calcMbdrHeadroom,
  calcEsppExpectedValue,
  calcEsppLookbackBonusValue,
  calcHsaValue,
  calcHealthPremiumValue,
  calcTotalEmployerValue,
  type FourOhOneKMatchValue,
} from "@/lib/benefits/value";
import type { CompanyBenefit } from "@/types";

const PLAN_YEAR = 2026;

describe("calc401kMatch", () => {
  it("computes a simple flat match up to a percent-of-salary cap", () => {
    const value: FourOhOneKMatchValue = {
      formula_type: "percent_of_employee_contribution",
      match_percent: 50,
      cap_type: "percent_of_salary",
      cap_value: 6,
    };
    // Contributing exactly the cap: 50% match on 6% of $200,000 = $6,000
    const line = calc401kMatch(value, 200_000, 6, PLAN_YEAR);
    expect(line.amountUsd).toBe(6_000);
  });

  it("caps the match even if the employee contributes more than the cap", () => {
    const value: FourOhOneKMatchValue = {
      formula_type: "percent_of_employee_contribution",
      match_percent: 50,
      cap_type: "percent_of_salary",
      cap_value: 6,
    };
    const line = calc401kMatch(value, 200_000, 15, PLAN_YEAR);
    expect(line.amountUsd).toBe(6_000);
  });

  it("computes a tiered match across bands (100% first 3%, 50% next 2%)", () => {
    const value: FourOhOneKMatchValue = {
      formula_type: "tiered",
      cap_type: "percent_of_salary",
      cap_value: 5,
      tiers: [
        { up_to_percent_of_salary: 3, match_percent: 100 },
        { up_to_percent_of_salary: 5, match_percent: 50 },
      ],
    };
    // Contributing 5% of $100,000: 3% band at 100% = $3,000, next 2% at 50% = $1,000
    const line = calc401kMatch(value, 100_000, 5, PLAN_YEAR);
    expect(line.amountUsd).toBe(4_000);
  });

  it("only matches contribution within the reached tier when contribution is below the top tier", () => {
    const value: FourOhOneKMatchValue = {
      formula_type: "tiered",
      cap_type: "percent_of_salary",
      cap_value: 5,
      tiers: [
        { up_to_percent_of_salary: 3, match_percent: 100 },
        { up_to_percent_of_salary: 5, match_percent: 50 },
      ],
    };
    // Contributing only 2%: fully inside the first band, 100% match
    const line = calc401kMatch(value, 100_000, 2, PLAN_YEAR);
    expect(line.amountUsd).toBe(2_000);
  });

  it("converts a percent_of_irs_limit cap into an equivalent percent of salary", () => {
    const value: FourOhOneKMatchValue = {
      formula_type: "percent_of_employee_contribution",
      match_percent: 100,
      cap_type: "percent_of_irs_limit",
      cap_value: 50, // 50% of the 2026 elective deferral limit ($24,500) = $12,250
    };
    const line = calc401kMatch(value, 200_000, 20, PLAN_YEAR);
    // capPercent = 12,250 / 200,000 * 100 = 6.125%; match = 200,000 * 6.125% * 100% = 12,250
    expect(line.amountUsd).toBe(12_250);
  });
});

describe("simulateMatchWithFrontloading (true-up)", () => {
  const value: FourOhOneKMatchValue = {
    formula_type: "percent_of_employee_contribution",
    match_percent: 100,
    cap_type: "percent_of_salary",
    cap_value: 6,
    true_up: false,
  };

  it("loses match to front-loading when contributing aggressively without true-up", () => {
    // $200,000 salary contributing 20% hits the $24,500 2026 deferral limit
    // well before the last paycheck, so several late-year matches are missed.
    const result = simulateMatchWithFrontloading(value, 200_000, 20, PLAN_YEAR);
    expect(result.matchLostToFrontloadingUsd).toBeGreaterThan(0);
    expect(result.matchReceivedUsd).toBeLessThan(result.fullYearMatchUsd);
  });

  it("recovers the full match with true-up enabled, regardless of front-loading", () => {
    const trueUpValue: FourOhOneKMatchValue = { ...value, true_up: true };
    const result = simulateMatchWithFrontloading(trueUpValue, 200_000, 20, PLAN_YEAR);
    expect(result.matchReceivedUsd).toBe(result.fullYearMatchUsd);
  });

  it("matches the steady-state full-year figure when contributions never hit the deferral limit", () => {
    // 6% of $200,000 = $12,000/year, well under the $24,500 limit — no front-loading effect.
    const result = simulateMatchWithFrontloading(value, 200_000, 6, PLAN_YEAR);
    expect(result.matchReceivedUsd).toBe(result.fullYearMatchUsd);
    expect(result.matchLostToFrontloadingUsd).toBe(0);
  });
});

describe("calcMbdrHeadroom", () => {
  it("computes 415(c) limit minus deferral minus match, capped by the plan's after-tax limit", () => {
    // 2026 415(c) limit is $72,000. Deferral $24,500, match $12,250 -> raw headroom $35,250.
    // Plan caps after-tax contributions at 10% of a $200,000 salary = $20,000, which binds.
    const line = calcMbdrHeadroom({ supported: true, after_tax_cap_percent_of_salary: 10 }, 200_000, PLAN_YEAR, 24_500, 12_250);
    expect(line.amountUsd).toBe(20_000);
  });

  it("is not capped when the plan limit exceeds the raw 415(c) headroom", () => {
    const line = calcMbdrHeadroom({ supported: true, after_tax_cap_percent_of_salary: 50 }, 200_000, PLAN_YEAR, 24_500, 12_250);
    // raw headroom = 72,000 - 24,500 - 12,250 = 35,250; plan cap = 100,000 -> doesn't bind
    expect(line.amountUsd).toBe(35_250);
  });

  it("returns zero when the plan does not support it", () => {
    const line = calcMbdrHeadroom({ supported: false }, 200_000, PLAN_YEAR, 24_500, 12_250);
    expect(line.amountUsd).toBe(0);
  });

  it("never returns a negative headroom", () => {
    // deferral + match alone already exceed the $72,000 415(c) limit
    const line = calcMbdrHeadroom({ supported: true, after_tax_cap_percent_of_salary: 50 }, 200_000, PLAN_YEAR, 24_500, 50_000);
    expect(line.amountUsd).toBe(0);
  });
});

describe("calcEsppExpectedValue (guaranteed discount)", () => {
  it("computes the immediate-sale gain from the discount", () => {
    // 15% discount on a $20,000 contribution: gain = 20,000 * (0.15/0.85)
    const line = calcEsppExpectedValue({ offered: true, discount_percent: 15, contribution_cap_percent_of_salary: 10 }, 200_000);
    expect(line.amountUsd).toBe(Math.round(20_000 * (0.15 / 0.85)));
  });

  it("returns zero when not offered", () => {
    const line = calcEsppExpectedValue({ offered: false }, 200_000);
    expect(line.amountUsd).toBe(0);
  });
});

describe("calcEsppLookbackBonusValue", () => {
  it("returns zero when there is no lookback provision", () => {
    const line = calcEsppLookbackBonusValue(
      { offered: true, discount_percent: 15, contribution_cap_percent_of_salary: 10, lookback: false, offering_period_months: 6 },
      200_000
    );
    expect(line.amountUsd).toBe(0);
  });

  it("estimates a positive bonus using the 0.4*sigma*sqrt(T) rule of thumb when lookback is offered", () => {
    const contributionUsd = 20_000;
    const line = calcEsppLookbackBonusValue(
      { offered: true, discount_percent: 15, contribution_cap_percent_of_salary: 10, lookback: true, offering_period_months: 6 },
      200_000,
      0.3
    );
    const expected = Math.round(contributionUsd * 0.4 * 0.3 * Math.sqrt(0.5));
    expect(line.amountUsd).toBe(expected);
    expect(line.amountUsd).toBeGreaterThan(0);
  });

  it("scales up with a longer offering period", () => {
    const shortLine = calcEsppLookbackBonusValue(
      { offered: true, contribution_cap_percent_of_salary: 10, lookback: true, offering_period_months: 6 },
      200_000
    );
    const longLine = calcEsppLookbackBonusValue(
      { offered: true, contribution_cap_percent_of_salary: 10, lookback: true, offering_period_months: 24 },
      200_000
    );
    expect(longLine.amountUsd).toBeGreaterThan(shortLine.amountUsd);
  });
});

describe("calcHsaValue", () => {
  it("uses the family contribution amount for family coverage", () => {
    const line = calcHsaValue({ offered: true, annual_employee_only_usd: 750, annual_family_usd: 1500 }, 3);
    expect(line.amountUsd).toBe(1500);
  });

  it("uses the employee-only amount for single coverage", () => {
    const line = calcHsaValue({ offered: true, annual_employee_only_usd: 750, annual_family_usd: 1500 }, 1);
    expect(line.amountUsd).toBe(750);
  });
});

describe("calcHealthPremiumValue", () => {
  it("derives the employer's annual dollar share from the coverage percent and employee cost", () => {
    // Employee pays $45/mo which is 10% of the premium (employer covers 90%)
    const line = calcHealthPremiumValue(
      { employer_covers_percent_employee_only: 90, monthly_employee_cost_usd_employee_only: 45 },
      1
    );
    const employerMonthly = (45 * 0.9) / 0.1;
    expect(line.amountUsd).toBe(Math.round(employerMonthly * 12));
  });
});

describe("calcTotalEmployerValue", () => {
  it("sums countable benefit lines and excludes non-countable / unrecognized data", () => {
    const benefits: CompanyBenefit[] = [
      {
        id: "1",
        company_id: "c1",
        benefit_key: "401k_match",
        plan_year: PLAN_YEAR,
        country: "US",
        value: { formula_type: "percent_of_employee_contribution", match_percent: 100, cap_type: "percent_of_salary", cap_value: 6 },
        notes: null,
        confidence: "official",
        last_verified_at: null,
        created_at: "",
        updated_at: "",
      },
      {
        id: "2",
        company_id: "c1",
        benefit_key: "lifestyle_stipend",
        plan_year: PLAN_YEAR,
        country: "US",
        value: { offered: true, annual_usd: 1200 },
        notes: null,
        confidence: "official",
        last_verified_at: null,
        created_at: "",
        updated_at: "",
      },
    ];
    const result = calcTotalEmployerValue(benefits, { salaryUsd: 100_000, contributionPercent: 6, familySize: 1, age: 30, planYear: PLAN_YEAR });
    // 401k match: 100,000 * 6% * 100% = 6,000. Lifestyle stipend: 1,200. Total: 7,200.
    expect(result.totalUsd).toBe(7_200);
    expect(result.lines).toHaveLength(2);
  });
});
