import {
  calc401kMatch,
  calcEsppExpectedValue,
  calcHealthPremiumValue,
  calcHsaValue,
  type EsppValue,
  type FourOhOneKMatchValue,
  type HealthPremiumValue,
  type HsaValue,
  type MegaBackdoorRothValue,
} from "./value";

const REFERENCE_SALARY = 200_000;
const REFERENCE_CONTRIBUTION_PERCENT = 6;
const PLAN_YEAR = 2026;

/**
 * A single comparable number per benefit type, used only to decide which
 * company's cell gets the "best in this row" marker on /compare and to sort
 * /benefits/[type] leaderboards. Higher is always better. Returns null when
 * a benefit type has no sensible single-number ordering (e.g. equity
 * refresh cadence) — those rows still render, just without highlighting.
 */
export function compareScore(benefitKey: string, value: Record<string, unknown>): number | null {
  switch (benefitKey) {
    case "401k_match":
      return calc401kMatch(value as unknown as FourOhOneKMatchValue, REFERENCE_SALARY, REFERENCE_CONTRIBUTION_PERCENT, PLAN_YEAR).amountUsd;
    case "mega_backdoor_roth":
      return (value as unknown as MegaBackdoorRothValue).supported ? 1 : 0;
    case "espp":
      return calcEsppExpectedValue(value as unknown as EsppValue, REFERENCE_SALARY).amountUsd;
    case "hsa_contribution":
      return calcHsaValue(value as unknown as HsaValue, 1).amountUsd;
    case "health_premium":
      return calcHealthPremiumValue(value as unknown as HealthPremiumValue, 1).amountUsd;
    case "childcare_stipend":
    case "lifestyle_stipend":
    case "education_reimbursement":
      return numberField(value, ["annual_usd", "annual_cap_usd"]);
    case "commuter_benefit":
      return numberField(value, ["monthly_subsidy_usd"]);
    case "parental_leave":
      return numberField(value, ["birthing_parent_weeks_paid"]);
    case "fertility_benefit":
      return numberField(value, ["lifetime_max_usd"]);
    case "sabbatical":
      return numberField(value, ["weeks_paid"]);
    case "life_insurance":
      return numberField(value, ["basic_life_multiple_of_salary"]);
    case "pto_policy":
      if (value.policy_type === "unlimited") return Number.POSITIVE_INFINITY;
      return numberField(value, ["days_per_year"]);
    default:
      return null;
  }
}

function numberField(value: Record<string, unknown>, keys: string[]): number | null {
  for (const key of keys) {
    const v = value[key];
    if (typeof v === "number") return v;
  }
  return null;
}
