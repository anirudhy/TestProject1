import { getIrsLimits } from "./irs-limits";
import type { CompanyBenefit } from "@/types";

/**
 * Pure arithmetic on stated plan terms. No projections of future stock
 * performance, no allocation advice, no ranking of companies as a
 * financial decision — see §4's guardrail. Every function here takes
 * explicit inputs and returns a number plus a human-readable explanation
 * of the formula used, so the UI can always show its work.
 */

export interface CalculatorInputs {
  salaryUsd: number;
  contributionPercent: number; // employee 401(k) deferral, % of salary
  familySize: number; // 1 = self only, >1 = family coverage
  age: number;
  planYear: number;
}

export interface CalcLine {
  benefitKey: string;
  label: string;
  amountUsd: number;
  formula: string;
}

// ---------------- 401(k) match ----------------

export interface FourOhOneKMatchTier {
  up_to_percent_of_salary: number;
  match_percent: number;
}

export interface FourOhOneKMatchValue {
  formula_type: "percent_of_employee_contribution" | "percent_of_salary" | "tiered" | "fixed_usd";
  match_percent?: number;
  cap_type?: "percent_of_irs_limit" | "percent_of_salary" | "usd";
  cap_value?: number;
  tiers?: FourOhOneKMatchTier[];
  true_up?: boolean;
}

function capAsPercentOfSalary(value: FourOhOneKMatchValue, salaryUsd: number, planYear: number): number {
  if (!value.cap_type || value.cap_value === undefined) return Infinity;
  const { cap_type, cap_value } = value;
  if (cap_type === "percent_of_salary") return cap_value;
  if (cap_type === "usd") return salaryUsd > 0 ? (cap_value / salaryUsd) * 100 : 0;
  if (cap_type === "percent_of_irs_limit") {
    const capUsd = (getIrsLimits(planYear).electiveDeferralUsd * cap_value) / 100;
    return salaryUsd > 0 ? (capUsd / salaryUsd) * 100 : 0;
  }
  return Infinity;
}

/** Match earned on a given contribution %, assuming even contributions all year. */
function matchForContributionPercent(value: FourOhOneKMatchValue, salaryUsd: number, contributionPercent: number, planYear: number): number {
  const capPercent = capAsPercentOfSalary(value, salaryUsd, planYear);
  const matchable = Math.min(contributionPercent, capPercent);

  if (value.formula_type === "tiered" && value.tiers?.length) {
    let matchUsd = 0;
    let previousUpTo = 0;
    for (const tier of value.tiers) {
      const bandWidth = tier.up_to_percent_of_salary - previousUpTo;
      const contributionInBand = Math.min(Math.max(matchable - previousUpTo, 0), bandWidth);
      matchUsd += (salaryUsd * contributionInBand * tier.match_percent) / 10_000;
      previousUpTo = tier.up_to_percent_of_salary;
    }
    return matchUsd;
  }

  const matchPercent = value.match_percent ?? 0;
  return (salaryUsd * matchable * matchPercent) / 10_000;
}

export function calc401kMatch(value: FourOhOneKMatchValue, salaryUsd: number, contributionPercent: number, planYear: number): CalcLine {
  const amountUsd = matchForContributionPercent(value, salaryUsd, contributionPercent, planYear);
  return {
    benefitKey: "401k_match",
    label: "401(k) match",
    amountUsd: Math.round(amountUsd),
    formula: `Employer match on a ${contributionPercent}% employee contribution, evenly spread across the year.`,
  };
}

/**
 * Simulates the plan paycheck-by-paycheck to quantify what true-up is
 * actually worth: an employee who front-loads contributions can hit the
 * IRS elective deferral limit before year-end, and payroll then stops
 * withholding — which means no further *employee* contribution to match.
 * Without true-up, the employer never makes up the difference; with
 * true-up, the plan tops it back up to the full-year figure.
 */
export function simulateMatchWithFrontloading(
  value: FourOhOneKMatchValue,
  salaryUsd: number,
  contributionPercent: number,
  planYear: number,
  payPeriodsPerYear = 26
): { matchReceivedUsd: number; matchLostToFrontloadingUsd: number; fullYearMatchUsd: number } {
  const limits = getIrsLimits(planYear);
  const perPeriodSalary = salaryUsd / payPeriodsPerYear;
  const perPeriodContribution = (perPeriodSalary * contributionPercent) / 100;

  let cumulativeEmployeeContribution = 0;
  let matchReceivedNoTrueUp = 0;
  for (let period = 0; period < payPeriodsPerYear; period++) {
    const remainingRoom = Math.max(limits.electiveDeferralUsd - cumulativeEmployeeContribution, 0);
    const actualContribution = Math.min(perPeriodContribution, remainingRoom);
    if (actualContribution <= 0) continue;
    cumulativeEmployeeContribution += actualContribution;
    const effectiveContributionPercentThisPeriod = (actualContribution / perPeriodSalary) * 100;
    matchReceivedNoTrueUp += matchForContributionPercent(value, perPeriodSalary, effectiveContributionPercentThisPeriod, planYear);
  }

  const fullYearMatchUsd = matchForContributionPercent(value, salaryUsd, contributionPercent, planYear);
  const matchReceivedUsd = value.true_up ? fullYearMatchUsd : matchReceivedNoTrueUp;

  return {
    matchReceivedUsd: Math.round(matchReceivedUsd),
    matchLostToFrontloadingUsd: Math.round(Math.max(fullYearMatchUsd - matchReceivedNoTrueUp, 0)),
    fullYearMatchUsd: Math.round(fullYearMatchUsd),
  };
}

// ---------------- Mega Backdoor Roth headroom ----------------

export interface MegaBackdoorRothValue {
  supported: boolean;
  after_tax_cap_percent_of_salary?: number;
}

/**
 * 415(c) limit minus employee deferral minus employer match, further capped
 * by whatever the plan itself allows for after-tax contributions. Catch-up
 * contributions are excluded from the 415(c) limit by IRS rule, so they're
 * not part of this calculation.
 */
export function calcMbdrHeadroom(
  value: MegaBackdoorRothValue,
  salaryUsd: number,
  planYear: number,
  employeeDeferralUsd: number,
  employerMatchUsd: number
): CalcLine {
  if (!value.supported) {
    return { benefitKey: "mega_backdoor_roth", label: "Mega Backdoor Roth headroom", amountUsd: 0, formula: "Plan does not support after-tax contributions." };
  }
  const limit415c = getIrsLimits(planYear).annualAddition415cUsd;
  const rawHeadroomUsd = Math.max(limit415c - employeeDeferralUsd - employerMatchUsd, 0);
  const planCapUsd = salaryUsd * ((value.after_tax_cap_percent_of_salary ?? 100) / 100);
  const amountUsd = Math.max(Math.min(rawHeadroomUsd, planCapUsd), 0);
  return {
    benefitKey: "mega_backdoor_roth",
    label: "Mega Backdoor Roth headroom",
    amountUsd: Math.round(amountUsd),
    formula: `415(c) limit ($${limit415c.toLocaleString()}) − employee deferral ($${Math.round(employeeDeferralUsd).toLocaleString()}) − employer match ($${Math.round(employerMatchUsd).toLocaleString()}), capped by the plan's after-tax contribution limit.`,
  };
}

// ---------------- ESPP ----------------

export interface EsppValue {
  offered: boolean;
  discount_percent?: number;
  contribution_cap_percent_of_salary?: number;
  lookback?: boolean;
  offering_period_months?: number;
}

/**
 * Guaranteed value only: buying at a discount and selling immediately (an
 * "ESPP flip"). Deliberately ignores the lookback provision's added
 * optionality value, since quantifying that requires a stock-price/
 * volatility assumption — which would cross from "arithmetic on stated
 * plan terms" into projecting returns.
 */
export function calcEsppExpectedValue(value: EsppValue, salaryUsd: number): CalcLine {
  if (!value.offered || !value.discount_percent) {
    return { benefitKey: "espp", label: "ESPP", amountUsd: 0, formula: "Not offered." };
  }
  const contributionUsd = (salaryUsd * (value.contribution_cap_percent_of_salary ?? 0)) / 100;
  const d = value.discount_percent / 100;
  const gainRate = d / (1 - d);
  const amountUsd = contributionUsd * gainRate;
  return {
    benefitKey: "espp",
    label: "ESPP (guaranteed discount only)",
    amountUsd: Math.round(amountUsd),
    formula: `${value.discount_percent}% discount on a $${Math.round(contributionUsd).toLocaleString()} contribution, assuming an immediate sale at purchase. Excludes any lookback upside, which depends on future stock price.`,
  };
}

/**
 * Illustrative estimate of the *additional* value a lookback provision adds
 * on top of the guaranteed discount, using the standard rule-of-thumb
 * approximation for an at-the-money option: value ≈ 0.4 · σ · √T (from
 * Black-Scholes). This requires a volatility assumption about the future,
 * which is why it's kept separate from — and excluded from — the total
 * employer value figure returned by calcTotalEmployerValue: that total is
 * pure arithmetic on stated plan terms, this is a labeled, assumption-
 * dependent estimate the UI must present as such, not as a promised return.
 */
export function calcEsppLookbackBonusValue(value: EsppValue, salaryUsd: number, assumedAnnualVolatility = 0.3): CalcLine {
  if (!value.offered || !value.lookback) {
    return { benefitKey: "espp_lookback", label: "ESPP lookback bonus (estimate)", amountUsd: 0, formula: "No lookback provision." };
  }
  const contributionUsd = (salaryUsd * (value.contribution_cap_percent_of_salary ?? 0)) / 100;
  const offeringYears = (value.offering_period_months ?? 6) / 12;
  const optionValueRate = 0.4 * assumedAnnualVolatility * Math.sqrt(offeringYears);
  const amountUsd = contributionUsd * optionValueRate;
  return {
    benefitKey: "espp_lookback",
    label: "ESPP lookback bonus (estimate)",
    amountUsd: Math.round(amountUsd),
    formula: `Illustrative only: assumes ${Math.round(assumedAnnualVolatility * 100)}% annualized stock volatility over a ${value.offering_period_months ?? 6}-month offering period (0.4·σ·√T rule of thumb). Not investment advice; excluded from the total employer value figure.`,
  };
}

// ---------------- HSA ----------------

export interface HsaValue {
  offered: boolean;
  annual_employee_only_usd?: number;
  annual_family_usd?: number;
}

export function calcHsaValue(value: HsaValue, familySize: number): CalcLine {
  const amountUsd = !value.offered ? 0 : familySize > 1 ? value.annual_family_usd ?? 0 : value.annual_employee_only_usd ?? 0;
  return {
    benefitKey: "hsa_contribution",
    label: "HSA employer contribution",
    amountUsd: Math.round(amountUsd),
    formula: familySize > 1 ? "Employer HSA seed for family coverage." : "Employer HSA seed for employee-only coverage.",
  };
}

// ---------------- Health premium subsidy ----------------

export interface HealthPremiumValue {
  employer_covers_percent_employee_only?: number;
  employer_covers_percent_family?: number;
  monthly_employee_cost_usd_employee_only?: number;
  monthly_employee_cost_usd_family?: number;
}

export function calcHealthPremiumValue(value: HealthPremiumValue, familySize: number): CalcLine {
  const coverPercent = familySize > 1 ? value.employer_covers_percent_family : value.employer_covers_percent_employee_only;
  const employeeMonthlyCost = familySize > 1 ? value.monthly_employee_cost_usd_family : value.monthly_employee_cost_usd_employee_only;
  if (!coverPercent || coverPercent >= 100 || employeeMonthlyCost === undefined) {
    return { benefitKey: "health_premium", label: "Health premium subsidy", amountUsd: 0, formula: "Insufficient data to derive employer share." };
  }
  const employerMonthlyShare = (employeeMonthlyCost * (coverPercent / 100)) / (1 - coverPercent / 100);
  return {
    benefitKey: "health_premium",
    label: "Health premium subsidy",
    amountUsd: Math.round(employerMonthlyShare * 12),
    formula: `Employer covers ${coverPercent}% of premium; derived from your $${employeeMonthlyCost}/month share.`,
  };
}

// ---------------- Flat annual stipends (childcare, lifestyle, commuter, education, perk_misc) ----------------

export function calcFlatAnnualUsd(benefitKey: string, label: string, annualUsd: number | undefined, formula: string): CalcLine {
  return { benefitKey, label, amountUsd: Math.round(annualUsd ?? 0), formula };
}

// ---------------- Aggregate across a company's benefits ----------------

export interface EmployerValueResult {
  lines: CalcLine[];
  totalUsd: number;
}

export function calcTotalEmployerValue(benefits: CompanyBenefit[], inputs: CalculatorInputs): EmployerValueResult {
  const byKey = new Map(benefits.map((b) => [b.benefit_key, b]));
  const lines: CalcLine[] = [];

  const match401k = byKey.get("401k_match");
  let employerMatchUsd = 0;
  if (match401k) {
    const line = calc401kMatch(match401k.value as unknown as FourOhOneKMatchValue, inputs.salaryUsd, inputs.contributionPercent, inputs.planYear);
    employerMatchUsd = line.amountUsd;
    lines.push(line);
  }

  const mbdr = byKey.get("mega_backdoor_roth");
  if (mbdr) {
    const employeeDeferralUsd = Math.min(
      (inputs.salaryUsd * inputs.contributionPercent) / 100,
      getIrsLimits(inputs.planYear).electiveDeferralUsd
    );
    lines.push(calcMbdrHeadroom(mbdr.value as unknown as MegaBackdoorRothValue, inputs.salaryUsd, inputs.planYear, employeeDeferralUsd, employerMatchUsd));
  }

  const hsa = byKey.get("hsa_contribution");
  if (hsa) lines.push(calcHsaValue(hsa.value as unknown as HsaValue, inputs.familySize));

  const healthPremium = byKey.get("health_premium");
  if (healthPremium) lines.push(calcHealthPremiumValue(healthPremium.value as unknown as HealthPremiumValue, inputs.familySize));

  const espp = byKey.get("espp");
  if (espp) lines.push(calcEsppExpectedValue(espp.value as unknown as EsppValue, inputs.salaryUsd));

  const childcare = byKey.get("childcare_stipend");
  if (childcare) {
    const v = childcare.value as { annual_usd?: number };
    lines.push(calcFlatAnnualUsd("childcare_stipend", "Childcare stipend", v.annual_usd, "Stated annual stipend."));
  }

  const lifestyle = byKey.get("lifestyle_stipend");
  if (lifestyle) {
    const v = lifestyle.value as { annual_usd?: number };
    lines.push(calcFlatAnnualUsd("lifestyle_stipend", "Lifestyle stipend", v.annual_usd, "Stated annual stipend."));
  }

  const commuter = byKey.get("commuter_benefit");
  if (commuter) {
    const v = commuter.value as { monthly_subsidy_usd?: number };
    lines.push(calcFlatAnnualUsd("commuter_benefit", "Commuter benefit", (v.monthly_subsidy_usd ?? 0) * 12, "Monthly subsidy × 12."));
  }

  const education = byKey.get("education_reimbursement");
  if (education) {
    const v = education.value as { annual_cap_usd?: number };
    lines.push(calcFlatAnnualUsd("education_reimbursement", "Education reimbursement", v.annual_cap_usd, "Stated annual cap; actual value depends on use."));
  }

  const perkMisc = benefits.filter((b) => b.benefit_key === "perk_misc");
  for (const perk of perkMisc) {
    const v = perk.value as { name?: string; est_annual_value_usd?: number };
    lines.push(calcFlatAnnualUsd("perk_misc", v.name ?? "Other perk", v.est_annual_value_usd, "Self-reported estimated annual value."));
  }

  const totalUsd = lines.reduce((sum, l) => sum + l.amountUsd, 0);
  return { lines, totalUsd };
}
