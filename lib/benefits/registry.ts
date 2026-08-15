import type { BenefitCategory } from "@/types";
import { type FieldSpec, fieldsToZodObject, jsonSchemaDocument } from "./fields";

export interface BenefitTypeDef {
  key: string;
  category: BenefitCategory;
  label: string;
  description: string;
  fields: readonly FieldSpec[];
  isComparable: boolean;
  isCountable: boolean;
  sortOrder: number;
}

const MATCH_TIER_FIELDS: readonly FieldSpec[] = [
  { key: "up_to_percent_of_salary", label: "Up to % of salary contributed", type: "number", unit: "%" },
  { key: "match_percent", label: "Match %", type: "number", unit: "%" },
] as const;

export const BENEFIT_TYPES: BenefitTypeDef[] = [
  {
    key: "401k_match",
    category: "retirement",
    label: "401(k) match",
    description: "Employer matching formula, cap, vesting schedule, and true-up policy for the 401(k) plan.",
    isComparable: true,
    isCountable: true,
    sortOrder: 10,
    fields: [
      { key: "formula_type", label: "Formula type", type: "enum", options: ["percent_of_employee_contribution", "percent_of_salary", "tiered", "fixed_usd"] },
      { key: "match_percent", label: "Match %", type: "number", unit: "%" },
      { key: "cap_type", label: "Cap type", type: "enum", options: ["percent_of_irs_limit", "percent_of_salary", "usd"] },
      { key: "cap_value", label: "Cap value", type: "number" },
      { key: "tiers", label: "Tiers", type: "array", itemFields: MATCH_TIER_FIELDS },
      { key: "vesting_type", label: "Vesting type", type: "enum", options: ["immediate", "cliff", "graded"] },
      { key: "vesting_years", label: "Vesting years", type: "number" },
      { key: "true_up", label: "True-up offered", type: "boolean", help: "Employer tops up the match at year-end for employees who front-load contributions." },
      { key: "eligibility_days", label: "Eligibility waiting period (days)", type: "number" },
    ],
  },
  {
    key: "mega_backdoor_roth",
    category: "retirement",
    label: "Mega Backdoor Roth",
    description: "After-tax 401(k) contributions plus a conversion path to Roth, beyond the standard employee deferral limit.",
    isComparable: true,
    isCountable: true,
    sortOrder: 20,
    fields: [
      { key: "supported", label: "Supported", type: "boolean" },
      { key: "after_tax_contributions_allowed", label: "After-tax contributions allowed", type: "boolean" },
      { key: "after_tax_cap_percent_of_salary", label: "After-tax cap (% of salary)", type: "number", unit: "%" },
      { key: "conversion_method", label: "Conversion method", type: "enum", options: ["automatic_daily", "automatic_quarterly", "manual_request", "in_service_withdrawal", "none"] },
      { key: "conversion_fee_usd", label: "Conversion fee", type: "number", unit: "$" },
      { key: "recordkeeper", label: "Recordkeeper", type: "string" },
      { key: "headroom_note", label: "Headroom note", type: "string" },
    ],
  },
  {
    key: "hsa_contribution",
    category: "retirement",
    label: "HSA employer contribution",
    description: "Employer seed money into a Health Savings Account for HDHP enrollees.",
    isComparable: true,
    isCountable: true,
    sortOrder: 30,
    fields: [
      { key: "offered", label: "Offered", type: "boolean" },
      { key: "annual_employee_only_usd", label: "Annual contribution — employee only", type: "number", unit: "$" },
      { key: "annual_family_usd", label: "Annual contribution — family coverage", type: "number", unit: "$" },
      { key: "requires_hdhp_enrollment", label: "Requires HDHP enrollment", type: "boolean" },
    ],
  },
  {
    key: "health_premium",
    category: "health",
    label: "Health insurance premium",
    description: "What the employer covers of monthly medical premiums.",
    isComparable: true,
    isCountable: true,
    sortOrder: 40,
    fields: [
      { key: "employer_covers_percent_employee_only", label: "Employer covers % — employee only", type: "number", unit: "%" },
      { key: "employer_covers_percent_family", label: "Employer covers % — family", type: "number", unit: "%" },
      { key: "monthly_employee_cost_usd_employee_only", label: "Monthly employee cost — employee only", type: "number", unit: "$" },
      { key: "monthly_employee_cost_usd_family", label: "Monthly employee cost — family", type: "number", unit: "$" },
      { key: "deductible_usd_individual", label: "Deductible — individual", type: "number", unit: "$" },
      { key: "out_of_pocket_max_usd_individual", label: "Out-of-pocket max — individual", type: "number", unit: "$" },
    ],
  },
  {
    key: "parental_leave",
    category: "family",
    label: "Parental leave",
    description: "Paid weeks off for birthing and non-birthing parents.",
    isComparable: true,
    isCountable: false,
    sortOrder: 50,
    fields: [
      { key: "birthing_parent_weeks_paid", label: "Birthing parent — weeks paid", type: "number", unit: "weeks" },
      { key: "non_birthing_parent_weeks_paid", label: "Non-birthing parent — weeks paid", type: "number", unit: "weeks" },
      { key: "pay_percent", label: "Pay %", type: "number", unit: "%" },
      { key: "covers_adoption_foster", label: "Covers adoption / foster", type: "boolean" },
      { key: "tenure_required_months", label: "Tenure required (months)", type: "number" },
    ],
  },
  {
    key: "fertility_benefit",
    category: "family",
    label: "Fertility & family-forming benefit",
    description: "Coverage for IVF, egg/embryo freezing, surrogacy, and adoption assistance.",
    isComparable: true,
    isCountable: true,
    sortOrder: 60,
    fields: [
      { key: "offered", label: "Offered", type: "boolean" },
      { key: "lifetime_max_usd", label: "Lifetime max", type: "number", unit: "$" },
      { key: "covers_ivf", label: "Covers IVF", type: "boolean" },
      { key: "covers_egg_freezing", label: "Covers egg/embryo freezing", type: "boolean" },
      { key: "covers_surrogacy", label: "Covers surrogacy", type: "boolean" },
      { key: "adoption_assistance_usd", label: "Adoption assistance", type: "number", unit: "$" },
      { key: "provider", label: "Provider (e.g. Carrot, Progyny)", type: "string" },
    ],
  },
  {
    key: "childcare_stipend",
    category: "family",
    label: "Childcare stipend",
    description: "Recurring or backup childcare support.",
    isComparable: true,
    isCountable: true,
    sortOrder: 70,
    fields: [
      { key: "offered", label: "Offered", type: "boolean" },
      { key: "annual_usd", label: "Annual amount", type: "number", unit: "$" },
      { key: "backup_care_days_per_year", label: "Backup care days / year", type: "number" },
      { key: "provider", label: "Provider (e.g. Bright Horizons)", type: "string" },
    ],
  },
  {
    key: "espp",
    category: "equity",
    label: "Employee Stock Purchase Plan",
    description: "Discounted purchase of company stock via payroll deduction.",
    isComparable: true,
    isCountable: true,
    sortOrder: 80,
    fields: [
      { key: "offered", label: "Offered", type: "boolean" },
      { key: "discount_percent", label: "Discount %", type: "number", unit: "%" },
      { key: "lookback", label: "Lookback provision", type: "boolean" },
      { key: "offering_period_months", label: "Offering period (months)", type: "number" },
      { key: "contribution_cap_percent_of_salary", label: "Contribution cap (% of salary)", type: "number", unit: "%" },
      { key: "holding_period_months", label: "Required holding period (months)", type: "number" },
    ],
  },
  {
    key: "equity_refresh",
    category: "equity",
    label: "Equity refresh policy",
    description: "Whether and how often employees receive additional equity grants after hire.",
    isComparable: true,
    isCountable: false,
    sortOrder: 90,
    fields: [
      { key: "offered", label: "Offered", type: "boolean" },
      { key: "cadence", label: "Cadence", type: "enum", options: ["annual", "biennial", "performance_based", "ad_hoc", "none"] },
      { key: "typical_vesting_years", label: "Typical vesting (years)", type: "number" },
      { key: "notes", label: "Notes", type: "string" },
    ],
  },
  {
    key: "pto_policy",
    category: "time_off",
    label: "PTO policy",
    description: "Vacation day structure — accrued, fixed, or unlimited.",
    isComparable: true,
    isCountable: false,
    sortOrder: 100,
    fields: [
      { key: "policy_type", label: "Policy type", type: "enum", options: ["accrued", "fixed_days", "unlimited", "flexible"] },
      { key: "days_per_year", label: "Days per year", type: "number", unit: "days" },
      { key: "rollover_allowed", label: "Rollover allowed", type: "boolean" },
      { key: "rollover_cap_days", label: "Rollover cap (days)", type: "number" },
      { key: "paid_holidays", label: "Paid holidays", type: "number", unit: "days" },
    ],
  },
  {
    key: "sabbatical",
    category: "time_off",
    label: "Sabbatical",
    description: "Extended paid leave after a tenure milestone.",
    isComparable: true,
    isCountable: false,
    sortOrder: 110,
    fields: [
      { key: "offered", label: "Offered", type: "boolean" },
      { key: "tenure_required_years", label: "Tenure required (years)", type: "number" },
      { key: "weeks_paid", label: "Weeks paid", type: "number", unit: "weeks" },
    ],
  },
  {
    key: "education_reimbursement",
    category: "education",
    label: "Education / tuition reimbursement",
    description: "Annual budget for courses, certifications, conferences, or degree programs.",
    isComparable: true,
    isCountable: true,
    sortOrder: 120,
    fields: [
      { key: "offered", label: "Offered", type: "boolean" },
      { key: "annual_cap_usd", label: "Annual cap", type: "number", unit: "$" },
      { key: "covers_degree_programs", label: "Covers degree programs", type: "boolean" },
      { key: "requires_manager_approval", label: "Requires manager approval", type: "boolean" },
    ],
  },
  {
    key: "lifestyle_stipend",
    category: "wellness",
    label: "Lifestyle / wellness stipend",
    description: "Recurring cash allowance for gym, wellness apps, or general lifestyle spend.",
    isComparable: true,
    isCountable: true,
    sortOrder: 130,
    fields: [
      { key: "offered", label: "Offered", type: "boolean" },
      { key: "annual_usd", label: "Annual amount", type: "number", unit: "$" },
      { key: "cadence", label: "Cadence", type: "enum", options: ["monthly", "quarterly", "annual"] },
      { key: "eligible_spend", label: "Eligible spend categories", type: "string" },
    ],
  },
  {
    key: "commuter_benefit",
    category: "logistics",
    label: "Commuter benefit",
    description: "Pre-tax transit/parking accounts or direct commuter subsidies.",
    isComparable: true,
    isCountable: true,
    sortOrder: 140,
    fields: [
      { key: "offered", label: "Offered", type: "boolean" },
      { key: "monthly_subsidy_usd", label: "Monthly subsidy", type: "number", unit: "$" },
      { key: "pretax_account_available", label: "Pre-tax account available", type: "boolean" },
    ],
  },
  {
    key: "life_insurance",
    category: "health",
    label: "Life & disability insurance",
    description: "Employer-paid life and disability coverage.",
    isComparable: true,
    isCountable: false,
    sortOrder: 150,
    fields: [
      { key: "basic_life_multiple_of_salary", label: "Basic life (× salary)", type: "number" },
      { key: "supplemental_life_available", label: "Supplemental life available", type: "boolean" },
      { key: "short_term_disability_covered", label: "Short-term disability covered", type: "boolean" },
      { key: "long_term_disability_covered", label: "Long-term disability covered", type: "boolean" },
    ],
  },
  {
    key: "perk_misc",
    category: "perks",
    label: "Other perk",
    description: "Escape hatch for perks that don't fit a structured category (free subscriptions, on-site amenities, etc).",
    isComparable: false,
    isCountable: true,
    sortOrder: 160,
    fields: [
      { key: "name", label: "Name", type: "string" },
      { key: "description", label: "Description", type: "string" },
      { key: "est_annual_value_usd", label: "Estimated annual value", type: "number", unit: "$" },
      { key: "conditions", label: "Conditions", type: "string" },
    ],
  },
];

export const BENEFIT_TYPE_MAP: Record<string, BenefitTypeDef> = Object.fromEntries(
  BENEFIT_TYPES.map((t) => [t.key, t])
);

export function getBenefitType(key: string): BenefitTypeDef | undefined {
  return BENEFIT_TYPE_MAP[key];
}

export function zodForBenefitType(key: string) {
  const def = getBenefitType(key);
  if (!def) return null;
  return fieldsToZodObject(def.fields);
}

export function jsonSchemaForBenefitType(key: string) {
  const def = getBenefitType(key);
  if (!def) return null;
  return jsonSchemaDocument(def.fields);
}

export const TOTAL_FIELD_COUNT = BENEFIT_TYPES.length;
