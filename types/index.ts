export type BenefitCategory =
  | "retirement"
  | "health"
  | "family"
  | "equity"
  | "time_off"
  | "education"
  | "wellness"
  | "perks"
  | "logistics";

export const BENEFIT_CATEGORIES: { key: BenefitCategory; label: string }[] = [
  { key: "retirement", label: "Retirement" },
  { key: "health", label: "Health" },
  { key: "family", label: "Family" },
  { key: "equity", label: "Equity" },
  { key: "time_off", label: "Time off" },
  { key: "education", label: "Education" },
  { key: "wellness", label: "Wellness" },
  { key: "perks", label: "Perks" },
  { key: "logistics", label: "Logistics" },
];

export type EmployeeBand = "1-50" | "51-500" | "501-5000" | "5001-50000" | "50000+";

export interface Company {
  id: string;
  slug: string;
  name: string;
  legal_name: string | null;
  logo_url: string | null;
  ticker: string | null;
  hq_country: string;
  employee_band: EmployeeBand | null;
  industry: string | null;
  careers_url: string | null;
  benefits_url: string | null;
  levels_fyi_url: string | null;
  email_domains: string[];
  is_published: boolean;
  created_at: string;
  updated_at: string;
}

export type ConfidenceLevel = "official" | "corroborated" | "single_report" | "unverified";

export type SourceType = "company_public_page" | "sec_filing" | "form_5500" | "press" | "news" | "user_report";

export interface BenefitSource {
  id: string;
  company_benefit_id: string;
  source_type: SourceType;
  url: string | null;
  title: string | null;
  retrieved_at: string;
  excerpt: string | null;
}

export interface CompanyBenefit {
  id: string;
  company_id: string;
  benefit_key: string;
  plan_year: number;
  country: string;
  value: Record<string, unknown>;
  notes: string | null;
  confidence: ConfidenceLevel;
  last_verified_at: string | null;
  created_at: string;
  updated_at: string;
  sources?: BenefitSource[];
}

export type EditStatus = "pending" | "approved" | "rejected" | "superseded";

export interface BenefitEdit {
  id: string;
  company_id: string;
  benefit_key: string;
  plan_year: number;
  proposed_value: Record<string, unknown>;
  current_value: Record<string, unknown> | null;
  source_type: SourceType;
  source_url: string | null;
  rationale: string | null;
  submitted_by: string | null;
  submitter_is_verified_employee: boolean;
  status: EditStatus;
  reviewed_by: string | null;
  review_note: string | null;
  created_at: string;
}

export interface Profile {
  id: string;
  handle: string;
  role: "user" | "trusted" | "moderator" | "admin";
  reputation: number;
  created_at: string;
}
