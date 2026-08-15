import "server-only";
import { SEED_COMPANIES } from "@/lib/seed/companies";
import { BENEFIT_TYPES } from "@/lib/benefits/registry";
import type { BenefitEdit, BenefitSource, Company, CompanyBenefit, EditStatus, Profile } from "@/types";

/**
 * In-memory fallback data store used when no Supabase project is configured
 * (`NEXT_PUBLIC_SUPABASE_URL` unset). Lets the app run, and every screen be
 * exercised, without live infra — this is what `pnpm dev` uses out of the box.
 * State resets on server restart / hot reload; that's fine for a demo store,
 * never used when a real Supabase URL is configured.
 */

let nextId = 1;
function id(prefix: string) {
  return `${prefix}_${nextId++}`;
}

const companies: Company[] = [];
const companyBenefits: CompanyBenefit[] = [];
const benefitSources: BenefitSource[] = [];
const benefitEdits: BenefitEdit[] = [];
const profiles: Profile[] = [
  { id: "profile_demo_admin", handle: "demo-admin", role: "admin", reputation: 100, created_at: "2026-01-01T00:00:00.000Z" },
];

function seedOnce() {
  if (companies.length > 0) return;

  for (const sc of SEED_COMPANIES) {
    const companyId = id("company");
    const now = new Date().toISOString();
    companies.push({
      id: companyId,
      slug: sc.slug,
      name: sc.name,
      legal_name: sc.legal_name,
      logo_url: sc.logo_url,
      ticker: sc.ticker,
      hq_country: sc.hq_country,
      employee_band: sc.employee_band,
      industry: sc.industry,
      careers_url: sc.careers_url,
      benefits_url: sc.benefits_url,
      levels_fyi_url: sc.levels_fyi_url,
      email_domains: sc.email_domains,
      is_published: sc.is_published,
      created_at: now,
      updated_at: now,
    });

    for (const b of sc.benefits) {
      const benefitId = id("benefit");
      companyBenefits.push({
        id: benefitId,
        company_id: companyId,
        benefit_key: b.benefit_key,
        plan_year: b.plan_year,
        country: "US",
        value: b.value,
        notes: b.notes ?? null,
        confidence: b.confidence,
        last_verified_at: b.last_verified_at,
        created_at: now,
        updated_at: now,
      });
      for (const s of b.sources) {
        benefitSources.push({
          id: id("source"),
          company_benefit_id: benefitId,
          source_type: s.source_type,
          url: s.url || null,
          title: s.title,
          retrieved_at: s.retrieved_at,
          excerpt: s.excerpt ?? null,
        });
      }
    }
  }
}
seedOnce();

export const localStore = {
  listPublishedCompanies(): Company[] {
    return companies.filter((c) => c.is_published);
  },
  getCompanyBySlug(slug: string): Company | undefined {
    return companies.find((c) => c.slug === slug && c.is_published);
  },
  getCompanyById(companyId: string): Company | undefined {
    return companies.find((c) => c.id === companyId);
  },
  listAllCompanySlugs(): string[] {
    return companies.filter((c) => c.is_published).map((c) => c.slug);
  },
  getBenefitsForCompany(companyId: string): CompanyBenefit[] {
    return companyBenefits
      .filter((b) => b.company_id === companyId)
      .map((b) => ({ ...b, sources: benefitSources.filter((s) => s.company_benefit_id === b.id) }));
  },
  getBenefitTypes() {
    return BENEFIT_TYPES;
  },
  getLeaderboard(benefitKey: string): { company: Company; benefit: CompanyBenefit }[] {
    const results: { company: Company; benefit: CompanyBenefit }[] = [];
    for (const b of companyBenefits) {
      if (b.benefit_key !== benefitKey) continue;
      const company = companies.find((c) => c.id === b.company_id);
      if (!company || !company.is_published) continue;
      const benefit: CompanyBenefit = { ...b, sources: benefitSources.filter((s) => s.company_benefit_id === b.id) };
      results.push({ company, benefit });
    }
    return results;
  },
  searchCompanies(query: string): Company[] {
    const q = query.trim().toLowerCase();
    if (!q) return companies.filter((c) => c.is_published);
    return companies.filter((c) => c.is_published && c.name.toLowerCase().includes(q));
  },

  // ---- contribution / moderation (in-memory, demo only) ----
  listEdits(status?: EditStatus): BenefitEdit[] {
    return status ? benefitEdits.filter((e) => e.status === status) : benefitEdits;
  },
  getEdit(editId: string): BenefitEdit | undefined {
    return benefitEdits.find((e) => e.id === editId);
  },
  submitEdit(input: Omit<BenefitEdit, "id" | "created_at" | "status" | "reviewed_by" | "review_note">): BenefitEdit {
    const edit: BenefitEdit = {
      ...input,
      id: id("edit"),
      created_at: new Date().toISOString(),
      status: "pending",
      reviewed_by: null,
      review_note: null,
    };
    benefitEdits.push(edit);
    return edit;
  },
  approveEdit(editId: string, reviewerId: string, note?: string): BenefitEdit | undefined {
    const edit = benefitEdits.find((e) => e.id === editId);
    if (!edit) return undefined;
    edit.status = "approved";
    edit.reviewed_by = reviewerId;
    edit.review_note = note ?? null;

    const now = new Date().toISOString();
    let benefit = companyBenefits.find(
      (b) => b.company_id === edit.company_id && b.benefit_key === edit.benefit_key && b.plan_year === edit.plan_year
    );
    if (benefit) {
      benefit.value = edit.proposed_value;
      benefit.confidence = edit.submitter_is_verified_employee ? "single_report" : benefit.confidence;
      benefit.last_verified_at = now;
      benefit.updated_at = now;
    } else {
      benefit = {
        id: id("benefit"),
        company_id: edit.company_id,
        benefit_key: edit.benefit_key,
        plan_year: edit.plan_year,
        country: "US",
        value: edit.proposed_value,
        notes: null,
        confidence: edit.submitter_is_verified_employee ? "single_report" : "unverified",
        last_verified_at: now,
        created_at: now,
        updated_at: now,
      };
      companyBenefits.push(benefit);
    }
    benefitSources.push({
      id: id("source"),
      company_benefit_id: benefit.id,
      source_type: edit.source_type,
      url: edit.source_url,
      title: `Approved edit — ${edit.rationale ?? "no rationale given"}`,
      retrieved_at: now,
      excerpt: edit.rationale?.slice(0, 200) ?? null,
    });

    const submitter = profiles.find((p) => p.id === edit.submitted_by);
    if (submitter) submitter.reputation += edit.submitter_is_verified_employee ? 10 : 3;

    return edit;
  },
  rejectEdit(editId: string, reviewerId: string, note?: string): BenefitEdit | undefined {
    const edit = benefitEdits.find((e) => e.id === editId);
    if (!edit) return undefined;
    edit.status = "rejected";
    edit.reviewed_by = reviewerId;
    edit.review_note = note ?? null;
    return edit;
  },
  getProfile(profileId: string): Profile | undefined {
    return profiles.find((p) => p.id === profileId);
  },
  getOrCreateDemoProfile(): Profile {
    let p = profiles.find((x) => x.id === "profile_demo_user");
    if (!p) {
      p = { id: "profile_demo_user", handle: "demo-user", role: "user", reputation: 0, created_at: new Date().toISOString() };
      profiles.push(p);
    }
    return p;
  },
};
