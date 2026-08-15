import "server-only";
import { getPublicSupabaseClient, getServiceSupabaseClient, isSupabaseConfigured } from "@/lib/db/client";
import { localStore } from "@/lib/db/local-store";
import type { BenefitEdit, Company, CompanyBenefit, EditStatus } from "@/types";

export function usingLocalStore() {
  return !isSupabaseConfigured();
}

export async function listPublishedCompanies(): Promise<Company[]> {
  if (usingLocalStore()) return localStore.listPublishedCompanies();
  const sb = getPublicSupabaseClient()!;
  const { data, error } = await sb.from("companies").select("*").eq("is_published", true).order("name");
  if (error) throw error;
  return data as Company[];
}

export async function listAllCompanySlugs(): Promise<string[]> {
  if (usingLocalStore()) return localStore.listAllCompanySlugs();
  const sb = getPublicSupabaseClient()!;
  const { data, error } = await sb.from("companies").select("slug").eq("is_published", true);
  if (error) throw error;
  return (data as { slug: string }[]).map((r) => r.slug);
}

export async function getCompanyBySlug(slug: string): Promise<Company | null> {
  if (usingLocalStore()) return localStore.getCompanyBySlug(slug) ?? null;
  const sb = getPublicSupabaseClient()!;
  const { data, error } = await sb.from("companies").select("*").eq("slug", slug).eq("is_published", true).maybeSingle();
  if (error) throw error;
  return (data as Company | null) ?? null;
}

export async function getBenefitsForCompany(companyId: string): Promise<CompanyBenefit[]> {
  if (usingLocalStore()) return localStore.getBenefitsForCompany(companyId);
  const sb = getPublicSupabaseClient()!;
  const { data, error } = await sb
    .from("company_benefits")
    .select("*, sources:benefit_sources(*)")
    .eq("company_id", companyId);
  if (error) throw error;
  return data as CompanyBenefit[];
}

export async function getLeaderboard(benefitKey: string): Promise<{ company: Company; benefit: CompanyBenefit }[]> {
  if (usingLocalStore()) return localStore.getLeaderboard(benefitKey);
  const sb = getPublicSupabaseClient()!;
  const { data, error } = await sb
    .from("company_benefits")
    .select("*, sources:benefit_sources(*), company:companies!inner(*)")
    .eq("benefit_key", benefitKey)
    .eq("company.is_published", true);
  if (error) throw error;
  return (data as (CompanyBenefit & { company: Company })[]).map((row) => ({ company: row.company, benefit: row }));
}

export async function searchCompanies(query: string): Promise<Company[]> {
  if (usingLocalStore()) return localStore.searchCompanies(query);
  const sb = getPublicSupabaseClient()!;
  const q = sb.from("companies").select("*").eq("is_published", true).order("name").limit(20);
  const { data, error } = query.trim() ? await q.ilike("name", `%${query.trim()}%`) : await q;
  if (error) throw error;
  return data as Company[];
}

// ---- contributions / moderation ----

export async function submitEdit(
  input: Omit<BenefitEdit, "id" | "created_at" | "status" | "reviewed_by" | "review_note">
): Promise<BenefitEdit> {
  if (usingLocalStore()) return localStore.submitEdit(input);
  const sb = getServiceSupabaseClient() ?? getPublicSupabaseClient()!;
  const { data, error } = await sb.from("benefit_edits").insert(input).select().single();
  if (error) throw error;
  return data as BenefitEdit;
}

export async function listEdits(status?: EditStatus): Promise<BenefitEdit[]> {
  if (usingLocalStore()) return localStore.listEdits(status);
  const sb = getServiceSupabaseClient() ?? getPublicSupabaseClient()!;
  const q = sb.from("benefit_edits").select("*").order("created_at", { ascending: false });
  const { data, error } = status ? await q.eq("status", status) : await q;
  if (error) throw error;
  return data as BenefitEdit[];
}

export async function approveEdit(editId: string, reviewerId: string, note?: string): Promise<BenefitEdit | null> {
  if (usingLocalStore()) return localStore.approveEdit(editId, reviewerId, note) ?? null;
  const sb = getServiceSupabaseClient();
  if (!sb) throw new Error("Approving an edit requires SUPABASE_SERVICE_ROLE_KEY to be configured.");

  const { data: edit, error: editErr } = await sb.from("benefit_edits").select("*").eq("id", editId).single();
  if (editErr) throw editErr;
  const e = edit as BenefitEdit;

  const { data: existing } = await sb
    .from("company_benefits")
    .select("id")
    .eq("company_id", e.company_id)
    .eq("benefit_key", e.benefit_key)
    .eq("plan_year", e.plan_year)
    .maybeSingle();

  const now = new Date().toISOString();
  let benefitId: string;
  if (existing) {
    benefitId = (existing as { id: string }).id;
    const { error } = await sb
      .from("company_benefits")
      .update({ value: e.proposed_value, last_verified_at: now })
      .eq("id", benefitId);
    if (error) throw error;
  } else {
    const { data: created, error } = await sb
      .from("company_benefits")
      .insert({
        company_id: e.company_id,
        benefit_key: e.benefit_key,
        plan_year: e.plan_year,
        value: e.proposed_value,
        confidence: e.submitter_is_verified_employee ? "single_report" : "unverified",
        last_verified_at: now,
      })
      .select()
      .single();
    if (error) throw error;
    benefitId = (created as { id: string }).id;
  }

  const { error: sourceErr } = await sb.from("benefit_sources").insert({
    company_benefit_id: benefitId,
    source_type: e.source_type,
    url: e.source_url,
    title: `Approved edit — ${e.rationale ?? "no rationale given"}`,
    excerpt: e.rationale?.slice(0, 200) ?? null,
  });
  if (sourceErr) throw sourceErr;

  const { data: updatedEdit, error: updateErr } = await sb
    .from("benefit_edits")
    .update({ status: "approved", reviewed_by: reviewerId, review_note: note ?? null })
    .eq("id", editId)
    .select()
    .single();
  if (updateErr) throw updateErr;

  if (e.submitted_by) {
    await sb.rpc("increment_reputation", { p_user_id: e.submitted_by, p_amount: e.submitter_is_verified_employee ? 10 : 3 });
  }

  return updatedEdit as BenefitEdit;
}

export async function rejectEdit(editId: string, reviewerId: string, note?: string): Promise<BenefitEdit | null> {
  if (usingLocalStore()) return localStore.rejectEdit(editId, reviewerId, note) ?? null;
  const sb = getServiceSupabaseClient();
  if (!sb) throw new Error("Rejecting an edit requires SUPABASE_SERVICE_ROLE_KEY to be configured.");
  const { data, error } = await sb
    .from("benefit_edits")
    .update({ status: "rejected", reviewed_by: reviewerId, review_note: note ?? null })
    .eq("id", editId)
    .select()
    .single();
  if (error) throw error;
  return data as BenefitEdit;
}
