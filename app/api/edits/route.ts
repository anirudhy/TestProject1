import { NextResponse } from "next/server";
import { z } from "zod";
import { getBenefitType, zodForBenefitType } from "@/lib/benefits/registry";
import { getBenefitsForCompany, submitEdit } from "@/lib/db/queries";
import { currentContributorId } from "@/lib/db/supabase-server";
import { isSupabaseConfigured } from "@/lib/db/client";

const requestSchema = z.object({
  companyId: z.string().min(1),
  benefitKey: z.string().min(1),
  planYear: z.number().int(),
  proposedValue: z.record(z.string(), z.unknown()),
  sourceType: z.enum(["company_public_page", "sec_filing", "form_5500", "press", "news", "user_report"]),
  sourceUrl: z.string().url().optional().or(z.literal("")),
  rationale: z.string().max(2000).optional(),
});

export async function POST(request: Request) {
  const submittedBy = await currentContributorId();
  // Demo mode always has a stand-in contributor; against a real Supabase
  // project, submitting requires a session (§9 open question #1: sign-in
  // only to contribute, browsing stays open).
  if (isSupabaseConfigured() && !submittedBy) {
    return NextResponse.json({ error: "Sign in to submit a contribution." }, { status: 401 });
  }

  const body = await request.json().catch(() => null);
  const parsed = requestSchema.safeParse(body);
  if (!parsed.success) {
    return NextResponse.json({ error: "Invalid request", details: parsed.error.flatten() }, { status: 400 });
  }
  const { companyId, benefitKey, planYear, proposedValue, sourceType, sourceUrl, rationale } = parsed.data;

  const type = getBenefitType(benefitKey);
  if (!type) return NextResponse.json({ error: `Unknown benefit type: ${benefitKey}` }, { status: 400 });

  // A source is non-negotiable for anything other than a raw employee report,
  // which is the one case where "I worked there and this is what I saw" is
  // itself the source — see §5's hard rules.
  if (sourceType !== "user_report" && !sourceUrl) {
    return NextResponse.json({ error: "A source URL is required for this source type." }, { status: 400 });
  }

  const valueSchema = zodForBenefitType(benefitKey)!;
  const valueParsed = valueSchema.safeParse(proposedValue);
  if (!valueParsed.success) {
    return NextResponse.json({ error: "Proposed value does not match the benefit's schema", details: valueParsed.error.flatten() }, { status: 400 });
  }

  const existingBenefits = await getBenefitsForCompany(companyId);
  const current = existingBenefits.find((b) => b.benefit_key === benefitKey && b.plan_year === planYear);

  const edit = await submitEdit({
    company_id: companyId,
    benefit_key: benefitKey,
    plan_year: planYear,
    proposed_value: valueParsed.data,
    current_value: current?.value ?? null,
    source_type: sourceType,
    source_url: sourceUrl || null,
    rationale: rationale ?? null,
    submitted_by: submittedBy,
    submitter_is_verified_employee: false,
  });

  return NextResponse.json({ edit });
}
