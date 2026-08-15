/**
 * Seeds a Supabase project from lib/seed/companies.ts.
 * Requires NEXT_PUBLIC_SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY.
 * Run migrations first (`supabase db push` or apply supabase/migrations/*.sql
 * in order — 0006_benefit_types_seed.sql populates the benefit_types
 * registry that this script's foreign keys depend on).
 */
import { createClient } from "@supabase/supabase-js";
import { SEED_COMPANIES } from "./companies";

async function main() {
  const url = process.env.NEXT_PUBLIC_SUPABASE_URL;
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!url || !key) {
    console.error("NEXT_PUBLIC_SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY must be set to seed a real database.");
    process.exit(1);
  }
  const sb = createClient(url, key, { auth: { persistSession: false } });

  for (const sc of SEED_COMPANIES) {
    console.log(`Seeding ${sc.name}...`);
    const { data: company, error: companyErr } = await sb
      .from("companies")
      .upsert(
        {
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
          // Published in a second pass, after every benefit row has a
          // source — the guard_company_publish trigger otherwise rejects it.
          is_published: false,
        },
        { onConflict: "slug" }
      )
      .select()
      .single();
    if (companyErr) throw companyErr;

    for (const b of sc.benefits) {
      const { data: benefit, error: benefitErr } = await sb
        .from("company_benefits")
        .upsert(
          {
            company_id: company.id,
            benefit_key: b.benefit_key,
            plan_year: b.plan_year,
            value: b.value,
            notes: b.notes ?? null,
            confidence: b.confidence,
            last_verified_at: b.last_verified_at,
          },
          { onConflict: "company_id,benefit_key,plan_year" }
        )
        .select()
        .single();
      if (benefitErr) throw benefitErr;

      for (const s of b.sources) {
        const { error: sourceErr } = await sb.from("benefit_sources").insert({
          company_benefit_id: benefit.id,
          source_type: s.source_type,
          url: s.url || null,
          title: s.title,
          retrieved_at: s.retrieved_at,
          excerpt: s.excerpt ?? null,
        });
        if (sourceErr) throw sourceErr;
      }
    }

    if (sc.is_published) {
      const { error: publishErr } = await sb.from("companies").update({ is_published: true }).eq("id", company.id);
      if (publishErr) throw publishErr;
    }
  }

  console.log(`Done. Seeded ${SEED_COMPANIES.length} companies.`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
