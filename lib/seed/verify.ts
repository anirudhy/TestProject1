/**
 * Reports every company_benefits field whose last_verified_at is older
 * than 12 months, per §5's mandate to track median data age. Run with
 * `pnpm seed:verify` against a real Supabase project.
 */
import { createClient } from "@supabase/supabase-js";

const STALE_MONTHS = 12;

async function main() {
  const url = process.env.NEXT_PUBLIC_SUPABASE_URL;
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!url || !key) {
    console.error("NEXT_PUBLIC_SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY must be set.");
    process.exit(1);
  }
  const sb = createClient(url, key, { auth: { persistSession: false } });

  const { data, error } = await sb
    .from("company_benefits")
    .select("benefit_key, last_verified_at, plan_year, companies(name, slug)")
    .order("last_verified_at", { ascending: true });
  if (error) throw error;

  const rows = data as unknown as {
    benefit_key: string;
    last_verified_at: string | null;
    plan_year: number;
    companies: { name: string; slug: string } | null;
  }[];

  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - STALE_MONTHS);

  const stale = rows.filter((r) => !r.last_verified_at || new Date(r.last_verified_at) < cutoff);
  const ages = rows
    .filter((r) => r.last_verified_at)
    .map((r) => (Date.now() - new Date(r.last_verified_at!).getTime()) / (1000 * 60 * 60 * 24 * 30));
  ages.sort((a, b) => a - b);
  const median = ages.length ? ages[Math.floor(ages.length / 2)] ?? 0 : 0;

  console.log(`Total fields: ${rows.length}`);
  console.log(`Median data age: ${median.toFixed(1)} months`);
  console.log(`Stale fields (>${STALE_MONTHS} months or never verified): ${stale.length}`);
  for (const r of stale) {
    console.log(
      `  - ${r.companies?.name ?? "?"} (${r.companies?.slug ?? "?"}) / ${r.benefit_key} / plan_year ${r.plan_year} — last verified ${r.last_verified_at ?? "never"}`
    );
  }

  if (stale.length > 0) process.exitCode = 1;
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
