import type { Metadata } from "next";
import { getBenefitsForCompany, getCompanyBySlug, listPublishedCompanies } from "@/lib/db/queries";
import { ComparePicker } from "@/components/compare/compare-picker";
import { CompareTable, type CompareCompany } from "@/components/compare/compare-table";

function parseSlugs(c: string | undefined): string[] {
  if (!c) return [];
  return Array.from(new Set(c.split(",").map((s) => s.trim()).filter(Boolean))).slice(0, 4);
}

export async function generateMetadata({ searchParams }: { searchParams: Promise<{ c?: string }> }): Promise<Metadata> {
  const { c } = await searchParams;
  const slugs = parseSlugs(c);
  const title = slugs.length > 0 ? `Compare: ${slugs.join(" vs ")}` : "Compare companies";
  const description = "Side-by-side, sourced benefits comparison — 401(k) match, Mega Backdoor Roth, ESPP, and more.";
  return {
    title,
    description,
    openGraph: {
      title,
      description,
      images: slugs.length > 0 ? [`/api/og?c=${encodeURIComponent(slugs.join(","))}`] : undefined,
    },
  };
}

export default async function ComparePage({ searchParams }: { searchParams: Promise<{ c?: string }> }) {
  const { c } = await searchParams;
  const slugs = parseSlugs(c);

  const allCompanies = await listPublishedCompanies();

  const resolved = await Promise.all(
    slugs.map(async (slug) => {
      const company = await getCompanyBySlug(slug);
      if (!company) return null;
      const benefits = await getBenefitsForCompany(company.id);
      return { company, benefits } satisfies CompareCompany;
    })
  );
  const companies = resolved.filter((c): c is CompareCompany => c !== null);

  return (
    <div className="mx-auto max-w-6xl px-4 py-10">
      <h1 className="text-2xl font-bold tracking-tight">Compare</h1>
      <p className="mt-1 text-muted-foreground">Pick 2–4 companies. The URL updates as you go, so it's shareable.</p>

      <div className="mt-6">
        <ComparePicker allCompanies={allCompanies} selectedSlugs={companies.map((c) => c.company.slug)} />
      </div>

      <div className="mt-8">
        {companies.length === 0 ? (
          <p className="text-muted-foreground">Add companies above to start comparing.</p>
        ) : (
          <CompareTable companies={companies} />
        )}
      </div>

      {companies.length > 0 && (
        <p className="mt-6 text-xs text-muted-foreground">
          Highlighted cells mark the best value in each row where an unambiguous comparison is possible. "Unknown"
          means no sourced data exists yet — not that the benefit isn't offered.
        </p>
      )}
    </div>
  );
}
