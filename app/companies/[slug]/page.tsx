import type { Metadata } from "next";
import { notFound } from "next/navigation";
import Link from "next/link";
import { getCompanyBySlug, getBenefitsForCompany, listAllCompanySlugs } from "@/lib/db/queries";
import { CategoryAccordion } from "@/components/company/category-accordion";
import { CompletenessMeter } from "@/components/company/completeness-meter";
import { EmployerValueCard } from "@/components/company/employer-value-card";
import { Badge } from "@/components/ui/badge";
import { formatDate } from "@/lib/utils";
import { getBenefitType } from "@/lib/benefits/registry";

export const revalidate = 3600;

export async function generateStaticParams() {
  const slugs = await listAllCompanySlugs();
  return slugs.map((slug) => ({ slug }));
}

export async function generateMetadata({ params }: { params: Promise<{ slug: string }> }): Promise<Metadata> {
  const { slug } = await params;
  const company = await getCompanyBySlug(slug);
  if (!company) return {};
  return {
    title: company.name,
    description: `${company.name} benefits: 401(k) match, Mega Backdoor Roth, ESPP, and more — structured and sourced.`,
  };
}

export default async function CompanyPage({ params }: { params: Promise<{ slug: string }> }) {
  const { slug } = await params;
  const company = await getCompanyBySlug(slug);
  if (!company) notFound();

  const benefits = await getBenefitsForCompany(company.id);
  const lastVerified = benefits
    .map((b) => b.last_verified_at)
    .filter((d): d is string => Boolean(d))
    .sort()
    .reverse()[0];

  const faqEntries = benefits
    .filter((b) => ["401k_match", "mega_backdoor_roth", "espp"].includes(b.benefit_key))
    .map((b) => {
      const type = getBenefitType(b.benefit_key);
      return {
        "@type": "Question",
        name: `Does ${company.name} support ${type?.label ?? b.benefit_key}?`,
        acceptedAnswer: { "@type": "Answer", text: b.notes ?? `See the ${type?.label ?? b.benefit_key} section on this page for full sourced details.` },
      };
    });

  const jsonLd = [
    {
      "@context": "https://schema.org",
      "@type": "Organization",
      name: company.name,
      url: company.careers_url,
      ...(company.logo_url ? { logo: company.logo_url } : {}),
    },
    ...(faqEntries.length > 0 ? [{ "@context": "https://schema.org", "@type": "FAQPage", mainEntity: faqEntries }] : []),
  ];

  return (
    <div className="mx-auto max-w-4xl px-4 py-10">
      {/* eslint-disable-next-line react/no-danger */}
      <script type="application/ld+json" dangerouslySetInnerHTML={{ __html: JSON.stringify(jsonLd) }} />

      <div className="flex flex-wrap items-start justify-between gap-4">
        <div>
          <p className="text-sm text-muted-foreground">
            <Link href="/companies" className="hover:underline">Companies</Link> / {company.name}
          </p>
          <h1 className="mt-1 text-3xl font-bold tracking-tight">{company.name}</h1>
          <div className="mt-2 flex flex-wrap items-center gap-2 text-sm text-muted-foreground">
            {company.industry && <Badge variant="outline">{company.industry}</Badge>}
            {company.employee_band && <Badge variant="outline">{company.employee_band} employees</Badge>}
            {lastVerified && <span>Data last verified {formatDate(lastVerified)}</span>}
          </div>
        </div>
        <div className="flex flex-col items-end gap-2">
          <CompletenessMeter benefits={benefits} />
          {company.careers_url && (
            <a href={company.careers_url} target="_blank" rel="noopener noreferrer" className="text-sm text-primary hover:underline">
              Careers page ↗
            </a>
          )}
        </div>
      </div>

      <div className="mt-6">
        <EmployerValueCard benefits={benefits} />
      </div>

      <div className="mt-8">
        <h2 className="text-lg font-semibold">Benefits</h2>
        <CategoryAccordion benefits={benefits} companySlug={company.slug} />
      </div>
    </div>
  );
}
