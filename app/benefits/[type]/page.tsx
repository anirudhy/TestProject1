import type { Metadata } from "next";
import { notFound } from "next/navigation";
import Link from "next/link";
import { BENEFIT_TYPES, getBenefitType } from "@/lib/benefits/registry";
import { getLeaderboard } from "@/lib/db/queries";
import { compareScore } from "@/lib/benefits/compare-score";
import { explainerParagraphs } from "@/lib/benefits/explainer";
import { formatFieldValue } from "@/components/benefits/format-value";
import { ConfidenceBadge } from "@/components/benefits/confidence-badge";
import { Card, CardContent } from "@/components/ui/card";

export const revalidate = 3600;

export async function generateStaticParams() {
  return BENEFIT_TYPES.map((t) => ({ type: t.key }));
}

export async function generateMetadata({ params }: { params: Promise<{ type: string }> }): Promise<Metadata> {
  const { type: typeKey } = await params;
  const type = getBenefitType(typeKey);
  if (!type) return {};
  return {
    title: `Companies with the best ${type.label.toLowerCase()}`,
    description: `Ranked, sourced comparison of ${type.label.toLowerCase()} across companies in the PerkStack directory.`,
  };
}

export default async function BenefitLeaderboardPage({ params }: { params: Promise<{ type: string }> }) {
  const { type: typeKey } = await params;
  const type = getBenefitType(typeKey);
  if (!type) notFound();

  const rows = await getLeaderboard(typeKey);
  const ranked = rows
    .map((r) => ({ ...r, score: compareScore(typeKey, r.benefit.value as Record<string, unknown>) }))
    .sort((a, b) => (b.score ?? -Infinity) - (a.score ?? -Infinity));

  const jsonLd = {
    "@context": "https://schema.org",
    "@type": "FAQPage",
    mainEntity: [
      {
        "@type": "Question",
        name: `What is ${type.label}?`,
        acceptedAnswer: { "@type": "Answer", text: type.description },
      },
    ],
  };

  return (
    <div className="mx-auto max-w-4xl px-4 py-10">
      {/* eslint-disable-next-line react/no-danger */}
      <script type="application/ld+json" dangerouslySetInnerHTML={{ __html: JSON.stringify(jsonLd) }} />

      <p className="text-sm text-muted-foreground">
        <Link href="/companies" className="hover:underline">Companies</Link> / Leaderboards
      </p>
      <h1 className="mt-1 text-3xl font-bold tracking-tight">Companies with the best {type.label.toLowerCase()}</h1>
      <p className="mt-2 text-muted-foreground">{type.description}</p>

      <div className="mt-6 space-y-3">
        {ranked.length === 0 && <p className="text-muted-foreground">No sourced data yet for this benefit.</p>}
        {ranked.map(({ company, benefit }, i) => (
          <Card key={company.id}>
            <CardContent className="flex items-start justify-between gap-4 p-5">
              <div>
                <p className="text-xs text-muted-foreground">#{i + 1}</p>
                <Link href={`/companies/${company.slug}`} className="font-semibold hover:underline">{company.name}</Link>
                <dl className="mt-2 grid gap-x-6 gap-y-1 text-sm sm:grid-cols-2">
                  {type.fields
                    .map((field) => ({ field, formatted: formatFieldValue(field, (benefit.value as Record<string, unknown>)[field.key]) }))
                    .filter((r) => r.formatted !== null)
                    .slice(0, 6)
                    .map(({ field, formatted }) => (
                      <div key={field.key} className="flex justify-between gap-2">
                        <dt className="text-muted-foreground">{field.label}</dt>
                        <dd className="font-medium">{formatted}</dd>
                      </div>
                    ))}
                </dl>
              </div>
              <ConfidenceBadge confidence={benefit.confidence} />
            </CardContent>
          </Card>
        ))}
      </div>

      <div className="mt-10 max-w-none space-y-4 text-sm leading-relaxed text-muted-foreground">
        <h2 className="text-lg font-semibold text-foreground">How {type.label.toLowerCase()} works</h2>
        {explainerParagraphs(type).map((p, i) => (
          <p key={i}>{p}</p>
        ))}
      </div>
    </div>
  );
}
