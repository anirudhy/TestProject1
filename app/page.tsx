import Link from "next/link";
import { listCompaniesWithBenefits } from "@/lib/db/company-with-benefits";
import { Card, CardContent } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { CompletenessMeter } from "@/components/company/completeness-meter";
import { calcTotalEmployerValue } from "@/lib/benefits/value";

export const revalidate = 3600;

export default async function HomePage() {
  const items = await listCompaniesWithBenefits();
  const featured = items.slice(0, 6);

  return (
    <div>
      <section className="border-b border-border bg-gradient-to-b from-secondary/60 to-background">
        <div className="mx-auto max-w-4xl px-4 py-20 text-center">
          <h1 className="text-4xl font-bold tracking-tight sm:text-5xl">
            Comp is what they pay you.
            <br />
            <span className="text-primary">Benefits are what they pay you that you forgot to claim.</span>
          </h1>
          <p className="mx-auto mt-4 max-w-2xl text-lg text-muted-foreground">
            A structured, sourced directory of company benefits — 401(k) match formulas, Mega Backdoor Roth support,
            ESPP terms, and more. Every number cited, every date tracked.
          </p>
          <div className="mt-8 flex justify-center gap-3">
            <Button asChild size="lg"><Link href="/companies">Browse companies</Link></Button>
            <Button asChild size="lg" variant="outline"><Link href="/calculator">Try the calculator</Link></Button>
          </div>
        </div>
      </section>

      <section className="mx-auto max-w-6xl px-4 py-14">
        <div className="flex items-center justify-between">
          <h2 className="text-xl font-semibold">Featured companies</h2>
          <Link href="/companies" className="text-sm text-primary hover:underline">View all →</Link>
        </div>
        <div className="mt-6 grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
          {featured.map(({ company, benefits }) => {
            const value = calcTotalEmployerValue(benefits, { salaryUsd: 200_000, contributionPercent: 6, familySize: 1, age: 30, planYear: 2026 });
            return (
              <Link key={company.id} href={`/companies/${company.slug}`}>
                <Card className="h-full transition-shadow hover:shadow-md">
                  <CardContent className="p-5">
                    <div className="flex items-start justify-between">
                      <p className="font-semibold">{company.name}</p>
                      {benefits.some((b) => b.benefit_key === "mega_backdoor_roth" && (b.value as { supported?: boolean }).supported) && (
                        <Badge variant="success">MBDR</Badge>
                      )}
                    </div>
                    <p className="text-sm text-muted-foreground">{company.industry}</p>
                    <p className="mt-3 text-sm">
                      Est. value @ $200k: <span className="font-medium">${value.totalUsd.toLocaleString()}</span>
                    </p>
                    <CompletenessMeter benefits={benefits} className="mt-3" />
                  </CardContent>
                </Card>
              </Link>
            );
          })}
        </div>
      </section>

      <section className="border-t border-border bg-secondary/40">
        <div className="mx-auto max-w-4xl px-4 py-14 text-center">
          <h2 className="text-xl font-semibold">Compare 2–3 companies side by side</h2>
          <p className="mt-2 text-muted-foreground">
            Same fields, same salary assumptions. See exactly where the numbers differ — and where the data is
            missing rather than guessed at.
          </p>
          <Button asChild className="mt-6"><Link href="/compare">Open compare</Link></Button>
        </div>
      </section>
    </div>
  );
}
