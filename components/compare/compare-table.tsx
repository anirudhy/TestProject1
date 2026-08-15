import { Fragment } from "react";
import Link from "next/link";
import { BENEFIT_CATEGORIES } from "@/types";
import { BENEFIT_TYPES } from "@/lib/benefits/registry";
import { formatFieldValue } from "@/components/benefits/format-value";
import { compareScore } from "@/lib/benefits/compare-score";
import { ConfidenceBadge } from "@/components/benefits/confidence-badge";
import type { Company, CompanyBenefit } from "@/types";
import { cn } from "@/lib/utils";

export interface CompareCompany {
  company: Company;
  benefits: CompanyBenefit[];
}

export function CompareTable({ companies }: { companies: CompareCompany[] }) {
  const comparableTypes = BENEFIT_TYPES.filter((t) => t.isComparable).sort((a, b) => a.sortOrder - b.sortOrder);

  return (
    <div className="overflow-x-auto">
      <table className="w-full border-collapse text-sm">
        <thead>
          <tr>
            <th className="sticky left-0 z-10 min-w-[220px] border-b border-border bg-background p-3 text-left align-bottom" />
            {companies.map(({ company }) => (
              <th key={company.id} className="min-w-[220px] border-b border-border p-3 text-left align-bottom">
                <Link href={`/companies/${company.slug}`} className="font-semibold hover:underline">
                  {company.name}
                </Link>
                <p className="text-xs font-normal text-muted-foreground">{company.industry}</p>
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {BENEFIT_CATEGORIES.map((category) => {
            const types = comparableTypes.filter((t) => t.category === category.key);
            if (types.length === 0) return null;
            return (
              <Fragment key={category.key}>
                <tr>
                  <td
                    colSpan={companies.length + 1}
                    className="sticky left-0 border-b border-border bg-muted/60 p-2 text-xs font-semibold uppercase tracking-wide text-muted-foreground"
                  >
                    {category.label}
                  </td>
                </tr>
                {types.map((type) => {
                  const cells = companies.map(({ benefits }) => benefits.find((b) => b.benefit_key === type.key));
                  const scores = cells.map((b) => (b ? compareScore(type.key, b.value as Record<string, unknown>) : null));
                  const validScores = scores.filter((s): s is number => s !== null && Number.isFinite(s));
                  const maxScore = validScores.length > 1 ? Math.max(...validScores) : null;

                  return (
                    <tr key={type.key} className="border-b border-border/60">
                      <td className="sticky left-0 z-10 bg-background p-3 font-medium">{type.label}</td>
                      {companies.map(({ company }, i) => {
                        const benefit = cells[i];
                        const score = scores[i];
                        const isBest = maxScore !== null && score !== null && score === maxScore && score > 0;

                        if (!benefit) {
                          return (
                            <td key={company.id} className="p-3 text-muted-foreground">
                              Unknown
                            </td>
                          );
                        }

                        const rows = type.fields
                          .map((field) => formatFieldValue(field, (benefit.value as Record<string, unknown>)[field.key]))
                          .filter((v): v is string => v !== null);

                        return (
                          <td key={company.id} className={cn("p-3 align-top", isBest && "bg-emerald-50 dark:bg-emerald-950/30")}>
                            <div className="space-y-0.5">
                              {rows.slice(0, 3).map((r, idx) => (
                                <p key={idx} className={idx === 0 ? "font-medium" : "text-xs text-muted-foreground"}>{r}</p>
                              ))}
                            </div>
                            <div className="mt-1.5">
                              <ConfidenceBadge confidence={benefit.confidence} />
                            </div>
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
              </Fragment>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}
