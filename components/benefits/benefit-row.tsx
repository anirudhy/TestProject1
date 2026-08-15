import Link from "next/link";
import { ExternalLink, HelpCircle } from "lucide-react";
import type { BenefitTypeDef } from "@/lib/benefits/registry";
import type { CompanyBenefit } from "@/types";
import { formatFieldValue } from "./format-value";
import { ConfidenceBadge } from "./confidence-badge";
import { formatDate } from "@/lib/utils";

export function BenefitRow({ type, benefit, companySlug }: { type: BenefitTypeDef; benefit?: CompanyBenefit; companySlug: string }) {
  if (!benefit) {
    return (
      <div className="flex items-start justify-between gap-4 py-4">
        <div>
          <p className="font-medium text-foreground">{type.label}</p>
          <p className="mt-1 flex items-center gap-1.5 text-sm text-muted-foreground">
            <HelpCircle className="h-3.5 w-3.5" /> Unknown — contribute this
          </p>
        </div>
        <Link
          href={`/contribute?company=${companySlug}&benefit=${type.key}`}
          className="whitespace-nowrap text-sm text-primary hover:underline"
        >
          Add data
        </Link>
      </div>
    );
  }

  const rows = type.fields
    .map((field) => ({ field, formatted: formatFieldValue(field, (benefit.value as Record<string, unknown>)[field.key]) }))
    .filter((r) => r.formatted !== null);

  return (
    <div className="py-4">
      <div className="flex flex-wrap items-start justify-between gap-2">
        <p className="font-medium text-foreground">{type.label}</p>
        <div className="flex items-center gap-2">
          <ConfidenceBadge confidence={benefit.confidence} />
          <Link href={`/contribute?company=${companySlug}&benefit=${type.key}`} className="text-sm text-primary hover:underline">
            Suggest an edit
          </Link>
        </div>
      </div>

      <dl className="mt-2 grid gap-x-6 gap-y-1 text-sm sm:grid-cols-2">
        {rows.map(({ field, formatted }) => (
          <div key={field.key} className="flex justify-between gap-2 border-b border-border/60 py-1 sm:border-none sm:py-0">
            <dt className="text-muted-foreground">{field.label}</dt>
            <dd className="text-right font-medium">{formatted}</dd>
          </div>
        ))}
      </dl>

      {benefit.notes && <p className="mt-2 text-sm italic text-muted-foreground">"{benefit.notes}"</p>}

      <div className="mt-2 flex flex-wrap items-center gap-x-4 gap-y-1 text-xs text-muted-foreground">
        {benefit.last_verified_at && <span>Last verified {formatDate(benefit.last_verified_at)}</span>}
        {benefit.sources?.map((s) =>
          s.url ? (
            <a key={s.id} href={s.url} target="_blank" rel="noopener noreferrer" className="inline-flex items-center gap-1 hover:text-foreground hover:underline">
              {s.title ?? "Source"} <ExternalLink className="h-3 w-3" />
            </a>
          ) : (
            <span key={s.id}>{s.title}</span>
          )
        )}
      </div>
    </div>
  );
}
