import { BENEFIT_TYPES } from "@/lib/benefits/registry";
import type { CompanyBenefit } from "@/types";
import { cn } from "@/lib/utils";

export function CompletenessMeter({ benefits, className }: { benefits: CompanyBenefit[]; className?: string }) {
  const total = BENEFIT_TYPES.length;
  const known = new Set(benefits.map((b) => b.benefit_key)).size;
  const percent = total === 0 ? 0 : Math.round((known / total) * 100);

  return (
    <div className={cn("flex items-center gap-2", className)}>
      <div className="h-1.5 w-24 overflow-hidden rounded-full bg-muted">
        <div className="h-full rounded-full bg-primary" style={{ width: `${percent}%` }} />
      </div>
      <span className="text-sm text-muted-foreground">
        {known}/{total} fields
      </span>
    </div>
  );
}
