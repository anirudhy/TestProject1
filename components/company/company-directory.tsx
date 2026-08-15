"use client";

import Link from "next/link";
import { useMemo, useState } from "react";
import { Input } from "@/components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Card, CardContent } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { CompletenessMeter } from "@/components/company/completeness-meter";
import type { CompanyWithBenefits } from "@/lib/db/company-with-benefits";
import { calc401kMatch, calcTotalEmployerValue, type FourOhOneKMatchValue } from "@/lib/benefits/value";
import { BENEFIT_TYPES } from "@/lib/benefits/registry";

type SortKey = "completeness" | "value" | "name";

const REFERENCE_SALARY = 200_000;

export function CompanyDirectory({ items }: { items: CompanyWithBenefits[] }) {
  const [query, setQuery] = useState("");
  const [industry, setIndustry] = useState<string>("all");
  const [employeeBand, setEmployeeBand] = useState<string>("all");
  const [hasMbdr, setHasMbdr] = useState(false);
  const [hasEspp, setHasEspp] = useState(false);
  const [minMatchPercent, setMinMatchPercent] = useState<string>("");
  const [sort, setSort] = useState<SortKey>("completeness");

  const industries = useMemo(() => Array.from(new Set(items.map((i) => i.company.industry).filter(Boolean))) as string[], [items]);

  const enriched = useMemo(
    () =>
      items.map((item) => {
        const totalTypes = BENEFIT_TYPES.length;
        const known = new Set(item.benefits.map((b) => b.benefit_key)).size;
        const completeness = totalTypes === 0 ? 0 : known / totalTypes;

        const mbdr = item.benefits.find((b) => b.benefit_key === "mega_backdoor_roth");
        const mbdrSupported = Boolean((mbdr?.value as { supported?: boolean } | undefined)?.supported);

        const espp = item.benefits.find((b) => b.benefit_key === "espp");
        const esppOffered = Boolean((espp?.value as { offered?: boolean } | undefined)?.offered);

        const match401k = item.benefits.find((b) => b.benefit_key === "401k_match");
        const matchLine = match401k
          ? calc401kMatch(match401k.value as unknown as FourOhOneKMatchValue, REFERENCE_SALARY, 6, 2026)
          : null;
        const matchPercentOfContribution = match401k ? (match401k.value as { match_percent?: number }).match_percent ?? null : null;

        const referenceValue = calcTotalEmployerValue(item.benefits, {
          salaryUsd: REFERENCE_SALARY,
          contributionPercent: 6,
          familySize: 1,
          age: 30,
          planYear: 2026,
        }).totalUsd;

        return { ...item, completeness, mbdrSupported, esppOffered, matchLine, matchPercentOfContribution, referenceValue };
      }),
    [items]
  );

  const filtered = enriched
    .filter((i) => (query.trim() ? i.company.name.toLowerCase().includes(query.trim().toLowerCase()) : true))
    .filter((i) => (industry === "all" ? true : i.company.industry === industry))
    .filter((i) => (employeeBand === "all" ? true : i.company.employee_band === employeeBand))
    .filter((i) => (hasMbdr ? i.mbdrSupported : true))
    .filter((i) => (hasEspp ? i.esppOffered : true))
    .filter((i) => (minMatchPercent ? (i.matchPercentOfContribution ?? 0) >= Number(minMatchPercent) : true));

  const sorted = [...filtered].sort((a, b) => {
    if (sort === "name") return a.company.name.localeCompare(b.company.name);
    if (sort === "value") return b.referenceValue - a.referenceValue;
    return b.completeness - a.completeness;
  });

  return (
    <div>
      <div className="flex flex-wrap items-end gap-3 rounded-lg border border-border bg-card p-4">
        <div className="min-w-[200px] flex-1">
          <label className="mb-1 block text-xs text-muted-foreground">Search</label>
          <Input placeholder="Company name…" value={query} onChange={(e) => setQuery(e.target.value)} />
        </div>
        <div>
          <label className="mb-1 block text-xs text-muted-foreground">Industry</label>
          <Select value={industry} onValueChange={setIndustry}>
            <SelectTrigger className="w-[160px]"><SelectValue /></SelectTrigger>
            <SelectContent>
              <SelectItem value="all">All industries</SelectItem>
              {industries.map((ind) => (
                <SelectItem key={ind} value={ind}>{ind}</SelectItem>
              ))}
            </SelectContent>
          </Select>
        </div>
        <div>
          <label className="mb-1 block text-xs text-muted-foreground">Company size</label>
          <Select value={employeeBand} onValueChange={setEmployeeBand}>
            <SelectTrigger className="w-[150px]"><SelectValue /></SelectTrigger>
            <SelectContent>
              <SelectItem value="all">Any size</SelectItem>
              {["1-50", "51-500", "501-5000", "5001-50000", "50000+"].map((b) => (
                <SelectItem key={b} value={b}>{b} employees</SelectItem>
              ))}
            </SelectContent>
          </Select>
        </div>
        <div>
          <label className="mb-1 block text-xs text-muted-foreground">Min match %</label>
          <Input
            type="number"
            className="w-[100px]"
            value={minMatchPercent}
            onChange={(e) => setMinMatchPercent(e.target.value)}
            placeholder="e.g. 50"
          />
        </div>
        <label className="flex items-center gap-2 text-sm">
          <input type="checkbox" checked={hasMbdr} onChange={(e) => setHasMbdr(e.target.checked)} /> Has MBDR
        </label>
        <label className="flex items-center gap-2 text-sm">
          <input type="checkbox" checked={hasEspp} onChange={(e) => setHasEspp(e.target.checked)} /> Has ESPP
        </label>
        <div className="ml-auto">
          <label className="mb-1 block text-xs text-muted-foreground">Sort by</label>
          <Select value={sort} onValueChange={(v) => setSort(v as SortKey)}>
            <SelectTrigger className="w-[180px]"><SelectValue /></SelectTrigger>
            <SelectContent>
              <SelectItem value="completeness">Data completeness</SelectItem>
              <SelectItem value="value">Employer value ($200k)</SelectItem>
              <SelectItem value="name">Name</SelectItem>
            </SelectContent>
          </Select>
        </div>
      </div>

      <p className="mt-3 text-sm text-muted-foreground">{sorted.length} companies</p>

      <div className="mt-3 grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
        {sorted.map((item) => (
          <Link key={item.company.id} href={`/companies/${item.company.slug}`}>
            <Card className="h-full transition-shadow hover:shadow-md">
              <CardContent className="p-5">
                <div className="flex items-start justify-between">
                  <div>
                    <p className="font-semibold">{item.company.name}</p>
                    <p className="text-sm text-muted-foreground">{item.company.industry} · {item.company.employee_band}</p>
                  </div>
                  {item.mbdrSupported && <Badge variant="success">MBDR</Badge>}
                </div>
                <div className="mt-3 flex items-center justify-between text-sm">
                  <span className="text-muted-foreground">Est. value @ $200k</span>
                  <span className="font-medium">${item.referenceValue.toLocaleString()}</span>
                </div>
                <CompletenessMeter benefits={item.benefits} className="mt-3" />
              </CardContent>
            </Card>
          </Link>
        ))}
      </div>
    </div>
  );
}
