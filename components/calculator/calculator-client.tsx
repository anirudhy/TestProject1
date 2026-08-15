"use client";

import { useMemo, useState } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { useCalculatorInputs, type FilingStatus } from "@/lib/hooks/use-calculator-inputs";
import { calcTotalEmployerValue } from "@/lib/benefits/value";
import { formatUsd } from "@/lib/utils";
import type { CompanyWithBenefits } from "@/lib/db/company-with-benefits";

export function CalculatorClient({ companies }: { companies: CompanyWithBenefits[] }) {
  const { inputs, setInputs, hydrated } = useCalculatorInputs();
  const [selectedSlugs, setSelectedSlugs] = useState<string[]>([]);

  const selected = companies.filter((c) => selectedSlugs.includes(c.company.slug));
  const results = useMemo(
    () => selected.map((c) => ({ company: c.company, result: calcTotalEmployerValue(c.benefits, inputs) })),
    [selected, inputs]
  );

  return (
    <div className="grid gap-6 lg:grid-cols-[340px_1fr]">
      <Card className="h-fit">
        <CardHeader><CardTitle>Your inputs</CardTitle></CardHeader>
        <CardContent className="space-y-4">
          <div>
            <Label htmlFor="salary">Base salary</Label>
            <Input id="salary" type="number" value={inputs.salaryUsd} onChange={(e) => setInputs({ ...inputs, salaryUsd: Number(e.target.value) || 0 })} />
          </div>
          <div>
            <Label htmlFor="bonus">Bonus %</Label>
            <Input id="bonus" type="number" value={inputs.bonusPercent} onChange={(e) => setInputs({ ...inputs, bonusPercent: Number(e.target.value) || 0 })} />
          </div>
          <div>
            <Label htmlFor="contribution">401(k) contribution %</Label>
            <Input id="contribution" type="number" value={inputs.contributionPercent} onChange={(e) => setInputs({ ...inputs, contributionPercent: Number(e.target.value) || 0 })} />
          </div>
          <div>
            <Label htmlFor="filing">Filing status</Label>
            <Select value={inputs.filingStatus} onValueChange={(v) => setInputs({ ...inputs, filingStatus: v as FilingStatus })}>
              <SelectTrigger id="filing"><SelectValue /></SelectTrigger>
              <SelectContent>
                <SelectItem value="single">Single</SelectItem>
                <SelectItem value="married_joint">Married filing jointly</SelectItem>
                <SelectItem value="married_separate">Married filing separately</SelectItem>
                <SelectItem value="head_of_household">Head of household</SelectItem>
              </SelectContent>
            </Select>
          </div>
          <div>
            <Label htmlFor="family">Family size</Label>
            <Input id="family" type="number" min={1} value={inputs.familySize} onChange={(e) => setInputs({ ...inputs, familySize: Math.max(1, Number(e.target.value) || 1) })} />
          </div>
          <div>
            <Label htmlFor="age">Age</Label>
            <Input id="age" type="number" value={inputs.age} onChange={(e) => setInputs({ ...inputs, age: Number(e.target.value) || 0 })} />
          </div>
          <p className="text-xs text-muted-foreground">Saved automatically in this browser and reused on company pages.</p>
        </CardContent>
      </Card>

      <div>
        <Card>
          <CardHeader><CardTitle>Compare companies</CardTitle></CardHeader>
          <CardContent>
            <Select onValueChange={(slug) => setSelectedSlugs((prev) => (prev.includes(slug) ? prev : [...prev, slug].slice(-3)))}>
              <SelectTrigger className="w-[240px]"><SelectValue placeholder="Add a company (up to 3)…" /></SelectTrigger>
              <SelectContent>
                {companies.map((c) => (
                  <SelectItem key={c.company.slug} value={c.company.slug}>{c.company.name}</SelectItem>
                ))}
              </SelectContent>
            </Select>

            {results.length === 0 ? (
              <p className="mt-4 text-sm text-muted-foreground">Add a company to see its estimated employer value with your inputs.</p>
            ) : (
              <div className="mt-4 grid gap-4 md:grid-cols-3">
                {results.map(({ company, result }) => (
                  <div key={company.slug} className="rounded-lg border border-border p-4">
                    <div className="flex items-center justify-between">
                      <p className="font-semibold">{company.name}</p>
                      <button
                        className="text-xs text-muted-foreground hover:text-foreground"
                        onClick={() => setSelectedSlugs((prev) => prev.filter((s) => s !== company.slug))}
                      >
                        Remove
                      </button>
                    </div>
                    <p className="mt-2 text-2xl font-bold" suppressHydrationWarning>{hydrated ? formatUsd(result.totalUsd) : "—"}</p>
                    <div className="mt-3 space-y-1.5">
                      {result.lines.map((line) => (
                        <div key={line.benefitKey} className="flex justify-between text-xs">
                          <span className="text-muted-foreground">{line.label}</span>
                          <span className="font-medium">{formatUsd(line.amountUsd)}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </CardContent>
        </Card>

        <p className="mt-4 text-xs text-muted-foreground">
          Estimates based on publicly reported plan terms. Not tax or investment advice. Verify against your own plan
          documents.
        </p>
      </div>
    </div>
  );
}
