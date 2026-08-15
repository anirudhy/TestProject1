"use client";

import { useMemo, useState } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Button } from "@/components/ui/button";
import { calcTotalEmployerValue } from "@/lib/benefits/value";
import { useCalculatorInputs } from "@/lib/hooks/use-calculator-inputs";
import { formatUsd } from "@/lib/utils";
import type { CompanyBenefit } from "@/types";

export function EmployerValueCard({ benefits }: { benefits: CompanyBenefit[] }) {
  const { inputs, setInputs, hydrated } = useCalculatorInputs();
  const [expanded, setExpanded] = useState(false);

  const result = useMemo(() => calcTotalEmployerValue(benefits, inputs), [benefits, inputs]);

  return (
    <Card>
      <CardHeader>
        <CardTitle>Estimated annual employer value</CardTitle>
      </CardHeader>
      <CardContent>
        <p className="text-4xl font-bold tracking-tight" suppressHydrationWarning>
          {hydrated ? formatUsd(result.totalUsd) : "—"}
        </p>
        <p className="mt-1 text-sm text-muted-foreground">
          Based on a ${inputs.salaryUsd.toLocaleString()} salary at a {inputs.contributionPercent}% 401(k) contribution rate.
        </p>

        <div className="mt-4 grid grid-cols-2 gap-3 sm:grid-cols-3">
          <div>
            <Label htmlFor="salary">Salary</Label>
            <Input
              id="salary"
              type="number"
              value={inputs.salaryUsd}
              onChange={(e) => setInputs({ ...inputs, salaryUsd: Number(e.target.value) || 0 })}
            />
          </div>
          <div>
            <Label htmlFor="contribution">401(k) contribution %</Label>
            <Input
              id="contribution"
              type="number"
              value={inputs.contributionPercent}
              onChange={(e) => setInputs({ ...inputs, contributionPercent: Number(e.target.value) || 0 })}
            />
          </div>
          <div>
            <Label htmlFor="family">Family size</Label>
            <Input
              id="family"
              type="number"
              min={1}
              value={inputs.familySize}
              onChange={(e) => setInputs({ ...inputs, familySize: Math.max(1, Number(e.target.value) || 1) })}
            />
          </div>
        </div>

        <Button variant="ghost" size="sm" className="mt-4" onClick={() => setExpanded((v) => !v)}>
          {expanded ? "Hide breakdown" : "Show breakdown"}
        </Button>

        {expanded && (
          <div className="mt-2 space-y-2 border-t border-border pt-3">
            {result.lines.map((line) => (
              <div key={line.benefitKey} className="text-sm">
                <div className="flex justify-between">
                  <span>{line.label}</span>
                  <span className="font-medium">{formatUsd(line.amountUsd)}</span>
                </div>
                <p className="text-xs text-muted-foreground">{line.formula}</p>
              </div>
            ))}
          </div>
        )}

        <p className="mt-4 text-xs text-muted-foreground">
          Estimates based on publicly reported plan terms. Not tax or investment advice. Verify against your own plan documents.
        </p>
      </CardContent>
    </Card>
  );
}
