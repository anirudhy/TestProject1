"use client";

import { useMemo, useState } from "react";
import { useRouter, useSearchParams } from "next/navigation";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Label } from "@/components/ui/label";
import { Input } from "@/components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Button } from "@/components/ui/button";
import { DynamicFieldInput, type FieldValue } from "./dynamic-field-input";
import { BENEFIT_TYPES, getBenefitType } from "@/lib/benefits/registry";
import { formatFieldValue } from "@/components/benefits/format-value";
import type { CompanyWithBenefits } from "@/lib/db/company-with-benefits";
import type { SourceType } from "@/types";

const SOURCE_TYPES: { value: SourceType; label: string }[] = [
  { value: "company_public_page", label: "Company public page" },
  { value: "sec_filing", label: "SEC filing (S-8, proxy)" },
  { value: "form_5500", label: "Form 5500" },
  { value: "press", label: "Press release" },
  { value: "news", label: "News article" },
  { value: "user_report", label: "I work here (employee report)" },
];

export function ContributeForm({ companies }: { companies: CompanyWithBenefits[] }) {
  const router = useRouter();
  const searchParams = useSearchParams();

  const [companySlug, setCompanySlug] = useState(searchParams.get("company") ?? "");
  const [benefitKey, setBenefitKey] = useState(searchParams.get("benefit") ?? "");
  const [formValues, setFormValues] = useState<Record<string, FieldValue>>({});
  const [sourceType, setSourceType] = useState<SourceType>("company_public_page");
  const [sourceUrl, setSourceUrl] = useState("");
  const [rationale, setRationale] = useState("");
  const [status, setStatus] = useState<"idle" | "submitting" | "done" | "error">("idle");
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  const selectedCompany = companies.find((c) => c.company.slug === companySlug);
  const type = getBenefitType(benefitKey);
  const currentBenefit = selectedCompany?.benefits.find((b) => b.benefit_key === benefitKey);

  const currentValueDisplay = useMemo(() => {
    if (!type || !currentBenefit) return null;
    return type.fields
      .map((f) => ({ f, v: formatFieldValue(f, (currentBenefit.value as Record<string, unknown>)[f.key]) }))
      .filter((r) => r.v !== null);
  }, [type, currentBenefit]);

  function selectBenefit(key: string) {
    setBenefitKey(key);
    const existing = selectedCompany?.benefits.find((b) => b.benefit_key === key);
    setFormValues((existing?.value as Record<string, FieldValue>) ?? {});
  }

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!selectedCompany || !type) return;
    setStatus("submitting");
    setErrorMessage(null);

    const res = await fetch("/api/edits", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        companyId: selectedCompany.company.id,
        benefitKey: type.key,
        planYear: new Date().getFullYear(),
        proposedValue: formValues,
        sourceType,
        sourceUrl,
        rationale,
      }),
    });

    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      setErrorMessage(data.error ?? "Something went wrong.");
      setStatus("error");
      return;
    }

    setStatus("done");
    router.refresh();
  }

  if (status === "done") {
    return (
      <Card>
        <CardContent className="p-8 text-center">
          <p className="text-lg font-semibold">Thanks — submitted for review.</p>
          <p className="mt-1 text-muted-foreground">
            A moderator will review it against your source before it goes live. Fast-track sources (company pages,
            SEC filings, Form 5500) usually clear quickly.
          </p>
          <Button className="mt-4" variant="outline" onClick={() => setStatus("idle")}>
            Submit another
          </Button>
        </CardContent>
      </Card>
    );
  }

  return (
    <form onSubmit={handleSubmit} className="space-y-6">
      <Card>
        <CardHeader><CardTitle>1. Company & benefit</CardTitle></CardHeader>
        <CardContent className="grid gap-4 sm:grid-cols-2">
          <div>
            <Label>Company</Label>
            <Select value={companySlug} onValueChange={(v) => { setCompanySlug(v); setBenefitKey(""); setFormValues({}); }}>
              <SelectTrigger><SelectValue placeholder="Select a company…" /></SelectTrigger>
              <SelectContent>
                {companies.map((c) => (
                  <SelectItem key={c.company.slug} value={c.company.slug}>{c.company.name}</SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
          <div>
            <Label>Benefit</Label>
            <Select value={benefitKey} onValueChange={selectBenefit} disabled={!companySlug}>
              <SelectTrigger><SelectValue placeholder="Select a benefit…" /></SelectTrigger>
              <SelectContent>
                {BENEFIT_TYPES.map((t) => (
                  <SelectItem key={t.key} value={t.key}>{t.label}</SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
        </CardContent>
      </Card>

      {type && selectedCompany && (
        <>
          {currentValueDisplay && currentValueDisplay.length > 0 && (
            <Card>
              <CardHeader><CardTitle className="text-base">Current value on file</CardTitle></CardHeader>
              <CardContent>
                <dl className="grid gap-x-6 gap-y-1 text-sm sm:grid-cols-2">
                  {currentValueDisplay.map(({ f, v }) => (
                    <div key={f.key} className="flex justify-between gap-2">
                      <dt className="text-muted-foreground">{f.label}</dt>
                      <dd className="font-medium">{v}</dd>
                    </div>
                  ))}
                </dl>
              </CardContent>
            </Card>
          )}

          <Card>
            <CardHeader><CardTitle>2. Proposed value</CardTitle></CardHeader>
            <CardContent className="grid gap-4 sm:grid-cols-2">
              {type.fields.map((field) => (
                <DynamicFieldInput
                  key={field.key}
                  field={field}
                  value={formValues[field.key]}
                  onChange={(v) => setFormValues((prev) => ({ ...prev, [field.key]: v }))}
                />
              ))}
            </CardContent>
          </Card>

          <Card>
            <CardHeader><CardTitle>3. Source</CardTitle></CardHeader>
            <CardContent className="grid gap-4 sm:grid-cols-2">
              <div>
                <Label>Source type</Label>
                <Select value={sourceType} onValueChange={(v) => setSourceType(v as SourceType)}>
                  <SelectTrigger><SelectValue /></SelectTrigger>
                  <SelectContent>
                    {SOURCE_TYPES.map((s) => (
                      <SelectItem key={s.value} value={s.value}>{s.label}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              <div>
                <Label htmlFor="source-url">Source URL {sourceType !== "user_report" && <span className="text-destructive">*</span>}</Label>
                <Input id="source-url" type="url" value={sourceUrl} onChange={(e) => setSourceUrl(e.target.value)} placeholder="https://…" />
              </div>
              <div className="sm:col-span-2">
                <Label htmlFor="rationale">Rationale / notes</Label>
                <textarea
                  id="rationale"
                  className="mt-1 w-full rounded-md border border-input bg-background px-3 py-2 text-sm shadow-sm"
                  rows={3}
                  value={rationale}
                  onChange={(e) => setRationale(e.target.value)}
                  placeholder="What changed, and why should this replace the current value?"
                />
              </div>
            </CardContent>
          </Card>

          {errorMessage && <p className="text-sm text-destructive">{errorMessage}</p>}

          <Button type="submit" disabled={status === "submitting"}>
            {status === "submitting" ? "Submitting…" : "Submit for review"}
          </Button>
        </>
      )}
    </form>
  );
}
