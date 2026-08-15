import type { Metadata } from "next";
import { listCompaniesWithBenefits } from "@/lib/db/company-with-benefits";
import { CalculatorClient } from "@/components/calculator/calculator-client";

export const metadata: Metadata = {
  title: "Calculator",
  description: "Estimate the total dollar value of a company's benefits — 401(k) match, MBDR headroom, HSA, ESPP, and stipends.",
};

export const revalidate = 3600;

export default async function CalculatorPage() {
  const companies = await listCompaniesWithBenefits();

  return (
    <div className="mx-auto max-w-6xl px-4 py-10">
      <h1 className="text-2xl font-bold tracking-tight">Calculator</h1>
      <p className="mt-1 text-muted-foreground">
        Enter your numbers once, then compare how much each company's benefits are actually worth.
      </p>
      <div className="mt-6">
        <CalculatorClient companies={companies} />
      </div>
    </div>
  );
}
