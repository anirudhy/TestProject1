import type { Metadata } from "next";
import { Suspense } from "react";
import { listCompaniesWithBenefits } from "@/lib/db/company-with-benefits";
import { ContributeForm } from "@/components/contribute/contribute-form";

export const metadata: Metadata = {
  title: "Contribute",
  description: "Add or correct sourced benefits data for a company.",
};

export default async function ContributePage() {
  const companies = await listCompaniesWithBenefits();

  return (
    <div className="mx-auto max-w-3xl px-4 py-10">
      <h1 className="text-2xl font-bold tracking-tight">Contribute</h1>
      <p className="mt-1 text-muted-foreground">
        Every submission needs a source — a company page, an SEC filing, a Form 5500, or (lowest confidence) your own
        experience as an employee. Nothing goes live without moderator review.
      </p>
      <div className="mt-6">
        <Suspense fallback={null}>
          <ContributeForm companies={companies} />
        </Suspense>
      </div>
    </div>
  );
}
