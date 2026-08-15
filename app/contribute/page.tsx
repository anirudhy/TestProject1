import type { Metadata } from "next";
import { Suspense } from "react";
import Link from "next/link";
import { listCompaniesWithBenefits } from "@/lib/db/company-with-benefits";
import { ContributeForm } from "@/components/contribute/contribute-form";
import { Button } from "@/components/ui/button";
import { isSupabaseConfigured } from "@/lib/db/client";
import { currentContributorId } from "@/lib/db/supabase-server";

export const metadata: Metadata = {
  title: "Contribute",
  description: "Add or correct sourced benefits data for a company.",
};

export default async function ContributePage() {
  const [companies, contributorId] = await Promise.all([listCompaniesWithBenefits(), currentContributorId()]);
  const requiresSignIn = isSupabaseConfigured() && !contributorId;

  return (
    <div className="mx-auto max-w-3xl px-4 py-10">
      <h1 className="text-2xl font-bold tracking-tight">Contribute</h1>
      <p className="mt-1 text-muted-foreground">
        Every submission needs a source — a company page, an SEC filing, a Form 5500, or (lowest confidence) your own
        experience as an employee. Nothing goes live without moderator review.
      </p>
      <div className="mt-6">
        {requiresSignIn ? (
          <div className="rounded-lg border border-border bg-card p-6 text-center">
            <p className="font-medium">Sign in to contribute</p>
            <p className="mt-1 text-sm text-muted-foreground">
              Browsing PerkStack is always open. Submitting or moderating data needs an account, so edits can be
              traced back to a submitter.
            </p>
            <Button asChild className="mt-4">
              <Link href={`/sign-in?next=/contribute`}>Sign in</Link>
            </Button>
          </div>
        ) : (
          <Suspense fallback={null}>
            <ContributeForm companies={companies} />
          </Suspense>
        )}
      </div>
    </div>
  );
}
