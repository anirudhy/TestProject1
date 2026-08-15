import type { Metadata } from "next";
import { listCompaniesWithBenefits } from "@/lib/db/company-with-benefits";
import { CompanyDirectory } from "@/components/company/company-directory";

export const metadata: Metadata = {
  title: "Companies",
  description: "Browse structured, sourced benefits data for every company in the PerkStack directory.",
};

export const revalidate = 3600;

export default async function CompaniesPage() {
  const items = await listCompaniesWithBenefits();

  return (
    <div className="mx-auto max-w-6xl px-4 py-10">
      <h1 className="text-2xl font-bold tracking-tight">Companies</h1>
      <p className="mt-1 text-muted-foreground">
        {items.length} companies, every field sourced and dated. Filter by industry, size, or specific benefits.
      </p>
      <div className="mt-6">
        <CompanyDirectory items={items} />
      </div>
    </div>
  );
}
