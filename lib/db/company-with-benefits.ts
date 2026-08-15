import "server-only";
import { getBenefitsForCompany, listPublishedCompanies } from "@/lib/db/queries";
import type { Company, CompanyBenefit } from "@/types";

export interface CompanyWithBenefits {
  company: Company;
  benefits: CompanyBenefit[];
}

export async function listCompaniesWithBenefits(): Promise<CompanyWithBenefits[]> {
  const companies = await listPublishedCompanies();
  return Promise.all(
    companies.map(async (company) => ({ company, benefits: await getBenefitsForCompany(company.id) }))
  );
}
