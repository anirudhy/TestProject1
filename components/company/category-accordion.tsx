import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from "@/components/ui/accordion";
import { BENEFIT_CATEGORIES } from "@/types";
import { BENEFIT_TYPES } from "@/lib/benefits/registry";
import { BenefitRow } from "@/components/benefits/benefit-row";
import type { CompanyBenefit } from "@/types";

export function CategoryAccordion({ benefits, companySlug }: { benefits: CompanyBenefit[]; companySlug: string }) {
  const byKey = new Map(benefits.map((b) => [b.benefit_key, b]));

  return (
    <Accordion type="multiple" defaultValue={BENEFIT_CATEGORIES.map((c) => c.key)}>
      {BENEFIT_CATEGORIES.map((category) => {
        const types = BENEFIT_TYPES.filter((t) => t.category === category.key).sort((a, b) => a.sortOrder - b.sortOrder);
        if (types.length === 0) return null;
        const knownCount = types.filter((t) => byKey.has(t.key)).length;

        return (
          <AccordionItem key={category.key} value={category.key}>
            <AccordionTrigger>
              <span className="flex items-center gap-2">
                {category.label}
                <span className="text-xs font-normal text-muted-foreground">
                  {knownCount}/{types.length}
                </span>
              </span>
            </AccordionTrigger>
            <AccordionContent>
              <div className="divide-y divide-border/60">
                {types.map((type) => (
                  <BenefitRow key={type.key} type={type} benefit={byKey.get(type.key)} companySlug={companySlug} />
                ))}
              </div>
            </AccordionContent>
          </AccordionItem>
        );
      })}
    </Accordion>
  );
}
