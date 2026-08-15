import type { BenefitTypeDef } from "./registry";

/**
 * Generic, mechanism-focused explainer paragraphs for a benefit type's
 * leaderboard page. Deliberately says nothing about any specific company —
 * it explains how the benefit works in general, which is safe to generate
 * from the field schema without individual research. Company-specific facts
 * always come from the sourced company_benefits rows, never from this.
 */
export function explainerParagraphs(type: BenefitTypeDef): string[] {
  const fieldList = type.fields.map((f) => f.label.toLowerCase()).join(", ");

  return [
    `${type.label} is one of the benefits PerkStack tracks in the ${type.category.replace(/_/g, " ")} category. ${type.description}`,
    `The terms that actually matter here are ${fieldList}. Two employers can both say they "offer" ${type.label.toLowerCase()}, but the details in those fields are what determine whether it's worth thousands of dollars a year or almost nothing — which is exactly why a plain "yes/no" on a careers page isn't enough, and why every row in this leaderboard is broken down field by field with a source.`,
    `Because these terms live in plan documents rather than marketing copy, they change from year to year and are easy to get wrong secondhand. PerkStack sources this data from company benefits pages, SEC filings (S-8s and proxy statements for equity-related terms), and Form 5500 filings with the Department of Labor — the same filings a benefits attorney would pull, just made searchable. Where the only available source is an employee report, that's shown explicitly via a confidence badge rather than presented as verified fact.`,
    `Use this leaderboard as a starting point, not a final answer. Confirm the specifics against your own plan documents or HR before making a decision that depends on them — terms can differ by employee class, location, or plan year, and PerkStack's calculator is arithmetic on the terms as reported, not tax or investment advice.`,
  ];
}
