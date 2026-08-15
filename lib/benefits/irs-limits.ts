/**
 * Publicly published IRS limits, not company-specific data — no per-row
 * source citation needed, but these change annually via IRS Revenue
 * Procedure. Verify against the current year's Rev. Proc. before relying on
 * this for a real financial decision (also see the calculator's disclaimer).
 */
export interface IrsLimits {
  electiveDeferralUsd: number;
  catchUp50PlusUsd: number;
  annualAddition415cUsd: number;
  hsaIndividualUsd: number;
  hsaFamilyUsd: number;
}

export const IRS_LIMITS: Record<number, IrsLimits> = {
  2025: {
    electiveDeferralUsd: 23_500,
    catchUp50PlusUsd: 7_500,
    annualAddition415cUsd: 70_000,
    hsaIndividualUsd: 4_300,
    hsaFamilyUsd: 8_550,
  },
  2026: {
    electiveDeferralUsd: 24_500,
    catchUp50PlusUsd: 8_000,
    annualAddition415cUsd: 72_000,
    hsaIndividualUsd: 4_400,
    hsaFamilyUsd: 8_750,
  },
};

const FALLBACK_LIMITS: IrsLimits = IRS_LIMITS[2026]!;

export function getIrsLimits(planYear: number): IrsLimits {
  return IRS_LIMITS[planYear] ?? FALLBACK_LIMITS;
}
