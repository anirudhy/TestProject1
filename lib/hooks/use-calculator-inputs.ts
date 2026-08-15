"use client";

import { useEffect, useState } from "react";
import type { CalculatorInputs } from "@/lib/benefits/value";

const STORAGE_KEY = "perkstack:calculator-inputs";

export type FilingStatus = "single" | "married_joint" | "married_separate" | "head_of_household";

export interface FullCalculatorInputs extends CalculatorInputs {
  bonusPercent: number;
  filingStatus: FilingStatus;
}

export const DEFAULT_CALCULATOR_INPUTS: FullCalculatorInputs = {
  salaryUsd: 200_000,
  bonusPercent: 10,
  contributionPercent: 6,
  familySize: 1,
  age: 30,
  filingStatus: "single",
  planYear: 2026,
};

export function useCalculatorInputs() {
  const [inputs, setInputs] = useState<FullCalculatorInputs>(DEFAULT_CALCULATOR_INPUTS);
  const [hydrated, setHydrated] = useState(false);

  useEffect(() => {
    try {
      const stored = window.localStorage.getItem(STORAGE_KEY);
      if (stored) setInputs({ ...DEFAULT_CALCULATOR_INPUTS, ...JSON.parse(stored) });
    } catch {
      // localStorage unavailable — fall back to defaults silently
    } finally {
      setHydrated(true);
    }
  }, []);

  useEffect(() => {
    if (!hydrated) return;
    try {
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(inputs));
    } catch {
      // ignore write failures (private browsing, quota, etc.)
    }
  }, [inputs, hydrated]);

  return { inputs, setInputs, hydrated };
}
