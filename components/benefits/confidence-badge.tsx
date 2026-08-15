import { Badge } from "@/components/ui/badge";
import type { ConfidenceLevel } from "@/types";

const CONFIG: Record<ConfidenceLevel, { label: string; variant: "success" | "secondary" | "warning" | "muted" }> = {
  official: { label: "Official source", variant: "success" },
  corroborated: { label: "Corroborated", variant: "secondary" },
  single_report: { label: "Single report", variant: "warning" },
  unverified: { label: "Unverified", variant: "muted" },
};

export function ConfidenceBadge({ confidence }: { confidence: ConfidenceLevel }) {
  const config = CONFIG[confidence];
  return <Badge variant={config.variant}>{config.label}</Badge>;
}
