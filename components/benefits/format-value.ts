import type { FieldSpec } from "@/lib/benefits/fields";
import { formatUsd } from "@/lib/utils";

function humanizeEnum(v: string) {
  return v.replace(/_/g, " ").replace(/\b\w/g, (c) => c.toUpperCase());
}

export function formatFieldValue(field: FieldSpec, raw: unknown): string | null {
  if (raw === null || raw === undefined || raw === "") return null;

  switch (field.type) {
    case "boolean":
      return raw ? "Yes" : "No";
    case "number": {
      const n = Number(raw);
      if (Number.isNaN(n)) return null;
      if (field.unit === "$") return formatUsd(n);
      if (field.unit === "%") return `${n}%`;
      if (field.unit) return `${n} ${field.unit}`;
      return String(n);
    }
    case "string":
      return String(raw);
    case "enum":
      return humanizeEnum(String(raw));
    case "array": {
      if (!Array.isArray(raw) || raw.length === 0) return null;
      return raw
        .map((item) =>
          field.itemFields
            .map((f) => {
              const v = formatFieldValue(f, (item as Record<string, unknown>)[f.key]);
              return v ? `${f.label}: ${v}` : null;
            })
            .filter(Boolean)
            .join(", ")
        )
        .join(" · ");
    }
  }
}
