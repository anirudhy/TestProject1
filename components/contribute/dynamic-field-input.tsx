"use client";

import type { FieldSpec } from "@/lib/benefits/fields";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Button } from "@/components/ui/button";
import { Plus, Trash2 } from "lucide-react";

export type FieldValue = string | number | boolean | Record<string, unknown>[] | null | undefined;

export function DynamicFieldInput({
  field,
  value,
  onChange,
}: {
  field: FieldSpec;
  value: FieldValue;
  onChange: (value: FieldValue) => void;
}) {
  if (field.type === "boolean") {
    return (
      <div className="flex items-center gap-2">
        <input
          id={field.key}
          type="checkbox"
          checked={Boolean(value)}
          onChange={(e) => onChange(e.target.checked)}
        />
        <Label htmlFor={field.key} className="font-normal">{field.label}</Label>
      </div>
    );
  }

  if (field.type === "number") {
    return (
      <div>
        <Label htmlFor={field.key}>{field.label}{field.unit ? ` (${field.unit})` : ""}</Label>
        <Input
          id={field.key}
          type="number"
          value={typeof value === "number" ? value : ""}
          onChange={(e) => onChange(e.target.value === "" ? null : Number(e.target.value))}
        />
      </div>
    );
  }

  if (field.type === "enum") {
    return (
      <div>
        <Label htmlFor={field.key}>{field.label}</Label>
        <Select value={typeof value === "string" ? value : undefined} onValueChange={(v) => onChange(v)}>
          <SelectTrigger id={field.key}><SelectValue placeholder="Select…" /></SelectTrigger>
          <SelectContent>
            {field.options.map((opt) => (
              <SelectItem key={opt} value={opt}>{opt.replace(/_/g, " ")}</SelectItem>
            ))}
          </SelectContent>
        </Select>
      </div>
    );
  }

  if (field.type === "array") {
    const rows: Record<string, unknown>[] = Array.isArray(value) ? value : [];
    return (
      <div>
        <div className="flex items-center justify-between">
          <Label>{field.label}</Label>
          <Button
            type="button"
            variant="ghost"
            size="sm"
            onClick={() => onChange([...rows, Object.fromEntries(field.itemFields.map((f) => [f.key, null]))])}
          >
            <Plus className="mr-1 h-3.5 w-3.5" /> Add row
          </Button>
        </div>
        <div className="space-y-2">
          {rows.map((row, i) => (
            <div key={i} className="flex items-end gap-2 rounded-md border border-border p-2">
              <div className="grid flex-1 grid-cols-2 gap-2">
                {field.itemFields.map((itemField) => (
                  <DynamicFieldInput
                    key={itemField.key}
                    field={itemField}
                    value={row[itemField.key] as FieldValue}
                    onChange={(v) => {
                      const next = [...rows];
                      next[i] = { ...row, [itemField.key]: v };
                      onChange(next);
                    }}
                  />
                ))}
              </div>
              <Button type="button" variant="ghost" size="icon" onClick={() => onChange(rows.filter((_, idx) => idx !== i))}>
                <Trash2 className="h-4 w-4" />
              </Button>
            </div>
          ))}
        </div>
      </div>
    );
  }

  return (
    <div>
      <Label htmlFor={field.key}>{field.label}</Label>
      <Input id={field.key} type="text" value={typeof value === "string" ? value : ""} onChange={(e) => onChange(e.target.value)} />
    </div>
  );
}
