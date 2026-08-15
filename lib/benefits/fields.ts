import { z } from "zod";

/**
 * Single source of truth for a benefit type's value shape. Drives three things:
 * the Zod validator (runtime + edit submission validation), the JSON Schema
 * stored in `benefit_types.value_schema` (documentation / external tooling),
 * and the dynamic contribution form in /contribute.
 */
export type FieldSpec =
  | { key: string; label: string; type: "boolean"; help?: string }
  | { key: string; label: string; type: "number"; unit?: string; min?: number; max?: number; help?: string }
  | { key: string; label: string; type: "string"; help?: string }
  | { key: string; label: string; type: "enum"; options: readonly string[]; help?: string }
  | { key: string; label: string; type: "array"; itemFields: readonly FieldSpec[]; help?: string };

function fieldToZod(field: FieldSpec): z.ZodTypeAny {
  switch (field.type) {
    case "boolean":
      return z.boolean();
    case "number":
      return field.min !== undefined || field.max !== undefined
        ? z.number().min(field.min ?? -Infinity).max(field.max ?? Infinity)
        : z.number();
    case "string":
      return z.string();
    case "enum":
      return z.enum(field.options as [string, ...string[]]);
    case "array":
      return z.array(fieldsToZodObject(field.itemFields));
  }
}

export function fieldsToZodObject(fields: readonly FieldSpec[]) {
  const shape: Record<string, z.ZodTypeAny> = {};
  for (const field of fields) {
    shape[field.key] = fieldToZod(field).nullable().optional();
  }
  return z.object(shape);
}

function fieldToJsonSchema(field: FieldSpec): Record<string, unknown> {
  switch (field.type) {
    case "boolean":
      return { type: "boolean", title: field.label };
    case "number":
      return {
        type: "number",
        title: field.label,
        ...(field.unit ? { "x-unit": field.unit } : {}),
        ...(field.min !== undefined ? { minimum: field.min } : {}),
        ...(field.max !== undefined ? { maximum: field.max } : {}),
      };
    case "string":
      return { type: "string", title: field.label };
    case "enum":
      return { type: "string", title: field.label, enum: [...field.options] };
    case "array":
      return {
        type: "array",
        title: field.label,
        items: { type: "object", properties: fieldsToJsonSchema(field.itemFields) },
      };
  }
}

export function fieldsToJsonSchema(fields: readonly FieldSpec[]): Record<string, unknown> {
  const properties: Record<string, unknown> = {};
  for (const field of fields) properties[field.key] = fieldToJsonSchema(field);
  return properties;
}

export function jsonSchemaDocument(fields: readonly FieldSpec[]) {
  return { type: "object", properties: fieldsToJsonSchema(fields) };
}
