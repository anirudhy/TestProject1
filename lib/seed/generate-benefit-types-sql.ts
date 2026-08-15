import { BENEFIT_TYPES } from "@/lib/benefits/registry";
import { jsonSchemaDocument } from "@/lib/benefits/fields";

function sqlString(value: string | null) {
  if (value === null) return "null";
  return `'${value.replace(/'/g, "''")}'`;
}

function sqlJson(value: unknown) {
  return `'${JSON.stringify(value).replace(/'/g, "''")}'::jsonb`;
}

const lines: string[] = [];
lines.push("-- Generated from lib/benefits/registry.ts via `tsx lib/seed/generate-benefit-types-sql.ts`.");
lines.push("-- Do not hand-edit; change the registry and regenerate instead.");
lines.push("insert into benefit_types (key, category, label, description, value_schema, is_comparable, is_countable, sort_order) values");

const rows = BENEFIT_TYPES.map((t) => {
  const schema = jsonSchemaDocument(t.fields);
  return `  (${sqlString(t.key)}, ${sqlString(t.category)}, ${sqlString(t.label)}, ${sqlString(t.description)}, ${sqlJson(schema)}, ${t.isComparable}, ${t.isCountable}, ${t.sortOrder})`;
});
lines.push(rows.join(",\n") + ";");

process.stdout.write(lines.join("\n") + "\n");
