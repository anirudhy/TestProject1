import type { MetadataRoute } from "next";
import { listAllCompanySlugs } from "@/lib/db/queries";
import { BENEFIT_TYPES } from "@/lib/benefits/registry";

export default async function sitemap(): Promise<MetadataRoute.Sitemap> {
  const base = process.env.NEXT_PUBLIC_SITE_URL ?? "http://localhost:3000";
  const slugs = await listAllCompanySlugs();

  return [
    { url: base, changeFrequency: "weekly", priority: 1 },
    { url: `${base}/companies`, changeFrequency: "daily", priority: 0.9 },
    { url: `${base}/compare`, changeFrequency: "weekly", priority: 0.6 },
    { url: `${base}/calculator`, changeFrequency: "monthly", priority: 0.6 },
    ...slugs.map((slug) => ({ url: `${base}/companies/${slug}`, changeFrequency: "monthly" as const, priority: 0.8 })),
    ...BENEFIT_TYPES.map((t) => ({ url: `${base}/benefits/${t.key}`, changeFrequency: "weekly" as const, priority: 0.7 })),
  ];
}
