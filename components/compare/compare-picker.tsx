"use client";

import { useRouter } from "next/navigation";
import { useState } from "react";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Badge } from "@/components/ui/badge";
import { X } from "lucide-react";
import type { Company } from "@/types";

export function ComparePicker({ allCompanies, selectedSlugs }: { allCompanies: Company[]; selectedSlugs: string[] }) {
  const router = useRouter();
  const [pending, setPending] = useState(false);

  function updateSlugs(next: string[]) {
    setPending(true);
    const params = new URLSearchParams();
    if (next.length > 0) params.set("c", next.join(","));
    router.push(`/compare${params.toString() ? `?${params}` : ""}`);
    setPending(false);
  }

  const available = allCompanies.filter((c) => !selectedSlugs.includes(c.slug));

  return (
    <div className="flex flex-wrap items-center gap-2">
      {selectedSlugs.map((slug) => {
        const company = allCompanies.find((c) => c.slug === slug);
        return (
          <Badge key={slug} variant="secondary" className="gap-1 py-1 pl-2.5 pr-1.5 text-sm">
            {company?.name ?? slug}
            <button
              onClick={() => updateSlugs(selectedSlugs.filter((s) => s !== slug))}
              className="rounded-full p-0.5 hover:bg-black/10 dark:hover:bg-white/10"
              aria-label={`Remove ${company?.name ?? slug}`}
            >
              <X className="h-3 w-3" />
            </button>
          </Badge>
        );
      })}
      {selectedSlugs.length < 4 && available.length > 0 && (
        <Select
          disabled={pending}
          onValueChange={(slug) => updateSlugs([...selectedSlugs, slug])}
        >
          <SelectTrigger className="w-[200px]"><SelectValue placeholder="Add a company…" /></SelectTrigger>
          <SelectContent>
            {available.map((c) => (
              <SelectItem key={c.slug} value={c.slug}>{c.name}</SelectItem>
            ))}
          </SelectContent>
        </Select>
      )}
    </div>
  );
}
