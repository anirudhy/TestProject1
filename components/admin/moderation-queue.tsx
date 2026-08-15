"use client";

import { useState } from "react";
import { useRouter } from "next/navigation";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";
import type { BenefitEdit } from "@/types";

export interface QueueEditItem {
  edit: BenefitEdit;
  companyName: string;
  benefitLabel: string;
}

function DiffRow({ label, before, after }: { label: string; before: unknown; after: unknown }) {
  const changed = JSON.stringify(before) !== JSON.stringify(after);
  return (
    <div className="grid grid-cols-[140px_1fr_1fr] gap-2 py-1 text-sm">
      <span className="text-muted-foreground">{label}</span>
      <span className={changed ? "text-destructive line-through" : ""}>{before === undefined || before === null ? "—" : String(before)}</span>
      <span className={changed ? "font-medium text-emerald-700 dark:text-emerald-400" : ""}>{after === undefined || after === null ? "—" : String(after)}</span>
    </div>
  );
}

function EditCard({ item, onDone }: { item: QueueEditItem; onDone: () => void }) {
  const { edit } = item;
  const [reviewerId, setReviewerId] = useState("profile_demo_admin");
  const [note, setNote] = useState("");
  const [busy, setBusy] = useState(false);

  const proposedKeys = Object.keys(edit.proposed_value ?? {});
  const currentKeys = Object.keys(edit.current_value ?? {});
  const allKeys = Array.from(new Set([...proposedKeys, ...currentKeys]));

  async function act(action: "approve" | "reject") {
    setBusy(true);
    const res = await fetch(`/api/edits/${edit.id}/${action}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ reviewerId, note }),
    });
    setBusy(false);
    if (res.ok) onDone();
  }

  return (
    <Card>
      <CardHeader>
        <div className="flex items-center justify-between">
          <CardTitle className="text-base">{item.companyName} — {item.benefitLabel}</CardTitle>
          <Badge variant={edit.source_type === "user_report" ? "warning" : "secondary"}>{edit.source_type.replace(/_/g, " ")}</Badge>
        </div>
      </CardHeader>
      <CardContent>
        <div className="grid grid-cols-[140px_1fr_1fr] gap-2 border-b border-border pb-1 text-xs font-semibold text-muted-foreground">
          <span>Field</span><span>Current</span><span>Proposed</span>
        </div>
        {allKeys.map((key) => (
          <DiffRow
            key={key}
            label={key}
            before={(edit.current_value as Record<string, unknown> | null)?.[key]}
            after={(edit.proposed_value as Record<string, unknown>)[key]}
          />
        ))}

        {edit.source_url && (
          <p className="mt-3 text-sm">
            Source: <a href={edit.source_url} target="_blank" rel="noopener noreferrer" className="text-primary hover:underline">{edit.source_url}</a>
          </p>
        )}
        {edit.rationale && <p className="mt-1 text-sm text-muted-foreground">&ldquo;{edit.rationale}&rdquo;</p>}

        <div className="mt-4 flex flex-wrap items-end gap-2 border-t border-border pt-3">
          <div>
            <label className="mb-1 block text-xs text-muted-foreground">Reviewer id</label>
            <Input value={reviewerId} onChange={(e) => setReviewerId(e.target.value)} className="w-[220px]" />
          </div>
          <div className="flex-1">
            <label className="mb-1 block text-xs text-muted-foreground">Note (optional)</label>
            <Input value={note} onChange={(e) => setNote(e.target.value)} />
          </div>
          <Button disabled={busy} onClick={() => act("approve")}>Approve</Button>
          <Button disabled={busy} variant="destructive" onClick={() => act("reject")}>Reject</Button>
        </div>
      </CardContent>
    </Card>
  );
}

export function ModerationQueue({ items }: { items: QueueEditItem[] }) {
  const router = useRouter();
  const [removed, setRemoved] = useState<Set<string>>(new Set());

  const visible = items.filter((i) => !removed.has(i.edit.id));

  if (visible.length === 0) return <p className="text-muted-foreground">Nothing pending. Nicely done.</p>;

  return (
    <div className="space-y-4">
      {visible.map((item) => (
        <EditCard
          key={item.edit.id}
          item={item}
          onDone={() => {
            setRemoved((prev) => new Set(prev).add(item.edit.id));
            router.refresh();
          }}
        />
      ))}
    </div>
  );
}
