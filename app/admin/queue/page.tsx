import type { Metadata } from "next";
import { listEdits, getCompanyById } from "@/lib/db/queries";
import { getBenefitType } from "@/lib/benefits/registry";
import { ModerationQueue, type QueueEditItem } from "@/components/admin/moderation-queue";

export const metadata: Metadata = { title: "Moderation queue", robots: { index: false, follow: false } };
export const dynamic = "force-dynamic";

export default async function AdminQueuePage() {
  const pending = await listEdits("pending");

  const items: QueueEditItem[] = await Promise.all(
    pending.map(async (edit) => {
      const company = await getCompanyById(edit.company_id);
      const type = getBenefitType(edit.benefit_key);
      return { edit, companyName: company?.name ?? edit.company_id, benefitLabel: type?.label ?? edit.benefit_key };
    })
  );

  return (
    <div className="mx-auto max-w-4xl px-4 py-10">
      <h1 className="text-2xl font-bold tracking-tight">Moderation queue</h1>
      <p className="mt-1 text-muted-foreground">{items.length} pending submissions.</p>
      <div className="mt-3 rounded-md border border-amber-300 bg-amber-50 p-3 text-sm text-amber-900 dark:border-amber-800 dark:bg-amber-950/40 dark:text-amber-200">
        Demo mode: this route has no authentication wired up yet. Before deploying, gate it behind Supabase Auth +
        a moderator/admin role check (see the comment in <code>app/api/edits/[id]/approve/route.ts</code>).
      </div>
      <div className="mt-6">
        <ModerationQueue items={items} />
      </div>
    </div>
  );
}
