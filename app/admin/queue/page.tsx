import type { Metadata } from "next";
import { redirect } from "next/navigation";
import { listEdits, getCompanyById } from "@/lib/db/queries";
import { getBenefitType } from "@/lib/benefits/registry";
import { ModerationQueue, type QueueEditItem } from "@/components/admin/moderation-queue";
import { isSupabaseConfigured } from "@/lib/db/client";
import { getAuthedProfile, isModeratorOrAdmin } from "@/lib/db/supabase-server";

export const metadata: Metadata = { title: "Moderation queue", robots: { index: false, follow: false } };
export const dynamic = "force-dynamic";

export default async function AdminQueuePage() {
  // middleware.ts already redirects unauthorized requests before they reach
  // this component; this is defense in depth, not the primary gate.
  if (isSupabaseConfigured()) {
    const profile = await getAuthedProfile();
    if (!isModeratorOrAdmin(profile)) redirect("/sign-in?next=/admin/queue");
  }

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
      {!isSupabaseConfigured() && (
        <div className="mt-3 rounded-md border border-amber-300 bg-amber-50 p-3 text-sm text-amber-900 dark:border-amber-800 dark:bg-amber-950/40 dark:text-amber-200">
          Demo mode: approvals here act as the fixed demo-admin identity — there&rsquo;s no real session to check
          because no Supabase project is configured. Against a real project this route requires a signed-in
          moderator/admin (see <code>middleware.ts</code> and <code>lib/db/supabase-server.ts</code>).
        </div>
      )}
      <div className="mt-6">
        <ModerationQueue items={items} />
      </div>
    </div>
  );
}
