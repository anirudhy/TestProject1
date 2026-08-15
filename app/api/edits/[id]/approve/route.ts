import { NextResponse } from "next/server";
import { revalidatePath } from "next/cache";
import { z } from "zod";
import { approveEdit, getCompanyById } from "@/lib/db/queries";

const bodySchema = z.object({ reviewerId: z.string().min(1), note: z.string().max(2000).optional() });

/**
 * NOTE for production: this route currently trusts `reviewerId` from the
 * request body because no auth session is wired up in this scaffold (see
 * Phase 0/4 in the build spec — Supabase Auth magic-link sign-in). Before
 * deploying, this must instead read the caller's identity from their
 * authenticated session server-side and verify profiles.role is
 * 'moderator' or 'admin' — never trust a client-supplied reviewer id.
 */
export async function POST(request: Request, { params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const parsed = bodySchema.safeParse(await request.json().catch(() => null));
  if (!parsed.success) return NextResponse.json({ error: "Invalid request" }, { status: 400 });

  const edit = await approveEdit(id, parsed.data.reviewerId, parsed.data.note);
  if (!edit) return NextResponse.json({ error: "Edit not found" }, { status: 404 });

  const company = await getCompanyById(edit.company_id);
  if (company) revalidatePath(`/companies/${company.slug}`);
  revalidatePath("/companies");
  revalidatePath(`/benefits/${edit.benefit_key}`);
  revalidatePath("/compare");

  return NextResponse.json({ edit });
}
