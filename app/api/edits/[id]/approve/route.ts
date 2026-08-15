import { NextResponse } from "next/server";
import { revalidatePath } from "next/cache";
import { z } from "zod";
import { approveEdit, getCompanyById } from "@/lib/db/queries";
import { requireModeratorId } from "@/lib/db/supabase-server";

const bodySchema = z.object({ note: z.string().max(2000).optional() });

export async function POST(request: Request, { params }: { params: Promise<{ id: string }> }) {
  const reviewerId = await requireModeratorId();
  if (!reviewerId) return NextResponse.json({ error: "Sign in as a moderator to approve edits." }, { status: 403 });

  const { id } = await params;
  const parsed = bodySchema.safeParse(await request.json().catch(() => ({})));
  if (!parsed.success) return NextResponse.json({ error: "Invalid request" }, { status: 400 });

  const edit = await approveEdit(id, reviewerId, parsed.data.note);
  if (!edit) return NextResponse.json({ error: "Edit not found" }, { status: 404 });

  const company = await getCompanyById(edit.company_id);
  if (company) revalidatePath(`/companies/${company.slug}`);
  revalidatePath("/companies");
  revalidatePath(`/benefits/${edit.benefit_key}`);
  revalidatePath("/compare");

  return NextResponse.json({ edit });
}
