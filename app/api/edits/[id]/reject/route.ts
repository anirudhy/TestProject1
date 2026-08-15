import { NextResponse } from "next/server";
import { z } from "zod";
import { rejectEdit } from "@/lib/db/queries";
import { requireModeratorId } from "@/lib/db/supabase-server";

const bodySchema = z.object({ note: z.string().max(2000).optional() });

export async function POST(request: Request, { params }: { params: Promise<{ id: string }> }) {
  const reviewerId = await requireModeratorId();
  if (!reviewerId) return NextResponse.json({ error: "Sign in as a moderator to reject edits." }, { status: 403 });

  const { id } = await params;
  const parsed = bodySchema.safeParse(await request.json().catch(() => ({})));
  if (!parsed.success) return NextResponse.json({ error: "Invalid request" }, { status: 400 });

  const edit = await rejectEdit(id, reviewerId, parsed.data.note);
  if (!edit) return NextResponse.json({ error: "Edit not found" }, { status: 404 });
  return NextResponse.json({ edit });
}
