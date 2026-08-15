import { NextResponse } from "next/server";
import { z } from "zod";
import { rejectEdit } from "@/lib/db/queries";

const bodySchema = z.object({ reviewerId: z.string().min(1), note: z.string().max(2000).optional() });

// See the approve route's note: reviewerId should come from an authenticated
// moderator session in production, not the request body.
export async function POST(request: Request, { params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const parsed = bodySchema.safeParse(await request.json().catch(() => null));
  if (!parsed.success) return NextResponse.json({ error: "Invalid request" }, { status: 400 });

  const edit = await rejectEdit(id, parsed.data.reviewerId, parsed.data.note);
  if (!edit) return NextResponse.json({ error: "Edit not found" }, { status: 404 });
  return NextResponse.json({ edit });
}
