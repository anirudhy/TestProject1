import { NextResponse } from "next/server";
import { getPublicSupabaseClient, isSupabaseConfigured } from "@/lib/db/client";

export async function GET() {
  if (!isSupabaseConfigured()) {
    return NextResponse.json({
      status: "ok",
      database: "local-seed-fallback",
      message: "NEXT_PUBLIC_SUPABASE_URL not set — serving from the in-memory demo dataset.",
    });
  }

  try {
    const sb = getPublicSupabaseClient()!;
    const { error } = await sb.from("benefit_types").select("key", { count: "exact", head: true });
    if (error) throw error;
    return NextResponse.json({ status: "ok", database: "supabase" });
  } catch (err) {
    return NextResponse.json(
      { status: "error", database: "supabase", message: err instanceof Error ? err.message : "unknown error" },
      { status: 503 }
    );
  }
}
