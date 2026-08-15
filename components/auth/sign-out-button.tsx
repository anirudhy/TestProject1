"use client";

import { useRouter } from "next/navigation";
import { getSupabaseBrowserClient } from "@/lib/db/supabase-browser";

export function SignOutButton() {
  const router = useRouter();

  async function handleSignOut() {
    const supabase = getSupabaseBrowserClient();
    if (!supabase) return;
    await supabase.auth.signOut();
    router.push("/");
    router.refresh();
  }

  return (
    <button onClick={handleSignOut} className="text-muted-foreground transition-colors hover:text-foreground">
      Sign out
    </button>
  );
}
