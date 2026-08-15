import "server-only";
import { cookies } from "next/headers";
import { createServerClient, type CookieOptions } from "@supabase/ssr";
import { isSupabaseConfigured } from "./client";
import { localStore } from "./local-store";
import type { Profile } from "@/types";

interface CookieToSet {
  name: string;
  value: string;
  options: CookieOptions;
}

/**
 * Session-aware Supabase client for Server Components and Route Handlers —
 * reads/writes the auth cookie, so `auth.getUser()` reflects who's actually
 * signed in. Distinct from lib/db/client.ts's clients, which are anon/service
 * clients with no notion of a session.
 */
export async function createSupabaseServerClient() {
  if (!isSupabaseConfigured()) return null;
  const cookieStore = await cookies();

  return createServerClient(process.env.NEXT_PUBLIC_SUPABASE_URL!, process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!, {
    cookies: {
      getAll() {
        return cookieStore.getAll();
      },
      setAll(cookiesToSet: CookieToSet[]) {
        try {
          for (const { name, value, options } of cookiesToSet) cookieStore.set(name, value, options);
        } catch {
          // Called from a Server Component render, where cookies are
          // read-only — middleware.ts is what actually refreshes the
          // session cookie on navigation. Safe to ignore here.
        }
      },
    },
  });
}

/**
 * The signed-in user's profile (including role), or null if there's no
 * session, no matching profile row, or Supabase isn't configured (demo
 * mode). This is the one function route handlers should call to decide
 * whether a request is authorized — never trust a client-supplied user id.
 */
export async function getAuthedProfile(): Promise<Profile | null> {
  const supabase = await createSupabaseServerClient();
  if (!supabase) return null;

  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) return null;

  const { data: profile } = await supabase.from("profiles").select("*").eq("id", user.id).maybeSingle();
  return (profile as Profile | null) ?? null;
}

export function isModeratorOrAdmin(profile: Profile | null): boolean {
  return profile?.role === "moderator" || profile?.role === "admin";
}

/**
 * The single choke point every moderation route (approve/reject) must call
 * before touching data. In demo mode there's no session to check, so it
 * returns the fixed demo-admin identity; against a real Supabase project it
 * requires a signed-in user whose profiles.role is moderator/admin — never
 * trusts a client-supplied id. Returns null when the caller is unauthorized.
 */
export async function requireModeratorId(): Promise<string | null> {
  if (!isSupabaseConfigured()) return localStore.getDemoAdminProfile().id;
  const profile = await getAuthedProfile();
  return isModeratorOrAdmin(profile) ? profile!.id : null;
}

/** submitted_by for a new contribution: the signed-in user, the demo user in demo mode, or null (anonymous). */
export async function currentContributorId(): Promise<string | null> {
  if (!isSupabaseConfigured()) return localStore.getOrCreateDemoProfile().id;
  const profile = await getAuthedProfile();
  return profile?.id ?? null;
}
