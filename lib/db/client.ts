import "server-only";
import { createClient, type SupabaseClient } from "@supabase/supabase-js";

export function isSupabaseConfigured(): boolean {
  return Boolean(process.env.NEXT_PUBLIC_SUPABASE_URL && process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY);
}

let publicClient: SupabaseClient | null | undefined;
export function getPublicSupabaseClient(): SupabaseClient | null {
  if (publicClient !== undefined) return publicClient;
  if (!isSupabaseConfigured()) {
    publicClient = null;
    return publicClient;
  }
  publicClient = createClient(process.env.NEXT_PUBLIC_SUPABASE_URL!, process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!, {
    auth: { persistSession: false },
  });
  return publicClient;
}

let serviceClient: SupabaseClient | null | undefined;
/** Server-only, bypasses RLS. Used for moderation writes and the seed script. */
export function getServiceSupabaseClient(): SupabaseClient | null {
  if (serviceClient !== undefined) return serviceClient;
  if (!process.env.NEXT_PUBLIC_SUPABASE_URL || !process.env.SUPABASE_SERVICE_ROLE_KEY) {
    serviceClient = null;
    return serviceClient;
  }
  serviceClient = createClient(process.env.NEXT_PUBLIC_SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY, {
    auth: { persistSession: false },
  });
  return serviceClient;
}
