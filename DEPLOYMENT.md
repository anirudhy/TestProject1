# Deployment

PerkStack is a standard Next.js 15 App Router project — it deploys to Vercel with no
special configuration.

## Vercel

1. Import this repo in Vercel.
2. Set the environment variables from `.env.example` in the Vercel project settings
   (Production and Preview). At minimum, set `NEXT_PUBLIC_SITE_URL` to the deployed
   URL so `sitemap.ts` and OG image URLs resolve correctly.
3. To go live with real data rather than the in-memory demo dataset, also set
   `NEXT_PUBLIC_SUPABASE_URL`, `NEXT_PUBLIC_SUPABASE_ANON_KEY`, and
   `SUPABASE_SERVICE_ROLE_KEY` — see the "Connecting a real Supabase project" section
   of the README. This also turns on real auth (see below) — magic-link sign-in
   only has something to authenticate against once these are set.
4. In the Supabase dashboard, under Authentication → URL Configuration, add your
   deployed origin (and `http://localhost:3000` for local dev) to the redirect
   allow-list, e.g. `https://your-domain.example/auth/callback`. Magic links will
   fail to redirect back without this.
5. Deploy. ISR (`revalidate = 3600` on company/benefit/calculator pages) works
   out of the box on Vercel.

## Before this is production-ready

This repo implements the MVP's UI, schema, calculation logic, and auth — one thing
called out in the build spec is intentionally still not done here:

- **Real seed data.** The shipped seed dataset is fictional (see
  `lib/seed/companies.ts` for why). Populate `benefit_types` (migration `0006`) and
  then research and cite real companies per `CONTRIBUTING.md` before launch.

Auth itself is wired up (Supabase magic-link, session-derived identity on every
write, role-gated `/admin/**` via `middleware.ts`) — see the README's "§9 open
questions" and "what's real vs. what's a placeholder" sections for exactly what
that does and doesn't cover. One manual step before anyone can moderate: promote
at least one account to `moderator` or `admin` via the Supabase SQL editor —
`update profiles set role = 'admin' where id = '<their auth.users id>'` — since
there's no self-service role escalation UI (deliberately: that would let anyone
approve their own edits).

## Supporting services (Sentry, PostHog, Resend)

These are declared as dependencies per the spec's stack but not yet initialized —
there's no deployed instance to send real error/analytics/email traffic from yet.
Follow each service's standard Next.js App Router setup guide when you're ready:
Sentry's `@sentry/nextjs` wizard, PostHog's Next.js quickstart (client init in a
provider, server init via `posthog-node` for backend events), and Resend for
work-email OTP verification + open enrollment reminder emails.
