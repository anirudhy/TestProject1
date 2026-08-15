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
   of the README.
4. Deploy. ISR (`revalidate = 3600` on company/benefit/calculator pages) works
   out of the box on Vercel.

## Before this is production-ready

This repo implements the MVP's UI, schema, and calculation logic, but two things
called out in the build spec are intentionally not done here — don't flip this live
without them:

- **Auth.** Phase 0 calls for Supabase Auth (magic link). `/admin/queue` and the
  edit approve/reject routes have no session check yet (see the comment in
  `app/api/edits/[id]/approve/route.ts`). Wire up Supabase Auth, add a
  `middleware.ts` that redirects unauthenticated/non-moderator users away from
  `/admin/*`, and have the approve/reject routes read the reviewer's identity from
  the session instead of the request body.
- **Real seed data.** The shipped seed dataset is fictional (see
  `lib/seed/companies.ts` for why). Populate `benefit_types` (migration `0006`) and
  then research and cite real companies per `CONTRIBUTING.md` before launch.

## Supporting services (Sentry, PostHog, Resend)

These are declared as dependencies per the spec's stack but not yet initialized —
there's no deployed instance to send real error/analytics/email traffic from yet.
Follow each service's standard Next.js App Router setup guide when you're ready:
Sentry's `@sentry/nextjs` wizard, PostHog's Next.js quickstart (client init in a
provider, server init via `posthog-node` for backend events), and Resend for
work-email OTP verification + open enrollment reminder emails.
