# PerkStack

A structured, sourced directory of company benefits — 401(k) match formulas, Mega
Backdoor Roth support, ESPP terms, and more — with side-by-side comparison and a
"total employer value" calculator. Every number is cited to a public source and
dated.

> Comp is what they pay you. Benefits are what they pay you that you forgot to claim.

This repo implements Phases 0–4 (the MVP) of the build spec.

## Stack

Next.js 15 (App Router, TypeScript) · Tailwind CSS + shadcn/ui-style primitives ·
Supabase (Postgres + RLS) · Zod · Vitest.

## Running locally

```bash
pnpm install
pnpm dev
```

No environment variables are required to start. Without `NEXT_PUBLIC_SUPABASE_URL`
set, the app serves everything from an in-memory demo dataset (`lib/db/local-store.ts`,
seeded from `lib/seed/companies.ts`) so every screen — directory, compare,
leaderboards, calculator, contribute, moderation queue — works immediately.

Check `/api/health` to see which data source is active.

### Connecting a real Supabase project

1. Create a Supabase project and copy `.env.example` to `.env.local`, filling in
   `NEXT_PUBLIC_SUPABASE_URL`, `NEXT_PUBLIC_SUPABASE_ANON_KEY`, and
   `SUPABASE_SERVICE_ROLE_KEY`.
2. Apply the migrations in `supabase/migrations/` in order (via the Supabase CLI's
   `supabase db push`, or paste them into the SQL editor in order — `0006` seeds the
   benefit type registry and must run after `0002`–`0005`).
3. Run `pnpm seed` to load `lib/seed/companies.ts` into the database.
4. Run `pnpm seed:verify` periodically to catch data older than 12 months.

**The seed data ships with fictional example companies** (`.example` TLD — reserved
by RFC 2606, never resolves). Real seeding — researching and citing real companies —
is the actual Phase 1 work described in the spec's §5 and is intentionally not faked
here; see `lib/seed/companies.ts` for why.

## Commands

| Command | What it does |
|---|---|
| `pnpm dev` | Start the dev server |
| `pnpm build` / `pnpm start` | Production build / serve |
| `pnpm typecheck` | `tsc --noEmit` |
| `pnpm lint` | ESLint |
| `pnpm test` | Vitest (unit tests for the benefits calculator) |
| `pnpm seed` | Seed a real Supabase project from `lib/seed/companies.ts` |
| `pnpm seed:verify` | Report benefit fields not re-verified in 12+ months |

## §9 open questions — resolved

The build spec left four decisions open for before Phase 4. Resolved as follows,
per the spec's own recommendation in each case:

1. **Anonymous browsing vs. sign-in wall** — fully open browsing, no wall.
   Everything under `/companies`, `/compare`, `/benefits`, `/calculator` renders
   with no session. `/contribute` requires sign-in against a real Supabase
   project (enforced in `app/contribute/page.tsx` and `app/api/edits/route.ts`);
   `/admin/**` requires a moderator/admin role (enforced in `middleware.ts`,
   with a redundant check in `app/admin/queue/page.tsx`). Neither check applies
   in demo mode, since there's no session concept without a real project.
2. **International plan variants** — `company_benefits.country` added
   (migration `0008_country_dimension.sql`), `text not null default 'US'`. Unused
   beyond the default until non-US data exists.
3. **Reputation → auto-approval threshold** — none. Every edit lands in
   `benefit_edits` with `status = 'pending'`; there is no code path that flips a
   row to `approved` without a moderator action, regardless of the submitter's
   `reputation`.
4. **Monetization** — deliberately not decided here. Both options the spec
   raises (job board, sponsored profiles) trade against data neutrality, which
   is the entire basis for user trust — that's a call for whoever owns the
   product, not something to default into via scaffolding.

## Project layout

```
/app                  routes (App Router)
/components/ui        shadcn-style primitives (button, card, table, ...)
/components/benefits   per-benefit-type renderers (value formatting, confidence, source links)
/components/company    directory + detail page pieces (compare-agnostic)
/components/compare    /compare table + picker
/components/contribute  dynamic form driven by the benefit type registry
/components/admin      moderation queue UI
/lib/benefits          the benefit type registry (single source of truth), Zod/JSON
                        Schema generation, and the pure calculator module (lib/benefits/value.ts)
/lib/db                Supabase clients + an in-memory fallback store, unified behind lib/db/queries.ts
/lib/seed              seed data + scripts
/supabase/migrations   SQL schema, RLS policies, and the source-required trigger
/types                 shared domain types
```

## What's real vs. what's a placeholder

- **Schema, RLS, and the source-enforcement trigger are production-shaped.** A
  `company_benefits` row cannot exist without a `benefit_sources` row (enforced by a
  deferred constraint trigger, not application code — see
  `supabase/migrations/0003_source_enforcement.sql`).
- **Auth is wired up: Supabase Auth magic-link sign-in.** `/sign-in` requests a
  one-time email link (`supabase.auth.signInWithOtp`); `/auth/callback` exchanges
  the code for a session; `auth.users` inserts auto-create a `profiles` row via a
  DB trigger (migration `0009_auto_create_profile.sql`) with a generated
  pseudonymous handle and `role = 'user'`. `middleware.ts` refreshes the session
  cookie and redirects unauthorized requests away from `/admin/**`; the
  approve/reject/submit-edit API routes derive the acting user from the session
  server-side (`lib/db/supabase-server.ts`) and never trust a client-supplied id.
  **Promoting a user to `moderator`/`admin` is a manual `update profiles set
  role = 'moderator' where id = '...'` in the Supabase SQL editor** — there is
  intentionally no self-service role escalation UI.
- **Seed data is fictional**, for the reasons above.
- **Sentry/PostHog/Resend are declared as dependencies** (per the spec's stack) but
  not yet initialized — there's no real error/analytics/email traffic to send from a
  repo with no deployed instance.

## Disclaimers (also shown in the product)

PerkStack is an unofficial, community-maintained directory, not affiliated with or
endorsed by the companies listed. Data may be inaccurate or outdated — verify with
your employer. The calculator performs arithmetic on stated plan terms; it is not
tax or investment advice.
