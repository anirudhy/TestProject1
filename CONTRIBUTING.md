# Contributing data to PerkStack

PerkStack's entire value is that every number is real and sourced. These rules are
non-negotiable — they're enforced partly by the database (see
`supabase/migrations/0003_source_enforcement.sql`) and partly by moderator review.

## Hard rules

1. **Never upload, host, or excerpt internal company documents** — Summary Plan
   Descriptions distributed only to employees, internal wikis, HR portals, screenshots
   of an internal benefits tool, etc. If you know a fact from an internal document,
   restate the *fact* ("match is 50% up to the IRS limit") and cite a **public**
   source for it instead — the company's own public benefits page is usually enough,
   and Form 5500 filings (public, searchable via EFAST2) cover most retirement plan
   terms.
2. **Do not scrape Levels.fyi, Glassdoor, or similar sites.** Link out to them
   instead. Their Terms of Service prohibit it.
3. **Every submission needs a source.** In order of preference:
   1. Company public benefits page (careers site) — cite the URL.
   2. Form 5500 + attached plan documents (DOL, via EFAST2) — the authoritative
      source for 401(k) match formulas and after-tax/in-plan-Roth provisions.
   3. SEC filings (S-8, proxy statements) — ESPP terms for public companies.
   4. Company press releases / benefits announcements.
   5. Your own experience as a verified employee — lowest confidence tier,
      requires corroboration from a second independent report or moderator
      approval before it affects the public value.
4. **Retrieval dates matter.** Benefits change at open enrollment. A four-year-old
   number presented as current is the failure mode that kills this product's
   credibility — always note when you retrieved a source.
5. **You are responsible for what you submit.** Don't submit anything covered by an
   NDA or confidentiality agreement, even if you personally have access to it.

## How review works

- Submissions citing a company public page, SEC filing, or Form 5500 go to a
  fast-track queue — one moderator click to approve.
- Employee reports need either a second corroborating report or explicit moderator
  approval, and are marked `single_report` confidence until then.
- Anonymous/unverified submissions are never auto-approved.

See `lib/benefits/registry.ts` for the exact fields each benefit type expects —
`/contribute` generates its form directly from that registry, so the fields you'll
be asked for there are the canonical schema.
