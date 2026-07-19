# Documentation to complete before AppSource submission

This is the single place that tracks every documentation item still needing a human to fill it in.
Each row says **where** to edit, **what** to enter, and **anywhere else the same value must change**.
Run **`npm run docs:pending`** at any time to list the placeholders that are still open — when it
reports zero, everything in the "Placeholder" section below is done.

> These are content/decision gaps only. The code and manifest changes for submission already landed
> (see `SUBMISSION.md` for the full submission checklist and hosted-URL table).

## Placeholders (`npm run docs:pending` tracks these)

- [x] **Governing-law jurisdiction** — ~~the one true blocker to publishing the Terms page.~~
  **Resolved:** `TERMS.md` § 8 (Governing law) was dropped rather than filled in, following standard
  practice for MIT-licensed open-source projects — the MIT License itself carries no choice-of-law
  clause, and Microsoft Partner Center does not require one. Contact is now § 8.

- [ ] **Publisher / `ProviderName` confirmation.**
  - **Where:** `SUBMISSION.md` § 1 (`<FILL IN: confirm>` on the Publisher row).
  - **What:** confirm the publisher name matches your Partner Center account's Publisher display
    name **exactly**. Current value everywhere is **William Barnhart**.
  - **Also update (only if the confirmed name differs):** `manifest.xml` `<ProviderName>`,
    `PRIVACY.md` ("published and maintained by …"), `TERMS.md` (intro), `SUBMISSION.md` § 1.
    A mismatch here is a top AppSource rejection cause.

- [ ] **Reviewer CourtListener API token.**
  - **Where:** `SUBMISSION.md` § 3 (`<FILL IN: provide a working token …>`).
  - **What:** paste a working read-only CourtListener token, or a free account reviewers can use,
    so the online Find Hallucinations / Online Lookup / Embed Cited Text features can be tested.
    Get one at <https://www.courtlistener.com/help/api/rest/>.

- [ ] **Sample test `.docx` for reviewers.**
  - **Where:** `SUBMISSION.md` § 3 (`<FILL IN: attach a sample .docx …>`).
  - **What:** attach (in Partner Center) and name a small Word document containing a few real case
    citations, so a reviewer can exercise the local features immediately.

## Reminders (not auto-tracked — no placeholder marker)

- [ ] **"Last updated" dates** — `PRIVACY.md` and `TERMS.md` both say **July 2026**. Bump each to
  the actual publication date when you finalize the documents.

## How the tracking works

- `npm run docs:pending` scans `SUBMISSION.md`, `TERMS.md`, `PRIVACY.md`, and `manifest.xml` for the
  deliberate markers (`<FILL IN: …>`, `to be specified`, `TODO: fill in`) and prints each remaining
  one as `file:line`. To add a new tracked gap, just drop one of those markers in the doc.
- The "Reminders" items above have no marker (they're valid text today, just dated), so complete
  them by hand — `docs:pending` won't flag them.
