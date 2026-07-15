# OpenClerk

## What This Is

OpenClerk is a Word add-in (task pane) for legal citation work: hyperlinking case-law and
parenthetical citations, checking citations for Bluebook formatting problems across three
editions (20th/21st/22nd), looking citations up live against public (CourtListener) and
enterprise (Westlaw, LexisNexis, Bloomberg Law) legal databases, detecting fabricated
("hallucinated") citations, and embedding cited opinion text — entirely client-side, with no
backend server of its own.

## Core Value

Legal citation checks and hyperlinks must be accurate and trustworthy — a hallucination check
must never falsely report a fabricated citation as "verified."

## Requirements

### Validated

- ✓ Hyperlink case-law and parenthetical citations from a source document — existing
- ✓ Bluebook citation-formatting checks (20th/21st/22nd editions) — existing
- ✓ Live citation lookup against CourtListener (free, no auth) and enterprise providers
  (Westlaw, LexisNexis, Bloomberg Law, auth required) — existing
- ✓ Hallucination detection (verify cited cases are real, with a name-mismatch signal for
  citations whose locator resolves to a different case) — existing
- ✓ Embed cited opinion text into the active document — existing
- ✓ Zero-network-calls-by-default architecture; entirely client-side, no backend — existing
- ✓ Sideloadable installer packages for Windows/macOS, online and offline variants — existing

### Active

This milestone is maintenance/compliance-focused: reduce tech debt flagged in
`.planning/codebase/CONCERNS.md` and prepare for Microsoft Partner Center submission.

- [ ] Split `src/taskpane/word.ts` (1,538 lines) into feature-scoped modules (hyperlinking,
      Bluebook checking, hallucination checking, opinion-text embedding)
- [ ] Harden hyperlink/HTML insertion so `escapeHtml`/`isSafeHyperlinkUrl` cannot be skipped at a
      new Office.js insertion call site (no compiler-enforced guard exists today)
- [ ] Dedupe `westlawProvider.ts`/`lexisNexisProvider.ts`/`bloombergLawProvider.ts` (near-identical
      69-line files) into shared logic in `providers/base.ts`
- [ ] Land PR #33 ("Depend on openclerk-core instead of vendoring its logic") — CI green and
      mergeable as of this milestone's start; fixed a stale `openclerk-core` dependency pin
      (git tag v0.2.1 → npm `^0.2.6`) that both broke installs under npm's `allow-scripts` policy
      and missed two ReDoS fixes plus a hallucination-verification bypass fix
- [ ] After PR #33 lands, audit the rest of the repo (`src/commands/`, `scripts/`) for any
      remaining logic that duplicates `openclerk-core` and remove it
- [ ] Land PR #27 ("Add Privacy Policy and Terms of Use for Partner Center submission") — open,
      currently behind `main`, needs rebase
- [ ] Resolve `TERMS.md` §8 Governing Law jurisdiction (currently a placeholder,
      `[jurisdiction to be specified]`) — deferred to this milestone rather than decided now;
      may need counsel input

### Out of Scope

- New citation providers (e.g., finishing the `usptoPatentCenterProvider.ts` stub) — a distinct
  feature milestone, not this cleanup/compliance one
- New Bluebook rules or editions — no gaps identified for this milestone
- Configuring the `PARTNER_CENTER_*` CI secrets / actual Partner Center account submission —
  business/ops task outside of code changes; this milestone only prepares the required documents

## Context

Brownfield project (OpenClerk, repo `OpenClerkProject/openclerk-word`) — mature, actively
maintained, in production use by legal professionals via sideloaded installers. A sibling repo,
`openclerk-core`, was extracted from this repo to hold shared citation-parsing/Bluebook-rule logic
so it can be reused (e.g., a planned `openclerk-gdocs` port) and fixed in one place instead of
drifting across copies — see `.planning/codebase/CONCERNS.md` for the full tech-debt audit this
milestone works from.

Two PRs already exist and were reviewed for merge-readiness at the start of this milestone:
- **PR #33** (`claude/depend-on-openclerk-core`) — switches this repo to depend on
  `openclerk-core` instead of vendoring it. Verified during this session: CI green, `tsc`/`jest`/
  production `webpack` build all pass locally, and a stale dependency pin was fixed and pushed.
- **PR #27** (`docs/privacy-policy-terms-of-use`) — drafts `PRIVACY.md` and `TERMS.md` for Partner
  Center submission. Open, behind `main`, with one intentional placeholder (governing-law
  jurisdiction) left for the repo owner.

The CI pipeline already has a Partner Center publish job (`.github/workflows/ci.yml`), currently a
no-op because `PARTNER_CENTER_*` secrets aren't configured — that account/secrets setup is
explicitly out of scope for this milestone.

## Constraints

- **Tech stack**: TypeScript targeting `es5`/`ie 11` (browserslist), no frontend framework, direct
  DOM manipulation — new module extractions must follow the existing `providers/`/`bluebook/`
  registry-plugin pattern rather than introducing a framework
- **Compatibility**: Must keep working inside the Office.js WebView host across Word
  desktop/online on Windows and macOS
- **Security**: This is a legal tool — citation verification must stay conservative (never claim
  "verified" when uncertain); any refactor of the hallucination-check or hyperlink-insertion paths
  must preserve or strengthen existing guards, not weaken them

## Key Decisions

| Decision | Rationale | Outcome |
|----------|-----------|---------|
| Bumped PR #33's `openclerk-core` dependency from git tag `v0.2.1` to npm registry `^0.2.6` | v0.2.1 predates two ReDoS fixes and a `caseNamesMatch` hallucination-verification bypass fix in `openclerk-core`; the git-tag dependency also fails to install under npm's `allow-scripts` restriction (reproduced locally) | ✓ Good — pushed, CI re-ran green |
| `TERMS.md` §8 governing-law jurisdiction left unresolved | Repo owner deferred the actual jurisdiction choice to a milestone task rather than deciding during project setup | — Pending |

## Evolution

This document evolves at phase transitions and milestone boundaries.

**After each phase transition** (via `/gsd-transition`):
1. Requirements invalidated? → Move to Out of Scope with reason
2. Requirements validated? → Move to Validated with phase reference
3. New requirements emerged? → Add to Active
4. Decisions to log? → Add to Key Decisions
5. "What This Is" still accurate? → Update if drifted

**After each milestone** (via `/gsd-complete-milestone`):
1. Full review of all sections
2. Core Value check — still the right priority?
3. Audit Out of Scope — reasons still valid?
4. Update Context with current state

---
*Last updated: 2026-07-15 after initialization*
