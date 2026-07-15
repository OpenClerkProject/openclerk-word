# Roadmap: OpenClerk

## Overview

This milestone is maintenance/compliance-focused, not new-feature work. It lands a
already-reviewed dependency fix, closes a real hyperlink/HTML-injection gap at compile time,
collapses three near-duplicate enterprise citation providers into one config-driven
implementation, and prepares the manifest, listing copy, and legal docs for a Microsoft Partner
Center submission. The `word.ts` monolith split (WORDTS-01/02) surfaced during research but was
explicitly deferred to v2 during requirements scoping — it is out of scope for this roadmap.
Dependency cleanup is sequenced first so later refactors build on a clean, non-vendored
`openclerk-core` baseline; escaping hardening precedes provider dedup because it touches the
shared insertion path; Partner Center prep is sequenced last because its longest pole (legal
jurisdiction resolution) is a business decision outside the engineering critical path.

## Phases

**Phase Numbering:**
- Integer phases (1, 2, 3): Planned milestone work
- Decimal phases (2.1, 2.2): Urgent insertions (marked with INSERTED)

Decimal phases appear between their surrounding integers in numeric order.

- [ ] **Phase 1: openclerk-core Dependency Cleanup** - Land PR #33 and remove any remaining logic duplicated in openclerk-core
- [ ] **Phase 2: Escaping Hardening** - Make hyperlink/HTML insertion bypass-proof at compile time
- [ ] **Phase 3: Provider Deduplication** - Collapse Westlaw/LexisNexis/Bloomberg Law into one config-driven provider
- [ ] **Phase 4: Partner Center Submission Prep** - Fix manifest/listing/legal docs and confirm the CI publish path

## Phase Details

### Phase 1: openclerk-core Dependency Cleanup
**Goal**: The codebase depends cleanly on the published `openclerk-core` npm package, with no vendored duplicate logic remaining anywhere in the repo.
**Depends on**: Nothing (first phase)
**Requirements**: CORE-01, CORE-02
**Success Criteria** (what must be TRUE):
  1. PR #33 ("Depend on openclerk-core instead of vendoring its logic") is merged to `main` with CI green.
  2. `npm install` succeeds cleanly against the `openclerk-core` `^0.2.6` npm dependency (no git-tag/`allow-scripts` install failure reproduced).
  3. `src/commands/` and `scripts/` are audited for logic duplicated in `openclerk-core`; any found is removed and replaced with an `openclerk-core` import, or the audit confirms none exists.
**Plans**: 2 plans

Plans:
- [ ] 01-01-PLAN.md — Merge PR #33 (human checkpoint) and verify post-merge dependency/build/CI state plus the hallucination-check Core Value guard (CORE-01)
- [ ] 01-02-PLAN.md — Audit `src/commands/` and `scripts/` for logic duplicated in `openclerk-core`, record findings (CORE-02)

### Phase 2: Escaping Hardening
**Goal**: Hyperlink/HTML insertion into the Word document cannot bypass the existing escaping guards, at any current or future Office.js insertion call site.
**Depends on**: Phase 1
**Requirements**: ESCAPE-01, ESCAPE-02, ESCAPE-03
**Success Criteria** (what must be TRUE):
  1. Branded `SafeHtml`/`SafeHyperlinkUrl` types exist in `utils.ts` and type-constrain the input parameter of every Office.js insertion helper.
  2. Every raw Office.js `insertHtml`/`insertHyperlink`/`insertComment` call in the codebase lives inside `safeInsertion.ts` and nowhere else.
  3. ESLint fails the build (`no-restricted-syntax`) if a raw `insertHtml`/`insertHyperlink`/`insertComment` call is added outside `safeInsertion.ts`.
  4. Hyperlinking, Bluebook checking, hallucination checking, and opinion-text embedding still work correctly against a real Word document after the wrapper refactor (manual smoke test).
**Plans**: TBD

### Phase 3: Provider Deduplication
**Goal**: Westlaw, LexisNexis, and Bloomberg Law enterprise citation lookups share one implementation instead of three near-identical 69-line files.
**Depends on**: Phase 1
**Requirements**: PROVIDER-01, PROVIDER-02
**Success Criteria** (what must be TRUE):
  1. A config-driven `GenericEnterpriseCitationProvider` class exists in `providers/base.ts`.
  2. `westlawProvider.ts`, `lexisNexisProvider.ts`, and `bloombergLawProvider.ts` each shrink to a small config object that instantiates the generic provider, with no duplicated request/parsing logic remaining between them.
  3. All three enterprise providers still self-register and appear in the existing citation-provider dropdown after the refactor (no silent tree-shaking regression).
  4. Test coverage is equivalent across all three providers — the same scenarios/assertions run for each.
**Plans**: TBD

### Phase 4: Partner Center Submission Prep
**Goal**: OpenClerk's manifest, listing copy, legal documents, and CI publish path are ready for Microsoft Partner Center/AppSource submission.
**Depends on**: Nothing further (independent of Phases 1-3; sequenced last per research because its longest pole is a legal/business decision, not an engineering one)
**Requirements**: MANIFEST-01, MANIFEST-02, SUBMIT-01, SUBMIT-02, SUBMIT-03
**Success Criteria** (what must be TRUE):
  1. `manifest.xml`'s `ProviderName`, `Description`, and `GetStarted` ribbon text show accurate OpenClerk copy — no Yeoman-template placeholder (`Contoso`, "A template to get started.") remains.
  2. A genuine 64×64 `HighResolutionIconUrl` icon asset is in place (replacing the mis-sized 80×80 asset).
  3. PR #27 ("Add Privacy Policy and Terms of Use for Partner Center submission") is rebased onto `main` and merged, with `TERMS.md` §8 (governing-law jurisdiction) tracked as a separate open task rather than blocking the PR.
  4. `office-addin-manifest validate -p` passes against the production-built package, and CI's Partner Center publish job is confirmed to upload that production-built package rather than the dev manifest.
  5. Certification/testing notes and a draft listing description exist, disclosing the optional external-service dependency (CourtListener / enterprise providers) for Microsoft's reviewer.
**Plans**: TBD

## Progress

**Execution Order:**
Phases execute in numeric order: 1 → 2 → 3 → 4

| Phase | Plans Complete | Status | Completed |
|-------|----------------|--------|-----------|
| 1. openclerk-core Dependency Cleanup | 0/TBD | Not started | - |
| 2. Escaping Hardening | 0/TBD | Not started | - |
| 3. Provider Deduplication | 0/TBD | Not started | - |
| 4. Partner Center Submission Prep | 0/TBD | Not started | - |
