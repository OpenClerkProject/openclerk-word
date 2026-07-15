# Requirements: OpenClerk

**Defined:** 2026-07-15
**Core Value:** Legal citation checks and hyperlinks must be accurate and trustworthy — a hallucination check must never falsely report a fabricated citation as "verified."

## v1 Requirements

Requirements for this milestone (tech-debt reduction + Microsoft Partner Center submission prep).
Each maps to roadmap phases.

### Escaping Hardening

- [ ] **ESCAPE-01**: Branded `SafeHtml`/`SafeHyperlinkUrl` types are added to `utils.ts` and used to
      type-constrain Office.js insertion helpers
- [ ] **ESCAPE-02**: A single wrapper module (`safeInsertion.ts`) owns all raw Office.js
      `insertHtml`/`insertHyperlink`/`insertComment` calls
- [ ] **ESCAPE-03**: An ESLint `no-restricted-syntax` rule bans direct calls to raw Office.js
      insertion APIs outside the wrapper module, failing the build on a bypass

### Provider Deduplication

- [ ] **PROVIDER-01**: A config-driven `GenericEnterpriseCitationProvider` in `providers/base.ts`
      replaces the near-duplicate logic in `westlawProvider.ts`, `lexisNexisProvider.ts`, and
      `bloombergLawProvider.ts`
- [ ] **PROVIDER-02**: Test coverage is extended so all three enterprise providers have equivalent
      coverage to each other

### openclerk-core Dependency Cleanup

- [ ] **CORE-01**: PR #33 ("Depend on openclerk-core instead of vendoring its logic") is merged to
      `main`
- [ ] **CORE-02**: `src/commands/` and `scripts/` are audited for any remaining logic duplicated in
      `openclerk-core`, and any found is removed

### Partner Center Manifest & Listing

- [ ] **MANIFEST-01**: `manifest.xml`'s `ProviderName`, `Description`, and `GetStarted` ribbon text
      are updated from Yeoman-template placeholders (`Contoso`, "A template to get started.") to
      accurate OpenClerk copy
- [ ] **MANIFEST-02**: A genuine 64×64 `HighResolutionIconUrl` icon asset is added (current asset
      is mis-sized at 80×80)

### Partner Center Submission Readiness

- [ ] **SUBMIT-01**: PR #27 ("Add Privacy Policy and Terms of Use for Partner Center submission")
      is rebased onto `main` and merged, with `TERMS.md` §8 (governing-law jurisdiction) tracked
      as a separate open task rather than blocking this PR
- [ ] **SUBMIT-02**: CI's Partner Center publish job is confirmed to validate the production-built
      package (`office-addin-manifest validate -p`) rather than the dev manifest
- [ ] **SUBMIT-03**: Certification/testing notes and a listing description are drafted, disclosing
      the optional external-service dependency (CourtListener / enterprise providers) for
      Microsoft's reviewer

## v2 Requirements

Deferred to future release. Tracked but not in current roadmap.

### word.ts Decomposition

- **WORDTS-01**: Split `src/taskpane/word.ts` into feature-scoped `workflows/*.ts` modules,
  starting with the low-risk pieces (`sourceDocument`, `documentActions`, `statusBar`,
  `bluebookCheck`, `embedOpinionText`)
- **WORDTS-02**: Split the high-risk workflows (`hyperlinking`, `onlineLookup`,
  `hallucinationCheck`) into their own modules, trimming `word.ts` to a thin composition root

### Manifest / Listing

- **MANIFEST-03**: Hide or disable the USPTO Patent Center provider stub from the UI until it's
  actually implemented (currently selectable, labeled `"USPTO Patent Center (TODO)"`)

### Legal

- **LEGAL-01**: Resolve `TERMS.md` §8 Governing Law jurisdiction (needs repo owner/counsel input)

## Out of Scope

Explicitly excluded. Documented to prevent scope creep.

| Feature | Reason |
|---------|--------|
| New citation providers (e.g. implementing USPTO Patent Center lookups) | Separate feature milestone, not this cleanup/compliance one |
| New Bluebook rules or editions | No gaps identified for this milestone |
| Configuring `PARTNER_CENTER_*` CI secrets / actual Partner Center submission | Business/ops task outside of code changes |
| `word.ts` module split (both low- and high-risk workflows) | Deferred to a future milestone — `word.ts` has zero existing test coverage (research finding), which raises the risk/effort for this pass; tracked as WORDTS-01/02 in v2 |

## Traceability

Which phases cover which requirements. Updated during roadmap creation.

| Requirement | Phase | Status |
|-------------|-------|--------|
| ESCAPE-01 | TBD | Pending |
| ESCAPE-02 | TBD | Pending |
| ESCAPE-03 | TBD | Pending |
| PROVIDER-01 | TBD | Pending |
| PROVIDER-02 | TBD | Pending |
| CORE-01 | TBD | Pending |
| CORE-02 | TBD | Pending |
| MANIFEST-01 | TBD | Pending |
| MANIFEST-02 | TBD | Pending |
| SUBMIT-01 | TBD | Pending |
| SUBMIT-02 | TBD | Pending |
| SUBMIT-03 | TBD | Pending |

**Coverage:**
- v1 requirements: 12 total
- Mapped to phases: 0 (filled by roadmap creation)
- Unmapped: 12 ⚠️

---
*Requirements defined: 2026-07-15*
*Last updated: 2026-07-15 after initial definition*
