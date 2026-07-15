# Project Research Summary

Project: WordClerk (OpenClerk) -- Office.js Word task-pane add-in
Domain: Compile-time HTML/hyperlink-escaping hardening plus monolithic controller decomposition plus Microsoft Partner Center/AppSource submission readiness
Researched: 2026-07-15
Confidence: HIGH

## Executive Summary

This milestone spans three distinct but interdependent efforts on a vanilla-TypeScript, es5/ie11-targeted Office.js Word add-in: (1) closing a real HTML/hyperlink-injection gap where escaping guards can be silently skipped at new Office.js insertion call sites, (2) splitting a 1538-line monolithic word.ts controller into feature-scoped workflows modules plus deduplicating three near-identical enterprise citation providers, and (3) preparing the manifest and listing artifacts for a Microsoft Partner Center (AppSource) submission. Confidence is HIGH throughout because every recommendation was verified against actual repo files.

The recommended approach is zero-new-dependency: branded TypeScript types combined with a core-ESLint no-restricted-syntax rule close the escaping gap at compile time; a feature-scoped workflows/shared directory split (following the existing providers/bluebook plugin-registry precedent) decomposes word.ts without a framework; and a short checklist of manifest and listing fixes closes the submission gap. Notably, the webpack build already strips localhost URLs from the packaged manifest and CI greps for their absence, so an initially-flagged localhost-in-manifest pitfall is a false alarm once cross-checked against FEATURES.md verified build-pipeline behavior -- the real gap is stale Yeoman placeholder text and icon sizing, not URL rewriting.

The dominant risk is sequencing: splitting word.ts before hardening escaping guards multiplies unguarded insertion call sites; splitting without first inventorying the seven module-level state variables in word.ts risks recreating tight coupling; and provider dedup must preserve self-registration side-effect imports or tree-shaking can silently drop a provider. All are addressed with concrete, low-cost mitigations rather than open unknowns.

## Key Findings

### Recommended Stack

No new dependencies. Escaping-hardening uses TypeScript branded types (TS 5.4.2, zero runtime cost, safe for es5/ie11) plus ESLint core no-restricted-syntax (already resolved via office-addin-lint). Branded types alone are insufficient -- TypeScript declaration merging cannot narrow the ambient Word.Range.insertHtml(html: string) signature, so an ESLint gate is required alongside the type.

Core technologies:
- TypeScript branded types (SafeHtml, SafeHyperlinkUrl) -- zero runtime cost, no es5/ie11 risk
- ESLint core no-restricted-syntax (AST selector) -- bans raw insertHtml/insertHyperlink/insertComment outside one wrapper module
- No runtime sanitizer (DOMPurify etc.) -- existing escapeHtml/isSafeHyperlinkUrl already sufficient

### Expected Features

Scoped to what a Partner Center/AppSource submission requires for this free, no-account, client-side add-in.

Must have (table stakes, blocks submission):
- Fix manifest.xml ProviderName (currently literal Contoso placeholder)
- Fix Description and GetStarted ribbon text (unedited Yeoman boilerplate)
- Add genuine 64x64 HighResolutionIconUrl asset (currently mis-sized at 80x80)
- Land PRIVACY.md/TERMS.md as live HTTPS URLs with TERMS.md section 8 jurisdiction resolved
- Confirm CI publish uploads the production-built package and runs office-addin-manifest validate -p
- Draft certification/testing notes and listing Description disclosing optional external-service dependency
- Correction to an initial pitfall flag: manifest.xml dev localhost:3000 URLs are NOT a live risk -- webpack already string-replaces them at build time and CI greps the packaged manifest to fail the build if localhost survives. This was raised as a critical pitfall in initial research but is resolved as a false alarm per FEATURES.md verified build-pipeline check.

Should have (differentiators):
- 3-5 real store screenshots beyond the required minimum
- Deliberate Marketplace category/industry selection

Defer (v2+):
- Unified (JSON) manifest migration -- no current forcing deadline
- Non-English localized listing
- Post-approval installation deep-link in README

### Architecture Approach

Split monolithic word.ts (1538 lines) into a thin composition root plus feature-scoped workflows/*.ts modules (hyperlinking, onlineLookup, bluebookCheck, hallucinationCheck, embedOpinionText) and a shared/* leaf layer, following the existing providers/bluebook plugin-registry precedent. Dependency direction is strictly one-way; each workflow self-wires via init()/refresh() and owns its own state. Separately, the three near-identical enterprise providers collapse into one GenericEnterpriseCitationProvider config-driven class.

Major components:
1. word.ts (shrinks to about 150-250 lines) -- Office.onReady composition root
2. workflows/(feature).ts -- one file per workflow, self-wiring, colocated state
3. shared/* -- leaf helpers used by more than one workflow
4. providers/base.ts -- new GenericEnterpriseCitationProvider

### Critical Pitfalls

1. Scattered module-level mutable state during the split -- inventory state ownership before moving code; use accessor functions, not exported let.
2. Stale or untracked Word.js proxy objects across new module boundaries -- pass identifiers, not live proxy objects, across boundaries.
3. Escaping guards remain manual discipline if the split happens before the hardening work -- sequence the guarded insertion wrapper before or atomically with the split.
4. Self-registering providers/rule-sets silently vanish from the production bundle if side-effect imports are refactored away -- keep registration imports static and top-level; add a registry-completeness test.
5. Registry population racing Office.onReady/UI wiring -- causes empty dropdowns on first load only.

## Implications for Roadmap

Escaping hardening must precede or accompany the word.ts split. Partner Center submission-prep work is independent and can run in parallel since it touches manifest.xml and docs, not controller code.

### Phase 1: Compile-Time Hyperlink/HTML Escaping Hardening
Rationale: Must land before or atomically with the word.ts split -- splitting first multiplies unguarded insertion call sites.
Delivers: Branded SafeHtml/SafeHyperlinkUrl types in utils.ts, a single wrapper module (safeInsertion.ts) owning all raw Office.js insertion calls, and an ESLint no-restricted-syntax rule banning direct calls outside it.
Addresses: PROJECT.md harden hyperlink/HTML insertion requirement.
Avoids: Escaping-stays-manual-discipline pitfall; unguarded-insertion-site security risk.

### Phase 2: Provider Deduplication
Rationale: Smallest, most self-contained refactor piece; already has partial test coverage; validates the config-driven-subclass pattern before the larger word.ts split.
Delivers: GenericEnterpriseCitationProvider in providers/base.ts; Westlaw/LexisNexis/BloombergLaw shrink to about 15-line configs; test coverage extended to all three.
Uses: Config-driven subclass pattern from ARCHITECTURE.md.
Implements: providers/base.ts generic class; self-registration unchanged.

### Phase 3: word.ts Split -- Low-Risk Workflows First
Rationale: Zero existing test coverage of word.ts means mechanical extraction of lowest-risk pieces first, each its own commit with manual smoke-test verification.
Delivers: shared/sourceDocument.ts (first real unit tests), shared/documentActions.ts, shared/statusBar.ts, then workflows/bluebookCheck.ts and workflows/embedOpinionText.ts.
Addresses: PROJECT.md goal of reducing merge-conflict risk.
Avoids: Scattered-state pitfall via explicit ownership inventory before moving code; big-bang single-PR extraction anti-pattern.

### Phase 4: word.ts Split -- High-Risk Workflows (Hyperlinking, Online Lookup, Hallucination Check)
Rationale: Higher risk due to hand-tuned context.sync() batching and ties to the core trust claim of the product; saved for last so the pattern is proven on prior extractions first.
Delivers: workflows/hyperlinking.ts and workflows/onlineLookup.ts together, then workflows/hallucinationCheck.ts last; final trim of word.ts to about 150-250 lines.
Addresses: Full completion of the module split described in PROJECT.md.
Avoids: Stale proxy-object pitfall; registry-vs-Office.onReady race; duplicate context.sync() performance trap.

### Phase 5: Partner Center / AppSource Submission Prep
Rationale: Independent of the controller refactor -- can run in parallel with Phases 1-4; sequenced last here because its longest pole (privacy/terms jurisdiction resolution) is a legal/business decision outside the engineering critical path.
Delivers: Fixed ProviderName/Description/GetStarted text, real 64x64 icon, published PRIVACY.md/TERMS.md with section 8 resolved, confirmed production-manifest CI publish path, draft certification/testing notes, draft listing Description, USPTO stub hidden from provider UI.
Delivers (features): All Table Stakes items from FEATURES.md.
Avoids: Wrong/dev manifest submitted (already mitigated by build pipeline, verify with pre-submission grep); unclear first-run experience; USPTO stub visible.

### Phase Ordering Rationale

- Escaping hardening (Phase 1) precedes the word.ts split because splitting multiplies insertion call sites -- doing it first means every extracted workflow is built against the safe wrapper from day one.
- Provider dedup (Phase 2) precedes the word.ts split because it is smaller, already partially tested, and validates the config-driven-subclass pattern in isolation.
- The word.ts split (Phases 3-4) is ordered low-risk-to-high-risk per the ARCHITECTURE.md Migration Order, with the state-ownership inventory as the mandatory first step of Phase 3.
- Partner Center submission prep (Phase 5) has no code dependency on the refactor phases; sequenced last only because its longest pole is out of the engineering critical path.

### Research Flags

Phases likely needing deeper research during planning:
- Phase 3-4 (word.ts split): ARCHITECTURE.md is HIGH confidence for codebase-specific structure but MEDIUM for general prior art (no Office.js-specific guidance exists on this decomposition problem) -- re-validate migration-order and proxy-object-tracking rules during plan-phase.
- Phase 5 (Partner Center submission): Category/industry tag selection depends on the live Partner Center picker UI, not confirmed in this research pass.

Phases with standard patterns (skip deep research-phase):
- Phase 1 (escaping hardening): HIGH confidence, concrete code sketch already in STACK.md, zero new dependencies.
- Phase 2 (provider dedup): HIGH confidence, standard Factory/config-object refactor pattern, existing test scaffolding to extend.

## Confidence Assessment

| Area | Confidence | Notes |
|------|------------|-------|
| Stack | HIGH | Verified directly against word.ts, utils.ts, resolved node_modules versions |
| Features | HIGH | Sourced from current learn.microsoft.com docs (2025-09 through 2026-07), cross-checked against manifest.xml/webpack.config.js/ci.yml in this repo |
| Architecture | HIGH (codebase-specific) / MEDIUM (general prior art) | Codebase findings HIGH; no official Microsoft guidance on this specific decomposition problem exists |
| Pitfalls | MEDIUM | Microsoft Learn docs cross-checked, but several sources are community-reported GitHub issues (LOW-MEDIUM) |

Overall confidence: HIGH

### Gaps to Address

- TERMS.md section 8 governing-law jurisdiction -- deferred to repo owner/counsel per PROJECT.md; not an engineering decision, but blocks Phase 5 completion.
- Partner Center category/industry tag picker contents -- no canonical Legal category confirmed; needs live check at submission time.
- Unified (JSON) manifest migration deadline -- no explicit deprecation statement found for XML-only submissions; correctly deferred, re-check if a future milestone adds Teams/Copilot surface.
- Accessibility conformance audit -- flagged as a real Marketplace requirement but out of scope for this research pass; track as a separate backlog item.

## Sources

### Primary (HIGH confidence)
- Direct codebase inspection: word.ts, utils.ts, providers/*, package.json, tsconfig.json, .eslintrc.json, manifest.xml, webpack.config.js, ci.yml, package-release*.js, .planning/codebase/ARCHITECTURE.md, STRUCTURE.md, CONCERNS.md, PROJECT.md
- Microsoft 365 app publishing checklist (learn.microsoft.com) -- updated 2025-09-25
- Microsoft Marketplace step-by-step submission guide -- updated 2025-09-25
- Certification policies for Microsoft Marketplace
- Publish your Office Add-in to Microsoft Marketplace -- updated 2026-07-10
- ESLint no-restricted-syntax rule docs

### Secondary (MEDIUM confidence)
- Top 5 AppSource validation errors for Office Add-ins submissions -- Microsoft 365 Developer Blog series
- Avoid using context.sync in loops -- Office Add-ins, Microsoft Learn
- Branded Types -- Learning TypeScript
- Duplicate Code -- Refactoring Guru

### Tertiary (LOW confidence)
- context.trackedObjects.add across Word.run calls -- OfficeDev/office-js issue 68 (community-reported)
- context.sync() taking progressively more time -- OfficeDev/office-js issue 3565 (community-reported)

---
Research completed: 2026-07-15
Ready for roadmap: yes
