---
phase: 02-escaping-hardening
plan: 03
subsystem: security
tags: [typescript, branded-types, office-js, safeInsertion, escaping, hyperlink-validation, migration]

# Dependency graph
requires:
  - phase: 02-escaping-hardening (Plan 02)
    provides: "src/taskpane/safeInsertion.ts exporting insertSafeHyperlink/insertSafeComment with branded-type-only parameters, and openclerk-core@0.3.0's toSafeHyperlinkUrl/toSafeHtml smart constructors"
provides:
  - "src/taskpane/word.ts contains zero raw insertHtml/insertHyperlink/insertComment calls -- every raw Office.js insertion call site now routes through safeInsertion.ts's insertSafeHyperlink/insertSafeComment"
  - "src/taskpane/word.ts contains zero references to the old plain (unbranded) escapeHtml/isSafeHyperlinkUrl functions -- all validation flows through toSafeHtml/toSafeHyperlinkUrl"
  - "All three hyperlink-insertion workflows (file-based, parenthetical, provider-lookup) and the pincite comment-embedding workflow verified passing npm test && npm run build after the migration"
affects: [02-04]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "Branded SafeHyperlinkUrl values threaded end-to-end through local map/filter pipelines (citationEntries, validEntries, matchedItems) instead of re-validating a plain string at the insertion call site -- single validation, reused downstream"
    - "word.ts is now purely a consumer of safeInsertion.ts's wrapper functions -- it no longer contains any Office.js insertion dispatch logic of its own (that logic lives exclusively in safeInsertion.ts per Plan 02)"

key-files:
  modified:
    - "src/taskpane/word.ts"

key-decisions:
  - "applyHyperlinkToItem deleted in its entirety (not deprecated/aliased) -- its 3-tier dispatch logic was already ported to safeInsertion.ts's insertSafeHyperlink in Plan 02, so this was a pure removal, not a rewrite."
  - "CitationMap's value type intentionally left as plain Map<string, string> in parseSourceDocument, per the plan's explicit scope boundary -- that guard is a parse-time quality filter, not the insertion security gate (toSafeHyperlinkUrl re-validates inside applyCaseLawHyperlinksFromSource, the actual gate before insertion)."

requirements-completed: [ESCAPE-01, ESCAPE-02]

coverage:
  - id: D1
    description: "word.ts contains zero calls to insertHtml/insertHyperlink/insertComment -- every raw Office.js insertion call now lives exclusively in safeInsertion.ts"
    requirement: "ESCAPE-02"
    verification:
      - kind: other
        ref: "grep -c \"applyHyperlinkToItem\" src/taskpane/word.ts == 0; grep -c \"\\.insertHtml(\\|\\.insertHyperlink(\\|\\.insertComment(\" src/taskpane/word.ts == 0"
        status: pass
    human_judgment: false
  - id: D2
    description: "word.ts contains zero references to the old plain HTML-escaping or URL-validation functions -- all validation now flows through toSafeHtml/toSafeHyperlinkUrl"
    requirement: "ESCAPE-01"
    verification:
      - kind: other
        ref: "grep -c \"escapeHtml\\|isSafeHyperlinkUrl\" src/taskpane/word.ts == 0; grep -c \"toSafeHyperlinkUrl\" src/taskpane/word.ts == 5; grep -c \"toSafeHtml\" src/taskpane/word.ts == 4 (Task 1 count, pre-Task-2 import-only additions unchanged); grep -c \"insertSafeHyperlink\" src/taskpane/word.ts == 4; grep -c \"insertSafeComment\" src/taskpane/word.ts == 2"
        status: pass
    human_judgment: false
  - id: D3
    description: "npx tsc --noEmit, npm test, and npm run build all exit 0 after the full migration -- no regressions introduced by rewiring the three hyperlink-insertion workflows and the comment-insertion call site onto branded types"
    requirement: "ESCAPE-01"
    verification:
      - kind: other
        ref: "npx tsc --noEmit (exit 0); npm test (3 suites, 11 passed / 2 skipped, 0 failed); npm run build (webpack production build succeeded, pre-existing bundle-size warnings only, unrelated to this plan)"
        status: pass
    human_judgment: false
  - id: D4
    description: "Hyperlinking, Bluebook checking, hallucination checking, and opinion-text embedding all still work correctly against a real Word document after the refactor (Roadmap Phase 2 Success Criterion 4)"
    requirement: "ESCAPE-02"
    verification:
      - kind: manual
        ref: "Open a real Word document (desktop or online) with OpenClerk sideloaded. Run through all four workflows end-to-end: (1) Manage Hyperlinks -- load a source .docx and apply case-law hyperlinks, confirm links are inserted correctly; (2) Bluebook Check -- run a check against document text, confirm issues render; (3) Hallucination Check -- connect a provider (e.g. CourtListener) and run a hallucination scan, confirm the \"possible hallucination\" guard still renders correctly for a deliberately-mismatched citation; (4) Embed Cited Text -- embed opinion text for a pincite citation, confirm a Word comment is inserted with the expected excerpt. Confirm no workflow regressed after the safeInsertion.ts wrapper refactor (Roadmap Phase 2 Success Criterion 4)."
        status: pending
    human_judgment: true

# Metrics
duration: ~20min
completed: 2026-07-16
status: complete
---

# Phase 2 Plan 3: Migrate word.ts's Insertion Call Sites onto safeInsertion.ts Summary

**Rewired every raw Office.js insertion call site and every URL/HTML validation call site in `src/taskpane/word.ts` onto `safeInsertion.ts`'s wrapper functions and `openclerk-core`'s branded smart constructors, deleting the local `applyHyperlinkToItem` dispatch function entirely.**

## Performance

- **Duration:** ~20 min
- **Tasks:** 2 of 2 completed
- **Files modified:** 1 (`src/taskpane/word.ts`)

## Accomplishments
- Updated the `openclerk-core` import block to swap `escapeHtml`/`isSafeHyperlinkUrl` for `toSafeHtml`/`toSafeHyperlinkUrl`/`SafeHyperlinkUrl`, and added a new `import { insertSafeHyperlink, insertSafeComment } from "./safeInsertion";`
- Deleted `applyHyperlinkToItem` in its entirety -- its 3-tier dispatch logic already lives in `safeInsertion.ts` (Plan 02)
- Migrated `applyCaseLawHyperlinksFromSource`: restructured `citationEntries` into a `.map(...).filter(...)` pipeline producing branded `SafeHyperlinkUrl` values, threaded the branded type through `matchedItems`, and replaced the insertion call with `insertSafeHyperlink(context, item, url, toSafeHtml(normalizedText))`
- Migrated `addParentheticalHyperlinks`: restructured `validEntries`'s filter to map each entry's URL through `toSafeHyperlinkUrl` and drop `null` results, threaded the branded type through `matchedItems`, and replaced the insertion call with `insertSafeHyperlink(context, item, entry.url, toSafeHtml(entry.citation))`
- Migrated `applyHyperlinksViaProvider`: replaced the combined guard with a two-step `safeUrl = match ? toSafeHyperlinkUrl(match.url) : null` check, preserving the exact rate-limit-vs-skip branching logic, and replaced the insertion call with `insertSafeHyperlink(context, item, safeUrl, toSafeHtml(raw))`
- Migrated the pincite opinion-text embedding workflow's comment-insertion call site to `insertSafeComment(context, searchResults.items[0], toSafeHtml(buildEmbeddedCommentContent(raw, excerpt)))`, removing the now-redundant standalone `context.sync()` call since the wrapper owns its own sync internally
- Migrated `parseSourceDocument`'s citation-map guard's plain URL-safety predicate to `toSafeHyperlinkUrl(url) !== null`, leaving `CitationMap`'s value type as plain `Map<string, string>` per the plan's explicit scope boundary
- Confirmed via grep that `word.ts` contains zero references to `applyHyperlinkToItem`, `escapeHtml`, `isSafeHyperlinkUrl`, and zero raw `.insertHtml(`/`.insertHyperlink(`/`.insertComment(` calls
- `npx tsc --noEmit`, `npm test` (3 suites, 11 passed / 2 skipped), and `npm run build` (production webpack build) all green after the full migration

## Task Commits

Each task was committed atomically:

1. **Task 1: Migrate imports and the three hyperlink-insertion workflows** - `9ede8d6` (feat)
2. **Task 2: Migrate the comment-insertion call site and the parseSourceDocument guard; verify no raw call or old validator remains** - `9562fd6` (feat)

## Files Created/Modified
- `src/taskpane/word.ts` - all raw Office.js insertion calls and unbranded validator calls migrated to `safeInsertion.ts`'s wrapper functions and `openclerk-core`'s branded smart constructors; `applyHyperlinkToItem` deleted

## Decisions Made
- `applyHyperlinkToItem` was deleted outright rather than deprecated/aliased -- its logic was already fully ported to `safeInsertion.ts` in Plan 02, so this plan's job was pure removal plus call-site rewiring, not a parallel implementation.
- `CitationMap`/`sourceCitationMap` module-level state intentionally kept as plain `Map<string, string>` -- threading branded types through that module-level state was explicitly out of this plan's scope per the plan's own action text; the actual insertion security gate is `toSafeHyperlinkUrl`'s re-validation inside `applyCaseLawHyperlinksFromSource`, which already runs on every URL immediately before insertion.

## Deviations from Plan

None - plan executed exactly as written. Every task's acceptance-criteria grep counts matched the plan's specified exact values on the first pass (`insertSafeHyperlink` == 4, `toSafeHyperlinkUrl` == 5, `toSafeHtml` == 4, `insertSafeComment` == 2, `applyHyperlinkToItem`/`escapeHtml`/`isSafeHyperlinkUrl`/raw-insertion-call counts == 0), and `npx tsc --noEmit`/`npm test`/`npm run build` all passed without needing any auto-fix.

## Issues Encountered

None.

## Human Verification Required

**Task 2's `<human-check>` was not run in this session** (per this project's `human_verify_mode: end-of-phase` config -- deferred to end-of-phase UAT harvest, not blocking this plan's execution). The automated checks (`npm test && npm run build`) were run and passed; `npx tsc --noEmit` also passed. The exact manual verification item to carry forward:

> Open a real Word document (desktop or online) with OpenClerk sideloaded. Run through all four workflows end-to-end: (1) Manage Hyperlinks -- load a source .docx and apply case-law hyperlinks, confirm links are inserted correctly; (2) Bluebook Check -- run a check against document text, confirm issues render; (3) Hallucination Check -- connect a provider (e.g. CourtListener) and run a hallucination scan, confirm the "possible hallucination" guard still renders correctly for a deliberately-mismatched citation; (4) Embed Cited Text -- embed opinion text for a pincite citation, confirm a Word comment is inserted with the expected excerpt. Confirm no workflow regressed after the safeInsertion.ts wrapper refactor (Roadmap Phase 2 Success Criterion 4).

This item is recorded in the `coverage` frontmatter above (id: D4, `human_judgment: true`, `status: pending`) so the phase verifier can harvest it into a UAT file at end-of-phase.

## User Setup Required

None - no external service configuration required.

## Next Phase Readiness

**Ready for Plan 04.** `word.ts` now contains zero raw Office.js insertion calls and zero references to the old unbranded validators; ESCAPE-03's ESLint `no-restricted-syntax` bypass guard (Plan 04's scope) can now be added without flagging any pre-existing violation in `word.ts`, since every call site already routes through `safeInsertion.ts`.

---
*Phase: 02-escaping-hardening*
*Completed: 2026-07-16*

## Self-Check: PASSED

All claimed changes verified: `src/taskpane/word.ts` modified (confirmed via `git status`/`git diff`), both commit hashes (`9ede8d6`, `9562fd6`) present in `git log --oneline --all`, and this SUMMARY.md written to disk.
