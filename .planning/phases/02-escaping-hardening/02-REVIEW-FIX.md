---
phase: 02-escaping-hardening
fixed_at: 2026-07-16T04:13:14Z
review_path: .planning/phases/02-escaping-hardening/02-REVIEW.md
iteration: 1
findings_in_scope: 7
fixed: 7
skipped: 0
status: all_fixed
---

# Phase 02: Code Review Fix Report

**Fixed at:** 2026-07-16T04:13:14Z
**Source review:** .planning/phases/02-escaping-hardening/02-REVIEW.md
**Iteration:** 1

**Summary:**
- Findings in scope: 7 (fix_scope: critical_warning -- CR-01, CR-02, WR-01 through WR-05; IN-01/IN-02 excluded from scope)
- Fixed: 7
- Skipped: 0

**Post-fix verification (run in the isolated worktree before writing this report):**
- `npx tsc --noEmit` -- clean, no errors
- `npm run lint` -- 0 errors, 3 pre-existing warnings (`office-addins/no-context-sync-in-loop`, unrelated to this fix batch)
- `npm test` -- 3 suites, 13 passed, 2 skipped (courtListener.live.test.ts is intentionally excluded from the default run), 0 failed
- `npm run build` -- webpack + build-docs.js completed successfully (only pre-existing bundle-size warnings)

## Fixed Issues

### CR-01: Unescaped URL enables HTML-attribute injection in the `insertHtml` fallback path

**Files modified:** `src/taskpane/safeInsertion.ts`
**Commit:** `9e0276c`
**Applied fix:** Imported `escapeHtml` from `openclerk-core` and wrapped the URL with it before splicing into the `href="..."` attribute in `insertSafeHyperlink`'s `insertHtml` branch: `` const html = `<a href="${escapeHtml(url)}">${displayText}</a>`; ``. `SafeHyperlinkUrl` only certifies scheme safety, not HTML-attribute safety, so this closes the gap the reviewer identified without weakening the branded-type guard itself. Verified empirically against `node_modules/openclerk-core`'s real `toSafeHyperlinkUrl`/`escapeHtml` that the malicious fixture URL from the review (`https://example.com/"><img src=x onerror=alert(1)>`) now escapes to a safe attribute value with no ability to break out of `href="..."`.

### CR-02: `SafeHtml` (HTML-escaped) content written into the plain-text-only `insertComment` API

**Files modified:** `src/taskpane/safeInsertion.ts`, `src/taskpane/word.ts`
**Commit:** `7d1ae9c`
**Applied fix:** Changed `insertSafeComment`'s `text` parameter type from `SafeHtml` to plain `string` (this is a corrected type-safety guarantee, not a weakening -- `Word.Range.insertComment` is a plain-text API and was never meant to receive escaped content). Updated the doc comment to explain the sink-specific rationale. Updated the call site at `word.ts:1327-1331` to pass `buildEmbeddedCommentContent(raw, excerpt)` directly, unescaped, instead of wrapped in `toSafeHtml(...)`. This also resolves the downstream `citationHasEmbeddedComment` matching bug the reviewer flagged (stored comment content now matches the unescaped `raw` string used for re-run detection).

### WR-01: ESLint exemption glob for `safeInsertion.ts` is unanchored to its directory

**Files modified:** `eslint.config.mjs`
**Commit:** `d873478`
**Applied fix:** Changed the exemption glob from `"**/safeInsertion.ts"` (basename-only match, matches any same-named file at any depth) to `"src/taskpane/safeInsertion.ts"` (path-anchored, matches only the real file). Updated the accompanying comment to describe the anchoring. Verified `npm run lint` still reports 0 errors for `safeInsertion.ts` (the exemption still applies) and the rest of `src/` is unaffected.

### WR-02: `insertOoxml` call in `word.ts` was outside both the wrapper and the ESLint guard

**Files modified:** `src/taskpane/safeInsertion.ts`, `src/taskpane/word.ts`, `eslint.config.mjs`
**Commit:** `778d92b`
**Applied fix:** Added a new `insertSafeOoxml(context, body, ooxml)` wrapper function to `safeInsertion.ts` (with a doc comment explaining its `ooxml` parameter is intentionally plain string since its only caller re-inserts document-derived, not attacker-influenced, OOXML). Updated `removeAllHyperlinks` in `word.ts` to call `insertSafeOoxml` instead of `body.insertOoxml(...)` directly. Extended `RAW_INSERTION_SELECTORS` in `eslint.config.mjs` to also flag `insertOoxml` outside `safeInsertion.ts`. Updated `safeInsertion.ts`'s header comment so it no longer overstates/understates the guard's coverage (now correctly lists all four raw insertion APIs and clarifies which two get compiler-enforced branded-type protection vs. which one doesn't need it). `npm run lint` confirms 0 errors post-change (no more unguarded raw-insertion call sites).

### WR-03: `safeInsertion.test.ts` never asserted on actual payload strings

**Files modified:** `tests/safeInsertion.test.ts`
**Commit:** `aff8ad8`
**Applied fix:** Added an assertion on the literal HTML string passed to `insertHtml` in the existing dispatch test. Added a new regression test using the review's own malicious-URL fixture (`https://example.com/"><img src=x onerror=alert(1)>`) asserting the exact escaped output, which would have caught CR-01 directly. Updated the `insertSafeComment` test to reflect the corrected plain-`string` signature, and added a new regression test with text containing `&`, `'`, and `"` asserting the stored value is byte-for-byte unchanged (no HTML entities), which would have caught CR-02 directly. All 8 tests in the suite pass (6 pre-existing + 2 new).

### WR-04: No compiler type-checking step anywhere in build/CI for `word.ts`

**Files modified:** `package.json`, `.github/workflows/ci.yml`
**Commit:** `6ef6dd0`
**Applied fix:** Added `"typecheck": "tsc --noEmit"` to `package.json` scripts. Added a new `typecheck` job to `.github/workflows/ci.yml`, structured identically to the existing `lint` job (`needs: build`, Node 18, `npm ci`, then `npm run typecheck` -- no `continue-on-error`, genuinely blocking). Added `typecheck` to the `publish` job's `needs` array alongside `lint`, so it gates release publishing the same way lint does. Verified locally: `npm run typecheck` passes cleanly against the whole `src/` tree (including `word.ts`, which was previously never type-checked by any pipeline step).

### WR-05: Stale comment claimed batched `context.sync()`, but every insertion syncs per item

**Files modified:** `src/taskpane/word.ts`
**Commit:** `c8972df`
**Applied fix:** Rewrote the comment preceding `applyCaseLawHyperlinksFromSource`'s `Word.run` block to accurately describe current behavior: the search/load/filter passes are batched into three `context.sync()` calls, but the insertion loop calls `insertSafeHyperlink` once per matched citation and `insertSafeHyperlink` unconditionally syncs internally, so the insertion phase itself still does roughly one round-trip per citation. Chose the "correct the comment" option (over "restore true batching") since it's the lower-risk fix that resolves the documentation-accuracy issue the reviewer flagged without touching the insertion/sync behavior itself (which fixes CR-01/CR-02 depend on and would be riskier to alter in the same pass); a follow-up phase can restore true batching if the freeze this comment originally described in fact reappears in practice.

## Skipped Issues

None -- all 7 in-scope findings were fixed.

---

_Fixed: 2026-07-16T04:13:14Z_
_Fixer: Claude (gsd-code-fixer)_
_Iteration: 1_
