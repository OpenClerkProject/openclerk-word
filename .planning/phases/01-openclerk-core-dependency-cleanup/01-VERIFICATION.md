---
phase: 01-openclerk-core-dependency-cleanup
verified: 2026-07-15T00:00:00Z
status: passed
score: 12/12 must-haves verified
behavior_unverified: 0
overrides_applied: 0
---

# Phase 01: openclerk-core Dependency Cleanup Verification Report

**Phase Goal:** The codebase depends cleanly on the published `openclerk-core` npm package, with no
vendored duplicate logic remaining anywhere in the repo.
**Verified:** 2026-07-15
**Status:** passed
**Re-verification:** No — initial verification

## Goal Achievement

### Observable Truths

| # | Truth | Status | Evidence |
|---|-------|--------|----------|
| 1 | PR #33 is merged to `main` with CI green (Roadmap SC #1) | ✓ VERIFIED | `git show --stat 0f48462` confirms merge commit `0f48462` ("Merge pull request #33 from OpenClerkProject/claude/depend-on-openclerk-core") is on `main` (`git merge-base --is-ancestor 0f48462 HEAD` → true). `gh run list --branch main` shows the CI workflow triggered by that exact merge commit (run `29445036558`, "Merge pull request #33...") completed with `success`. |
| 2 | Merge landed via a genuine merge commit (not squash/rebase), preserving PR #33's original commits (D-01) | ✓ VERIFIED | `git log --oneline 006fbf3..58b2e13` lists the PR branch's discrete commits (`afdcef2` "Depend on openclerk-core instead of vendoring its logic", `4fcac93` "Migrate Find Hallucinations onto openclerk-core's shared function", `58b2e13` "Switch openclerk-core dependency from git tag to npm registry", plus 2 intermediate `main`-sync merges) individually reachable from `0f48462`'s second parent — not squashed into one commit. |
| 3 | `npm install` succeeds cleanly against `openclerk-core@^0.2.6` — no git-tag/allow-scripts install failure (Roadmap SC #2) | ✓ VERIFIED | `package.json` `dependencies.openclerk-core` = `"^0.2.6"`. `npm ci` completed successfully ("added 1421 packages... found 0 vulnerabilities"); the `allow-scripts` warnings present are for unrelated pre-existing packages (`@azure/msal-node-extensions`, `sharp`, `core-js`, `keytar`, etc.), not `openclerk-core`, and are non-fatal warnings, not install failures. |
| 4 | `openclerk-core` resolves from the public npm registry, not a git-tag/tarball/local override | ✓ VERIFIED | `package-lock.json` `node_modules/openclerk-core` entry: `"resolved": "https://registry.npmjs.org/openclerk-core/-/openclerk-core-0.2.6.tgz"`. `npm ls openclerk-core` reports `openclerk-core@0.2.6`. |
| 5 | `src/commands/` and `scripts/` audited for logic duplicated in `openclerk-core`; any found removed and replaced, or audit confirms none exists (Roadmap SC #3) | ✓ VERIFIED | Independently re-ran the audit's grep commands: `grep -rn "parseCaseCitation\|ParsedCitation\|BluebookRuleSet\|escapeHtml\|isSafeHyperlinkUrl\|citationProviderRegistry\|bluebookRuleSetRegistry" src/commands/` → 0 matches. `test -f scripts/generate-bluebook-data.js` → absent (deleted by PR #33, the one file with real overlap). Same identifier grep across the 5 remaining `scripts/*.js` files → 0 matches. |
| 6 | No file anywhere in `src/` or `scripts/` still imports from the deleted local `./providers`, `./bluebook`, or `./utils` paths | ✓ VERIFIED | `grep -rn 'from "\./providers"\|from "\./bluebook"\|from "\./utils"\|from "\.\./providers\|from "\.\./bluebook\|from "\.\./utils' src/ scripts/` → 0 matches (independently re-run, not just trusted from the audit doc). `src/taskpane/providers/`, `src/taskpane/bluebook/`, `src/taskpane/utils.ts` all confirmed absent from the working tree. |
| 7 | `src/taskpane/word.ts` imports hallucination-check / citation / Bluebook symbols from `openclerk-core`, not deleted local paths | ✓ VERIFIED | `src/taskpane/word.ts:11-31` imports `citationProviderRegistry`, `checkCitationsForHallucinations`, `HallucinationCheckResult`, `bluebookRuleSetRegistry`, `escapeHtml`, `isSafeHyperlinkUrl`, `parseCaseCitation`, etc. all `from "openclerk-core"` (single import statement, no other local-path import of these symbols exists). |
| 8 | The `nameMismatch` "Possible hallucination" rendering branch (Core Value guard) is preserved verbatim post-merge | ✓ VERIFIED | `src/taskpane/word.ts:1067,1073`: `else if (result.nameMismatch) { ... status.textContent = \`Possible hallucination -- ${result.nameMismatch.provider} resolves this citation to a different case: "${result.nameMismatch.foundCaseName}".\`; }`. `result.verifiedVia` gates the "Verified" path (checked first), so an unverified/mismatched result can never be silently upgraded. |
| 9 | No orphaned old hand-rolled hallucination type/loop remains duplicated alongside `openclerk-core`'s `HallucinationCheckResult` | ✓ VERIFIED | `grep -n "interface HallucinationResult\|type HallucinationResult" src/` → 0 matches. The only `HallucinationResult`-substring hits are inside the current, correct `renderHallucinationResults` function name (not a duplicated old type) — consistent with 01-01-SUMMARY.md's documented false-positive investigation, independently re-confirmed here. |
| 10 | `npm run build` (production webpack) succeeds post-merge | ✓ VERIFIED | `npm run build` exits 0, produces `dist/index.html` and bundles `openclerk-core/lib/bluebook` + `openclerk-core/lib/providers` modules into `taskpane.js`; only non-blocking bundle-size warnings (411 KiB > 244 KiB recommended limit), no errors. |
| 11 | `npm test` (Jest) passes post-merge | ✓ VERIFIED | `npm test`: 2 suites (`installer.test.ts`, `manifest.test.ts`), 5 passed / 2 skipped, 7 total, exit 0. |
| 12 | A persistent, dated audit record exists documenting CORE-02's grep evidence (falsifiable, not verbal-only) | ✓ VERIFIED | `.planning/phases/01-openclerk-core-dependency-cleanup/01-DUPLICATION-AUDIT.md` exists, contains literal grep commands + output for both `src/commands/` and `scripts/`, and a "CORE-02: confirmed" conclusion. `.planning/codebase/CONCERNS.md`'s "`openclerk-core` logic duplicated, not shared" entry carries a `Resolved (2026-07-15)` line referencing the audit file and merge commit `0f48462`. |

**Score:** 12/12 truths verified (0 present, behavior-unverified)

### Required Artifacts

| Artifact | Expected | Status | Details |
|----------|----------|--------|---------|
| `.planning/phases/01-openclerk-core-dependency-cleanup/01-DUPLICATION-AUDIT.md` | CORE-02 audit record with grep evidence | ✓ VERIFIED | Present, contains literal commands/output, "CORE-02" appears 3×. |
| `.planning/codebase/CONCERNS.md` | Duplication tech-debt entry marked resolved | ✓ VERIFIED | `Resolved (2026-07-15)` line present, references audit file + merge SHA. |
| `src/taskpane/providers/`, `src/taskpane/bluebook/`, `src/taskpane/utils.ts` | Deleted (vendored logic removed) | ✓ VERIFIED (absence confirmed) | `ls` returns exit 2 / not-found for all three paths. |
| `scripts/generate-bluebook-data.js` | Deleted (the one `scripts/` file with real overlap) | ✓ VERIFIED (absence confirmed) | `test -f` returns non-zero. |

### Key Link Verification

| From | To | Via | Status | Details |
|------|-----|-----|--------|---------|
| `package.json` `dependencies.openclerk-core` | `node_modules/openclerk-core` | npm registry resolution | ✓ WIRED | `package-lock.json` resolved field = `https://registry.npmjs.org/openclerk-core/-/openclerk-core-0.2.6.tgz`; `npm ls openclerk-core` confirms `0.2.6` installed. |
| `src/taskpane/word.ts` import statement | `openclerk-core`'s `checkCitationsForHallucinations` | direct npm-package import (not local path) | ✓ WIRED | `word.ts:23` imports and `word.ts:1009` calls `checkCitationsForHallucinations(candidates, selectedProviders)`. |
| `HallucinationCheckResult.nameMismatch` | task-pane DOM status rendering | `renderHallucinationResults` | ✓ WIRED | `word.ts:1067-1073` renders the flagged branch before any "Verified" status can be shown. |
| CI workflow (`.github/workflows/ci.yml`) | merge commit `0f48462` | GitHub Actions push trigger | ✓ WIRED | `gh run list --branch main` shows the CI run triggered directly by the `0f48462` push completed with `success`. |
| Task 1/Task 2 grep evidence (Plan 01-02) | `01-DUPLICATION-AUDIT.md` | audit record | ✓ WIRED | Grep commands and literal output independently re-run during this verification produce identical (zero-match) results to those recorded in the audit file. |

### Behavioral Spot-Checks

| Behavior | Command | Result | Status |
|----------|---------|--------|--------|
| Clean install against `openclerk-core@^0.2.6` | `npm ci` | "added 1421 packages... found 0 vulnerabilities", exit 0 | ✓ PASS |
| Production build succeeds with `openclerk-core` bundled | `npm run build` | webpack compiled with 3 non-blocking size warnings, `dist/index.html` written, exit 0 | ✓ PASS |
| Full test suite passes | `npm test` | 2 suites, 5 passed / 2 skipped / 7 total, exit 0 | ✓ PASS |
| Zero duplication in `src/commands/` (independent re-run, not trusted from audit doc) | `grep -rn <identifier-list> src/commands/` | 0 matches | ✓ PASS |
| Zero residual imports of deleted local paths (independent re-run) | `grep -rn 'from "\./providers"...' src/ scripts/` | 0 matches | ✓ PASS |
| CI green on the exact merge commit | `gh run list --branch main` | Run for "Merge pull request #33..." push = `success` | ✓ PASS |

### Requirements Coverage

| Requirement | Source Plan | Description | Status | Evidence |
|-------------|-------------|-------------|--------|----------|
| CORE-01 | 01-01-PLAN.md | PR #33 merged to `main` | ✓ SATISFIED | Merge commit `0f48462` on `main`, CI green, `npm ci`/`build`/`test` all pass, hallucination guard intact. |
| CORE-02 | 01-02-PLAN.md | `src/commands/`/`scripts/` audited, duplication removed | ✓ SATISFIED | Fresh audit (independently re-verified) confirms zero duplication; `01-DUPLICATION-AUDIT.md` records evidence; `CONCERNS.md` closed. |

No orphaned requirements — REQUIREMENTS.md maps only CORE-01/CORE-02 to Phase 1, both are claimed by the two plans and both verified above.

### Anti-Patterns Found

| File | Line | Pattern | Severity | Impact |
|------|------|---------|----------|--------|
| `package.json` / `README.md` | `package.json:35`, `README.md:266-278` | `npm run test:live` script references `tests/courtListener.live.test.ts`, deleted by this phase's merge (PR #33 moved that test into `openclerk-core`) but the script/docs were not updated | ⚠️ Warning (info per task instructions) | Running the documented `test:live` command fails unconditionally ("No tests found"). Does not affect `npm install`/`build`/`test`/CI or any of the phase's roadmap success criteria — flagged as a documented code-review finding (`01-REVIEW.md` CR-01), explicitly out of scope for this verification per task instructions ("note as context, not a blocker"). |
| `src/taskpane/word.ts` | `1017-1022` | Hallucination-check *aggregate summary* counter (`flaggedCount`/`rateLimitedCount`) does not account for a citation that has both `nameMismatch` set and a non-empty `rateLimitedProviders` — under-counts in that narrow combination | ℹ️ Info | Reviewed (`01-REVIEW.md` WR-01) and confirmed the *per-citation* "Possible hallucination" render path (the actual Core Value guard) is unaffected — only the one-line summary count can be off. Not a regression introduced by vendoring cleanup; a pre-existing edge case in the migrated aggregate logic. Noted as context per task instructions, not a phase-blocking gap. |

No debt markers (TBD/FIXME/XXX) were found in any file modified by this phase's own tasks (this phase produced 0 `src/` changes of its own — all `src/`/`package.json` changes originated from PR #33, which this phase verifies). The one pre-existing `TODO` in `README.md` (USPTO Patent Center placeholder) is untouched by PR #33's diff and tracked separately as `MANIFEST-03` (v2, out of scope).

### Human Verification Required

None. All must-haves for this phase are process/dependency/audit-based and were fully verifiable via git history, `npm`/`gh` commands, and grep evidence — no UI/visual/real-time behavior is in scope for CORE-01/CORE-02.

### Gaps Summary

No gaps. All 12 derived truths (roadmap Success Criteria #1-#3 plus PLAN-frontmatter must-haves) verified against the live codebase, independently re-run rather than trusted from SUMMARY.md/audit-doc claims:

- PR #33 is merged to `main` as merge commit `0f48462`, ancestor of `HEAD`, with CI green on that exact commit.
- `openclerk-core@0.2.6` installs cleanly from the public npm registry (`npm ci`, `npm run build`, `npm test` all independently re-run and pass).
- `src/taskpane/providers/`, `src/taskpane/bluebook/`, and `src/taskpane/utils.ts` are confirmed deleted; `src/taskpane/word.ts` imports the equivalent symbols from `openclerk-core`.
- The Core Value hallucination guard (`nameMismatch` → "Possible hallucination", `verifiedVia`-gated "Verified") is present verbatim, with no orphaned old hand-rolled type/loop.
- `src/commands/` and `scripts/` contain zero duplicated logic (`scripts/generate-bluebook-data.js`, the one file with real overlap, is deleted); zero residual imports of deleted local paths anywhere in `src/`/`scripts/`.
- A persistent, falsifiable audit record (`01-DUPLICATION-AUDIT.md`) and a closed `CONCERNS.md` entry document the CORE-02 conclusion.

Two pre-existing code-review findings (`CR-01` broken `test:live` script, `WR-01` hallucination summary undercount edge case) are noted as context per this verification's task instructions — they do not block phase completion and are not part of the roadmap Success Criteria for this phase.

---

*Verified: 2026-07-15*
*Verifier: Claude (gsd-verifier)*
