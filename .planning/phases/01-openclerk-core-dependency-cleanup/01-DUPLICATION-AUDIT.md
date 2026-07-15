# CORE-02 Duplication Audit

**Date:** 2026-07-15
**Phase:** 01-openclerk-core-dependency-cleanup
**Plan:** 01-02
**Requirement:** CORE-02 -- audit `src/commands/` and `scripts/` for logic duplicated in `openclerk-core`, remove/replace anything found, and record the result.

**Baseline:** PR #33 ("Depend on openclerk-core instead of vendoring its logic") merged into `main` as merge commit `0f48462fe5a59eb4e75126fc9dadf60a4de3a4be` (`Merge pull request #33 from OpenClerkProject/claude/depend-on-openclerk-core`), confirmed an ancestor of the current `HEAD` (`955ee84`, local sync merge of `WordClerk/main`). This audit was run fresh, post-merge, directly against that state -- it is not a restatement of the pre-merge spot-check already recorded in `01-CONTEXT.md` decision D-03 / `01-PATTERNS.md`.

```
$ git log --merges -1 main --format="%H %s"
955ee84be605a3fe25c93149ddaace6972c3bcdf Merge remote-tracking branch 'WordClerk/main'

$ git merge-base --is-ancestor 0f48462 HEAD && echo "0f48462 is ancestor of HEAD"
0f48462 is ancestor of HEAD
```

## Task 1: `src/commands/` -- zero duplication

Grepped `src/commands/` for the identifier set that would indicate citation-parsing, Bluebook-rule, or HTML-escaping logic:

```
$ grep -rn "parseCaseCitation\|ParsedCitation\|BluebookRuleSet\|escapeHtml\|isSafeHyperlinkUrl\|citationProviderRegistry\|bluebookRuleSetRegistry" src/commands/
(no output, exit 1)

$ grep -rc "parseCaseCitation\|ParsedCitation\|BluebookRuleSet\|escapeHtml\|isSafeHyperlinkUrl\|citationProviderRegistry\|bluebookRuleSetRegistry" src/commands/commands.ts src/commands/commands.word.ts src/commands/commands.html
src/commands/commands.ts:0
src/commands/commands.word.ts:0
src/commands/commands.html:0
```

**Result:** 0 matches across all three files (sum = 0). Manual read of `src/commands/commands.word.ts` confirms it still contains only the `insertBlueParagraphInWord` "Hello World" ribbon-command handler and `Office.onReady`/`Office.actions.associate` wiring -- unmodified Office-Add-in-template boilerplate, matching the excerpt already captured in `01-PATTERNS.md` lines 102-133. PR #33's diff does not touch `src/commands/`.

## Task 2: `scripts/` -- zero duplication, zero residual imports

### 2a. `scripts/generate-bluebook-data.js` confirmed deleted

```
$ test -f scripts/generate-bluebook-data.js && echo "EXISTS (unexpected)" || echo "ABSENT (expected)"
ABSENT (expected)
```

This was the one file in `scripts/` with real overlap with `openclerk-core` (Bluebook reference-data generation); PR #33 deleted it (confirmed in its diff stat: `scripts/generate-bluebook-data.js | 172 -`).

### 2b. Remaining `scripts/*.js` files -- zero duplication

```
$ grep -rc "parseCaseCitation\|ParsedCitation\|BluebookRuleSet\|escapeHtml\|isSafeHyperlinkUrl\|citationProviderRegistry\|bluebookRuleSetRegistry" scripts/build-docs.js scripts/convert-logos.js scripts/install-openclerk.js scripts/package-release.js scripts/package-release-offline.js
scripts/build-docs.js:0
scripts/convert-logos.js:0
scripts/install-openclerk.js:0
scripts/package-release.js:0
scripts/package-release-offline.js:0
```

**Result:** 0 matches across all 5 remaining `scripts/*.js` files (sum = 0).

### 2c. No residual imports of deleted local `./providers`, `./bluebook`, or `./utils` paths anywhere in `src/` or `scripts/`

```
$ grep -rn 'from "\./providers"\|from "\./bluebook"\|from "\./utils"\|from "\.\./providers\|from "\.\./bluebook\|from "\.\./utils' src/ scripts/
(no output, exit 1)
```

**Result:** 0 matches. The whole `providers/`/`bluebook/` tree and `utils.ts` were deleted wholesale by PR #33 (per `01-PATTERNS.md` lines 135-157), and no orphaned import of those paths remains anywhere in `src/` or `scripts/`.

### 2d. Discretionary `tests/` deep-import spot-check (01-CONTEXT.md Claude's Discretion)

```
$ grep -rn '\.\./src/taskpane/providers\|\.\./src/taskpane/bluebook\|\.\./src/taskpane/utils' tests/
(no output, exit 1)

$ ls tests/
installer.test.ts
manifest.test.ts

$ for f in bluebook.test.ts providers.test.ts utils.test.ts opinionText.test.ts hyperlinks.test.ts courtListener.live.test.ts; do
    test -f "tests/$f" && echo "FOUND (unexpected): $f" || echo "absent (expected): $f"
  done
absent (expected): bluebook.test.ts
absent (expected): providers.test.ts
absent (expected): utils.test.ts
absent (expected): opinionText.test.ts
absent (expected): hyperlinks.test.ts
absent (expected): courtListener.live.test.ts
```

**Result:** 0 matches for the deep-import pattern; all 6 test files PR #33 removed (`bluebook.test.ts`, `providers.test.ts`, `utils.test.ts`, `opinionText.test.ts`, `hyperlinks.test.ts`, `courtListener.live.test.ts`) are confirmed absent from `tests/` -- none was missed or left with a dangling deep-import. Only `installer.test.ts` and `manifest.test.ts` remain, consistent with `01-01-SUMMARY.md`'s "Reduced test suite size" note (that coverage now lives in the `openclerk-core` sibling repo).

## Conclusion

**CORE-02: confirmed -- no logic in `src/commands/` or `scripts/` duplicates `openclerk-core`.**

- `src/commands/` (Task 1): 0 matches for citation-parsing/Bluebook/escaping identifiers. Unmodified Office-Add-in template boilerplate.
- `scripts/` (Task 2): the one file with real overlap (`generate-bluebook-data.js`) was deleted by PR #33 and confirmed absent; the 5 remaining scripts have 0 matches for the same identifier set.
- No residual imports of the deleted `./providers`, `./bluebook`, or `./utils` local paths remain anywhere in `src/` or `scripts/`.
- The discretionary `tests/` deep-import spot-check confirms all 6 removed test files are genuinely gone with no dangling deep-import left behind.

No duplicated logic was found; nothing required removal or replacement in this plan beyond what PR #33 already deleted.
