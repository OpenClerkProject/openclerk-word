# Phase 1: openclerk-core Dependency Cleanup - Pattern Map

**Mapped:** 2026-07-15
**Files analyzed:** 6 (audit/verify targets — no net-new files this phase)
**Analogs found:** N/A — this is a dependency-cleanup phase; "patterns" below are the current-state
excerpts and the already-drafted PR #33 diff, not analog files to imitate.

This phase does not create new files. It merges an already-open PR (#33) and re-verifies its
result. Below is the grounded current state (on `main`, pre-merge) and the exact diff PR #33
applies, so the planner can write verification tasks against real content instead of assumptions.

## File Classification

| File | Role | Data Flow | Current State (main, pre-merge) | Post-merge State (PR #33) |
|------|------|-----------|----------------------------------|----------------------------|
| `package.json` | config | N/A | `dependencies` has no `openclerk-core`; has `jszip`, `core-js`, `regenerator-runtime`, `@types/jszip` | Adds `"openclerk-core": "^0.2.6"` to `dependencies`; removes `bluebook:update-data` script |
| `src/taskpane/word.ts` | controller (UI orchestrator) | request-response + event-driven | Imports `citationProviderRegistry`, `parseCaseCitation`, etc. from local `./providers`, `./utils`, `./bluebook` | Same names imported from `openclerk-core` instead; also adopts `checkCitationsForHallucinations`/`HallucinationCheckResult` (replaces hand-rolled loop + local `HallucinationResult` type) |
| `src/taskpane/providers/index.ts` (+ whole `providers/` dir) | service registry (self-registering plugin) | request-response | Exists on `main`, vendored locally | Deleted entirely by PR #33 — logic now lives in `openclerk-core` |
| `src/taskpane/bluebook/index.ts` (+ whole `bluebook/` dir) | rule-engine registry (self-registering plugin) | transform | Exists on `main`, vendored locally | Deleted entirely by PR #33 |
| `src/taskpane/utils.ts` | utility | transform | Exists on `main` (`normalizeText`, `escapeHtml`, `isSafeHyperlinkUrl`, etc.) | Deleted entirely by PR #33 — re-exported from `openclerk-core` |
| `src/commands/commands.ts`, `commands.word.ts`, `commands.html` | route/controller (ribbon function-file) | request-response | Unmodified Yeoman/Office-Add-in-template "Hello World" boilerplate; zero citation/Bluebook logic | Untouched by PR #33 (confirmed no overlap) |
| `scripts/generate-bluebook-data.js` | utility (build script) | batch | Exists on `main`, 6,800 bytes, generates `bluebook/generated/*.generated.ts` from external sources | Deleted by PR #33 (equivalent generator now lives in `openclerk-core`) |
| `tests/bluebook.test.ts`, `tests/providers.test.ts`, `tests/utils.test.ts`, `tests/opinionText.test.ts`, `tests/hyperlinks.test.ts`, `tests/courtListener.live.test.ts` | test | N/A | Deep-import internals (`../src/taskpane/providers/...`, `../src/taskpane/bluebook/...`, `../src/taskpane/utils`) | All 6 removed by PR #33 (equivalent coverage now lives in `openclerk-core`'s own test suite) |

## Current State Excerpts (main, pre-merge)

### `package.json` dependencies block (lines 38-43)
```json
"dependencies": {
  "@types/jszip": "^3.4.0",
  "core-js": "^3.36.0",
  "jszip": "^3.10.1",
  "regenerator-runtime": "^0.14.1"
},
```
No `openclerk-core` entry yet on `main` — this phase's context doc's claim that it's "already
bumped to `^0.2.6` on PR #33's branch" refers to the PR branch
(`remotes/WordClerk/claude/depend-on-openclerk-core`), not `main`. Confirmed via
`git diff main remotes/WordClerk/claude/depend-on-openclerk-core -- package.json`.

### `src/taskpane/word.ts` current imports (lines 10-29, main)
```typescript
import JSZip from "jszip";
import {
  ... // normalizeText, escapeHtml, isSafeHyperlinkUrl, extractParentheticalCitations, etc.
} from "./utils";
import {
  citationProviderRegistry,
  parseCaseCitation,
  extractCaseCitations,
  expandPincitePages,
  supportsOpinionText,
  supportsRateLimitAwareness,
  CitationProvider,
  OpinionTextCapableProvider,
  ParsedCitation,
} from "./providers";
import { bluebookRuleSetRegistry, BluebookRuleSet, BluebookIssue } from "./bluebook";
```

### `src/taskpane/providers/index.ts` (whole file, main — self-registering plugin registry pattern)
```typescript
import { citationProviderRegistry } from "./registry";
import { CourtListenerProvider } from "./courtListenerProvider";
import { LexisNexisProvider } from "./lexisNexisProvider";
import { WestlawProvider } from "./westlawProvider";
import { BloombergLawProvider } from "./bloombergLawProvider";
import { UsptoPatentCenterProvider } from "./usptoPatentCenterProvider";

citationProviderRegistry.register(new CourtListenerProvider());
citationProviderRegistry.register(new LexisNexisProvider());
citationProviderRegistry.register(new WestlawProvider());
citationProviderRegistry.register(new BloombergLawProvider());
citationProviderRegistry.register(new UsptoPatentCenterProvider());

export { citationProviderRegistry } from "./registry";
export * from "./types";
export { parseCaseCitation, extractCaseCitations } from "./citationParser";
export { expandPincitePages } from "./pincitePages";
```

### `src/taskpane/bluebook/index.ts` (whole file, main — same registry pattern)
```typescript
import { bluebookRuleSetRegistry } from "./registry";
import { Bluebook20thEdition } from "./edition20th";
import { Bluebook21stEdition } from "./edition21st";
import { Bluebook22ndEdition } from "./edition22nd";

bluebookRuleSetRegistry.register(new Bluebook22ndEdition());
bluebookRuleSetRegistry.register(new Bluebook21stEdition());
bluebookRuleSetRegistry.register(new Bluebook20thEdition());

export { bluebookRuleSetRegistry } from "./registry";
export * from "./types";
```
**Post-merge implication:** after PR #33 merges, both `index.ts` files (and their entire
directories) are deleted, so the "self-registration by import side-effect" pattern moves inside
`openclerk-core`'s own package — verify at audit time that `openclerk-core`'s public API
(imported in `word.ts`) still exposes `citationProviderRegistry`/`bluebookRuleSetRegistry` with all
5 providers / 3 editions registered, since `word.ts` no longer does the registration itself.

### `src/commands/commands.word.ts` (whole file, main — confirmed zero overlap)
```typescript
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office Word console */

export async function insertBlueParagraphInWord(event: Office.AddinCommands.Event) {
  try {
    await Word.run(async (context) => {
      const paragraph = context.document.body.insertParagraph(
        "Hello World",
        Word.InsertLocation.end
      );
      paragraph.font.color = "blue";
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
  event.completed();
}

Office.onReady(async () => {
  Office.actions.associate("action", insertBlueParagraphInWord);
});
```
Unmodified Office-Add-in-template boilerplate — no citation parsing, no Bluebook logic, no HTML
escaping. `commands.ts` is a one-line `import "./commands.word";`. Confirms D-03's spot-check:
nothing here duplicates `openclerk-core`.

### Remaining local cross-imports on `main` that PR #33's deletions must fully account for
```
src/taskpane/bluebook/checkCaseNameAbbreviations.ts:1:import { ParsedCitation } from "../providers/types";
src/taskpane/bluebook/commonRules.ts:1:import { ParsedCitation } from "../providers/types";
src/taskpane/bluebook/courtRules.ts:1:import { ParsedCitation } from "../providers/types";
src/taskpane/bluebook/edition20th.ts:1:import { ParsedCitation } from "../providers/types";
src/taskpane/bluebook/edition21st.ts:1:import { ParsedCitation } from "../providers/types";
src/taskpane/bluebook/edition22nd.ts:1:import { ParsedCitation } from "../providers/types";
src/taskpane/bluebook/pageRangeRules.ts:1:import { ParsedCitation } from "../providers/types";
src/taskpane/bluebook/pageRangeRules.ts:2:import { reconstructFullPageNumber } from "../providers/pincitePages";
src/taskpane/bluebook/reporterRules.ts:1:import { ParsedCitation } from "../providers/types";
src/taskpane/bluebook/types.ts:1:import { ParsedCitation } from "../providers/types";
src/taskpane/providers/citationParser.ts:1:import { normalizeText } from "../utils";
src/taskpane/word.ts:17: (from "./utils")
src/taskpane/word.ts:28: (from "./providers")
src/taskpane/word.ts:29: (from "./bluebook")
```
All of these are inside `bluebook/`, `providers/`, and `utils.ts` — the same three directories/file
PR #33 deletes wholesale. Since the whole tree is removed together, no orphaned import is left
pointing at a partially-deleted path. **Audit task:** after merge, grep the *entire* `src/` and
`scripts/` trees for any remaining `from "./providers"`, `from "./bluebook"`, `from "./utils"`,
`from "../providers`, `from "../bluebook`, or `from "../utils` pattern — expected result is zero
matches outside files also being deleted.

## PR #33 Diff Excerpts (`main` → `remotes/WordClerk/claude/depend-on-openclerk-core`)

### `package.json` diff
```diff
-    "bluebook:update-data": "node scripts/generate-bluebook-data.js",
     "test": "jest",
@@
     "jszip": "^3.10.1",
+    "openclerk-core": "^0.2.6",
     "regenerator-runtime": "^0.14.1"
```

### `src/taskpane/word.ts` diff (import block, lines 14-29)
```diff
-} from "./utils";
-import {
   citationProviderRegistry,
   parseCaseCitation,
   extractCaseCitations,
   expandPincitePages,
   supportsOpinionText,
   supportsRateLimitAwareness,
+  checkCitationsForHallucinations,
   CitationProvider,
   OpinionTextCapableProvider,
   ParsedCitation,
-} from "./providers";
-import { bluebookRuleSetRegistry, BluebookRuleSet, BluebookIssue } from "./bluebook";
+  HallucinationCheckResult,
+  bluebookRuleSetRegistry,
+  BluebookRuleSet,
+  BluebookIssue,
+} from "openclerk-core";
```
Single combined import from `openclerk-core` replaces three separate local-path imports
(`./utils`, `./providers`, `./bluebook`).

### `src/taskpane/word.ts` diff (hallucination-check core logic — the real bug-fix commit)
```diff
-      const results: HallucinationResult[] = [];
-      for (const raw of candidates) {
-        const parsed = parseCaseCitation(raw) || { raw };
-        let verifiedVia: string | null = null;
-        const skippedProviders: string[] = [];
-        const rateLimitedProviders: string[] = [];
-        for (const provider of selectedProviders) {
-          if (provider.requiresAuth && !provider.isAuthenticated()) {
-            skippedProviders.push(provider.name);
-            continue;
-          }
-          const match = await provider.lookupCitation(parsed);
-          if (match) {
-            verifiedVia = provider.name;
-            break;
-          }
-          if (supportsRateLimitAwareness(provider) && provider.wasLastRequestRateLimited()) {
-            rateLimitedProviders.push(provider.name);
-          }
-        }
-        results.push({ raw, verifiedVia, skippedProviders, rateLimitedProviders });
-      }
+      const results: HallucinationCheckResult[] = await checkCitationsForHallucinations(candidates, selectedProviders);
```
Replaces a hand-rolled "first provider that resolves the locator wins" loop with
`openclerk-core`'s shared `checkCitationsForHallucinations`, which additionally verifies the
resolved case *name* matches (`caseNamesMatch`) — closing the "real citation locator, fabricated
case name" gap the hand-rolled loop couldn't detect. This is the "real bug fix" referenced in D-01.

### `src/taskpane/word.ts` diff (rendering — new `nameMismatch` signal)
```diff
+    } else if (result.nameMismatch) {
+      status.classList.add("issue-flagged");
+      status.textContent = `Possible hallucination -- ${result.nameMismatch.provider} resolves this citation to a different case: "${result.nameMismatch.foundCaseName}".`;
     } else if (result.rateLimitedProviders.length > 0) {
```
Core-value-relevant: this is exactly the kind of guard-strengthening (never silently treat a
locator match as "verified" when the case name doesn't match) called out in the project's
Constraints — confirm this branch's rendering logic is preserved verbatim on merge, don't let it
regress to the old "verifiedVia present ⇒ verified" logic.

## Shared Patterns

### Self-registering plugin registry (pattern being relocated, not removed)
**Source (main, pre-merge):** `src/taskpane/providers/index.ts`, `src/taskpane/bluebook/index.ts`
**Post-merge:** the same registration side-effect pattern must still hold, just inside
`openclerk-core`'s package instead of this repo. Nothing in `word.ts` calls `.register(...)`
itself — it only consumes `citationProviderRegistry`/`bluebookRuleSetRegistry` already populated.
**Apply to:** verification task — after merge + `npm install`, add a smoke check (or rely on
existing test coverage in `openclerk-core`) confirming `citationProviderRegistry` has all 5
providers and `bluebookRuleSetRegistry` has all 3 editions registered when imported from
`openclerk-core`.

### "Move on" error handling (unchanged, now inside openclerk-core)
**Source:** `src/taskpane/providers/base.ts` (pre-merge) — network/lookup failures converted to
`null` rather than thrown, per project CLAUDE.md's documented Error Handling conventions.
**Apply to:** no new code needed this phase; just confirm (by reading `openclerk-core`'s
published source/types, not this repo) that this convention wasn't weakened when the logic moved
package.

## No Analog Found

Not applicable — every file in scope for this phase is being audited/verified against its own
prior state and the already-drafted PR #33 diff, not built fresh against an unrelated analog.

## Metadata

**Analog search scope:** `src/taskpane/word.ts`, `src/taskpane/providers/`, `src/taskpane/bluebook/`,
`src/taskpane/utils.ts`, `src/commands/`, `scripts/`, `package.json`, `tests/` (all read from
`main` at HEAD `13b2e78`, and diffed against `remotes/WordClerk/claude/depend-on-openclerk-core`
at HEAD `58b2e13`, which is PR #33's branch).
**Files scanned:** 20+ (full `providers/`, `bluebook/` directory listings; grep across `src/` for
remaining local cross-imports).
**Pattern extraction date:** 2026-07-15
