# Phase 2: Escaping Hardening - Research

**Researched:** 2026-07-15
**Domain:** Cross-repo TypeScript hardening (branded/nominal types + compiler-enforced wrapper module + ESLint bypass gate) for an Office.js Word task-pane add-in
**Confidence:** HIGH

<user_constraints>
## User Constraints (from CONTEXT.md)

### Locked Decisions

**Branded-type home (cross-repo)**
- **D-01:** The branded `SafeHtml`/`SafeHyperlinkUrl` types are added to **`openclerk-core`**, not
  a local `utils.ts` in `openclerk-word`. `utils.ts` no longer exists in this repo — Phase 1
  moved `escapeHtml`/`isSafeHyperlinkUrl` into `openclerk-core`'s `src/utils.ts`
  (`C:\Users\willi\openclerk-core`, sibling checkout, currently at package version `0.2.7`).
  `escapeHtml`/`isSafeHyperlinkUrl` become (or are wrapped by) smart constructors that return the
  branded types, so the type system — not convention — enforces that only validated/escaped
  values reach an Office.js insertion call. Rationale: `openclerk-core` is the shared logic home
  across Word/GDocs/LibreOffice per its own README; the branded types belong at the source of
  truth, not duplicated per consumer.
- **D-02 (sequencing):** This is **one phase spanning both repos**, not a separate prerequisite
  effort in `openclerk-core`'s own planning. The plan should include an early wave that: (1) adds
  the branded types + smart constructors to `openclerk-core`, (2) bumps its version, (3) stops at
  a **human checkpoint** before the actual `npm publish` / release-tag push (same pattern as PR
  #33's merge checkpoint in Phase 1 — publishing is an externally-visible, hard-to-reverse
  action), then (4) once published, bumps the `openclerk-core` dependency in `openclerk-word`'s
  `package.json` and proceeds to build `safeInsertion.ts` against the new branded exports.
- **D-03:** Publish authority is a **human checkpoint** — the executor prepares the version bump
  and any changelog/PR in `openclerk-core`, but a person triggers the actual publish
  (`npm publish` or the release tag that fires `openclerk-core`'s existing `publish.yml`).

**Failure behavior at the validation boundary**
- **D-04:** When a URL or string fails validation (e.g. `isSafeHyperlinkUrl`-equivalent check
  returns false), the smart constructor **returns `null`** rather than throwing. The caller skips
  that citation, matching the codebase's existing "move on, don't crash" pattern for
  best-effort/network-adjacent operations (documented in project conventions and already how
  `word.ts:250`/`346`/`556`/`1442` filter unsafe URLs today). This phase makes that filtering
  compiler-enforced (you cannot call the insertion wrapper without a branded value) rather than
  changing the runtime behavior users see.

**CI enforcement**
- **D-05:** Add an `npm run lint` (ESLint) step to `.github/workflows/ci.yml`. Confirmed by
  reading the workflow directly: CI currently runs build, `npm test` (jest), and manifest
  `xmllint` validation, but **never runs ESLint**. Without this, the `no-restricted-syntax`
  bypass-guard (ESCAPE-03) only protects a developer who happens to lint locally — a bad PR would
  still go green. This is now in scope for Phase 2, not deferred.

**Test coverage**
- **D-06:** `safeInsertion.ts` gets **dedicated unit tests** in `openclerk-word`'s `tests/`
  directory. Confirmed `word.ts` has zero automated test coverage today (only
  `installer.test.ts` and `manifest.test.ts` remain in `tests/` — citation/provider/utils tests
  moved to `openclerk-core` in Phase 1). Scope is limited to the new `safeInsertion.ts` module
  (branded-type constructors + the insertion dispatch/fallback logic) — it does **not** extend to
  testing the rest of the untested `word.ts` monolith, which stays deferred to v2 (WORDTS-01/02).

### Claude's Discretion
- Exact branding mechanism in TypeScript (e.g., `string & { readonly __brand: unique symbol }`)
  and smart-constructor function naming/signatures in `openclerk-core`.
- Whether `safeInsertion.ts`'s wrapper functions take a `Word.RequestContext` and perform
  `context.sync()` internally (as today's `applyHyperlinkToItem` does) or leave sync to the
  caller (as today's `insertComment` call site does) — reconcile the two existing patterns into
  one consistent wrapper API.
- Whether the existing 3-tier fallback in `applyHyperlinkToItem` (`insertHyperlink` →
  `insertHtml` with escaped anchor → plain `insertText`) is preserved as-is inside
  `safeInsertion.ts` — default to preserving existing behavior unless research surfaces a reason
  not to.
- `src/commands/commands.word.ts:15`'s `insertParagraph(atext, ...)` call (Yeoman-boilerplate
  "Hello World" paragraph, hardcoded string, no dynamic/user-derived content) is **out of scope**
  for the ESLint bypass-guard — ESCAPE-02/03 name only `insertHtml`/`insertHyperlink`/
  `insertComment`, and this call carries no injection risk (fixed literal, no interpolation).
- Exact ESLint `no-restricted-syntax` selector/override mechanism (e.g. `overrides` glob
  excluding `safeInsertion.ts` from the restricted-syntax rule) is an implementation detail for
  the planner/researcher. **Resolved by this research — see Common Pitfalls and Code Examples
  below: a project-root flat `eslint.config.mjs` is required, not `.eslintrc.json`.**
- Version-number choice for the new `openclerk-core` release (patch vs. minor) given it adds new
  exported types/functions — likely minor per semver, but left to the executor at bump time.
  **Research finding: the current `0.2.7` is already the published/tagged registry version (see
  Common Pitfalls), so any valid semver bump above `0.2.7` works; `0.3.0` is the semver-correct
  choice for new public exports, but this remains the executor's call.**

### Deferred Ideas (OUT OF SCOPE)
None — discussion stayed within phase scope. `word.ts`'s broader monolith split (WORDTS-01/02)
was already deferred to v2 before this discussion began (per `REQUIREMENTS.md`) and was not
reopened here.
</user_constraints>

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|------------------|
| ESCAPE-01 | Branded `SafeHtml`/`SafeHyperlinkUrl` types are added (per D-01, to `openclerk-core`'s `src/utils.ts`, not a local `utils.ts`) and used to type-constrain Office.js insertion helpers | Branded-type pattern + smart-constructor design documented in Code Examples; exact `openclerk-core` file/export layout documented in Architecture Patterns and Code Context below |
| ESCAPE-02 | A single wrapper module (`safeInsertion.ts`) owns all raw Office.js `insertHtml`/`insertHyperlink`/`insertComment` calls | All 3 current call sites identified with exact line numbers (word.ts:223-227, 1232) plus the 4 `isSafeHyperlinkUrl` pre-filter call sites that become smart-constructor call sites (word.ts:250, 342, 552, 1426); wrapper API design options documented in Architecture Patterns |
| ESCAPE-03 | An ESLint `no-restricted-syntax` rule bans direct calls to raw Office.js insertion APIs outside the wrapper module, failing the build on a bypass | Critical finding: `npm run lint` currently ignores `.eslintrc.json` entirely (see Common Pitfalls) — a new project-root `eslint.config.mjs` (flat config) is required for the rule to take effect; exact selector syntax and per-file exemption pattern verified against ESLint's own docs and a closed "works as intended" GitHub issue confirming the exemption mechanism is safe to use |
</phase_requirements>

## Summary

This phase is a compiler-enforced hardening pass on an already-correct runtime guard: `openclerk-core`'s `escapeHtml`/`isSafeHyperlinkUrl` (published, tested, allowlist-based) are working today, but nothing stops a future call site in `word.ts` from calling Office.js's `insertHtml`/`insertHyperlink`/`insertComment` directly without them. The fix is two-layered — (1) branded/nominal TypeScript types (`SafeHtml`, `SafeHyperlinkUrl`) produced only by smart constructors that wrap the existing validators and return `null` on failure, so the type system rejects any unvalidated string at the insertion boundary; (2) a single `safeInsertion.ts` wrapper module that is the only place in `openclerk-word` allowed to call the raw Office.js insertion APIs, enforced by an ESLint `no-restricted-syntax` AST-selector rule.

The work spans two repositories with a hard sequencing dependency: `openclerk-core` must publish a new version to the public npm registry (registry-verified: `openclerk-core@0.2.7` is *already* published — this is not a pending-but-unpublished local bump, so the branded-type change requires a new version number above `0.2.7`) before `openclerk-word` can depend on the branded exports. `openclerk-core`'s `publish.yml` fires only on a `v*` git tag push and requires a human to trigger it (D-03) — the plan must include a `checkpoint:human-action`/`gate="blocking-human"` task at that exact point, matching the pattern already used for Phase 1's PR #33 merge checkpoint.

The single most important and non-obvious research finding is that **`npm run lint` in this repo does not currently read `.eslintrc.json` at all.** `office-addin-lint check` (what `npm run lint` invokes) passes an explicit `-c <path>` flag to ESLint pointing at either a project-root `eslint.config.mjs` (flat config) if one exists, or its own bundled flat config if not — `.eslintrc.json`-format legacy configs are never consulted by this toolchain, regardless of ESLint version. This means ESCAPE-03's `no-restricted-syntax` rule **must** be added to a new project-root `eslint.config.mjs`, not `.eslintrc.json`, or `npm run lint` (and therefore the new CI lint step from D-05) will silently never enforce it.

**Primary recommendation:** Add a project-root `eslint.config.mjs` that spreads `eslint-plugin-office-addins`'s recommended flat config, adds a `no-restricted-syntax` rule with `CallExpression[callee.property.name='insertHtml'|'insertHyperlink'|'insertComment']` selectors as a global rule, and a second config object scoped to `files: ["**/safeInsertion.ts"]` that sets `no-restricted-syntax: "off"` — verified as the correct, ESLint-documented exemption mechanism (not merged/appended, but a full per-file override, which is exactly what's needed here).

## Architectural Responsibility Map

| Capability | Primary Tier | Secondary Tier | Rationale |
|------------|-------------|----------------|-----------|
| HTML escaping / URL-scheme validation logic | Shared library (`openclerk-core`) | — | Already the cross-platform (Word/GDocs/LibreOffice) shared logic home per its own README; Phase 1 already migrated the plain functions here (D-01 confirms, not reopens, this placement) |
| Branded type definitions + smart constructors | Shared library (`openclerk-core`) | — | D-01 locked: types must live at the validation source of truth so every consumer (not just Word) gets the compile-time guarantee |
| Office.js insertion dispatch (feature-detect `insertHyperlink`/`insertHtml`/`insertText` fallback, `insertComment`) | Browser/Client (task-pane WebView) | — | Office.js `Word.Range`/`Word.Comment` APIs only exist inside the Word WebView host; this is inherently client-tier code, cannot move to a shared/server tier |
| Compile-time bypass prevention (ESLint rule) | Build tooling (client repo, dev-time only) | — | ESLint runs at build/CI time against `openclerk-word`'s TypeScript source; has no runtime component and is scoped to this repo only (branded types make the *type system* the runtime-adjacent guard; ESLint is the static-analysis backstop for anyone who bypasses the type system with `any`) |
| CI enforcement (new lint job) | Build tooling (GitHub Actions, `openclerk-word`) | — | `.github/workflows/ci.yml` already orchestrates build/test/manifest-validate jobs; the new lint job is an additive peer job |
| Release/publish authority | Human (out of tooling entirely) | — | D-03 locks this as a human checkpoint; no tier owns "publish to npm" as an automated capability in this phase |

## Standard Stack

No new runtime dependencies are introduced by this phase. Both repos already have every tool needed installed.

### Core
| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| TypeScript | `^5.4.2` (already a devDependency in both repos) [VERIFIED: package.json] | Branded/nominal type definitions via intersection types | Already the project's language; no new tooling needed — branded types are a pure type-level pattern, not a library |
| ESLint | `9.39.4` installed (transitively via `office-addin-lint`) [VERIFIED: node_modules/eslint/package.json] | `no-restricted-syntax` AST-selector rule enforcement | Core ESLint rule, no plugin required; already the project's linter |

### Supporting
| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| `office-addin-lint` | `3.0.9` (already installed) [VERIFIED: node_modules/office-addin-lint/package.json] | Wraps ESLint invocation for `npm run lint`/`npm run lint:fix` | Already the project's lint runner — see Common Pitfalls for how it resolves config |
| `eslint-plugin-office-addins` | `^4.0.3` declared / `4.0.9` resolved transitively by `office-addin-lint` [VERIFIED: package.json, office-addin-lint's own dependency] | Office-addin-specific recommended rule set (`call-sync-after-load`, etc.) | Already extended by `.eslintrc.json`; must also be spread into the new flat `eslint.config.mjs` so existing office-addin rules keep firing, not just the new `no-restricted-syntax` rule |

### Alternatives Considered
| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| TypeScript branded types (`string & { readonly __brand: unique symbol }`) | A wrapper class (`class SafeHtml { constructor(private value: string) {} }`) | Class wrapper adds a runtime allocation per call and requires `.value`/`.toString()` unwrapping at every call site; branded primitives are zero-runtime-cost and unwrap for free since they *are* the primitive at runtime — branded types are the better fit for an `es5` target where every extra allocation/prototype matters |
| ESLint `no-restricted-syntax` (AST selector) | A custom `ts-morph`/AST-walking script run as a separate CI step | `no-restricted-syntax` is already a core ESLint rule with zero new dependencies and integrates with the existing `npm run lint` step (D-05); a custom script would be a second, disconnected enforcement path with its own maintenance burden — clear over-engineering for a 3-selector ban |
| Flat `eslint.config.mjs` | Keep `.eslintrc.json` and force `office-addin-lint` to use it | Not viable: `office-addin-lint`'s `getEsLintBaseCommand` (`node_modules/office-addin-lint/lib/lint.js:32-38`) hardcodes a check for `eslint.config.mjs` specifically and passes ESLint an explicit `-c` flag pointing at either that file or its own bundled flat config — there is no code path in the installed version that reads `.eslintrc.json` |

**Installation:**
```bash
# No new packages required. Only file additions:
#   openclerk-word:  eslint.config.mjs  (new, project root)
#   openclerk-core:  branded types + smart constructors added to existing src/utils.ts
```

**Version verification:** `openclerk-core` on the npm registry is confirmed at `0.2.7` — `npm view openclerk-core version` returns `0.2.7`, and `npm view openclerk-core versions --json` shows `["0.2.5","0.2.6","0.2.7"]` [VERIFIED: npm registry]. The sibling checkout's git tags (`v0.2.0` … `v0.2.7`) match exactly, with a clean working tree at `v0.2.7`'s commit [VERIFIED: `git tag -l`, `git status --short` in `C:\Users\willi\openclerk-core`]. This means `0.2.7` is **not** an unpublished local bump waiting on a phase-1 leftover — it is already live on the registry. The branded-type work in this phase must target a *new* version number (e.g. `0.3.0`), not simply "publish what's already at 0.2.7."

## Package Legitimacy Audit

**Not applicable — this phase installs no new external packages.** It adds a flat ESLint config file and TypeScript source (branded types + one wrapper module) using tooling already present in both repos' `node_modules`/`package.json`. The only `package.json` change in `openclerk-word` is a version-range bump of the already-installed `openclerk-core` dependency (from `^0.2.6` to whatever the new `openclerk-core` release version is), and the only `package.json` change in `openclerk-core` is its own `version` field bump — neither is a new package being introduced.

## Architecture Patterns

### System Architecture Diagram

```
┌─────────────────────────── openclerk-core (sibling repo, published to npm) ───────────────────────────┐
│                                                                                                          │
│  src/utils.ts                                                                                            │
│    escapeHtml(str) ───────────┐                                                                          │
│    isSafeHyperlinkUrl(url) ───┤──> smart constructors (NEW, this phase)                                  │
│                                │       toSafeHtml(str): SafeHtml | null                                  │
│                                │       toSafeHyperlinkUrl(url): SafeHyperlinkUrl | null                  │
│                                └──> branded types (NEW): SafeHtml, SafeHyperlinkUrl                       │
│                                                                                                            │
│  src/index.ts  ──exports──>  SafeHtml, SafeHyperlinkUrl, toSafeHtml, toSafeHyperlinkUrl (+ existing)      │
│                                                                                                            │
│  package.json version bump  ──git tag v0.x.y──>  [HUMAN CHECKPOINT: npm publish / push tag] (D-03)       │
│                                                                                                            │
└───────────────────────────────────────────┬────────────────────────────────────────────────────────────┘
                                              │ published npm package
                                              ▼
┌─────────────────────────── openclerk-word (this repo) ─────────────────────────────────────────────────┐
│                                                                                                            │
│  package.json  dependencies.openclerk-core  bumped to new version  ──npm install──>  node_modules/       │
│                                                                                                            │
│  src/taskpane/safeInsertion.ts  (NEW — the ONLY file allowed to call raw Office.js insertion APIs)        │
│    import { SafeHtml, SafeHyperlinkUrl, toSafeHtml, toSafeHyperlinkUrl } from "openclerk-core"            │
│    insertSafeHyperlink(context, item, url: SafeHyperlinkUrl, text: SafeHtml): Promise<void>               │
│      -- feature-detects insertHyperlink -> insertHtml -> insertText fallback (preserves current 3-tier)   │
│    insertSafeComment(context, range, text: SafeHtml): Promise<void>                                       │
│      -- wraps Word.Range.insertComment                                                                    │
│                                                                                                            │
│  src/taskpane/word.ts  (call sites migrate INTO safeInsertion.ts, not out of it)                          │
│    applyHyperlinkToItem (word.ts:217-233)      ──replaced by──>  safeInsertion.insertSafeHyperlink(...)   │
│    embedCitedText's insertComment (word.ts:1232) ──replaced by──> safeInsertion.insertSafeComment(...)    │
│    isSafeHyperlinkUrl pre-filters (word.ts:250,342,552,1426) ──replaced by──> toSafeHyperlinkUrl(...)      │
│      calls that now ALSO produce the branded value passed into the wrapper (single validation, reused)    │
│                                                                                                            │
│  eslint.config.mjs  (NEW, project root — see Common Pitfalls: .eslintrc.json is NOT read by npm run lint) │
│    global rule: no-restricted-syntax bans CallExpression[callee.property.name='insertHtml'|               │
│                 'insertHyperlink'|'insertComment'] everywhere                                              │
│    override:    files: ["**/safeInsertion.ts"]  ->  no-restricted-syntax: "off"                           │
│                                                                                                            │
│  .github/workflows/ci.yml  (NEW lint job, D-05)  ──runs `npm run lint`──>  fails build on any bypass      │
│                                                                                                            │
│  tests/safeInsertion.test.ts  (NEW, D-06)  ──unit tests──>  smart-constructor null-on-invalid behavior +   │
│                                                              insertion dispatch/fallback logic (mocked      │
│                                                              Word.Range-shaped objects, no real Office.js) │
│                                                                                                            │
└────────────────────────────────────────────────────────────────────────────────────────────────────────┘
```

### Recommended Project Structure
```
openclerk-core/
└── src/
    └── utils.ts          # existing escapeHtml/isSafeHyperlinkUrl + NEW branded types & smart constructors

openclerk-word/
├── eslint.config.mjs     # NEW — flat config, required by office-addin-lint for no-restricted-syntax to fire
├── src/taskpane/
│   ├── safeInsertion.ts  # NEW — sole owner of raw insertHtml/insertHyperlink/insertComment calls
│   └── word.ts           # call sites updated to import from safeInsertion.ts + openclerk-core
└── tests/
    └── safeInsertion.test.ts  # NEW
```

### Pattern 1: Branded (nominal) type with smart constructor returning `null`
**What:** A branded type is a structural type (e.g. `string`) intersected with a unique, uninhabited "brand" property that exists only at compile time. Values of the branded type can only be produced by a "smart constructor" function that performs runtime validation, so a value typed `SafeHyperlinkUrl` is a compile-time guarantee that it already passed `isSafeHyperlinkUrl`.
**When to use:** Any boundary where a plain primitive type (`string`) is structurally indistinguishable from a "validated" version of itself, and you want the compiler — not code review — to catch a missing validation call.
**Example:**
```typescript
// openclerk-core/src/utils.ts
// Source pattern: community-standard TypeScript branded-type + smart-constructor idiom
// [ASSUMED — general pattern, not from openclerk-core's own official docs; cross-checked against
//  multiple TS community references, see Sources]

export type SafeHyperlinkUrl = string & { readonly __brand: "SafeHyperlinkUrl" };
export type SafeHtml = string & { readonly __brand: "SafeHtml" };

// Smart constructor: only way to produce a SafeHyperlinkUrl. Returns null on invalid input
// (D-04: never throws) so callers can "move on" the same way isSafeHyperlinkUrl callers do today.
export function toSafeHyperlinkUrl(url: string): SafeHyperlinkUrl | null {
  return isSafeHyperlinkUrl(url) ? (url as SafeHyperlinkUrl) : null;
}

// Smart constructor for pre-escaped HTML content (wraps the existing escapeHtml, which cannot
// itself fail validation -- escaping is total -- so this always returns a SafeHtml, never null).
export function toSafeHtml(raw: string): SafeHtml {
  return escapeHtml(raw) as SafeHtml;
}
```
**Why this compiles under this project's `tsconfig.json` (no `strict`, no `strictNullChecks`):** both repos' `tsconfig.json` omit `"strict": true` [VERIFIED: `C:\Users\willi\WordClerk\tsconfig.json`, `C:\Users\willi\openclerk-core\tsconfig.json` — both read directly, neither sets `strict`]. Without `strictNullChecks`, `SafeHyperlinkUrl | null` is still a valid, checkable union return type — TypeScript still flags `insertHyperlink(url)` as an error if `url: SafeHyperlinkUrl | string` isn't narrowed to exactly `SafeHyperlinkUrl` first, because the brand intersection makes `string` structurally incompatible with `SafeHyperlinkUrl` regardless of strictness mode. The `| null` return only loses *implicit-any-null-assignment* checking (a non-issue here since callers must branch on the value anyway to "move on" per D-04), not the branding guarantee itself.

### Pattern 2: Single wrapper module as the only raw-API caller
**What:** All raw Office.js `insertHtml`/`insertHyperlink`/`insertComment` calls live in one file (`safeInsertion.ts`); every other file imports typed wrapper functions from it instead of calling Office.js insertion APIs directly.
**When to use:** Whenever a static-analysis rule (ESLint) needs a single, greppable enforcement boundary — `no-restricted-syntax` bans the raw call pattern everywhere *except* the one file where it's the wrapper's own implementation.
**Example:**
```typescript
// src/taskpane/safeInsertion.ts
import { SafeHtml, SafeHyperlinkUrl } from "openclerk-core";

/* eslint-disable-next-line -- not needed if eslint.config.mjs scopes the override to this whole
   file via `files: ["**/safeInsertion.ts"]`; shown here only if a narrower per-line exemption is
   preferred instead -- see Common Pitfalls for why file-level scoping is the recommended approach */

export async function insertSafeHyperlink(
  context: Word.RequestContext,
  item: Word.Range,
  url: SafeHyperlinkUrl,
  displayText: SafeHtml
): Promise<void> {
  // Preserves the existing 3-tier fallback (Claude's Discretion item: preserve unless research
  // says otherwise -- no reason found to change it; insertHyperlink is a newer/preview Word API
  // not yet present in @types/office-js 1.0.598, hence the `as any` bridge below is unavoidable
  // and pre-existing, not new risk introduced by this refactor).
  if (typeof (item as any).insertHyperlink === "function") {
    (item as any).insertHyperlink(url, displayText, Word.InsertLocation.replace);
  } else if (typeof (item as any).insertHtml === "function") {
    const html = `<a href="${url}">${displayText}</a>`; // both already branded/escaped -- no
                                                          // further escapeHtml call needed here
    (item as any).insertHtml(html, Word.InsertLocation.replace);
  } else {
    item.insertText(displayText, Word.InsertLocation.replace);
  }
  await context.sync();
}

export function insertSafeComment(range: Word.Range, text: SafeHtml): Word.Comment {
  // Office.js's insertComment takes plain text (Word.Comment.insertComment(commentText: string)),
  // not HTML -- confirmed via @types/office-js:109028. No HTML-escaping risk exists here today,
  // but the call still moves into this file per ESCAPE-02's "all raw calls live in
  // safeInsertion.ts" requirement, regardless of whether escaping applies to this particular API.
  return range.insertComment(text);
}
```

### Anti-Patterns to Avoid
- **Escaping/validating at the call site instead of the type boundary:** the current pre-Phase-2 code (`escapeHtml(url)` inlined at `word.ts:226`) is exactly the pattern being eliminated — it relies on every future call site remembering to do it. Don't replicate that pattern inside `safeInsertion.ts` either; the smart constructors should be the *only* place validation logic runs, and `safeInsertion.ts`'s functions should only ever accept already-branded parameters.
- **Widening the branded type back to `string` inside `safeInsertion.ts`:** e.g. `function insertSafeHyperlink(url: string, ...)` would silently defeat ESCAPE-01's "type-constrain the input parameter" requirement — the parameter types on every exported function in `safeInsertion.ts` must be the branded types, not `string`.
- **Adding a second `no-restricted-syntax` config object that also declares selectors for the same file range as another entry:** confirmed via a closed ESLint GitHub issue (#19239, "works as intended") that flat config does NOT merge `no-restricted-syntax` selector arrays across multiple matching config objects — the *last* matching config's `no-restricted-syntax` value fully replaces earlier ones for that rule. This is desired for the `safeInsertion.ts` exemption (full "off" override) but would be a bug if two *different* restriction lists were meant to add together — keep all three selectors (`insertHtml`/`insertHyperlink`/`insertComment`) in one single rule-array entry, not split across multiple config objects.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Detecting "raw Office.js insertion call outside safeInsertion.ts" | A custom `ts-morph`/regex-based CI script | ESLint core `no-restricted-syntax` rule with AST selectors | Already a zero-dependency built-in ESLint rule; integrates with the existing `npm run lint`/D-05 CI step instead of adding a second enforcement mechanism to maintain |
| HTML escaping | A new escaping function in `safeInsertion.ts` or `openclerk-word` | `openclerk-core`'s existing `escapeHtml` (already tested, complete for text + quoted-attribute contexts per `SECURITY_AUDIT.md`) | Reimplementing escaping risks reintroducing exactly the drift/duplication bug Phase 1 (CORE-01/02) just fixed |
| URL-scheme validation | A denylist of dangerous schemes (`javascript:`, `vbscript:`, etc.) | `openclerk-core`'s existing `isSafeHyperlinkUrl` (allowlist-based: `http:`/`https:`/`mailto:` only) | Denylists are the classic bypass-prone pattern (miss a scheme variant); the existing allowlist implementation is already audited and tested — wrap it, don't replace it |
| Nominal typing in TypeScript | A runtime wrapper class with `instanceof` checks | Branded/intersection types (`string & { __brand }`) | TypeScript has no native nominal typing; the branded-primitive pattern is the established zero-runtime-cost idiom, and a class wrapper would add allocation overhead unsuited to the `es5`/`ie 11` target and require unwrapping (`.value`) at every Office.js call site, which defeats the point of a transparent branded string |

**Key insight:** Every piece of "new" logic this phase needs (escaping, URL validation) already exists and is already tested in `openclerk-core`. The entire phase is about *placement and enforcement* — moving the guarantee from "convention, verified by manual code review" (documented explicitly as fragile in `CONCERNS.md`) to "compiler + linter enforced" — not about writing new validation logic.

## Common Pitfalls

### Pitfall 1: `.eslintrc.json` is silently ignored by `npm run lint`
**What goes wrong:** A developer adds `no-restricted-syntax` to `.eslintrc.json` (the file currently present in this repo), runs `npm run lint` locally, sees it pass, and assumes ESCAPE-03 is satisfied — but the rule was never actually evaluated.
**Why it happens:** `package.json`'s `"lint"` script runs `office-addin-lint check`, not `eslint` directly. `office-addin-lint`'s `lint.js` (`node_modules/office-addin-lint/lib/lint.js:32-38`, read directly) constructs its ESLint invocation with an explicit `-c <configFilePath>` flag, where `configFilePath` is `eslint.config.mjs` at the project root **if it exists**, otherwise `office-addin-lint`'s own bundled `node_modules/office-addin-lint/config/eslint.config.mjs` (confirmed by reading that bundled file — it re-exports `eslint-plugin-office-addins`'s recommended flat config). An explicit `-c` flag bypasses ESLint's automatic `.eslintrc.*`/`eslint.config.*` discovery entirely, at any ESLint major version — this is not an ESLint-9-flat-config-migration issue, it is `office-addin-lint`'s own design. [VERIFIED: direct read of installed `node_modules/office-addin-lint/lib/lint.js` and `node_modules/office-addin-lint/config/eslint.config.mjs`]
**How to avoid:** Create `eslint.config.mjs` at the `openclerk-word` project root (not `eslint.config.js`/`.ts`, not `.eslintrc.json` — the exact filename `eslint.config.mjs` is hardcoded in `office-addin-lint`'s lookup). Spread `eslint-plugin-office-addins`'s recommended config (so existing office-addin-specific rules keep firing) and add the `no-restricted-syntax` rule plus the `safeInsertion.ts` exemption in this new file. Confirmed the default lint file glob (`src/**/*.{ts,tsx,js,jsx}`, from `office-addin-lint`'s own `defaults.js`) already covers every current call site (`src/taskpane/word.ts`, `src/commands/commands.word.ts`) with no additional config needed.
**Warning signs:** `npm run lint` passes even when a raw `insertHtml`/`insertHyperlink`/`insertComment` call is added directly to `word.ts` outside `safeInsertion.ts` — that's the signal the rule isn't wired up. Verify by intentionally adding a violating call site during plan verification and confirming `npm run lint` exits non-zero before removing it again.

### Pitfall 2: `insertHyperlink` has no type declaration at all
**What goes wrong:** Attempting to give `insertHyperlink`'s parameters proper branded types on a *typed* `Word.Range` reference fails to compile, because `Word.Range` (from `@types/office-js@1.0.598`, the version installed in this repo) has no `insertHyperlink` member — it's a newer/preview Office.js API not yet reflected in the public type definitions. [VERIFIED: `grep -n "insertHyperlink" node_modules/@types/office-js/index.d.ts` returns zero matches, while `insertHtml`/`insertComment` are present and typed]
**Why it happens:** The existing code already works around this with `(item as any).insertHyperlink(...)` (`word.ts:223-224`) — this is pre-existing, not a new gap introduced by this phase.
**How to avoid:** Preserve the `as any` cast inside `safeInsertion.ts`'s wrapper implementation exactly as it exists today; only the *parameter types the wrapper function itself exposes* (`url: SafeHyperlinkUrl`, `displayText: SafeHtml`) need to be branded — the internal cast to reach the untyped Office.js preview API is an accepted, unavoidable exception, not a hole in the guard (the branding still ensures nothing unvalidated reaches that cast in the first place).
**Warning signs:** A TS2339 "Property 'insertHyperlink' does not exist on type 'Range'" compile error if the `as any` cast is accidentally removed while refactoring into `safeInsertion.ts`.

### Pitfall 3: `openclerk-core@0.2.7` is already published — there is no "unpublished bump" to build on
**What goes wrong:** Assuming (per CONTEXT.md's framing that local `0.2.7` is "one ahead of" the `^0.2.6` dependency `openclerk-word` currently uses) that the sibling checkout's `0.2.7` is a locally-bumped-but-not-yet-released version the planner can simply publish as-is once branded types are added.
**Why it happens:** `0.2.7`'s `package.json` version and its own `v0.2.7` git tag are already live: `npm view openclerk-core version` returns `0.2.7`, `npm view openclerk-core versions --json` includes `"0.2.7"`, and the sibling repo's `git status --short` is clean at that tag. [VERIFIED: npm registry + direct `git tag -l`/`git status` in `C:\Users\willi\openclerk-core`]
**How to avoid:** The plan's `openclerk-core` version-bump task must target a version **above** `0.2.7` (e.g. `0.3.0`), and the "prepare version bump" step must include editing `package.json`'s `version` field forward from `0.2.7`, not merely "run `npm publish` on the current checkout."
**Warning signs:** `publish.yml`'s own guard (`Verify tag matches package.json version` step, comparing `$GITHUB_REF_NAME` against `package.json`'s `version`) would fail the moment someone tries to push a `v0.2.7` tag a second time — npm also rejects republishing an existing version outright.

### Pitfall 4: AST selector must target `callee.property.name`, not `callee.name`
**What goes wrong:** A selector written as `CallExpression[callee.name='insertHtml']` (the form shown in ESLint's own top-level rule documentation example, which targets bare function calls like `insertHtml(...)`) matches nothing in this codebase, because every current call site is a *method* call on an object (`item.insertHtml(...)`, `(item as any).insertHtml(...)`, `range.insertComment(...)`) — the callee is a `MemberExpression`, not a plain `Identifier`.
**Why it happens:** ESLint's official rule-doc example (`no-restricted-syntax` docs, fetched directly) uses the bare-identifier form for illustration; it doesn't cover the member-call case this project needs.
**How to avoid:** Use `CallExpression[callee.property.name='insertHtml']` (and the equivalent for `insertHyperlink`/`insertComment`) — `callee.property.name` reaches the method name regardless of what the object expression is, so it matches `item.insertHtml(...)`, `(item as any).insertHtml(...)`, and any future variable name equally.
**Warning signs:** The rule reports zero violations even when a raw call is deliberately added during verification — that's the signal the selector is targeting the wrong AST shape.

## Code Examples

### `eslint.config.mjs` (new, project root)
```javascript
// Source: eslint.org/docs/latest/rules/no-restricted-syntax (selector syntax, fetched directly);
// exemption-via-second-config-object pattern confirmed against eslint/eslint#19239 (closed,
// "works as intended" -- a later flat-config entry fully replaces an earlier no-restricted-syntax
// value for files it matches, which is exactly the desired "off" override for safeInsertion.ts)
import officeAddins from "eslint-plugin-office-addins";
import tsParser from "@typescript-eslint/parser";

const RAW_INSERTION_SELECTORS = [
  {
    selector: "CallExpression[callee.property.name='insertHtml']",
    message: "Raw insertHtml calls must go through src/taskpane/safeInsertion.ts.",
  },
  {
    selector: "CallExpression[callee.property.name='insertHyperlink']",
    message: "Raw insertHyperlink calls must go through src/taskpane/safeInsertion.ts.",
  },
  {
    selector: "CallExpression[callee.property.name='insertComment']",
    message: "Raw insertComment calls must go through src/taskpane/safeInsertion.ts.",
  },
];

export default [
  ...officeAddins.configs.recommended,
  {
    plugins: { "office-addins": officeAddins },
    languageOptions: { parser: tsParser },
    rules: {
      "no-restricted-syntax": ["error", ...RAW_INSERTION_SELECTORS],
    },
  },
  {
    files: ["**/safeInsertion.ts"],
    rules: {
      "no-restricted-syntax": "off",
    },
  },
];
```

### `.github/workflows/ci.yml` new lint job (D-05)
```yaml
# Source: mirrors the existing "test" job's structure in this same file (build/test job pattern
# already established at ci.yml:53-71), read directly.
  lint:
    name: Lint
    needs: build
    runs-on: ubuntu-latest
    steps:
      - name: Checkout
        uses: actions/checkout@v4
      - name: Setup Node.js
        uses: actions/setup-node@v4
        with:
          node-version: '18'
          cache: 'npm'
      - name: Install dependencies
        run: npm ci
      - name: Run lint
        run: npm run lint
```

### `safeInsertion.test.ts` skeleton (D-06 — no existing Word.js mock pattern in this repo to follow; none found in `tests/`)
```typescript
// Source: original — no Word.js/Office.js mocking pattern currently exists in this repo's tests/
// directory (confirmed: only installer.test.ts and manifest.test.ts remain post-Phase-1, neither
// touches Word.js). Smart-constructor tests need no mock at all; dispatch/fallback tests need only
// a plain object matching the duck-typed shape the existing feature-detection already relies on
// (typeof x.insertHyperlink === "function") -- no real @types/office-js Word.RequestContext needed.
import { toSafeHyperlinkUrl, toSafeHtml } from "openclerk-core";
import { insertSafeHyperlink } from "../src/taskpane/safeInsertion";

describe("toSafeHyperlinkUrl", () => {
  test("returns null for a javascript: URL", () => {
    expect(toSafeHyperlinkUrl("javascript:alert(1)")).toBeNull();
  });
  test("returns a branded value for an https URL", () => {
    expect(toSafeHyperlinkUrl("https://example.com")).toBe("https://example.com");
  });
});

describe("insertSafeHyperlink dispatch", () => {
  test("prefers insertHyperlink when available", async () => {
    const insertHyperlink = jest.fn();
    const item = { insertHyperlink, insertHtml: jest.fn(), insertText: jest.fn() };
    const context = { sync: jest.fn().mockResolvedValue(undefined) };
    const url = toSafeHyperlinkUrl("https://example.com")!;
    const text = toSafeHtml("Smith v. Jones");
    await insertSafeHyperlink(context as any, item as any, url, text);
    expect(insertHyperlink).toHaveBeenCalled();
  });

  test("falls back to insertHtml when insertHyperlink is unavailable", async () => {
    const insertHtml = jest.fn();
    const item = { insertHtml, insertText: jest.fn() }; // no insertHyperlink property
    const context = { sync: jest.fn().mockResolvedValue(undefined) };
    const url = toSafeHyperlinkUrl("https://example.com")!;
    const text = toSafeHtml("Smith v. Jones");
    await insertSafeHyperlink(context as any, item as any, url, text);
    expect(insertHtml).toHaveBeenCalled();
  });
});
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|---------------|--------|
| `escapeHtml`/`isSafeHyperlinkUrl` called by convention at each of 4+ call sites in `word.ts`, verified only by manual security-audit code review (`SECURITY_AUDIT.md`, `CONCERNS.md` "Fragile Areas") | Branded types (`SafeHtml`, `SafeHyperlinkUrl`) make the guard a compile-time requirement; `no-restricted-syntax` makes bypassing the wrapper a build-breaking lint error | This phase | A future call site that forgets the guard now fails `npm run build`/`npm run lint`, not just a future manual audit |
| `.eslintrc.json` (legacy eslintrc format) — present in repo, extended `plugin:office-addins/recommended`, but **not actually consulted** by `npm run lint` in the currently-installed toolchain | Flat `eslint.config.mjs` at project root, the only format `office-addin-lint@3.0.9`'s `-c` invocation will pick up over its own bundled default | Toolchain fact as of this research (`office-addin-lint@3.0.9`, `eslint@9.39.4`, both currently installed) | `.eslintrc.json` should be considered dead config in this repo going forward for anything invoked via `npm run lint`; the planner should not assume adding rules there has any effect |
| No CI lint step | `npm run lint` added as a peer CI job (D-05) | This phase | Bypass-guard (ESCAPE-03) becomes enforceable on every PR, not just locally |

**Deprecated/outdated:**
- `.eslintrc.json` as the active lint config for this repo's `npm run lint` path: superseded in practice by `office-addin-lint`'s flat-config-first resolution — not formally deprecated by ESLint itself, but non-functional here regardless.

## Assumptions Log

| # | Claim | Section | Risk if Wrong |
|---|-------|---------|---------------|
| A1 | The general TypeScript branded-type + smart-constructor idiom (`string & { readonly __brand: unique symbol }`, constructor returns `null` on failure) is the community-standard pattern for this use case | Standard Stack (Alternatives), Architecture Patterns (Pattern 1) | Low — this is a well-established, widely-documented TS idiom cross-checked against multiple independent community sources (Effect docs, learningtypescript.com, kourge/ts-brand); even if a different exact brand-property name/shape were chosen, the underlying mechanism (intersection type + gatekept constructor) is sound and low-risk to adjust during planning/execution |
| A2 | `0.3.0` is the semver-correct next version for `openclerk-core` (new public exports = minor bump, not patch) | User Constraints (Claude's Discretion), Common Pitfalls (Pitfall 3) | Low — explicitly left to the executor's discretion per CONTEXT.md; any version above `0.2.7` is functionally valid, this is a semver-convention recommendation only |

**All other claims in this research are `[VERIFIED]`** — either confirmed by direct file reads of installed tooling (`office-addin-lint`, `@types/office-js`, both repos' `tsconfig.json`/`package.json`), direct `npm view`/`git tag`/`git status` registry checks, or ESLint's own official documentation (fetched directly) plus a closed, resolved GitHub issue confirming flat-config override behavior.

## Open Questions (RESOLVED)

1. **Exact minor/patch version number for the `openclerk-core` release**
   - What we know: Must be above `0.2.7` (currently published); `0.3.0` is the semver-conventional choice for new public exports.
   - What's unclear: Whether the repo owner has any other pending `openclerk-core` work queued that should land in the same release (would affect whether this is a clean single-purpose release or bundled with other changes).
   - Recommendation: Default to `0.3.0`, single-purpose release (branded types + smart constructors only); confirm with repo owner at the D-03 human checkpoint before the publish is actually triggered.
   - **RESOLVED:** Plan 02-01 targets `0.3.0`, single-purpose release, per the recommendation above.

2. **Whether `safeInsertion.ts`'s wrapper functions accept `context: Word.RequestContext` and call `context.sync()` internally, or leave sync to the caller**
   - What we know: Today's two call sites use different patterns — `applyHyperlinkToItem` takes `context` and syncs internally (`word.ts:232`); the `insertComment` call site syncs in the surrounding `Word.run` block *after* the insert (`word.ts:1233`), not inside a dedicated function.
   - What's unclear: Which pattern the planner should standardize on for both wrapper functions.
   - Recommendation: This is explicitly listed as Claude's Discretion in CONTEXT.md. Research suggests taking `context` and syncing internally for both (matching `applyHyperlinkToItem`'s existing pattern) is slightly safer for the ESLint-enforcement goal of this phase — it keeps `context.sync()` calls colocated with the insertion they follow, reducing the chance a future caller forgets to sync — but either choice satisfies ESCAPE-02/03 equally; this is a design-consistency call, not a correctness one.
   - **RESOLVED:** Plan 02-02 standardizes both wrapper functions on taking `context: Word.RequestContext` and calling `context.sync()` internally, matching the recommendation.

## Environment Availability

| Dependency | Required By | Available | Version | Fallback |
|------------|------------|-----------|---------|----------|
| Node.js | Build/test/lint in both repos | ✓ | v24.18.0 (local); CI pins Node 18 (build/test/lint jobs) and Node 22 (`publish` job) [VERIFIED: `.github/workflows/ci.yml`] | — |
| npm | Dependency install, `npm publish` | ✓ | 11.18.0 | — |
| git | Version tagging (`openclerk-core` release), CI checkout | ✓ | 2.55.0 | — |
| npm registry access | Verifying `openclerk-core` published version, eventual publish | ✓ | `npm view openclerk-core version` succeeded live | — |
| ESLint (via `office-addin-lint`) | ESCAPE-03 enforcement | ✓ | 9.39.4 (transitively installed) | — |
| `@types/office-js` | Type-checking `safeInsertion.ts` against `Word.Range`/`Word.Comment` | ✓ | 1.0.598 — lacks `insertHyperlink` typings (see Common Pitfalls Pitfall 2), has `insertHtml`/`insertComment` | `as any` cast for `insertHyperlink` only (pre-existing pattern, not new) |

**Missing dependencies with no fallback:** None.

**Missing dependencies with fallback:** None — `insertHyperlink`'s missing type declaration is not a missing *dependency*, it's a known, already-worked-around gap in the installed `@types/office-js` version; documented above for completeness, not a blocker.

## Validation Architecture

### Test Framework
| Property | Value |
|----------|-------|
| Framework | Jest 29 (`ts-jest` preset) [VERIFIED: `package.json`] — identical framework in both repos |
| Config file | `package.json`'s embedded `"jest"` block (`testMatch: **/tests/**/*.test.ts`, `testEnvironment: node`) — no standalone `jest.config.*` |
| Quick run command | `npx jest tests/safeInsertion.test.ts` |
| Full suite command | `npm test` (excludes `tests/courtListener.live.test.ts` implicitly — that file is invoked separately via `npm run test:live` and lives outside this phase's scope) |

### Phase Requirements → Test Map
| Req ID | Behavior | Test Type | Automated Command | File Exists? |
|--------|----------|-----------|-------------------|-------------|
| ESCAPE-01 | Smart constructors return branded value on valid input, `null` on invalid input (D-04) | unit | `npx jest tests/safeInsertion.test.ts -t "toSafeHyperlinkUrl"` | ❌ Wave 0 — new file |
| ESCAPE-01 | Branded-type parameter constraint actually rejects a plain `string` at compile time | type-check (compile-time, not runtime-testable via Jest) | `npm run build` (or `npx tsc --noEmit` on `safeInsertion.ts`) must fail if a wrapper function's parameter is accidentally typed `string` instead of the branded type — verify by transient negative test during plan verification (intentionally widen a type, confirm build fails, then revert) | ❌ Wave 0 — no existing compile-time negative-test convention in this repo; document as a manual/verification-step check rather than a permanent automated test (permanent `// @ts-expect-error` fixtures are an option but add ongoing maintenance for a one-time guarantee already covered by ESLint + code review) |
| ESCAPE-02 | `insertSafeHyperlink` dispatches to `insertHyperlink` → `insertHtml` → `insertText` in the existing fallback order | unit | `npx jest tests/safeInsertion.test.ts -t "dispatch"` | ❌ Wave 0 — new file |
| ESCAPE-02 | `insertSafeComment` calls `Word.Range.insertComment` with the given text | unit | `npx jest tests/safeInsertion.test.ts -t "insertSafeComment"` | ❌ Wave 0 — new file |
| ESCAPE-03 | `no-restricted-syntax` fails the build when a raw insertion call exists outside `safeInsertion.ts` | integration (lint-as-test) | `npm run lint` against a deliberately-reintroduced violation, confirmed to exit non-zero, during plan verification (not a permanent Jest test — ESLint itself is the enforcement mechanism, redundant to also unit-test it) | N/A — enforced by `npm run lint`, not Jest |
| Success Criterion 4 (manual smoke test: hyperlinking/Bluebook/hallucination-check/embed-text still work) | manual-only | Documented in CONTEXT.md as explicitly manual (real Word document, real Office.js host) — no automated equivalent exists or is expected given `word.ts` has zero existing automated coverage (D-06 scope boundary) | manual | — |

### Sampling Rate
- **Per task commit:** `npx jest tests/safeInsertion.test.ts`
- **Per wave merge:** `npm test` (full suite, both repos where applicable) + `npm run lint` (openclerk-word)
- **Phase gate:** Full suite green in both repos, `npm run lint` green, plus the manual Word smoke test (Success Criterion 4) before `/gsd-verify-work`

### Wave 0 Gaps
- [ ] `openclerk-word/tests/safeInsertion.test.ts` — new file, covers ESCAPE-01/02 (no existing test in this repo touches `word.ts`'s insertion logic at all — this is genuinely new coverage, not an extension of existing tests)
- [ ] `openclerk-word/eslint.config.mjs` — new file, required for ESCAPE-03 to be enforceable at all (see Pitfall 1)
- [ ] `openclerk-core/tests/utils.test.ts` — extend with `toSafeHyperlinkUrl`/`toSafeHtml` smart-constructor cases (existing file, read directly: currently only tests the plain `isSafeHyperlinkUrl`/`normalizeText`/etc. functions, not any branded-type constructors)
- [ ] No shared Word.js/Office.js mock helper exists anywhere in `openclerk-word`'s `tests/` directory — the plan should keep mocks minimal/inline (plain objects matching the duck-typed `insertHyperlink`/`insertHtml`/`insertText` shape) rather than introducing a new mocking framework/library, consistent with "Don't Hand-Roll" guidance and the project's `es5`/no-framework conventions

## Security Domain

### Applicable ASVS Categories

| ASVS Category | Applies | Standard Control |
|---------------|---------|-------------------|
| V2 Authentication | No | This phase touches no authentication logic (enterprise provider OAuth is untouched, out of scope) |
| V3 Session Management | No | No session state introduced or modified |
| V4 Access Control | No | No access-control logic in scope |
| V5 Input Validation | Yes | `openclerk-core`'s existing `isSafeHyperlinkUrl` (allowlist-based URL-scheme check, already ASVS-aligned per `SECURITY_AUDIT.md`'s prior review) becomes compiler-enforced via the `SafeHyperlinkUrl` branded type; `escapeHtml` (complete escaping of `&`/`<`/`>`/`"`/`'`) becomes compiler-enforced via `SafeHtml` — this phase strengthens V5 controls that already existed, per CLAUDE.md's explicit requirement that this refactor "preserve or strengthen existing guards, not weaken them" |
| V6 Cryptography | No | No cryptographic operations in scope |

### Known Threat Patterns for this stack

| Pattern | STRIDE | Standard Mitigation |
|---------|--------|----------------------|
| DOM/HTML injection via `insertHtml` with unescaped `url`/`displayText` interpolation | Tampering | `escapeHtml` (existing, tested) wrapped by the `SafeHtml` smart constructor — compiler now refuses to compile a call site that skips it, closing the "fragile, convention-only" gap `CONCERNS.md` documented before this phase |
| `javascript:`/`vbscript:`/`data:` scheme hyperlink injection | Tampering / Elevation of Privilege (XSS-adjacent, since a crafted hyperlink could execute script in the WebView context if inserted unvalidated) | `isSafeHyperlinkUrl`'s existing allowlist (`http:`/`https:`/`mailto:` only) wrapped by the `SafeHyperlinkUrl` smart constructor — same compiler-enforcement strengthening as above |
| Bypass of the wrapper module via a new raw Office.js call added directly to `word.ts` (regression risk this whole phase exists to close) | Tampering | `no-restricted-syntax` ESLint rule + CI lint gate (D-05) — turns a silent regression into a build failure; this is the *new* control this phase adds beyond what existed before |
| Silently non-functional lint gate (Pitfall 1: `.eslintrc.json` not read by `npm run lint`) | Tampering (of the control itself — a security gate that appears to exist but doesn't enforce anything) | Must use `eslint.config.mjs`, not `.eslintrc.json`, per Common Pitfalls — verified during plan execution by deliberately introducing and then confirming rejection of a violating call site before considering ESCAPE-03 complete |

## Sources

### Primary (HIGH confidence)
- Direct reads of installed tooling source in this repo and the sibling repo: `node_modules/office-addin-lint/lib/lint.js`, `node_modules/office-addin-lint/lib/commands.js`, `node_modules/office-addin-lint/lib/defaults.js`, `node_modules/office-addin-lint/config/eslint.config.mjs`, `node_modules/@types/office-js/index.d.ts`, `node_modules/eslint-plugin-office-addins` (recommended config inspected via `node -e`), both repos' `tsconfig.json`/`package.json`/`.eslintrc.json`
- `npm view openclerk-core version` / `npm view openclerk-core versions --json` — live npm registry query [VERIFIED: npm registry]
- `git tag -l "v0.2.*"` / `git status --short` in `C:\Users\willi\openclerk-core` [VERIFIED: local git]
- `src/taskpane/word.ts` (this repo, read directly, all cited line numbers confirmed against current file content)
- `src/utils.ts`, `tests/utils.test.ts`, `src/index.ts`, `package.json`, `.github/workflows/publish.yml` (openclerk-core, all read directly)
- eslint.org official docs: `no-restricted-syntax` rule page (fetched directly)

### Secondary (MEDIUM confidence)
- `eslint/eslint#19239` GitHub issue (fetched directly) — closed as "works as intended," confirms flat-config `no-restricted-syntax` per-file override behavior
- `.planning/codebase/SECURITY_AUDIT.md`, `.planning/codebase/CONCERNS.md`, `.planning/codebase/STRUCTURE.md` — prior project research, cross-referenced against current code state

### Tertiary (LOW confidence)
- WebSearch results on general TypeScript branded-type/smart-constructor community patterns (Effect docs, learningtypescript.com, kourge/ts-brand, oneuptime.com, etc.) — pattern cross-checked across multiple independent sources but not from an "official" TypeScript-language source (branded types are a community idiom, not a language feature with official docs); tagged `[ASSUMED]` in Assumptions Log (A1)

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH — no new packages, all versions confirmed via direct `node_modules`/registry inspection
- Architecture: HIGH — every call site, line number, and toolchain behavior (especially the `.eslintrc.json`-is-ignored finding) verified by direct file reads and live commands, not inferred
- Pitfalls: HIGH — the two most consequential findings (ESLint config resolution, `openclerk-core@0.2.7` already published) were both independently confirmed via direct tool execution, not assumed from CONTEXT.md's framing

**Research date:** 2026-07-15
**Valid until:** 30 days (stable toolchain; re-verify `openclerk-core`'s published version and `office-addin-lint`/`eslint` versions if this research is reused after a `npm install`/dependency bump in either repo)
