# Stack Research

**Domain:** Compile-time-enforced HTML/hyperlink escaping boundary for a vanilla-TypeScript,
es5/ie11-targeted Office.js Word task-pane add-in
**Researched:** 2026-07-15
**Confidence:** HIGH

## Codebase Facts That Drive This Recommendation

Verified directly in `WordClerk` (not assumed):

- The unsafe-insertion boundary is narrow and already identified: `src/taskpane/word.ts` has
  exactly one wrapper (`applyHyperlinkToItem`, `word.ts:221-237`) that calls the Office.js
  `Word.Range` insertion APIs — `insertHyperlink`, `insertHtml` (with manual `escapeHtml()` calls
  at `word.ts:230`), and a plain-text `insertText` fallback. Two more raw call sites exist outside
  that wrapper: `insertComment` (`word.ts:1248`) and `insertOoxml` (`word.ts:1497`, already fed
  pre-stripped OOXML). The guards being bypassed are `escapeHtml`/`isSafeHyperlinkUrl` in
  `src/taskpane/utils.ts:43-63`, both plain string→string/boolean functions with **no type-level
  signal** that their output is safe.
- The task-pane **DOM-rendering** side (the other classic HTML-injection surface) already follows
  a safe pattern — `createElement()` + `.textContent =` — and only uses `innerHTML =` for clearing
  (`""`) or static string literals (verified: `word.ts:397-855`, no dynamic `innerHTML` sink
  found). So this hardening effort is scoped correctly: the fragile surface is exclusively the
  **Office.js `Word.Range` insertion APIs**, not the DOM.
- Toolchain actually resolved in `node_modules` today: TypeScript `5.4.2`, ESLint `9.39.4`
  (pulled in transitively by `office-addin-lint@3.0.6`, which also transitively pulls in
  `typescript-eslint@^8.4.0`), config format is legacy `.eslintrc.json` (`plugin:office-addins/
  recommended`). `eslint-plugin-office-addins`'s rule set (`call-sync-after-load`,
  `no-context-sync-in-loop`, `no-navigational-load`, etc.) is entirely about `context.sync()`
  ordering — it has **no rule that touches HTML/URL safety**, and no
  `eslint-plugin-no-unsanitized`-style plugin is installed.
- `tsconfig.json` targets `es5`, `lib: ["es2015","dom"]`, and does not set `"strict": true`.
  `browserslist` includes `ie 11`, with `core-js`/`regenerator-runtime` as production deps for
  polyfilling — a real runtime-dependency-weight signal.

## Recommended Approach

**Combine two zero-new-dependency mechanisms — neither is sufficient alone:**

1. **Branded (nominal) `SafeHtml`/`SafeUrl` types** so a function that requires "already-escaped"
   input cannot type-check against a raw `string`.
2. **An ESLint `no-restricted-syntax` rule (core ESLint, already resolved — no plugin needed)**
   that bans direct calls to the raw Office.js insertion methods (`insertHtml`, `insertHyperlink`,
   `insertComment`) anywhere in `src/taskpane/**` except inside one designated wrapper module.

**Why both are required, not either alone (HIGH confidence — this is a direct consequence of how
`@types/office-js` is authored, verified by reading the pattern, not a stylistic preference):**

`@types/office-js` declares `Word.Range.insertHtml(html: string, insertLocation: InsertLocation)`
with a plain `string` parameter. A branded type `SafeHtml = string & { readonly __brand: unique
symbol }` is structurally still a `string` (an intersection with a plain object literal is
assignable *to* `string`), so **nothing stops a raw string from being passed straight into the
real `Word.Range.insertHtml`** — branding only protects the boundary of *your own* wrapper
function's parameter type, never the ambient third-party method itself. TypeScript interface
merging can only *add* overloads to an ambient class, never narrow/remove the existing
`(html: string, ...)` overload, so "augment the Office.js `.d.ts` to require `SafeHtml`" is not a
viable path (and would be fragile against `@types/office-js` version bumps regardless). This means
the type system can only fully close the hole at a **single, deliberately narrow call site** — the
wrapper module — and ESLint must be the mechanism that guarantees every other call site is forced
through it.

### Core Technologies

| Technology | Version | Purpose | Why Recommended |
|------------|---------|---------|-----------------|
| TypeScript branded/nominal types | Already at TS 5.4.2, no bump needed | `SafeHtml`, `SafeHyperlinkUrl` opaque string subtypes returned only by `escapeHtml`/`isSafeHyperlinkUrl` | Purely a compile-time construct — the brand field is erased entirely during transpilation, so it has **zero runtime cost and zero es5/ie11 compatibility risk** (no polyfill, no new syntax feature, nothing shipped to the bundle). This is the standard 2025-era pattern for "tainted vs. sanitized string" boundaries (Trusted Types' `TrustedHTML`, Closure's `goog.html.SafeHtml`, and hand-rolled branded types are all the same idea; branded types are the zero-dependency version). |
| ESLint core rule `no-restricted-syntax` (AST selector) | ESLint 9.39.4 (already resolved via `office-addin-lint`) | Ban `CallExpression` nodes whose `callee.property.name` is `insertHtml`/`insertHyperlink`/`insertComment` project-wide, with a per-file `overrides` exemption for the one wrapper module | Ships with ESLint core — **no new dependency at all**. AST-selector matching on `callee.property.name` catches the call regardless of the receiver variable's name (`item.insertHtml(...)`, `range.insertHtml(...)`, `searchResults.items[0].insertComment(...)`), which a name-based rule like `no-restricted-properties` cannot do cleanly since the object identifier varies per call site in this codebase. |

### Supporting Libraries

None needed. Both mechanisms above are implemented with tooling the project already resolves
(`typescript`, `eslint` via `office-addin-lint`). Do not add a runtime sanitizer package — the
existing hand-rolled `escapeHtml`/`isSafeHyperlinkUrl` in `utils.ts` are already correct for this
project's narrow needs (escaping into an `<a href>`/text node, and allow-listing `http:`/`https:`/
`mailto:` schemes) and adding e.g. DOMPurify would be a large, unnecessary runtime dependency for
an es5 bundle that has no need for full DOM sanitization.

### Development Tools

| Tool | Purpose | Notes |
|------|---------|-------|
| ESLint `.eslintrc.json` `overrides` array (legacy config format, already in use) | Scope the `no-restricted-syntax` ban so the designated wrapper module is exempt | Add an `overrides` entry keyed to the new wrapper file's path (e.g. `src/taskpane/safeInsertion.ts`) that disables the rule only there — this is the single legitimate place the raw Office.js methods may be called. |
| TypeScript compiler (`tsc --noEmit`, already run via `npm run build`) | Enforces the branded-type boundary — a raw `string` passed where `SafeHtml`/`SafeHyperlinkUrl` is required fails the build, not just lint | No config change needed beyond the new type declarations; `noEmitOnError: true` (already set) means a bad call site cannot ship. |

## Concrete Pattern (code sketch)

**1. Brand the escaping/validation functions' return types** (`src/taskpane/utils.ts`):

```typescript
// Nominal/branded types: structurally still `string` at runtime (zero cost), but the compiler
// will not accept a plain `string` where these are required -- only these two functions can
// produce them.
export type SafeHtml = string & { readonly __brand: "SafeHtml" };
export type SafeHyperlinkUrl = string & { readonly __brand: "SafeHyperlinkUrl" };

export function escapeHtml(str: string): SafeHtml {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;") as SafeHtml;
}

// Returns the branded type only on the validated path; callers must handle the `null` case,
// which forces a decision at every call site instead of silently trusting an unchecked URL.
export function toSafeHyperlinkUrl(url: string): SafeHyperlinkUrl | null {
  return isSafeHyperlinkUrl(url) ? (url as SafeHyperlinkUrl) : null;
}
```

**2. One wrapper module owns every raw Office.js insertion call** (new file, e.g.
`src/taskpane/safeInsertion.ts` — the *only* file exempted by the ESLint rule below):

```typescript
import { SafeHtml, SafeHyperlinkUrl, escapeHtml } from "./utils";

export async function insertSafeHyperlink(
  context: Word.RequestContext,
  item: Word.Range,
  url: SafeHyperlinkUrl,   // <-- compiler rejects a plain string here
  displayText: string
): Promise<void> {
  if (typeof (item as any).insertHyperlink === "function") {
    (item as any).insertHyperlink(url, displayText, Word.InsertLocation.replace);
  } else if (typeof (item as any).insertHtml === "function") {
    const html: SafeHtml = escapeHtml(`<a href="${url}">${escapeHtml(displayText)}</a>`);
    (item as any).insertHtml(html, Word.InsertLocation.replace);
  } else {
    item.insertText(displayText, Word.InsertLocation.replace);
  }
  await context.sync();
}
```

Callers in `word.ts` now call `insertSafeHyperlink(...)` with a `SafeHyperlinkUrl` obtained only
via `toSafeHyperlinkUrl()` — a raw `entry.url: string` will not compile.

**3. ESLint gate that makes bypassing the wrapper a build failure**, in `.eslintrc.json`:

```jsonc
{
  "rules": {
    "no-restricted-syntax": [
      "error",
      {
        "selector": "CallExpression[callee.property.name=/^(insertHtml|insertHyperlink|insertComment)$/]",
        "message": "Do not call Word.Range insertion APIs directly -- route through insertSafeHyperlink()/insertSafeHtml() in safeInsertion.ts so escapeHtml/isSafeHyperlinkUrl cannot be skipped."
      }
    ]
  },
  "overrides": [
    {
      "files": ["src/taskpane/safeInsertion.ts"],
      "rules": { "no-restricted-syntax": "off" }
    }
  ]
}
```

This is the piece that actually satisfies "impossible... to skip" — the branded type alone only
protects calls *into* the wrapper; this rule is what prevents a future contributor from writing a
brand-new `item.insertHtml(rawString, ...)` call somewhere else in `word.ts` (or a new feature
module) that never goes through the wrapper at all.

## Alternatives Considered

| Recommended | Alternative | When to Use Alternative |
|-------------|-------------|--------------------------|
| Hand-rolled branded `SafeHtml`/`SafeHyperlinkUrl` types | `io-ts` / `zod` / `newtype-ts` branded-type + runtime-validation libraries | Only if the project later needs runtime (not just compile-time) validation of *many* different domain types with parsing/decoding — overkill here for two string-shaped guards, and each is a real added dependency this project's minimal-dependency philosophy argues against. |
| ESLint core `no-restricted-syntax` | `eslint-plugin-no-unsanitized` | Only if the codebase were sanitizing native DOM sinks (`element.innerHTML =`, `insertAdjacentHTML`, `document.write`). It does not inspect Office.js's proprietary `Word.Range.insertHtml`/`insertHyperlink`/`insertComment` RPC methods at all (they are not DOM APIs), so installing it here would add a dependency that provides **zero coverage** of the actual insertion sites in `word.ts`. Not recommended for this milestone. |
| A single shared wrapper module + ESLint ban on raw calls elsewhere | Monkey-patch/augment the `@types/office-js` ambient declarations to require `SafeHtml` on `Word.Range.insertHtml` directly | TypeScript declaration merging can only *add* overloads to an ambient class, not remove/narrow the existing `(html: string, ...)` overload, so a plain string would still type-check against the original overload. This path does not work and would also silently break on `@types/office-js` version bumps. Do not pursue. |
| Compile-time (TypeScript + ESLint) enforcement | Browser Trusted Types (`require-trusted-types-for 'script'` CSP directive + `TrustedHTML`) | Trusted Types intercepts native DOM sinks (`Element.innerHTML`, `document.write`, etc.) at the browser level. `Word.Range.insertHtml` is not a DOM sink — it is a custom Office.js message sent to the Word host — so Trusted Types has **no effect on the actual APIs this milestone is hardening**. It would be relevant only if/when the task pane starts writing dynamic content via `innerHTML =` (verified: it currently does not; see Codebase Facts above), and even then is a browser-platform feature, not an npm dependency, so it's compatible with the minimal-dependency philosophy if ever needed for that separate surface. |

## What NOT to Use

| Avoid | Why | Use Instead |
|-------|-----|--------------|
| `eslint-plugin-no-unsanitized` | Targets `innerHTML`/`insertAdjacentHTML`/`document.write` DOM sinks only; does not recognize Office.js's `Word.Range.insertHtml`/`insertHyperlink`/`insertComment` methods as sinks at all, so it would add a dependency with no effect on the bug class this milestone targets | ESLint core `no-restricted-syntax` with a `callee.property.name` AST selector (no new dependency) |
| DOMPurify or any general HTML sanitizer library | Full sanitization is the wrong tool for a narrow "build one `<a href>` tag from an already-scheme-validated URL and escaped text" need; adds real weight to an es5-targeted bundle that already carries `core-js`/`regenerator-runtime` polyfill overhead, for a case the existing 20-line `escapeHtml`/`isSafeHyperlinkUrl` already covers correctly | Keep the existing hand-rolled `escapeHtml`/`isSafeHyperlinkUrl`; only brand their return types |
| Augmenting/monkey-patching `@types/office-js`'s `Word.Range` class to change `insertHtml`'s parameter type | TS declaration merging can only add overloads, not narrow/remove the ambient `(html: string, ...)` overload — a plain string still type-checks against the original signature, so this gives a false sense of safety and breaks silently on `@types/office-js` upgrades | Route all insertion through one hand-written wrapper module with branded parameter types, gated by ESLint (see Recommended Approach) |
| Branded types alone, with no ESLint gate | A branded type only constrains inputs to *your own* wrapper function; it cannot stop a new call site from calling the raw ambient `Word.Range.insertHtml(rawString, ...)` directly, because that ambient method still accepts plain `string` | Pair branded types with the `no-restricted-syntax` ESLint rule banning direct calls outside the wrapper module |
| Trusted Types / CSP `require-trusted-types-for` for this specific bug | Does not intercept Office.js's custom `insertHtml`/`insertHyperlink`/`insertComment` RPC methods (not native DOM sinks); would require a WebView2/Word-host CSP change with no effect on the actual insertion path being hardened | ESLint `no-restricted-syntax` + branded types (see above) |

## Stack Patterns by Variant

**If a future insertion call site needs to write OOXML directly (like the existing `insertOoxml`
at `word.ts:1497`) rather than HTML:**
- Extend the same wrapper module with an `insertSafeOoxml(context, target, ooxml: SafeOoxml)`
  function and add `insertOoxml` to the `no-restricted-syntax` selector's banned-property regex.
- Because that existing call site already only receives already-stripped/validated OOXML
  (`strippedOoxml` from `stripHtmlHyperlinks`), branding its input type as `SafeOoxml` (produced
  only by `stripHtmlHyperlinks`/an equivalent validator) closes the same class of gap there too.

**If the module split (this milestone's other deliverable) moves hyperlink logic into its own
feature module:**
- Put `safeInsertion.ts` at the same level as the new feature modules (e.g.
  `src/taskpane/hyperlinking/safeInsertion.ts` or a shared `src/taskpane/shared/`), not inside
  `word.ts` — it should be a leaf dependency every feature module imports, never the reverse, to
  keep it trivial to update the ESLint `overrides.files` glob to a single stable path.

## Version Compatibility

| Package A | Compatible With | Notes |
|-----------|------------------|-------|
| TypeScript 5.4.2 (already pinned via `^5.4.2`) | Branded types via `string & { readonly __brand: unique symbol }` | No version bump required; the intersection-type branding pattern has worked unchanged since TS 2.x and is target-independent (compiles away regardless of `target: es5`). |
| ESLint 9.39.4 (resolved transitively via `office-addin-lint@^3.0.6`) | `no-restricted-syntax` in legacy `.eslintrc.json` format with `overrides` | Confirmed resolved in `node_modules` at the time of this research; `no-restricted-syntax` and file-scoped `overrides` are long-stable core ESLint features present in both eslintrc and flat config, so this survives a future `office-addin-lint`/ESLint flat-config migration with only a config-syntax translation, not a logic change. |
| `es5` / `ie 11` target (`tsconfig.json`, `browserslist`) | Branded types (zero runtime footprint) | No interaction — branded types produce no new syntax or runtime helper, so they impose no additional Babel/`core-js` polyfill burden beyond what the project already ships. |

## Sources

- Direct codebase inspection (`src/taskpane/word.ts`, `src/taskpane/utils.ts`, `package.json`,
  `tsconfig.json`, `.eslintrc.json`, resolved `node_modules` versions) — HIGH confidence, verified
  against the actual repository state on 2026-07-15, not assumed from training data.
- [ESLint `no-restricted-syntax` rule docs](https://eslint.org/docs/latest/rules/no-restricted-syntax) — confirms current (ESLint 9-era) AST-selector syntax and behavior — HIGH confidence, official docs.
- [Branded Types — Learning TypeScript](https://www.learningtypescript.com/articles/branded-types) — confirms the `string & { __brand }` pattern and its compile-time-only, zero-runtime-cost nature — MEDIUM-HIGH confidence, community reference consistent with known TS semantics.
- [TypeScript Application Security from A to Z — DEV Community](https://dev.to/devsdaddy/typescript-application-security-from-a-to-z-a-guide-to-protecting-against-obvious-and-55nh) — corroborates the `SafeHtml`-via-branding pattern as the standard TypeScript-native (no-new-dependency) mitigation for this exact bug class — MEDIUM confidence, community source, cross-checked against official ESLint/TS docs above.
- Verified in `node_modules`: ESLint `9.39.4` and `typescript-eslint@^8.4.0` are already resolved transitively through `office-addin-lint@^3.0.6`, and no `eslint-plugin-no-unsanitized`-equivalent package is installed — HIGH confidence, direct inspection.

---
*Stack research for: Compile-time HTML/hyperlink-escaping enforcement, Office.js vanilla-TS task-pane add-in*
*Researched: 2026-07-15*
