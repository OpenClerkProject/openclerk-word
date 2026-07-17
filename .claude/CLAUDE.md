<!-- GSD:project-start source:PROJECT.md -->

## Project

**OpenClerk**

OpenClerk is a Word add-in (task pane) for legal citation work: hyperlinking case-law and
parenthetical citations, checking citations for Bluebook formatting problems across three
editions (20th/21st/22nd), looking citations up live against public (CourtListener) and
enterprise (Westlaw, LexisNexis, Bloomberg Law) legal databases, detecting fabricated
("hallucinated") citations, and embedding cited opinion text — entirely client-side, with no
backend server of its own.

**Core Value:** Legal citation checks and hyperlinks must be accurate and trustworthy — a hallucination check
must never falsely report a fabricated citation as "verified."

### Constraints

- **Tech stack**: TypeScript targeting `es5`/`ie 11` (browserslist), no frontend framework, direct
  DOM manipulation — new module extractions must follow the existing `providers/`/`bluebook/`
  registry-plugin pattern rather than introducing a framework

- **Compatibility**: Must keep working inside the Office.js WebView host across Word
  desktop/online on Windows and macOS

- **Security**: This is a legal tool — citation verification must stay conservative (never claim
  "verified" when uncertain); any refactor of the hallucination-check or hyperlink-insertion paths
  must preserve or strengthen existing guards, not weaken them
<!-- GSD:project-end -->

<!-- GSD:stack-start source:codebase/STACK.md -->

## Technology Stack

## Languages

- TypeScript (compiled with `target: es5`, `lib: es2015, dom`) - `src/**/*.ts`, used for all add-in logic (taskpane UI, Word interop, Bluebook citation engine, citation-provider integrations)
- HTML - `src/taskpane/taskpane.html`, `src/commands/commands.html` (Office Add-in task pane and function-file entry points)
- CSS - `src/taskpane/taskpane.css`
- PowerShell / Bash / Batch - `scripts/install-openclerk.ps1`, `scripts/install-openclerk.sh`, `scripts/install-openclerk.cmd`, `scripts/local-server/*.ps1` (end-user installer/offline-server tooling, not part of the add-in bundle)
- XML - `manifest.xml` (Office Add-in manifest, uses Microsoft's `OfficeApp`/`VersionOverridesV1_0` schema)

## Runtime

- Node.js 18 (CI: `.github/workflows/ci.yml` uses `actions/setup-node@v4` with `node-version: '18'`, except the `publish` job which uses Node 22 for an ESM-only manifest-validation dependency)
- Browser runtime: Microsoft Office WebView (Word desktop/online), targeting `es5`/`ie 11` per `browserslist` in `package.json` for maximum Office host compatibility
- npm
- Lockfile: present (`package-lock.json`)

## Frameworks

- Office JS API (`@types/office-js`, `@types/office-runtime`) - Word document interop via `src/taskpane/word.ts`
- No frontend UI framework (React/Vue/Angular) - taskpane UI is built with vanilla TS/DOM manipulation (`src/taskpane/taskpane.ts`, `src/taskpane/utils.ts`)
- `marked` (^18.0.5) - Markdown rendering, used in `scripts/build-docs.js` to generate `dist/index.html` from `README.md`
- Jest 29 with `ts-jest` preset - config embedded in `package.json` (`testMatch: **/tests/**/*.test.ts`, `testEnvironment: node`)
- Test suite: `tests/*.test.ts` (bluebook citation formatting, hyperlinks, installer scripts, manifest validation, opinion-text extraction, provider behavior, utils)
- `tests/courtListener.live.test.ts` is excluded from the default `npm test` run; invoked separately via `npm run test:live` (hits the real CourtListener API)
- Webpack 5 (`webpack.config.js`) - bundles `taskpane` and `commands` entry points, dev-serves via `webpack-dev-server` with `office-addin-dev-certs` for HTTPS
- Babel (`@babel/core`, `@babel/preset-env`, `@babel/preset-typescript`) via `babel-loader` - transpiles TS/JS for the `es5`/`ie 11` browserslist target
- `office-addin-cli` / `office-addin-debugging` / `office-addin-dev-settings` / `office-addin-manifest` / `office-addin-lint` - Microsoft's Office Add-in tooling for sideloading, manifest validation, and linting
- `sharp` - image processing, used in `scripts/convert-logos.js`
- `archiver` - zips release packages in `scripts/package-release.js` / `scripts/package-release-offline.js`
- ESLint with `eslint-plugin-office-addins` (`plugin:office-addins/recommended`) - `.eslintrc.json`
- Prettier via `office-addin-prettier-config` (`"prettier"` field in `package.json`)

## Key Dependencies

- `jszip` (^3.10.1) + `@types/jszip` - reads/writes `.docx` (OOXML zip) structure directly, used for citation/hyperlink manipulation independent of the Office JS API
- `core-js` / `regenerator-runtime` - polyfills for `es5`/`ie 11` target compatibility
- `copy-webpack-plugin`, `html-webpack-plugin`, `html-loader`, `file-loader`, `source-map-loader` - webpack asset/HTML pipeline
- `os-browserify`, `process` - Node built-in shims for browser bundling

## Configuration

- No `.env` file present in the repo; the add-in has no server-side environment variables — it runs entirely client-side inside Word
- Per-provider credentials (CourtListener API token, LexisNexis/Westlaw/Bloomberg Law API base URL + client id/secret) are supplied by the end user at runtime through the taskpane UI and held in memory only for the session — see `src/taskpane/providers/base.ts` and `src/taskpane/providers/courtListenerProvider.ts`
- `package.json` `config` block sets Office debugging defaults (`app_to_debug: word`, `dev_server_port: 3000`)
- `webpack.config.js` - dev server on port 3000 with HTTPS via `office-addin-dev-certs`, separate `taskpane`/`commands` bundles
- `tsconfig.json` - `es5` target, `jsx: react` (unused, artifact of the Office Add-in template), `outDir: lib`
- `babel.config.json` - preset-env + preset-typescript
- `manifest.xml` - Office Add-in manifest; production variant generated at package time replaces `https://localhost:3000` URLs (verified by CI's "Verify production manifest URLs" step)
- `.clasp.json` / `appsscript.json` - Google Apps Script project config; present in the repo root but no corresponding `.gs`/Apps Script source files were found under `src/` — appears to be a leftover/unused scaffold rather than an active integration

## Platform Requirements

- Node.js 18+ and npm
- Windows/macOS with Word desktop (or Word on the web) for sideloading via `manifest.xml`
- HTTPS dev certs (`office-addin-dev-certs`) required for local Office Add-in development
- Distributed as a sideloaded Office Add-in (no app-store/AppSource listing referenced in code) via:
- Optional Microsoft Partner Center publishing step in CI (`publish` job), gated on `PARTNER_CENTER_*` secrets being configured

<!-- GSD:stack-end -->

<!-- GSD:conventions-start source:CONVENTIONS.md -->

## Conventions

## Naming Patterns

- Source files: `camelCase.ts` (e.g. `src/taskpane/providers/citationParser.ts`, `src/taskpane/bluebook/checkCaseNameAbbreviations.ts`)
- Generated data files: `<name>.generated.ts` under `src/taskpane/bluebook/generated/` — never hand-edit, produced by `scripts/generate-bluebook-data.js` (npm script `bluebook:update-data`)
- Test files: `<subject>.test.ts` in top-level `tests/` directory, mirroring the module under test (not co-located with source)
- `camelCase`, verb-first: `parseCaseCitation`, `extractCaseCitations`, `checkCommonCaseCitationRules`, `applyManualReporterOverrides`, `stripHtmlHyperlinks`
- Boolean-returning functions read as predicates: `isAuthenticated()`, `isReadyForOpinionText()`, `wasLastRequestRateLimited()`
- `camelCase` for locals/params; `UPPER_SNAKE_CASE` for module-level constants (`API_BASE`, `SITE_ORIGIN`, `EXAMPLE_CITATION` in tests)
- `PascalCase` for interfaces and classes: `CitationProvider`, `ParsedCitation`, `CourtListenerCluster`, `CourtListenerProvider`
- Provider-vendor-specific response shapes are declared as local `interface`s at the top of the provider file (e.g. `CourtListenerCitationResult`, `CourtListenerOpinionsResponse` in `src/taskpane/providers/courtListenerProvider.ts`), not shared globally

## Code Style

- Prettier via `office-addin-prettier-config` (`prettier` field in `package.json`); run with `npm run prettier`
- No local `.prettierrc` overrides — the shared office-addin config is authoritative
- ESLint via `eslint-plugin-office-addins`, `plugin:office-addins/recommended` (`.eslintrc.json`)
- Run with `npm run lint`; `npm run lint:fix` for auto-fixable issues
- Office-addin-specific rules are enforced (e.g. `call-sync-after-load`, `no-office-initialize`, `test-for-null-using-isNullObject`) — relevant when touching any Word/Office.js API calls in `src/taskpane/word.ts` or `src/commands/`
- `tsconfig.json` targets `es5` with `lib: ["es2015", "dom"]`, `jsx: "react"`, `experimentalDecorators: true`
- `allowJs: true`, `esModuleInterop: true`; no `strict` mode enabled — do not assume strict-null-checks safety when writing new code
- Double-quoted string literals are common in provider/type files (`base.ts`, `courtListenerProvider.ts`); single-quoted strings appear in test files. Prettier governs actual enforcement — do not hand-wrangle quote style.

## Import Organization

- None. Use relative paths consistently with existing files.

## Error Handling

- Providers throw `Error` with a user-facing message on invalid/missing credentials, rather than returning error codes: see `EnterpriseCitationProvider.authenticate()` in `src/taskpane/providers/base.ts` (`throw new Error(\`Missing required field(s): ...\`)`)
- Network/lookup failures are swallowed and converted to `null` return values rather than propagated exceptions — "move on" semantics are explicit in test names (`tests/providers.test.ts`: "moves on (returns null) instead of throwing on a network failure"). This is a deliberate pattern: citation lookups are best-effort and must never crash the taskpane UI.
- Rate-limiting (HTTP 429) is tracked via a dedicated boolean flag (`lastRequestWasRateLimited` in `CourtListenerProvider`) rather than a thrown/typed error, so callers can distinguish "rate limited" from "not found" without try/catch.
- Validation errors (e.g. non-HTTPS API base URL in `EnterpriseCitationProvider.authenticate`) are thrown synchronously/early, before any network call is attempted — fail fast on bad input.

## Comments

- JSDoc-style block comments precede exported classes/functions to explain *why*, not *what* — see `src/taskpane/providers/base.ts` (explains why credentials are never persisted) and `src/taskpane/providers/courtListenerProvider.ts` (explains CourtListener's auth requirement and links to API docs)
- Inline comments justify non-obvious behavior or document known limitations, especially near regex/parsing logic (`src/taskpane/providers/citationParser.ts` and its test file `tests/bluebook.test.ts` document "known limitation" cases directly in test names/comments)
- Regression-motivated tests include a comment explaining the bug that was fixed and why the assertion matters (see `tests/bluebook.test.ts` lines ~98-117)
- Used selectively on exported abstractions (base classes, provider entry points), not on every function. Favor documenting *design rationale* (security/legal/compliance reasoning) over restating parameter types.

## Function Design

## Module Design

<!-- GSD:conventions-end -->

<!-- GSD:architecture-start source:ARCHITECTURE.md -->

## Architecture

## System Overview

```text

```

## Component Responsibilities

| Component | Responsibility | File |
|-----------|----------------|------|
| Task pane entry | Bootstraps the UI bundle by importing the controller | `src/taskpane/taskpane.ts` |
| UI controller / workflow orchestrator | Office.onReady wiring, all DOM event handlers, hyperlinking workflows, source-document (.docx) parsing, Bluebook check orchestration, hallucination check orchestration, cited-text embedding | `src/taskpane/word.ts` |
| Citation provider plugin system | Defines `CitationProvider` interface and registry for pluggable case-law lookup sources | `src/taskpane/providers/types.ts`, `src/taskpane/providers/registry.ts`, `src/taskpane/providers/index.ts` |
| Built-in providers | Concrete lookup implementations (free CourtListener API; enterprise-gated Westlaw/Lexis/Bloomberg/USPTO stubs requiring user-supplied credentials) | `src/taskpane/providers/courtListenerProvider.ts`, `westlawProvider.ts`, `lexisNexisProvider.ts`, `bloombergLawProvider.ts`, `usptoPatentCenterProvider.ts` |
| Enterprise provider base class | Shared OAuth2 client-credentials auth flow for paid/contract-gated APIs | `src/taskpane/providers/base.ts` |
| Citation parsing/formatting | Parses raw citation text into structured `ParsedCitation`; extracts pincite page ranges; extracts opinion text from API responses | `src/taskpane/providers/citationParser.ts`, `pincitePages.ts`, `opinionTextExtractor.ts` |
| Bluebook rule engine | Registry of per-edition rule sets (20th/21st/22nd) that validate/format case-name and reporter citations | `src/taskpane/bluebook/registry.ts`, `edition20th.ts`, `edition21st.ts`, `edition22nd.ts`, `commonRules.ts`, `courtRules.ts`, `reporterRules.ts`, `pageRangeRules.ts`, `checkCaseNameAbbreviations.ts`, `manualCorrections.ts` |
| Generated reference data | Large static lookup tables (case-name/reporter/state abbreviations) generated from external sources by a build script, not hand-edited | `src/taskpane/bluebook/generated/*.generated.ts` (produced by `scripts/generate-bluebook-data.js`) |
| Ribbon/function commands | Minimal Office "commands" entry point for ribbon-triggered actions (function-file bundle, separate from the task pane bundle) | `src/commands/commands.ts`, `src/commands/commands.word.ts`, `src/commands/commands.html` |
| Shared utilities | Small cross-cutting helpers (string/DOM helpers used by the task pane) | `src/taskpane/utils.ts` |
| Installer/build tooling | Node scripts for packaging, release, and local install of the add-in (not part of the runtime add-in) | `scripts/*.js`, `scripts/install-openclerk.*` |
| Standalone PDF extraction tool | Separate self-contained utility (own `node_modules`, own `dist`) for extracting text from PDFs, decoupled from the Office add-in build | `tools/pdf-extract/` |

## Pattern Overview

- No framework (no React/Vue/Angular) — direct DOM manipulation via `document.getElementById` and manual event listeners, all concentrated in `src/taskpane/word.ts`.
- Two independent webpack entry bundles: `taskpane` (the full UI) and `commands` (a lightweight function-file for ribbon buttons) — see `webpack.config.js` entry block.
- Extensibility via two parallel registry/plugin patterns: `CitationProviderRegistry` (`src/taskpane/providers/registry.ts`) and `bluebookRuleSetRegistry` (`src/taskpane/bluebook/registry.ts`). New providers/rule sets self-register by import side-effect in their respective `index.ts`.
- Office Open XML (OOXML) is read/written directly in places (via JSZip) for operations Office.js does not expose cleanly, e.g. hyperlink relationship parsing (`parseRelationships` in `src/taskpane/word.ts:1508`) and source `.docx` parsing (`parseSourceDocument`, `word.ts:1401`).
- Credentials for enterprise providers are held only in memory for the session (never persisted to disk/localStorage) — enforced by convention in `EnterpriseCitationProvider` (`src/taskpane/providers/base.ts`).

## Layers

- Purpose: DOM wiring, user workflows (hyperlinking case-law citations, Bluebook checks, hallucination checks, embedding cited opinion text), talks to Office.js Word API
- Location: `src/taskpane/taskpane.ts`, `src/taskpane/taskpane.html`, `src/taskpane/word.ts`, `src/taskpane/taskpane.css`
- Contains: Office.onReady bootstrap, event handlers, rendering functions, OOXML parsing helpers
- Depends on: providers layer, bluebook layer, `src/taskpane/utils.ts`, Office.js/Word.js globals
- Used by: Office task pane runtime (loaded via manifest `SourceLocation`)
- Purpose: Abstracts "look up a citation" / "get opinion text" across multiple case-law data sources behind one interface
- Location: `src/taskpane/providers/`
- Contains: `CitationProvider` interface (`types.ts`), registry, base class for auth, concrete provider implementations
- Depends on: `fetch` (browser), nothing else internal
- Used by: `src/taskpane/word.ts`
- Purpose: Encodes citation-formatting rules per Bluebook edition, validates document citations against them
- Location: `src/taskpane/bluebook/`
- Contains: `BluebookRuleSet` interface (`types.ts`), registry, per-edition classes, shared rule modules, generated data tables
- Depends on: generated data (`generated/*.generated.ts`)
- Used by: `src/taskpane/word.ts` (Bluebook-check workflow)
- Purpose: Separate, minimal bundle Office loads for ribbon-button function commands (kept small/fast per Office add-in convention)
- Location: `src/commands/`
- Depends on: Office.js
- Used by: manifest ribbon action bindings
- Purpose: Generate static data, package releases, install the add-in locally
- Location: `scripts/`
- Depends on: Node.js, npm packages (archiver, etc.)

## Data Flow

### Primary Request Path (apply case-law hyperlinks from a source document)

### Bluebook Check Flow

### Online Citation Lookup Flow (hallucination check / embed opinion text)

- All state is in-memory, module-level variables inside `src/taskpane/word.ts` (selected provider, parsed citation map, last scan results) — reset on task-pane reload. No global store, no persistence layer, no server-side session.

## Key Abstractions

- Purpose: Uniform interface for "authenticate, then look up a citation" across free and paid case-law APIs
- Examples: `src/taskpane/providers/courtListenerProvider.ts` (free, no auth), `src/taskpane/providers/westlawProvider.ts`, `lexisNexisProvider.ts`, `bloombergLawProvider.ts` (enterprise, auth required)
- Pattern: Interface (`providers/types.ts`) + registry (`providers/registry.ts`) + self-registering built-ins (`providers/index.ts`)
- Purpose: Encapsulates one edition's citation-formatting rules as a pluggable, swappable object
- Examples: `src/taskpane/bluebook/edition20th.ts`, `edition21st.ts`, `edition22nd.ts`
- Pattern: Interface (`bluebook/types.ts`) + registry (`bluebook/registry.ts`) + self-registering editions (`bluebook/index.ts`)
- Purpose: Shared OAuth2 client-credentials handshake and in-memory-only credential storage for paid providers
- Examples: subclassed by `westlawProvider.ts`, `lexisNexisProvider.ts`, `bloombergLawProvider.ts`
- Pattern: Template method (`authenticate()` calls abstract `verifyCredentials()`)

## Entry Points

- Location: `src/taskpane/taskpane.html` + `src/taskpane/taskpane.ts` (imports `./word`)
- Triggers: Office `ShowTaskpane` ribbon action defined in `manifest.xml`
- Responsibilities: Full UI — all hyperlinking, Bluebook checking, hallucination checking, opinion-text embedding workflows
- Location: `src/commands/commands.html` + `src/commands/commands.ts` (imports `./commands.word`)
- Triggers: Ribbon `ExecuteFunction` actions defined in `manifest.xml`
- Responsibilities: Lightweight commands not requiring the full task pane UI

## Architectural Constraints

- **Threading:** Single-threaded browser/webview context (standard for Office.js task panes); all document mutations are batched through `Word.run(async (context) => { ... context.sync() })` calls.
- **Global state:** Module-level mutable state lives entirely in `src/taskpane/word.ts` (selected citation map, active provider, last check results) — no external state library.
- **Circular imports:** None detected; dependency direction is strictly `word.ts` → `providers/` and `word.ts` → `bluebook/`, with no reverse imports.
- **No backend:** All network calls are made directly from the browser to third-party APIs (CourtListener, or firm-supplied enterprise API base URLs); OpenClerk operates no server component.

## Anti-Patterns

### Monolithic controller file

### Direct OOXML manipulation alongside Office.js calls

## Error Handling

- `setStatus(message)` (`src/taskpane/word.ts:1394`) writes user-visible status/progress text into the task pane
- Provider `authenticate()`/`verifyCredentials()` throw `Error` with human-readable messages that are surfaced directly in the UI (`providers/base.ts:26-41`)
- Bounded reads guard against unbounded memory use, e.g. `readZipEntryWithLimit` (`word.ts:1452`)

## Cross-Cutting Concerns

<!-- GSD:architecture-end -->

<!-- GSD:skills-start source:skills/ -->

## Project Skills

No project skills found. Add skills to any of: `.claude/skills/`, `.agents/skills/`, `.cursor/skills/`, `.github/skills/`, or `.codex/skills/` with a `SKILL.md` index file.
<!-- GSD:skills-end -->

<!-- GSD:workflow-start source:GSD defaults -->

## GSD Workflow Enforcement

Before using Edit, Write, or other file-changing tools, start work through a GSD command so planning artifacts and execution context stay in sync.

Use these entry points:

- `/gsd-quick` for small fixes, doc updates, and ad-hoc tasks
- `/gsd-debug` for investigation and bug fixing
- `/gsd-execute-phase` for planned phase work

Do not make direct repo edits outside a GSD workflow unless the user explicitly asks to bypass it.
<!-- GSD:workflow-end -->

<!-- GSD:profile-start -->

## Developer Profile

> Profile not yet configured. Run `/gsd-profile-user` to generate your developer profile.
> This section is managed by `generate-claude-profile` -- do not edit manually.
<!-- GSD:profile-end -->
