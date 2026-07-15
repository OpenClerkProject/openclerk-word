# Architecture Research

**Domain:** Office.js task-pane add-in (vanilla TypeScript, no framework) — decomposing a monolithic workflow-controller file
**Researched:** 2026-07-15
**Confidence:** HIGH (codebase-specific findings — read `word.ts`, `providers/base.ts`, all three enterprise provider files, the registry, and the full `tests/` import graph directly) / MEDIUM (general prior art — Microsoft's official Office Add-ins docs have no specific guidance on this; recommendations below lean on general software-architecture principles, which are well-established but not Office.js-specific)

This is a targeted design document for one milestone task, not a broad ecosystem survey — see `.planning/codebase/ARCHITECTURE.md` and `STRUCTURE.md` for the current-state baseline this builds on.

## Standard Architecture

### System Overview (target state)

```text
┌──────────────────────────────────────────────────────────────────────┐
│  src/taskpane/word.ts  (composition root, target ~150-250 lines)      │
│  Office.onReady: DOM lookups for tab/select elements → calls each     │
│  workflow module's init(); owns ONLY tab-switching (setActiveTab)     │
│  and cross-tab refresh dispatch. No workflow logic, no shared state.  │
└───────────────┬─────────────────────────────────────────────────────┘
                │ imports + calls init()/refresh() (one-way only)
                ▼
┌──────────────────────────────────────────────────────────────────────┐
│  src/taskpane/workflows/   (NEW — one file per feature, self-wiring)  │
│  ┌────────────────┐ ┌───────────────┐ ┌──────────────────┐          │
│  │ hyperlinking.ts │ │ onlineLookup.ts│ │ bluebookCheck.ts  │          │
│  │ (file+parenth.) │ │ (provider auth │ │                   │          │
│  │                 │ │  + API apply)  │ │                   │          │
│  └────────┬────────┘ └───────┬────────┘ └─────────┬─────────┘          │
│  ┌────────┴──────────┐ ┌─────┴─────────────┐       │                  │
│  │ hallucinationCheck.ts │ embedOpinionText.ts│       │                  │
│  └────────────────────┘ └────────────────────┘       │                  │
└───────────────┬───────────────────────────────────────┬───────────────┘
                │ imports only (never re-imported by word.ts or shared/)  │
                ▼                                                        ▼
┌───────────────────────────────┐   ┌─────────────────────────────────────┐
│ src/taskpane/shared/  (NEW)    │   │ src/taskpane/providers/ (existing)   │
│ documentActions.ts             │   │ src/taskpane/bluebook/  (existing)   │
│ (applyHyperlinkToItem,         │   │ Both untouched by this split, both   │
│  goToCitationInDocument,       │   │ already leaf plugin systems — no     │
│  MAX_SEARCH_TEXT_LENGTH)       │   │ change to registry/self-registration │
│ sourceDocument.ts              │   └─────────────────────────────────────┘
│ (OOXML .docx parsing, JSZip)   │
│ statusBar.ts (setStatus)       │
└───────────────┬─────────────────┘
                │
                ▼
┌──────────────────────────────────────────────────────────────────────┐
│                    Office.js Word API (Word.run / OOXML)               │
└──────────────────────────────────────────────────────────────────────┘
```

Dependency direction is strictly top-to-bottom: `word.ts → workflows/* → shared/* , providers/, bluebook/`. Nothing below imports anything above it. This preserves the "no circular imports" constraint already documented in `.planning/codebase/ARCHITECTURE.md`.

### Component Responsibilities

| Component | Responsibility | Typical Implementation |
|-----------|----------------|-------------------------|
| `word.ts` (shrunk) | Composition root: look up the ~20 DOM elements, wire tab-switching, call each workflow's `init()`/`refresh()` | Office.onReady body only; no business logic, no OOXML, no module-level workflow state |
| `workflows/<feature>.ts` | One file per user-facing workflow named in PROJECT.md (hyperlinking, Bluebook checking, hallucination checking, opinion-text embedding), plus one for the online-provider-driven hyperlink apply flow | Owns its own DOM element lookups, its own event listeners (self-wiring `init()`), its own module-level state, its own render/status functions |
| `shared/documentActions.ts` | Word.js-coupled actions used by more than one workflow (insert a hyperlink into a Range, search-and-select a citation to "jump to" it) plus the shared `MAX_SEARCH_TEXT_LENGTH` Word `Range.search()` constraint | Leaf module — imported by workflows, imports nothing from workflows/word.ts |
| `shared/sourceDocument.ts` | OOXML `.docx` parsing (JSZip unzip, relationship parsing, zip-bomb size guards) — already flagged as a self-contained candidate for extraction in the existing `ARCHITECTURE.md` anti-patterns section | Pure functions except `File`/`JSZip` I/O; no DOM, no `Word.run` — the easiest and highest-value module to unit-test first |
| `shared/statusBar.ts` | `setStatus()` — single DOM status element write, used by every workflow | Trivial, but centralizing it avoids each workflow reaching into `word.ts` for it |
| `providers/`, `bluebook/` | Unchanged — already the plugin-registry precedent this split follows | No changes needed beyond the provider-dedup work described below |

## Recommended Project Structure

```
src/taskpane/
├── taskpane.ts                    # unchanged — entry point, imports ./word
├── word.ts                        # SHRINKS from 1538 → ~150-250 lines: Office.onReady,
│                                   #   DOM lookups, tab-switching, calls workflow init()s
├── workflows/                     # NEW — one file per feature-scoped workflow
│   ├── hyperlinking.ts            #   file-source + parenthetical hyperlink apply/remove
│   ├── onlineLookup.ts            #   provider select/connect/disconnect + apply-via-API
│   ├── bluebookCheck.ts           #   edition select, check, render, "report as wrong" link
│   ├── hallucinationCheck.ts      #   provider ordering UI, check, render
│   └── embedOpinionText.ts        #   pincite lookup, comment embed/remove, render
├── shared/                        # NEW — leaf helpers used across >1 workflow
│   ├── documentActions.ts         #   applyHyperlinkToItem, goToCitationInDocument,
│   │                               #   MAX_SEARCH_TEXT_LENGTH
│   ├── sourceDocument.ts          #   parseSourceDocument, parseRelationships,
│   │                               #   getElementText, readZipEntryWithLimit + the three
│   │                               #   zip-bomb size-limit constants
│   └── statusBar.ts               #   setStatus
├── utils.ts                       # unchanged — pure string/DOM helpers, no Word.js dep
├── providers/                     # unchanged directory shape; base.ts gains a generic
│   │                               #   config-driven enterprise-provider class (see below)
│   ├── types.ts / registry.ts / index.ts   # unchanged
│   ├── base.ts                    # + new `GenericEnterpriseCitationProvider` class
│   ├── westlawProvider.ts         # shrinks to a ~15-line config + constructor call
│   ├── lexisNexisProvider.ts      # shrinks to a ~15-line config + constructor call
│   ├── bloombergLawProvider.ts    # shrinks to a ~15-line config + constructor call
│   └── courtListenerProvider.ts / usptoPatentCenterProvider.ts / citationParser.ts / ...  # unchanged
└── bluebook/                      # unchanged
```

### Structure Rationale

- **`workflows/` groups by feature, not by layer.** A layer-based split (e.g. `dom/`, `logic/`, `parsers/`) would scatter each workflow's DOM wiring, business logic, and render function across three directories — exactly the "one change touches many files" problem the milestone is trying to fix, and it breaks the codebase's own precedent (`providers/`, `bluebook/` are both organized by *what integration/edition they are*, not by *what kind of code they contain*). Feature-scoped files keep "everything about hallucination checking" in one place, matching PROJECT.md's own phrasing of the split ("hyperlinking, Bluebook checking, hallucination checking, opinion-text embedding").
- **`shared/` is new and deliberately thin.** Only code genuinely used by more than one workflow goes here (`applyHyperlinkToItem` is called from both `hyperlinking.ts` and `onlineLookup.ts`; `goToCitationInDocument` is called from `bluebookCheck.ts`, `hallucinationCheck.ts`, and `embedOpinionText.ts`'s render functions). Putting these in `shared/` rather than leaving them in `word.ts` (which would force every workflow to import from the composition root) is what makes the dependency graph acyclic.
- **`sourceDocument.ts` is split out of `shared/` conceptually but lives there** because, while today only `hyperlinking.ts` uses it, it has zero DOM/Word.run coupling and is the highest-value first extraction (see Migration Order below) — keeping it beside `documentActions.ts` avoids inventing a fourth top-level directory for one file.
- **`utils.ts` stays as-is.** It currently has no Office.js/Word.js imports (confirmed by reading it) and existing tests (`tests/utils.test.ts`) depend on that. Don't add `Word.run`-coupled helpers to it — that's what `shared/documentActions.ts` is for. Keeping the pure-string-helpers module pure also keeps it usable from `tools/pdf-extract/` or any future non-Office context without dragging in Office typings.

## Architectural Patterns

### Pattern 1: Self-wiring feature module (`init()` composition)

**What:** Each `workflows/<feature>.ts` exports an `init(): void` function that does its own `document.getElementById` lookups and `addEventListener` calls for only the elements that workflow owns. `word.ts`'s `Office.onReady` becomes a flat list of `hyperlinking.init(); onlineLookup.init(); bluebookCheck.init(); hallucinationCheck.init(); embedOpinionText.init();` plus the handful of DOM lookups/listeners that are genuinely cross-cutting (the tab `<select>` itself).
**When to use:** Any vanilla-JS/TS controller where DOM wiring for logically distinct features currently lives in one giant `Office.onReady`/`DOMContentLoaded` block. This is the standard "feature module" refactor for framework-less UI code — the vanilla-JS analogue of a React component owning its own `useEffect`/event bindings instead of a parent wiring child behavior for it.
**Trade-offs:** Avoids passing 15+ individual `HTMLElement | null` references as function parameters between `word.ts` and workflow modules (brittle, hard to review). Costs a small amount of duplication in "look up `document.getElementById(...)`, null-check, addEventListener" boilerplate per module — acceptable, matches the existing per-vendor-file duplication precedent already accepted in `providers/`.

**Example:**
```typescript
// workflows/bluebookCheck.ts
export function init(): void {
  const editionSelect = document.getElementById("bluebook-edition-select") as HTMLSelectElement | null;
  const checkButton = document.getElementById("check-bluebook-citations") as HTMLButtonElement | null;
  const flaggedOnlyCheckbox = document.getElementById("bluebook-show-flagged-only") as HTMLInputElement | null;

  editionSelect?.addEventListener("change", () => { renderEditionDescription(); invalidateResults(); });
  checkButton?.addEventListener("click", checkBluebookCitations);
  flaggedOnlyCheckbox?.addEventListener("change", () => { showFlaggedOnly = flaggedOnlyCheckbox.checked; renderResults(); });

  populateEditionSelect();
  renderResults();
}

// Called only from another workflow's tab-switch hook, never from a reverse import of word.ts:
export function refresh(): void { /* re-render without losing state, e.g. on tab activation */ }
```

### Pattern 2: Colocated module-level state (no shared globals)

**What:** Each workflow module keeps its own module-level mutable state (`sourceCitationMap`/`parentheticalEntries`/`hyperlinkScope`/`caseLawSource` → `hyperlinking.ts`; `lastBluebookResults`/`bluebookShowFlaggedOnly` → `bluebookCheck.ts`; `hallucinationProviderOrder` → `hallucinationCheck.ts`). Nothing outside a workflow module reaches into its state directly.
**When to use:** Whenever splitting a monolith that currently has several unrelated pieces of `let` state living at the top of one file. This is the root cause of the circular-import risk in this specific refactor: if `word.ts` keeps owning the state and workflow modules need to read/write it, you get `word.ts → workflows/*` (to wire clicks) **and** `workflows/* → word.ts` (to touch state) simultaneously — a cycle. Moving state to live with the workflow that owns it collapses that to one direction.
**Trade-offs:** Cross-workflow reads become explicit function calls instead of shared variable access — e.g., if `onlineLookup.ts` ever needed to know whether a source file was already loaded, it would call an exported `hyperlinking.hasSourceCitations()` rather than reading `sourceCitationMap` directly. Today's codebase doesn't actually need any such cross-workflow reads (verified: each of the four workflows' state is only ever read/written within its own function block in the current `word.ts`), so this pattern costs nothing here and is purely upside.

### Pattern 3: Cross-tab refresh via exported hook, not shared render function

**What:** `setActiveTab()` in `word.ts` currently calls `renderHallucinationProviderList()` and `renderEmbedTextProviderStatus()` directly when switching to those tabs (so a provider connected on another tab shows up-to-date status). After the split, `word.ts` instead calls each workflow's exported `refresh()` (Pattern 1's example above) from inside `setActiveTab`.
**When to use:** Any time the composition root needs to trigger a re-render in a feature module it doesn't own the internals of.
**Trade-offs:** None significant — this is just Pattern 1 applied to the one piece of cross-workflow coordination that already exists in the code today.

### Pattern 4: Config-driven subclass instead of copy-pasted subclasses (provider dedup)

**What:** `WestlawProvider`, `LexisNexisProvider`, and `BloombergLawProvider` are structurally identical (same OAuth2 client-credentials `verifyCredentials`, same `lookupCitation` body, same `credentialFields` shape) and differ only in six data values: `id`, `name`, `description`, `apiBaseUrlPlaceholder`, `TOKEN_PATH`, `SEARCH_PATH`. This is the textbook "duplicate code across sibling subclasses" smell — the standard fix (Refactoring Guru's "Duplicate Code" catalog entry, and the Factory-with-config pattern) is to pull the shared behavior up into one concrete class parameterized by a config object, not to keep three copies of the logic in sync by hand.
**When to use:** Sibling subclasses of the same abstract base whose *only* differences are literal values (not behavior). If any of the three providers ever needs genuinely different request/response shaping (e.g. a different auth grant type, or nested JSON response parsing), split it back out at that point — don't force a future divergent provider through the generic class.
**Trade-offs:** One indirection layer to read through (config → generic class → instance), but eliminates the risk this milestone is explicitly trying to close (a bug fixed in `westlawProvider.ts` not being copied into the other two — this already nearly happened once: all three files are byte-identical except for the six values, so a fix applied to only one is an easy mistake). Crucially, **this does not touch the self-registration pattern**: `providers/index.ts` still does `citationProviderRegistry.register(new WestlawProvider())` etc., unchanged.

**Example:**
```typescript
// providers/base.ts — new addition alongside EnterpriseCitationProvider
export interface EnterpriseOAuth2ProviderConfig {
  id: string;
  name: string;
  description: string;
  apiBaseUrlPlaceholder: string;
  tokenPath: string;   // e.g. "/oauth/token"
  searchPath: string;  // e.g. "/content/search/v1/cases"
}

export class GenericEnterpriseCitationProvider extends EnterpriseCitationProvider {
  readonly id: string;
  readonly name: string;
  readonly description: string;
  readonly credentialFields: ProviderCredentialField[];
  private readonly tokenPath: string;
  private readonly searchPath: string;
  private accessToken: string | null = null;

  constructor(config: EnterpriseOAuth2ProviderConfig) {
    super();
    this.id = config.id;
    this.name = config.name;
    this.description = config.description;
    this.tokenPath = config.tokenPath;
    this.searchPath = config.searchPath;
    this.credentialFields = [
      { key: "apiBaseUrl", label: `API base URL (from your ${config.name} contract)`, type: "text", placeholder: config.apiBaseUrlPlaceholder },
      { key: "clientId", label: "Client ID", type: "text" },
      { key: "clientSecret", label: "Client secret", type: "password" },
    ];
  }

  protected async verifyCredentials(credentials: Record<string, string>): Promise<void> {
    const baseUrl = trimTrailingSlash(credentials.apiBaseUrl);
    this.accessToken = await fetchClientCredentialsToken(`${baseUrl}${this.tokenPath}`, credentials.clientId, credentials.clientSecret);
  }

  signOut(): void {
    super.signOut();
    this.accessToken = null;
  }

  async lookupCitation(citation: ParsedCitation): Promise<CitationMatch | null> {
    // identical body to the current three files, using this.searchPath / this.accessToken
  }
}

// providers/westlawProvider.ts — shrinks to:
import { GenericEnterpriseCitationProvider } from "./base";

export class WestlawProvider extends GenericEnterpriseCitationProvider {
  constructor() {
    super({
      id: "westlaw",
      name: "Westlaw",
      description: "Looks up citations through your organization's Westlaw / Thomson Reuters API subscription. Requires the API base URL and client credentials issued under your firm's Westlaw contract.",
      apiBaseUrlPlaceholder: "https://your-tenant.api.thomsonreuters.com",
      tokenPath: "/oauth/token",
      searchPath: "/content/search/v1/cases",
    });
  }
}
```
`lexisNexisProvider.ts` and `bloombergLawProvider.ts` become the same ~15-line shape with their own config. `id`/`name`/`description`/`credentialFields` remain readable as plain properties on the instance (TypeScript allows an `abstract readonly` property to be satisfied by a concrete `readonly` property assigned in the constructor), so nothing that reads `provider.id`, `provider instanceof EnterpriseCitationProvider`, etc. — including `tests/providers.test.ts` and the registry — needs to change.

**Alternative considered and rejected:** collapsing all three into config *data* (no per-vendor files at all — e.g. a `PROVIDER_CONFIGS` array that `index.ts` loops over to `new GenericEnterpriseCitationProvider(config)` directly). This is more DRY but breaks the documented "add a new provider = add a new `<vendorName>Provider.ts` file" onboarding convention in `.planning/codebase/STRUCTURE.md`, and removes the natural place to put vendor-specific overrides if one provider later needs a small behavioral tweak. Not recommended unless the team also wants to change that documented convention.

## Data Flow

### Event-wiring flow (composition)

```
Office.onReady (word.ts)
    ↓ calls
hyperlinking.init() → wires its own DOM elements/listeners, sets up its own state
onlineLookup.init()      "
bluebookCheck.init()     "
hallucinationCheck.init()"
embedOpinionText.init()  "
    ↓
word.ts wires ONLY the tab <select> + workflow <select> listeners, calling
setActiveTab() → which calls each workflow's exported refresh() for the
newly-active tab only (Pattern 3) — never reads/writes workflow-internal state
```

### Runtime request flow (unchanged from current system, just re-homed)

1. User clicks a workflow's action button → the click handler registered by that workflow's own `init()` fires (was: registered in `word.ts`'s single `Office.onReady`).
2. The handler calls `Word.run(...)`, using `shared/documentActions.ts` helpers (`applyHyperlinkToItem`, `goToCitationInDocument`) where the action is shared across workflows, or workflow-local helpers where it isn't.
3. The handler calls into `providers/` or `bluebook/` exactly as it does today — these two directories are untouched by the `word.ts` split.
4. The handler updates its own module-level state and calls its own local `render*()` function to update the DOM.

**State Management:** Still entirely in-memory, still reset on task-pane reload — no change to the existing "no external state library, no persistence" architecture. The only change is *where* each piece of state lives (colocated with its workflow, per Pattern 2, instead of all in `word.ts`).

## Migration Order

**Key finding that shapes this order:** grepping every file under `tests/` for imports from `src/taskpane/` shows **zero tests currently import anything from `word.ts`** — every existing test targets `utils.ts`, `providers/*`, or `bluebook/*`. There is also no Word.js/Office mock anywhere in the test suite (`jest.mock`/`global.Word`/`global.Office` all return no matches). This means:
- Splitting `word.ts` cannot "break" any existing test that exercises `word.ts` directly, because none exist.
- The real risk is a *silent behavior regression* in Word.run-coupled code that has never been automatically tested, so the safest sequence prioritizes (a) extracting the pure, currently-untestable-only-because-it's-buried logic first and giving it real tests immediately, and (b) treating every DOM/Word.run extraction as a **mechanical move, not a rewrite** — same function bodies, same call order, same batching of `context.sync()` calls (several workflows have hand-tuned sync batching with comments explaining *why*, e.g. "50-100+ citations... taskpane visibly freeze" — preserve these verbatim) — verified by a manual sideload smoke test after each step, since that is the only verification method this behavior has ever had.

Recommended order, each step its own commit/PR, `tsc`/`jest`/`webpack build` green before moving to the next:

1. **Provider dedup first** (`providers/base.ts` + the three enterprise provider files). Smallest, self-contained, already partially covered by `tests/providers.test.ts` (which imports `LexisNexisProvider` directly and exercises the shared `authenticate()`/`EnterpriseCitationProvider` flow). While here, add the parametrized test coverage `tests/providers.test.ts` is currently missing for `WestlawProvider`/`BloombergLawProvider` (today only `LexisNexisProvider` is asserted on) — collapsing to one generic class is the natural moment to write one test that runs against all three configs instead of leaving two untested. Doing this first also validates the "config-driven subclass" pattern in isolation before the larger, riskier `word.ts` split.
2. **Extract `shared/sourceDocument.ts`** (`parseSourceDocument`, `parseRelationships`, `getElementText`, `readZipEntryWithLimit`, the three zip-bomb size constants). Zero DOM coupling, zero `Word.run` coupling, pure functions over `File`/strings — the highest-value first extraction from `word.ts` because it's the easiest to give real unit tests to (fixture `.docx` zip buffers), establishing a safety net where currently there is none. This is also exactly the module the existing `.planning/codebase/ARCHITECTURE.md` anti-patterns section already flagged as a good extraction candidate.
3. **Extract `shared/documentActions.ts`** (`applyHyperlinkToItem`, `goToCitationInDocument`, `MAX_SEARCH_TEXT_LENGTH`) and **`shared/statusBar.ts`** (`setStatus`). Small, but extracting these *before* the workflow modules avoids rework — every workflow extraction after this point can import from `shared/` on day one instead of temporarily importing from `word.ts` and needing a follow-up cleanup.
4. **Extract the two lowest-risk, most self-contained workflows next:** `workflows/bluebookCheck.ts` and `workflows/embedOpinionText.ts`. Bluebook check's core rule-matching is already covered by `tests/bluebook.test.ts` (even though `checkBluebookCitations` itself isn't — the risk surface is smaller). These two extractions prove out the `init()`/state-colocation pattern (Patterns 1-3) on workflows with no reverse-dependency complications before touching anything load-bearing for the product's core trust claim.
5. **Extract `workflows/hyperlinking.ts`** (file-source + parenthetical apply/remove) and **`workflows/onlineLookup.ts`** (provider connect/disconnect + apply-via-API) together, since `applyHyperlinkToItem` is shared between them and it's easier to verify the shared-helper boundary (`shared/documentActions.ts`) is correct when both consumers move at once. These carry more risk than step 4 because of the hand-tuned `context.sync()` batching mentioned above — budget extra manual smoke-test time (apply/remove hyperlinks on a document with 50+ citations, not just a trivial one, to exercise the batching path).
6. **Extract `workflows/hallucinationCheck.ts` last.** This is deliberately saved for last: it's the workflow most directly tied to PROJECT.md's stated core value ("a hallucination check must never falsely report a fabricated citation as verified"). By this point the `init()`/state-colocation/shared-helper pattern is proven on five prior extractions, so this final one is the most mechanical, lowest-novelty step — exactly what you want for the highest-stakes code. Before merging, manually re-verify the `verifiedVia`/`skippedProviders`/`rateLimitedProviders` branching still distinguishes "not connected," "rate-limited," and "genuinely not found" the same way it does today (these three states must never collapse into each other — that's the correctness property the milestone constraints call out explicitly).
7. **Final step: trim `word.ts`** to just `Office.onReady`, the DOM lookups for tab/workflow `<select>` elements, `setActiveTab`/`TAB_PANEL_IDS`/`updateManageHyperlinksVisibility`, and the five `init()` calls. Confirm final line count lands in the ~150-250 line range estimated above; if it doesn't, something that should have moved into a workflow or `shared/` module didn't.

Steps 2-3 can be combined into one PR if preferred (both are small, low-risk, and step 3 has no dependents yet); steps 4-6 should each stay separate so a regression is bisectable to one workflow.

## Anti-Patterns

### Anti-Pattern 1: Extracting workflow logic but leaving DOM wiring in `word.ts`

**What people do:** Move `checkBluebookCitations`, `renderBluebookResults`, etc. into `workflows/bluebookCheck.ts`, but leave the `document.getElementById(...)` + `addEventListener(...)` calls for those elements in `word.ts`'s `Office.onReady`, with `word.ts` importing the handler functions to pass to `addEventListener`.
**Why it's wrong:** This is a halfway split — `word.ts` still has to know about every DOM element for every workflow (no line-count reduction of the part that's actually hard to read), and it re-introduces the exact cross-file coupling this refactor is meant to remove (a change to `bluebookCheck.ts`'s element IDs still requires a matching edit in `word.ts`).
**Do this instead:** Pattern 1 — each workflow module owns its own element lookups and its own `addEventListener` calls behind a no-argument `init()` that `word.ts` simply calls.

### Anti-Pattern 2: Shared mutable state re-exported from `word.ts`

**What people do:** Keep `sourceCitationMap`, `lastBluebookResults`, etc. declared in `word.ts` and `export` them so the new workflow modules can `import { sourceCitationMap } from "../word"`.
**Why it's wrong:** Creates the exact circular-import risk called out above (`word.ts → workflows/*` to wire clicks, `workflows/* → word.ts` to touch state) and defeats the purpose of the split — the state remains globally shared and any workflow can still mutate any other workflow's data.
**Do this instead:** Pattern 2 — state lives inside the workflow module that owns it; if a future workflow genuinely needs to read another workflow's state, expose a narrow getter function from that workflow, not the raw variable.

### Anti-Pattern 3: "Big bang" single-PR extraction of all five workflows at once

**What people do:** Do the whole `word.ts` split in one commit/PR since it's "all mechanical moves anyway."
**Why it's wrong:** Given there is no automated test coverage for any of this code today (confirmed above), a single large diff makes it very hard to isolate which extracted workflow introduced a regression if the manual smoke test catches one problem after the fact — you'd be re-reviewing all five workflows' diffs to find it.
**Do this instead:** The step-by-step order above, one (or at most two closely-related) workflow(s) per PR, with a manual sideload smoke test of that specific workflow before moving to the next.

### Anti-Pattern 4: Hand-copying a bug fix across the three provider files (the status quo)

**What people do:** Fix a bug in `westlawProvider.ts` (e.g., the auth-failure error message, or a change to how the search request body is built) and manually apply the identical edit to `lexisNexisProvider.ts` and `bloombergLawProvider.ts`.
**Why it's wrong:** This is the exact situation the milestone is trying to close — three files that are byte-identical except for six values will drift the moment one fix is applied to only one of them; there is no compiler or test that will catch the other two falling out of sync.
**Do this instead:** Pattern 4 — one `GenericEnterpriseCitationProvider` in `base.ts`; a fix applied once, config objects for the three vendors carry only the genuinely-different values.

## Scaling Considerations

This is a client-side, single-user-session Office.js add-in with no backend — there is no "concurrent users" axis to scale against (each install runs entirely in one person's Word process). The relevant scaling axis here is **codebase growth (number of workflows / providers / rule editions over time)**, not traffic:

| Scale | Architecture Adjustments |
|-------|--------------------------|
| Current (5 workflows, 5 citation providers, 3 Bluebook editions) | The structure recommended above (flat `workflows/`, flat `providers/`, flat `bluebook/`) is sufficient — no further nesting needed. |
| Growth to ~10+ workflows | Consider whether some workflows naturally group under a shared sub-concern (e.g. all provider-auth-adjacent workflows) — but resist introducing a deeper directory hierarchy until there's a concrete file that doesn't fit the flat layout; premature nesting adds navigation cost without a matching benefit at this scale. |
| Growth to ~10+ citation providers (mentioned as a distinct future milestone: `usptoPatentCenterProvider.ts` and beyond) | The `GenericEnterpriseCitationProvider` pattern from this milestone should absorb any new OAuth2-client-credentials vendor with just a new config + thin file, per the existing "add a `<vendorName>Provider.ts`" convention — no registry changes needed, since `CitationProviderRegistry` (`providers/registry.ts`) is already a flat `Map` with no scaling concern at this size. |

### Scaling Priorities

1. **First bottleneck: file navigability, not runtime performance.** At this project's scale (single-session, client-side, no concurrent load), the thing that actually degrades over time is a human's ability to find and safely change one workflow's code — which is precisely what this milestone's split addresses. There is no performance bottleneck to design around here.
2. **Second bottleneck (further out): the `EnterpriseCitationProvider`/`GenericEnterpriseCitationProvider` base class assuming OAuth2 client-credentials for every enterprise vendor.** If a future provider (e.g. a firm's internal API) uses a different auth scheme, `EnterpriseCitationProvider` (the abstract base, unchanged by this milestone) already accommodates that via its abstract `verifyCredentials` — only vendors using the *generic* OAuth2 shape opt into `GenericEnterpriseCitationProvider`; a divergent vendor can still subclass `EnterpriseCitationProvider` directly, exactly as today.

## Integration Points

### External Services

Unchanged by this refactor — see `.planning/codebase/ARCHITECTURE.md`'s existing "Data Flow" section for the CourtListener/enterprise-provider lookup flow. No new integration surface is introduced by splitting `word.ts` or deduping the provider files.

### Internal Boundaries

| Boundary | Communication | Notes |
|----------|---------------|-------|
| `word.ts` ↔ `workflows/*` | Direct function calls: `word.ts` calls each workflow's exported `init()`/`refresh()`; workflows never call back into `word.ts` | One-way only — enforce via code review, not tooling (no ESLint import-boundary rule exists in this repo today per `.eslintrc.json`; adding one, e.g. `eslint-plugin-boundaries` or a simple `no-restricted-imports` rule forbidding `workflows/*` → `../word`, would be a reasonable low-cost addition alongside this milestone but isn't required for correctness) |
| `workflows/*` ↔ `shared/*` | Direct function/constant imports | One-way — `shared/*` must never import from `workflows/*` or `word.ts` |
| `workflows/*` ↔ `providers/`, `bluebook/` | Direct imports, exactly as `word.ts` does today | Unchanged — both are already leaf plugin systems with no internal dependencies on task-pane code |
| `providers/index.ts` ↔ `<vendorName>Provider.ts` | Self-registration via import side-effect (`citationProviderRegistry.register(new WestlawProvider())`) | Unchanged by the dedup — `GenericEnterpriseCitationProvider` is an implementation detail inside each provider file's constructor call; `index.ts` still instantiates by concrete class name, one line per vendor |

## Sources

- Direct reads of `src/taskpane/word.ts` (1538 lines, full file), `src/taskpane/providers/base.ts`, `westlawProvider.ts`, `lexisNexisProvider.ts`, `bloombergLawProvider.ts`, `providers/index.ts`, `providers/registry.ts`, `src/taskpane/utils.ts` — confidence HIGH (primary source, this session).
- Grep of all `tests/*.test.ts` import statements and of `Word.run`/`global.Word`/`global.Office`/`jest.mock` usage across `tests/` — confirmed zero test coverage of `word.ts` and zero Office.js mocking anywhere in the suite — confidence HIGH (direct verification, this session).
- `.planning/codebase/ARCHITECTURE.md` and `.planning/codebase/STRUCTURE.md` — baseline for existing registry/plugin conventions this proposal extends rather than replaces — confidence HIGH (project-authoritative).
- [Duplicate Code — Refactoring Guru](https://refactoring.guru/smells/duplicate-code) and general Factory/config-object pattern guidance — confidence MEDIUM (well-established general software-engineering practice, not Office.js-specific; no single canonical source, synthesized from general web search results on TypeScript factory/config-driven subclass deduplication).
- [Microsoft Learn — Task panes in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/design/task-pane-add-ins) and [OfficeDev/office-js-docs-pr — task-pane-add-ins.md](https://github.com/OfficeDev/office-js-docs-pr/blob/main/docs/design/task-pane-add-ins.md) — confidence LOW-MEDIUM for this specific question: these describe task-pane UX/manifest concerns, not internal code-organization patterns for large vanilla-TS controllers; no official Microsoft guidance on this specific decomposition problem was found. Flagged as a gap below rather than treated as authoritative for the module-boundary recommendations, which are instead derived directly from this codebase's own existing `providers/`/`bluebook/` precedent (HIGH confidence, primary source).

---
*Architecture research for: Office.js task-pane add-in monolith decomposition (OpenClerk `word.ts` split + provider dedup)*
*Researched: 2026-07-15*
