# Pitfalls Research

**Domain:** Office.js Word task-pane add-in refactor (monolithic controller split, provider dedup) + Microsoft Partner Center/AppSource submission
**Researched:** 2026-07-15
**Confidence:** MEDIUM (Microsoft Learn docs, `OfficeDev/office-js` GitHub issues, and Microsoft 365 Developer Blog's recurring "Top 5 AppSource validation errors" series cross-checked across multiple posts; verified against this repo's actual `word.ts`, `providers/`, `bluebook/`, `manifest.xml`, and `scripts/package-release*.js`)

## Critical Pitfalls

### Pitfall 1: Splitting `word.ts` scatters module-level mutable state, breaking implicit invariants

**What goes wrong:**
`word.ts` currently holds seven module-level `let`/mutable bindings at file scope — `sourceCitationMap`, `parentheticalEntries`, `hallucinationProviderOrder`, `hyperlinkScope`, `caseLawSource`, `lastBluebookResults`, `bluebookShowFlaggedOnly` (verified at `src/taskpane/word.ts:75-81`). These act as de facto shared session state read and written by multiple functions in the file. A naive split-by-feature refactor (hyperlinking / Bluebook checking / hallucination checking / opinion-text embedding, per PROJECT.md) either (a) duplicates this state per module, causing two modules to disagree about e.g. `hyperlinkScope`, or (b) creates a new shared-state module that every feature module imports, silently recreating the same "everything touches everything" coupling the split was meant to remove — just moved one file over.

**Why it happens:**
Module-level `let` state is easy to reach for when the code was one file (no import friction). When splitting, the state's true ownership (which feature actually needs to read/write each variable, and when) was never mapped, so it gets carried over structurally instead of being redesigned.

**How to avoid:**
Before splitting, inventory each module-level variable: which functions write it, which read it, and whether it represents (a) per-operation input that should be a function parameter/return value instead, or (b) genuine cross-cutting session state (e.g., `hyperlinkScope`/`caseLawSource` as user-selected settings) that belongs in one explicit state module with a narrow read/write API — not re-declared `let` in each feature module. Prefer passing state explicitly through function signatures over new shared mutable modules; if a shared state module is unavoidable, give it accessor functions (`getHyperlinkScope()`/`setHyperlinkScope()`) rather than exported `let` bindings, so every read/write site is greppable.

**Warning signs:**
- Two new modules both importing the same `let` variable from a "state" module and mutating it in different call orders.
- A feature module's exported function silently depends on another module having run first to populate `sourceCitationMap` or `lastBluebookResults`.
- Tests for one feature module needing to manually seed state that "belongs" to a different feature.

**Phase to address:**
word.ts split — do this inventory as the first step of the split, before moving any code.

---

### Pitfall 2: Word.js proxy objects passed across the new module boundaries become stale or throw `InvalidObjectPath`

**What goes wrong:**
Office.js Word objects (`Range`, `Paragraph`, `ContentControl`, etc.) obtained inside a `Word.run(context => ...)` callback are proxy objects: property reads/writes are queued and only reflect the document after `context.sync()`. If the word.ts split moves citation-matching (which locates a `Range`/`Paragraph` in the document) into one module and hyperlink-insertion (which acts on that range) into another, and the object is handed from one module's `Word.run` call to a separate `Word.run` call (or held across an intervening `sync()` outside the same sequential batch), Office.js throws `InvalidObjectPath` unless the object was explicitly tracked via `context.trackedObjects.add(range)` (or `.track()`) when first created — and its parent collection tracked too if it came from one.

**Why it happens:**
This is invisible in a monolithic file where the whole hyperlink-insertion flow (find range → insert hyperlink → sync) typically runs inside a single `Word.run`/`context.sync()` sequence, so tracking was never needed. Splitting by feature creates a natural temptation to also split by "one `Word.run` per module," which is exactly the pattern that breaks proxy-object validity.

**How to avoid:**
When designing the split's module boundaries, decide explicitly whether a `Range`/`Paragraph` reference ever needs to survive past the `Word.run` call that created it. If it does (e.g., a hallucination-check module locates suspect citations, then a separate insertion module later marks them), either (a) keep both operations inside the same `Word.run`/single `context` so the object stays valid without tracking, or (b) explicitly call `context.trackedObjects.add()` on creation and `.untrack()` when done, and add a code comment explaining why — don't let this become implicit/undocumented. Prefer re-locating the range by a stable identifier (paragraph index, search text, bookmark) in the second module over passing a live proxy object across the boundary at all; it's more robust to document mutation between the two calls.

**Warning signs:**
- A newly split module receives a `Word.Range`/`Word.Paragraph` object as a function parameter from another module (not created via its own `Word.run`).
- `InvalidObjectPath` or "the object is invalid" errors appearing on manual QA of hyperlink/comment insertion after the split, especially only on the second document edit in a session (first works, later ones fail because the object went stale).

**Phase to address:**
word.ts split — establish a rule during module-boundary design: pass identifiers/data, not live Office.js proxy objects, across module boundaries. Verify with manual QA of a multi-step flow (find citations → check hallucinations → insert hyperlinks in sequence) after the split, not just each feature in isolation.

---

### Pitfall 3: Self-registering provider/rule-set modules silently vanish if imported only for their type, not their side effect

**What goes wrong:**
This repo already uses a self-registration pattern for both citation providers and Bluebook rule sets: `providers/index.ts` and `bluebook/registry.ts`'s companion index import each concrete provider/rule-set class and call `registry.register(new X())` as a top-level side effect (verified: `src/taskpane/providers/index.ts:8-12` registers `CourtListenerProvider`, `LexisNexisProvider`, `WestlawProvider`, `BloombergLawProvider`, `UsptoPatentCenterProvider`). The dedup work (merging the three near-identical enterprise providers into shared `providers/base.ts` logic) and the broader duplication audit against `openclerk-core` both risk restructuring these imports. If a refactor changes an import from `import "./providers"` (side-effect import, runs registration) to `import type { ... } from "./providers"` or a lazy/dynamic import, or if a future `"sideEffects": false` is added to `package.json` for webpack bundle-size optimization (not currently present, but a plausible follow-up given IE11/bundle-size is already flagged in CONCERNS.md), webpack's tree-shaking can drop the "unused" registration modules entirely from the production bundle. The provider silently disappears from the UI/registry with no compile error — only a runtime "provider not found" or an empty provider list.

**Why it happens:**
Side-effect-only imports look like dead code to both humans and bundlers ("nothing exports symbols I use, why keep it?"). This pattern trades that visibility for the plugin-registry benefit PROJECT.md's constraints call for, so it needs a compensating safeguard.

**How to avoid:**
Keep `package.json` free of `"sideEffects": false` (or if added for other reasons, explicitly whitelist `providers/index.ts` and `bluebook/registry.ts`'s registration module in the `sideEffects` array). During the dedup and duplication-audit work, if any provider files move or the registration import chain changes, add/keep a test that asserts `citationProviderRegistry.list()` and `bluebookRuleSetRegistry.list()` return the full expected set of IDs after a fresh module load — this catches both accidental tree-shaking and accidental omission from a refactored index file. `tests/providers.test.ts` already exists; confirm it asserts registry completeness, not just individual provider behavior.

**Warning signs:**
- A provider or rule set that works in `npm start` (dev, unminified, all imports live) but is missing in the packaged production build (`npm run build` + smoke test) — a classic tree-shaking-only symptom.
- Adding `"sideEffects": false` to `package.json` in an unrelated bundle-size PR.
- Refactored `providers/index.ts`/`bluebook` index file changed from concrete imports to a barrel re-export without the registration calls preserved.

**Phase to address:**
Provider dedup (`providers/base.ts` consolidation) and the openclerk-core duplication audit — both touch the files where self-registration happens. Verify with a registry-completeness test plus one production-build smoke test.

---

### Pitfall 4: Escaping/URL-validation guards stay "manual discipline" after the split instead of becoming compiler-enforced

**What goes wrong:**
CONCERNS.md already flags this: `escapeHtml`/`isSafeHyperlinkUrl` (`src/taskpane/utils.ts:43-63`) must be called explicitly at every Office.js insertion call site (`applyHyperlinkToItem`/`insertHyperlink`/`insertHtml`) in `word.ts`, and there's no compiler-enforced wrapper preventing a new call site from skipping the guard. PROJECT.md's Active requirements list "harden hyperlink/HTML insertion so escapeHtml/isSafeHyperlinkUrl cannot be skipped at a new insertion call site" as its own item — separate from the word.ts split — but if the split happens first and the hardening happens second (or not at all, if scope gets compressed), the split creates more insertion call sites spread across more files, each one a fresh opportunity to forget the guard, with no build-time signal.

**Why it happens:**
Splitting a file by feature naturally multiplies the number of places that "do the same category of thing" (insert HTML/hyperlinks), because each feature module (hyperlinking, hallucination-flagging comments, opinion-text embedding) independently needs to insert content into the document.

**How to avoid:**
Sequence the hardening work before or atomically with the split: introduce a single typed insertion helper (e.g., a `SafeHtml`/`SafeUrl` branded type, or a function like `insertEscapedHyperlink(context, range, url, text)` that internally calls `escapeHtml`/`isSafeHyperlinkUrl` and is the *only* function allowed to call the underlying Office.js insertion API) in `utils.ts` or a new `wordInsertion.ts` module, then have every split-out feature module call only that helper — never the raw Office.js insertion methods directly. This converts "remember to call the guard" into "there is no other way to insert content," which survives the split by construction rather than by discipline.

**Warning signs:**
- Any split-out module directly calling `Office`/`Word` insertion APIs (`insertHtml`, `insertText` with HTML mode, hyperlink field insertion) instead of a shared guarded helper.
- Code review of the split PR finding more than one place that calls `escapeHtml`/`isSafeHyperlinkUrl` inline rather than through a single choke point.
- `tests/utils.test.ts` still only testing `isSafeHyperlinkUrl` in isolation, with no test asserting that each feature module's insertion path is guarded (this test gap is already noted in CONCERNS.md).

**Phase to address:**
Both the escapeHtml/isSafeHyperlinkUrl hardening item and the word.ts split — do the hardening (single guarded insertion helper) first, or in the same phase as the split, so newly created modules are built against the safe API from day one rather than needing a follow-up migration.

---

### Pitfall 5: Splitting `word.ts` changes `Office.onReady`/event-handler registration order, breaking ribbon-button wiring

**What goes wrong:**
Office.js task panes call `Office.onReady(...)` once at startup, and ribbon/task-pane button handlers are typically wired inside or after that callback via `Office.actions.associate(...)` or direct DOM event listeners set up in `taskpane.ts`. If the split moves functions currently defined in `word.ts` (and directly referenced by `taskpane.ts`'s button-click wiring) into feature modules, and those feature modules have their own side-effect imports (e.g., importing `bluebook`'s or `providers`' self-registering index modules), the *order* in which `taskpane.ts` imports the new feature modules can change when registries are populated relative to when `Office.onReady` fires and UI is first rendered. A button wired before its provider registry is populated can render an empty dropdown or throw on first click, then work fine on a second render — a timing bug that's easy to miss in dev (fast reloads mask it) and only shows up intermittently in production.

**Why it happens:**
ES module side effects run in import order, and TypeScript/webpack import ordering across newly split files is easy to get subtly wrong (e.g., a feature module importing `./providers` lazily inside a function instead of at module top level, for tree-shaking or circular-import-avoidance reasons introduced during the split).

**How to avoid:**
Keep all registry-populating imports (`providers/index.ts`, `bluebook`'s registration index) as static top-level imports in a single, early-loaded entry module (`taskpane.ts` or a new `bootstrap.ts`) that runs before `Office.onReady`'s callback wires any UI, rather than letting each split feature module import them independently and non-deterministically. Add an explicit comment at that entry point documenting that import order matters here.

**Warning signs:**
- A dropdown/list driven by `citationProviderRegistry.list()` or `bluebookRuleSetRegistry.list()` that's empty on first render but populated after a UI state change or second interaction.
- Circular-import warnings from webpack/TypeScript introduced during the split (a common cause of developers moving registry imports into function bodies to "fix" the cycle, which reintroduces this ordering problem).

**Phase to address:**
word.ts split — verify with manual QA that all provider/rule-set-driven UI (provider selection, edition selection) is fully populated on the very first task-pane load after the split, not just after interaction.

---

### Pitfall 6: Submitting the dev manifest (with `localhost` URLs) instead of the production build's manifest to Partner Center

**What goes wrong:**
The root-level `manifest.xml` in this repo has `IconUrl`, `HighResolutionIconUrl`, and `SourceLocation` pointing at `https://localhost:3000/...` (verified: `manifest.xml:9-11,24`) — this is the dev manifest used with `office-addin-debugging`. The actual production manifest is generated by webpack into `dist/manifest.xml` with URLs rewritten to the real hosting location (confirmed by `scripts/package-release.js:7-8,25-26`, which explicitly notes "manifest's URLs point at GitHub Pages"). Full AppSource/Marketplace validation explicitly rejects `localhost` references. If anyone manually zips/uploads the root `manifest.xml` (e.g., by copy-pasting a "known good" file into the Partner Center UI, or by pointing the CI publish job at the wrong path) instead of the built `dist/manifest.xml`, the submission fails validation — or worse, if some CI/manual step is inconsistent about which manifest it uses, this could pass locally (dev manifest happens to validate its XML schema fine) but fail Partner Center's endpoint-reachability checks.

**Why it happens:**
Having two manifests (source-of-truth root one for dev, generated one for prod) is correct and necessary, but it's an easy mixup point precisely because both are named `manifest.xml` and only differ in URL values, which aren't obviously wrong at a glance.

**How to avoid:**
Since the CI publish job (`.github/workflows/ci.yml`) already exists as a no-op pending `PARTNER_CENTER_*` secrets, when that job is eventually enabled, confirm it packages from `dist/manifest.xml` (post-`npm run build`), not the root one — mirror the same sourcing logic `scripts/package-release.js` already uses. If any manual submission happens before the CI job is live, explicitly document in the submission runbook: "upload `dist/manifest.xml` (after `npm run build`), never the repo-root `manifest.xml`."

**Warning signs:**
- Partner Center validation error mentioning `localhost` or an unreachable `SourceLocation`.
- Any script or documentation instructing a manual `manifest.xml` upload without a preceding `npm run build` step.

**Phase to address:**
Partner Center submission prep — add this as an explicit checklist item / assertion in the submission runbook, verified by re-reading `dist/manifest.xml` immediately before upload to confirm no `localhost` strings remain.

---

## Technical Debt Patterns

| Shortcut | Immediate Benefit | Long-term Cost | When Acceptable |
|----------|-------------------|-----------------|------------------|
| Splitting `word.ts` by copy-pasting function groups into new files without redesigning the shared `let` state | Fast, low-risk-looking diff | Recreates tight coupling in a new shape; doesn't actually reduce merge-conflict risk (PROJECT.md's stated goal) since all modules still fight over shared state | Never — defeats the purpose of this milestone's split |
| Deduping the three enterprise providers by extracting only the *identical* lines into `base.ts`, leaving near-duplicate (not identical) auth/error-handling logic in place | Smaller diff, faster to land | The exact "bug fixed in one, missed in the other two" risk CONCERNS.md warns about persists for anything not extracted | Acceptable only as an interim step if immediately followed by extracting the remaining near-duplicate logic (e.g., auth-header construction, error normalization) in the same phase |
| Treating Partner Center's automated validation pass as sufficient sign-off | Submission "looks done" once the automated checks are green | Automated checks don't catch FRE clarity, disclosure completeness, or reviewer-testing-instructions issues — the top human-reviewer rejection reasons per Microsoft's own published data | Never for a first submission — always do a manual FRE/disclosure/testing-instructions pass before submitting, regardless of automated validator status |

## Integration Gotchas

| Integration | Common Mistake | Correct Approach |
|-------------|-----------------|-------------------|
| Office.js `Word.run`/`context.sync()` across split modules | Passing a live `Range`/`Paragraph` proxy object as a function argument between modules that each open their own `Word.run` | Pass a stable identifier (search text, paragraph index) and re-locate the object inside the receiving module's own `Word.run`, or explicitly `context.trackedObjects.add()` if a live reference must cross the boundary |
| webpack tree-shaking vs. self-registering provider/rule-set modules | Assuming an unused-looking side-effect import is safe to remove or lazy-load during refactor | Keep registration imports static and top-level in one early-loaded entry point; add/keep a registry-completeness test |
| Partner Center manifest submission | Uploading the repo-root dev `manifest.xml` (localhost URLs) instead of the webpack-built `dist/manifest.xml` | Always build (`npm run build`) immediately before packaging/submitting, and grep the manifest for `localhost` as a pre-submission check |
| Partner Center privacy policy / terms of use listing fields | Treating `PRIVACY.md`/`TERMS.md` existing in the repo (PR #27) as sufficient — Partner Center listing requires linkable, hosted URLs to these documents entered in specific listing fields, reviewed for matching actual data-handling behavior | Publish `PRIVACY.md`/`TERMS.md` to a stable public URL (e.g., GitHub Pages, same host as the add-in) before submission, and cross-check the policy text actually describes this add-in's real behavior (zero-network-by-default, CourtListener/enterprise provider calls only when configured) rather than generic boilerplate |

## Performance Traps

| Trap | Symptoms | Prevention | When It Breaks |
|------|----------|------------|-----------------|
| Calling `context.sync()` inside a per-citation loop when processing many citations (e.g., "Find Hallucinations" across a long brief) | Task pane UI freezes/becomes unresponsive on documents with many citations | Use the split-loop pattern: queue all reads/writes across citations in one batch, then a single `context.sync()`, then process results | Documents with dozens+ of citations; queue also hard-caps at 50 pending batch jobs, which can itself throw |
| Module split introducing incidental duplicate `Word.run`/`context.sync()` calls per feature (e.g., hyperlinking module and hallucination module each independently re-loading the same document range) | Slower perceived performance after the split vs. before, even though "nothing changed" logically | When designing module boundaries, note where multiple features need the same document data and consider one shared read (via a small orchestration function in `taskpane.ts`) rather than N independent `Word.run` calls | As citation count / document length grows |

## Security Mistakes

| Mistake | Risk | Prevention |
|---------|------|------------|
| New insertion call site added during/after the split without the escaping/URL-validation guard | HTML/URL injection into the user's Word document (the exact risk CONCERNS.md's "Fragile Areas" section flags) | Route all insertion through a single guarded helper (see Pitfall 4); add a lint rule or code-review checklist item forbidding direct calls to raw Office.js insertion APIs outside that helper |
| Provider dedup accidentally weakening the one differentiator between enterprise providers (e.g., a provider-specific credential-scoping check gets "simplified away" when merging into `base.ts`) | Auth/credential handling regression across all three enterprise providers at once, since they'd now share the buggy consolidated logic | Diff each provider's auth/error-handling path against the consolidated `base.ts` implementation line-by-line before merging, not just the identical boilerplate; add/verify equivalent test coverage for all three post-dedup (CONCERNS.md notes this parity is currently unverified) |
| Assuming AppSource certification implies a security audit of the add-in's code | False confidence that certification = "security reviewed"; certification is primarily a policy/listing/functionality review, not a code security audit | Continue relying on this repo's own `SECURITY_AUDIT.md` process independent of Partner Center's review; don't treat AppSource approval as a substitute |

## UX Pitfalls

| Pitfall | User Impact | Better Approach |
|---------|-------------|-------------------|
| Task pane's first-run screen doesn't explain what OpenClerk does or that it requires no account for CourtListener but does for enterprise providers, before any sign-in prompt | AppSource reviewers reject for unclear First-Run Experience (the #2 most common rejection reason across Microsoft's published data); real users also churn without understanding value prop | Add/verify a brief first-run explanation on task-pane load: what the add-in does, that CourtListener works with zero setup, and that Westlaw/LexisNexis/Bloomberg Law require the firm's own credentials |
| USPTO Patent Center provider stub (`"USPTO Patent Center (TODO)"`, per CONCERNS.md) left selectable in the provider UI at submission time | A reviewer selecting it during testing hits a non-functional, confusingly-labeled option — directly matches rejection reason #3, "functionality doesn't match description" / add-in showing errors | Hide or disable the USPTO stub entry from the provider-selection UI before Partner Center submission (already flagged in CONCERNS.md as out-of-scope-to-*implement* this milestone, but hiding it is a near-zero-cost prerequisite for submission specifically) |

## "Looks Done But Isn't" Checklist

- [ ] **word.ts split:** Looks done when the file compiles and `tsc`/`jest` pass — verify by also running a manual multi-step QA pass (locate citations → check hallucinations → insert hyperlinks → insert comments in one session) to catch cross-module state and proxy-object staleness that unit tests alone won't surface.
- [ ] **Provider dedup:** Looks done when `westlawProvider.ts`/`lexisNexisProvider.ts`/`bloombergLawProvider.ts` shrink in line count — verify each of the three still has equivalent test coverage to the others post-refactor (CONCERNS.md notes this parity was never confirmed even pre-refactor).
- [ ] **escapeHtml/isSafeHyperlinkUrl hardening:** Looks done when existing call sites are updated — verify by adding a test that would fail if a *new* insertion call site were added without the guard (e.g., asserting there's exactly one function in the codebase that calls the raw Office.js insertion API), not just re-testing existing sites.
- [ ] **PR #33 (openclerk-core dependency):** Looks done when CI is green and it merges — verify the post-merge duplication audit (src/commands/, scripts/) actually happens as its own follow-through step, not skipped because "the main dependency landed."
- [ ] **PR #27 (privacy policy/terms of use):** Looks done when `PRIVACY.md`/`TERMS.md` exist and merge — verify they are also published to a stable public URL and entered into the actual Partner Center listing fields; a merged markdown file in the repo alone does not satisfy AppSource's requirement.
- [ ] **Partner Center manifest submission:** Looks done when `manifest.xml` passes `office-addin-manifest validate` locally — verify that check ran against `dist/manifest.xml` (production, post-build), not the root dev manifest with `localhost` URLs.
- [ ] **TERMS.md §8 governing-law jurisdiction:** Looks done once any jurisdiction string replaces the placeholder — verify it's an actual considered choice (ideally with counsel input, per PROJECT.md's own note) rather than a placeholder-filling reflex, since this is a real legal document being submitted to a commercial marketplace.

## Recovery Strategies

| Pitfall | Recovery Cost | Recovery Steps |
|---------|-----------------|-------------------|
| Shared mutable state recreated across split modules (Pitfall 1) | MEDIUM | Consolidate into one explicit state module with accessor functions; migrate call sites incrementally, module by module, with tests passing at each step rather than a single big-bang rewrite |
| `InvalidObjectPath` from untracked proxy objects post-split (Pitfall 2) | LOW | Add `context.trackedObjects.add()` at the object's creation site, or refactor the two call sites to re-locate the object by identifier instead of passing the live reference — localized fix, doesn't require re-architecting the split |
| Provider/rule-set silently missing from production bundle (Pitfall 3) | LOW | Revert the offending import change or `sideEffects` config; add the registry-completeness test so this can't regress silently again |
| Partner Center rejection for FRE clarity, disclosure, or testing instructions (Critical Pitfalls list, general) | LOW | These are listing-content fixes, not code fixes — update the AppSource listing description/notes-for-certification and resubmit; Microsoft's stated review turnaround is 3-5 business days |
| Wrong manifest (dev/localhost) submitted to Partner Center (Pitfall 6) | LOW | Rebuild (`npm run build`), regenerate `dist/manifest.xml`, resubmit — no code change needed, just process correction |

## Pitfall-to-Phase Mapping

| Pitfall | Prevention Phase | Verification |
|---------|-------------------|----------------|
| Scattered module-level mutable state (Pitfall 1) | word.ts split | Code review confirms no re-declared `let` state duplicated across new modules; state ownership documented before code moves |
| Stale/untracked Word.js proxy objects across module boundaries (Pitfall 2) | word.ts split | Manual QA of a multi-step citation workflow after split; no `InvalidObjectPath` errors on second+ operations in a session |
| Self-registering providers/rule-sets dropped by refactor or tree-shaking (Pitfall 3) | Provider dedup + openclerk-core duplication audit | Registry-completeness test (`citationProviderRegistry.list()`/`bluebookRuleSetRegistry.list()` return full expected ID sets) run against a production build, not just dev |
| Escaping guard remains "manual discipline" after split (Pitfall 4) | escapeHtml/isSafeHyperlinkUrl hardening (sequenced before/with word.ts split) | Single guarded insertion helper exists; grep/lint confirms no direct raw Office.js insertion API calls outside it |
| Registry population racing `Office.onReady`/UI wiring after split (Pitfall 5) | word.ts split | Manual QA: provider/edition dropdowns fully populated on very first task-pane load, not just after interaction |
| Dev manifest with `localhost` URLs submitted instead of production `dist/manifest.xml` (Pitfall 6) | Partner Center submission prep | Pre-submission checklist step: grep `dist/manifest.xml` for `localhost`, confirm built post-`npm run build` |
| USPTO stub visible in provider UI at submission time (UX Pitfalls) | Partner Center submission prep | Manual check of provider-selection UI immediately before submission |
| Unclear first-run experience (UX Pitfalls) | Partner Center submission prep | First-run screen reviewed against Microsoft's FRE guidance before submission |
| PRIVACY.md/TERMS.md merged but not hosted/linked in listing ("Looks Done But Isn't") | Partner Center submission prep (PR #27 follow-through) | Confirm public URL exists and is entered in Partner Center listing fields, not just merged to `main` |

## Sources

- [Avoid using the context.sync method in loops — Office Add-ins, Microsoft Learn](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/correlated-objects-pattern) — MEDIUM confidence (official docs)
- [OfficeExtension.TrackedObjects class — Office Add-ins, Microsoft Learn](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.trackedobjects?view=common-js-preview) — MEDIUM confidence (official docs)
- [context.trackedObjects.add doesn't support invocation across Word.run calls — OfficeDev/office-js#68](https://github.com/OfficeDev/office-js/issues/68) — MEDIUM confidence (maintainer-engaged issue thread, cross-checked against official docs)
- [context.sync() taking progressively more time to run — OfficeDev/office-js#3565](https://github.com/OfficeDev/office-js/issues/3565) — LOW-MEDIUM confidence (community-reported issue)
- [Top 5 AppSource validation errors for Office Add-ins submissions — December 2023, Microsoft 365 Developer Blog](https://devblogs.microsoft.com/microsoft365dev/top-5-appsource-validation-errors-for-office-add-ins-submissions-december-2023/) — MEDIUM confidence (official Microsoft blog, cross-checked against March/June/September 2023 entries in the same series)
- [Top 5 AppSource validation errors for Office Add-ins submissions — September 2023, Microsoft 365 Developer Blog](https://devblogs.microsoft.com/microsoft365dev/top-5-appsource-validation-errors-for-office-add-ins-submissions-september-2023/) — MEDIUM confidence
- [Validate an Office Add-in's manifest — Office Add-ins, Microsoft Learn](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-manifest) — MEDIUM confidence (official docs)
- [Office Add-ins manifest — Office Add-ins, Microsoft Learn](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests) — MEDIUM confidence (official docs)
- [Add-in manifests fail to validate when they contain AppDomains tag — OfficeDev/office-js#5057](https://github.com/OfficeDev/office-js/issues/5057) — LOW-MEDIUM confidence (community-reported issue, illustrates a real failure mode rather than a guaranteed current bug)
- [Privacy and security for Office Add-ins — Office Add-ins, Microsoft Learn](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/privacy-and-security) — MEDIUM confidence (official docs)
- [Publish your Office Add-in to Microsoft Marketplace — Office Add-ins, Microsoft Learn](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-office-add-ins-to-appsource) — MEDIUM confidence (official docs)
- [Microsoft Marketplace submission FAQ — Partner Center, Microsoft Learn](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/appsource-submission-faq) — MEDIUM confidence (official docs)
- In-repo verification: `src/taskpane/word.ts:75-81` (module-level mutable state), `src/taskpane/providers/registry.ts` and `index.ts` (self-registration pattern), `src/taskpane/bluebook/registry.ts` (parallel self-registration pattern), `manifest.xml:9-11,24` (dev localhost URLs), `scripts/package-release.js:7-8,25-26` and `scripts/package-release-offline.js` (production manifest sourcing from `dist/`), `.planning/codebase/CONCERNS.md` (existing fragile-area findings on escaping guards and provider duplication) — HIGH confidence (direct source read)

---
*Pitfalls research for: Office.js Word add-in controller refactor + Microsoft Partner Center/AppSource submission*
*Researched: 2026-07-15*
