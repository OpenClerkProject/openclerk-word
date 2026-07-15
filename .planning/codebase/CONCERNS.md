# Codebase Concerns

**Analysis Date:** 2026-07-15

## Tech Debt

**Unimplemented enterprise provider (USPTO Patent Center):**
- Issue: `usptoPatentCenterProvider.ts` is a stub — its own header comment says "TODO: USPTO Patent Center lookups are not implemented yet," and its `name` field literally reads `"USPTO Patent Center (TODO)"`, meaning the placeholder is user-visible in any provider-selection UI.
- Files: `src/taskpane/providers/usptoPatentCenterProvider.ts`
- Impact: If this provider is selectable in the task pane, users can pick it and get a non-functional lookup with a confusing label instead of a disabled/hidden option.
- Fix approach: Either hide/disable the provider entry until implemented, or complete the lookup implementation and rename it.

**`openclerk-core` logic duplicated, not shared:**
- Issue: Per `SECURITY_AUDIT.md` (finding A), this repo does not depend on the sibling `openclerk-core` package despite that package's README describing itself as holding logic "extracted from `openclerk-word` so logic doesn't have to be duplicated (or drift out of sync)." `escapeHtml`/`isSafeHyperlinkUrl` (`src/taskpane/utils.ts:43-63`) and `manualCorrections.ts` are maintained as independent copies.
- Files: `src/taskpane/utils.ts`, `src/taskpane/bluebook/manualCorrections.ts`
- Impact: Any future security fix landed in `openclerk-core`'s equivalent code will not automatically reach this shipped add-in — it must be manually ported or the repo migrated to depend on `openclerk-core`, and the two implementations can silently drift.
- Fix approach: Either vendor `openclerk-core` as a real dependency, or establish a process (e.g., a CI diff check) to keep the copies in sync.
- **Resolved (2026-07-15):** PR #33 merged (`main` merge commit `0f48462`), switching the repo to depend on the published `openclerk-core@0.2.6` npm package instead of vendoring `src/taskpane/utils.ts`/`bluebook/manualCorrections.ts` as independent copies. A fresh post-merge audit of `src/commands/` and `scripts/` confirmed zero remaining logic duplicated from `openclerk-core` and zero residual imports of the deleted local `./providers`/`./bluebook`/`./utils` paths — see `.planning/phases/01-openclerk-core-dependency-cleanup/01-DUPLICATION-AUDIT.md` (CORE-02) for the full evidence trail.

**Generated data files committed and very large:**
- Issue: `src/taskpane/bluebook/generated/reporterAbbreviations.generated.ts` is 10,352 lines (the single largest file in the repo by a wide margin, ~75% of all `src/` LOC), produced by `scripts/generate-bluebook-data.js` (`npm run bluebook:update-data`).
- Files: `src/taskpane/bluebook/generated/reporterAbbreviations.generated.ts`, `src/taskpane/bluebook/generated/caseNameAbbreviations.generated.ts`, `src/taskpane/bluebook/generated/stateAbbreviations.generated.ts`
- Impact: Large generated files inflate diffs when regenerated, slow down IDE tooling/search, and risk manual edits being silently overwritten on next regeneration (no visible guard comment observed against hand-editing).
- Fix approach: Confirm each generated file has a clear "DO NOT EDIT" header; consider excluding from routine `grep`/search-heavy workflows and documenting the regeneration process in CONTRIBUTING.md.

**`word.ts` is a large, monolithic file:**
- Issue: `src/taskpane/word.ts` is 1,538 lines and appears to hold most of the task-pane's Word-interaction logic (hyperlink insertion, "Find Hallucinations," citation lookup orchestration, comment insertion) based on line references cited in `SECURITY_AUDIT.md` (lines 221-237, 254, 346, 556, 980-1058, 1442).
- Files: `src/taskpane/word.ts`
- Impact: A single file this size increases merge-conflict risk and makes it harder to reason about which code paths touch the DOM/document vs. network calls; also raises the odds that future security-relevant changes (e.g., new insertion call sites) are made without updating the escaping/URL-validation guards documented in the security audit.
- Fix approach: Consider splitting by feature (hyperlink insertion, hallucination detection, comment insertion, provider orchestration) into separate modules under `src/taskpane/`.

## Known Bugs

No known bugs were identified in the current source or issue tracker during this pass. `SECURITY_AUDIT.md` records one real fixed issue (CI log leak, see Security Considerations) and one fixed installer hardening issue (self-signed cert trust scope), both already remediated on `main`.

## Security Considerations

**Two independent audit rounds already performed and largely clean — see `SECURITY_AUDIT.md` for full detail:**
- Risk: `SECURITY_AUDIT.md` documents a fixed CI credential-leak (verbose `curl -v` printing a derived Bearer token to Actions logs in the Partner Center publish step) and a fixed installer issue (self-signed cert previously trusted machine-wide via `Cert:\LocalMachine\Root`, requiring UAC elevation).
- Files: `.github/workflows/ci.yml` (fix applied), `scripts/local-server/setup-local-server.ps1`, `scripts/local-server/serve-openclerk.ps1` (rewritten to per-user cert + `TcpListener`/`SslStream`)
- Current mitigation: Both issues are fixed on `main` as of the audit dates (2026-07-09 and 2026-07-12 respectively).
- Recommendations: None outstanding from the audit's "fixed" section; treat as resolved. Re-run `npm audit` periodically since the last confirmed-clean run was 2026-07-11.

**`connect-src` CSP is the real enforcement boundary for enterprise API domains — easy to weaken by accident:**
- Risk: A user-typed `apiBaseUrl` for enterprise providers (LexisNexis, Westlaw, Bloomberg Law) is validated only for `https://` scheme in app code (`src/taskpane/providers/base.ts:34-37`); there is no in-app host allowlist. The actual security boundary is the task-pane's CSP `connect-src` directive in `src/taskpane/taskpane.html:14`, which must be manually widened (along with `manifest.xml`'s `AppDomains`) per enterprise deployment.
- Files: `src/taskpane/providers/base.ts`, `src/taskpane/taskpane.html`, `manifest.xml`
- Current mitigation: Default `connect-src` is scoped to `https://www.courtlistener.com` only, matching the single `AppDomains` entry.
- Recommendations: Any deployment fork that wildcards `connect-src` (e.g., `https://*`) to simplify enabling enterprise providers would silently defeat this control. Document this loudly for firms doing custom deployments, and consider adding a build-time lint/check that fails if `connect-src` contains a wildcard.

**Offline packaging vendors Microsoft CDN assets without integrity pinning:**
- Risk: `scripts/package-release-offline.js` fetches `office.js` and Fabric's `fabric.min.css` from Microsoft's CDN at packaging time with no Subresource-Integrity hash/checksum.
- Files: `scripts/package-release-offline.js`
- Current mitigation: Runs only in maintainer CI/dev environment over HTTPS from a first-party CDN, not on end-user machines.
- Recommendations: Pin `fabric.min.css` by version+hash since it already has a version-stamped URL (per `SECURITY_AUDIT.md` finding E — low priority, cheap hardening).

## Performance Bottlenecks

No specific measured performance bottlenecks were identified. The single large generated data file (`reporterAbbreviations.generated.ts`, 10,352 lines) is bundled into the task-pane webpack build (`webpack.config.js`) and could affect initial bundle size/parse time, but no profiling data exists in-repo to confirm impact.

## Fragile Areas

**Hyperlink/HTML insertion call sites depend on manual discipline, not a shared abstraction:**
- Files: `src/taskpane/word.ts:221-237, 254, 346, 556, 1442`, `src/taskpane/utils.ts:43-63`
- Why fragile: `escapeHtml`/`isSafeHyperlinkUrl` must be called explicitly at every insertion call site before `applyHyperlinkToItem`/`insertHyperlink`/`insertHtml`. There is no compiler-enforced wrapper type that forces escaping before these Office.js calls — a future call site added to `word.ts` (currently 1,538 lines and growing) could omit the guard without a build-time error.
- Safe modification: Any new hyperlink/HTML insertion code path must call `escapeHtml`/`isSafeHyperlinkUrl` from `src/taskpane/utils.ts` before touching Office.js insertion APIs; consider wrapping insertion in a single helper function that always escapes, so the guard cannot be forgotten.
- Test coverage: `tests/utils.test.ts` covers `isSafeHyperlinkUrl` against `javascript:`/`vbscript:`/`data:`/`file:` schemes, but no test appears to assert that every `insertHtml`/`applyHyperlinkToItem` call site in `word.ts` actually invokes the guard (i.e., no regression test would catch a new unguarded call site).

**Enterprise provider files are near-duplicates of each other:**
- Files: `src/taskpane/providers/westlawProvider.ts` (69 lines), `src/taskpane/providers/lexisNexisProvider.ts` (69 lines), `src/taskpane/providers/bloombergLawProvider.ts` (69 lines)
- Why fragile: Three providers of identical line count strongly suggest copy-pasted structure. A bug or missing validation fixed in one (e.g., an auth/credential-handling issue) may not be propagated to the other two.
- Safe modification: When changing shared behavior (auth flow, error handling, URL validation) in one enterprise provider, apply the same change to all three, or refactor common logic into `src/taskpane/providers/base.ts`.
- Test coverage: Not verified in this pass whether all three have equivalent test coverage in `tests/`.

## Scaling Limits

Not applicable in the traditional sense — this is a client-side Office.js task-pane add-in with no backend server component, so conventional server-scaling concerns (request throughput, DB capacity) do not apply. The one exception is the optional local offline server (`scripts/local-server/serve-openclerk.ps1`), which is loopback-only and single-user by design, not intended to scale.

## Dependencies at Risk

**`overrides` block forcing transitive dependency versions:**
- Risk: `package.json` forces `uuid@^14.0.1` and `serialize-javascript@^7.0.7` via `overrides`, documented as a fix for two transitive-dependency advisories (GHSA-w5hq-g745-h8pq, GHSA-5c6j-r48x-rmvq/GHSA-qj8w-gfj5-8c6v) in upstream devDependencies that haven't bumped their own ranges yet.
- Impact: These overrides need to be revisited once the upstream packages (whichever devDependency pulls in `uuid`/`serialize-javascript` transitively) release fixed versions natively — otherwise the override becomes permanent unmaintained debt.
- Migration plan: Periodically check if the overrides can be removed once upstream devDependencies (`office-addin-*` toolchain, `webpack`, etc.) update their own transitive pins; `npm audit` (last run 2026-07-11, 0 vulnerabilities) should be re-run after any dependency bump to confirm overrides are still necessary.

**Browserslist includes `ie 11`:**
- Risk: `package.json`'s `browserslist` targets `"last 2 versions", "ie 11"`. Since Office.js task panes run in WebView2/embedded browser contexts on modern Windows/Mac, IE11 support may be legacy cruft that forces Babel to emit more conservative, larger output (`core-js`/`regenerator-runtime` polyfills are both production dependencies, consistent with this).
- Impact: Larger bundle size and unnecessary polyfill overhead if IE11 is no longer a real target environment for the shipped add-in.
- Migration plan: Confirm current Office/Word host WebView requirements and drop `ie 11` from `browserslist` if no longer needed, which would also allow removing `core-js`/`regenerator-runtime` if nothing else needs them.

## Missing Critical Features

**USPTO Patent Center provider not implemented:** see Tech Debt above — this is the one visible gap in the provider matrix (CourtListener, Westlaw, LexisNexis, Bloomberg Law are implemented; USPTO Patent Center is a stub).

## Test Coverage Gaps

**No test assertions tying escaping guards to every insertion call site:**
- What's not tested: Whether every `insertHtml`/`applyHyperlinkToItem`/`insertHyperlink` call site in `src/taskpane/word.ts` actually calls `escapeHtml`/`isSafeHyperlinkUrl` before invoking the Office.js API — the security audit verified this by manual code read, not by an automated test.
- Files: `src/taskpane/word.ts`, `tests/utils.test.ts`
- Risk: A future call site could omit the guard and ship without any test failing.
- Priority: Medium — low likelihood of regression given the audit's manual verification, but the blast radius (DOM/HTML injection) is high if it happens.

**Enterprise provider parity not verified in this pass:**
- What's not tested: Whether `westlawProvider.ts`, `lexisNexisProvider.ts`, and `bloombergLawProvider.ts` have equivalent test coverage to each other and to `courtListenerProvider.ts` (which is the largest/most-tested provider at 273 lines vs. 69 lines each for the enterprise providers).
- Files: `src/taskpane/providers/westlawProvider.ts`, `src/taskpane/providers/lexisNexisProvider.ts`, `src/taskpane/providers/bloombergLawProvider.ts`, `tests/`
- Risk: A shared bug across the three near-duplicate providers (see Fragile Areas) could go undetected if test coverage is uneven.
- Priority: Low-Medium — worth a follow-up pass cross-referencing `tests/` file names against each provider file.

---

*Concerns audit: 2026-07-15*
