# Security audit — openclerk-word

Date: 2026-07-09
Scope: `openclerk-word` only (the Office.js Word add-in for legal citation
work). A companion audit of `openclerk-core` — a platform-agnostic library
with similar citation/Bluebook logic that this repo does not currently
depend on — was performed alongside this one; see that repo's
`SECURITY_AUDIT.md`. `openclerk-gdocs` and `openclerk-libreoffice` were out
of scope: both are empty placeholder repos with no commits and no code.

## Methodology

Manual read-through plus targeted checks of: all outbound network calls
(`fetch`) and how enterprise-provider credentials (LexisNexis, Westlaw,
Bloomberg Law) are collected, stored, and transmitted; `manifest.xml`
permissions/`AppDomains` and the task-pane's Content-Security-Policy; every
DOM/HTML-injection surface (`innerHTML` usage, Office.js `insertHtml`
calls) and whether network- or document-derived content ever reaches them
unescaped; hyperlink URL-scheme validation at every insertion call site;
`eval`/`Function`/dynamic-code-execution usage; a by-eye pass over
`package.json` dependencies; and the GitHub Actions CI/release pipeline for
credential-handling issues. The existing "Security & IT review" section of
`README.md` was cross-checked against the actual source rather than taken
on faith.

## Findings fixed in this audit

### 1. Bearer token could leak into CI logs via verbose curl — fixed
**File:** `.github/workflows/ci.yml` (Partner Center publish step, was
line 324)

The optional "Publish to Partner Center" CI step acquires an OAuth2 access
token from Azure AD (client-credentials flow) and then uploads the release
package with `curl -v -X POST -H "Authorization: Bearer $access_token" ...`.
GitHub Actions automatically redacts literal `secrets.*` references from
logs, but `access_token` here is a **derived** value (extracted at runtime
via `jq -r .access_token` from the token endpoint's JSON response) and so
is not on that auto-mask list. `curl -v` prints request headers — including
`Authorization` — to stderr, which lands in the public/team-visible Actions
log. This is a real, credential-adjacent risk with no upside: `-v` is not
needed for the step to function.

**Fix:**
- Dropped `-v` in favor of `-sS` (silent but still shows errors) on the
  upload `curl` call.
- Added `echo "::add-mask::$access_token"` immediately after the token is
  extracted, as defense-in-depth so this specific derived value is masked
  by GitHub Actions' log redaction going forward even if it's referenced
  elsewhere in the step later.
- Verified the edited `ci.yml` is still well-formed YAML (`yaml.safe_load`).

This step only runs when the Partner Center secrets are configured (it
early-exits otherwise), so this fix has no effect on default/fork CI runs.

## Findings documented only (no code change)

### A. `escapeHtml`/`isSafeHyperlinkUrl` are self-contained here, not imported from `openclerk-core`
`openclerk-core`'s README describes itself as holding "shared" logic
"extracted from `openclerk-word` so logic doesn't have to be duplicated (or
drift out of sync) across each platform integration." Verified this repo
does not actually depend on `openclerk-core`: no entry in `package.json`,
nothing under `node_modules/openclerk-core`. `openclerk-word` maintains its
own separate implementations — `src/taskpane/utils.ts` has its own
`escapeHtml`/`isSafeHyperlinkUrl`, and
`src/taskpane/bluebook/manualCorrections.ts` is a distinct file from
`openclerk-core`'s `src/bluebook/manualCorrections.ts`. Reviewed this
repo's own copies directly (not assuming parity with `openclerk-core`):
- `isSafeHyperlinkUrl` (`src/taskpane/utils.ts:43-54`) uses the `URL` API
  with an `http:`/`https:`/`mailto:` allowlist (not a `javascript:`/etc.
  denylist) — the correct pattern, avoiding classic scheme-denylist-bypass
  bugs. Tested against `javascript:`/`vbscript:`/`data:`/`file:` in
  `tests/utils.test.ts` — all correctly rejected.
- `escapeHtml` (`src/taskpane/utils.ts:56-63`) escapes `&`, `<`, `>`, `"`,
  `'` — complete for both HTML text and quoted-attribute contexts.
- Enforced at every hyperlink-insertion call site before
  `applyHyperlinkToItem`/`insertHyperlink`/`insertHtml`
  (`src/taskpane/word.ts:254, 346, 556, 1442`).

No bypass was found in either function as currently implemented here.
Flagging the drift risk itself: any future security fix landed in
`openclerk-core`'s equivalent code (see that repo's audit, items 1-3) will
**not** reach this shipped add-in unless it's independently ported or this
repo is migrated to actually depend on `openclerk-core`. Recommend the
maintainers resolve this one way or the other rather than maintaining two
copies indefinitely.

### B. Enterprise `apiBaseUrl` enforcement lives in CSP `connect-src`, not app code
A user-typed `apiBaseUrl` for enterprise providers is checked only for the
`https://` scheme in code (`src/taskpane/providers/base.ts:34-37`) — there
is no in-app host allowlist. The README correctly identifies that the real
enforcement boundary is the task-pane's Content-Security-Policy
`connect-src` directive (`src/taskpane/taskpane.html:14`), which defaults
to `https://www.courtlistener.com` only; enabling an enterprise provider
requires a firm to manually widen `connect-src` (and `manifest.xml`'s
`AppDomains`) to that provider's specific domain. The README explicitly
warns against wildcarding `connect-src` for this reason. This is a sound
design — flagging only so it's clear the CSP is a security-load-bearing
config, not incidental: any deployment fork that loosens `connect-src` to
a wildcard would defeat this control. Verified default `connect-src` in
`taskpane.html:14` matches the single `AppDomains` entry in `manifest.xml`
(`https://www.courtlistener.com`).

### C. `manifest.xml` permission scope
`manifest.xml:26` requests exactly `ReadWriteDocument` — the minimal tier
needed for the add-in's hyperlink/comment features, no mail/other-app/
filesystem access, no `WebApplicationInfo`/SSO block. Consistent with the
README's claims. No change needed.

## Verified-clean areas (no findings)

- **No AI/LLM calls anywhere.** "Find Hallucinations" (`src/taskpane/word.ts:980-1058`)
  re-uses the same `CitationProvider.lookupCitation()` interface as Online
  Lookup — it does not call any generative-AI service.
- **Credential handling:** enterprise OAuth credentials are held in-memory
  only (`EnterpriseCitationProvider.credentials`,
  `src/taskpane/providers/base.ts:20`), never written to `localStorage`,
  cookies, or Office roaming settings; secret-typed fields render as
  password inputs; tokens are transmitted only via `Authorization` headers,
  never query strings; cleared on sign-out.
  `src/taskpane/providers/base.ts:68-76` sends the OAuth client secret in
  the token-request **body**, not the URL.
- **DOM injection surface:** all network- or document-derived content is
  rendered via `document.createElement` + `.textContent`; `innerHTML` is
  used only to clear a container or set static hardcoded helper text (no
  variable interpolation). The one dynamic `insertHtml` use
  (`applyHyperlinkToItem`, `src/taskpane/word.ts:221-237`) HTML-escapes its
  interpolated `url`/`displayText` via `escapeHtml` and is gated behind
  `isSafeHyperlinkUrl` at every call site. "Embed Cited Text" writes opinion
  excerpts via Office.js's plain-text `Range.insertComment()` API — not
  HTML-coerced — so CourtListener response text can't inject markup even
  though it's not independently escaped.
- **No `eval`/`Function()`/dynamic code loading** anywhere in `src/`. The
  one `child_process.spawnSync` use (`scripts/package-release-offline.js`)
  is a build-time-only Node script (invokes `npx webpack`), not shipped in
  the browser bundle.
- **Dependencies:** production dependencies (`jszip`, `core-js`,
  `regenerator-runtime`) are common and actively maintained; `sharp` and
  `marked` are dev/build-time only. `package.json`'s `overrides` block
  (forcing `uuid`/`serialize-javascript`) is a documented, legitimate fix
  for transitive toolchain advisories, not a smell.

## Out of scope

`openclerk-gdocs` and `openclerk-libreoffice` — both empty repositories
with no commits on the `claude/security-audit-r5lbyz` branch or elsewhere;
nothing to audit.

Independent verification recommended but not performed as part of this
audit: an actual `npm audit --omit=dev` run against the locked dependency
tree (README claims zero known vulnerabilities in production dependencies —
worth confirming against current advisory databases rather than the
snapshot this audit worked from), and inspection of the *deployed*
`taskpane.html`/`manifest.xml` (e.g. the GitHub Pages-hosted copy) in case
a firm's local deployment fork has changed the CSP/AppDomains from what's
in this repo.
