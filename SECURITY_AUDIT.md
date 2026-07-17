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

---

# Round 2 — 2026-07-11

Delta audit covering the end-user installer / offline-server scripts that
landed just before round 1 merged (they were only reviewed for merge
cleanliness in round 1, not audited), plus the two items round 1 deferred
(`npm audit`, and re-checking the CSP/manifest claims against current
`main`). No new application code shipped to `main` since round 1 beyond
those scripts.

## Deferred round-1 items — now done

- **`npm audit` run.** `npm audit --omit=dev` → **0 vulnerabilities**;
  full tree `npm audit` (incl. dev) → **0 vulnerabilities**. The README's
  "zero known vulnerabilities in production dependencies" claim checks out
  against the current locked tree.
- **CSP / manifest re-verified on current `main`.** `taskpane.html`'s CSP
  still `default-src 'none'` with `connect-src https://www.courtlistener.com`
  only; `manifest.xml` still `<Permissions>ReadWriteDocument</Permissions>`
  with `AppDomains` = `https://www.courtlistener.com` only. Unchanged from
  round 1.

## Installer / offline-server scripts — audited, no code changes needed

Scope: `scripts/install-openclerk.{sh,cmd,ps1}` (shipped to end users by
`package-release.js`), `scripts/local-server/{setup-local-server,serve-openclerk}.ps1`
(shipped by `package-release-offline.js`), and both packaging scripts.
These run on end-user machines, so they got a close read for command
injection, unsafe download-execute, path traversal, and credential
handling. They are notably well-hardened; specific checks:

- **No command injection / no shell string-building.** The bash installer
  runs under `set -euo pipefail` and quotes every variable expansion; the
  PowerShell scripts pass arguments as typed parameters (no `Invoke-Expression`,
  no string-built command lines). No user/download-derived value is ever
  passed to a shell.
- **No download-and-execute.** The manifest installers only `cp`/`Copy-Item`
  a local `manifest.xml` into Word's WEF sideload folder. Nothing is
  fetched at install time. (The only network fetch in the whole
  installer/packaging surface is `package-release-offline.js` vendoring
  Microsoft's `office.js`/Fabric CSS at *packaging* time in CI — see the
  documented note below.)
- **Path traversal is correctly guarded** in `serve-openclerk.ps1`: the
  requested path is resolved with `[Path]::GetFullPath` and required to
  start with `ContentRoot` **plus a trailing separator**, which also
  closes the classic sibling-prefix bypass (`/rootEVIL` vs `/root/`), and
  must resolve to an existing leaf file.
- **Local server is loopback-only and authenticated.** `serve-openclerk.ps1`
  binds `https://127.0.0.1:$Port/` only, and gates every request on a
  per-install GUID secret — supplied once via `?k=` then carried in an
  `HttpOnly; Secure` cookie. The secret is stored in an ACL-restricted
  file (inheritance stripped, restricted to the installing user) and
  passed to the scheduled task via `-SecretFile`, deliberately *not* as a
  task argument (task definitions are readable by same-user processes).
  The `.cmd` wrapper's `-ExecutionPolicy Bypass` is scoped to the bundled
  local `.ps1` only (standard for a double-clickable installer of a
  first-party script), not a remote payload.

## Findings fixed (round 2 follow-up, 2026-07-12)

Prompted by a follow-up question about finding D below — is there a better
way to avoid the machine-wide Trusted Root install? There was, and it's
now implemented.

### D. Offline setup installed a self-signed cert into the machine Trusted Root store — fixed
**Files:** `scripts/local-server/setup-local-server.ps1`,
`scripts/local-server/serve-openclerk.ps1`

Previously, `setup-local-server.ps1` (`Install-LocalhostCertBinding`)
created a self-signed `CN=OpenClerkLocalServer` certificate in
`Cert:\LocalMachine\My` with an **exportable** private key, trusted it via
`Cert:\LocalMachine\Root` (machine-wide), and bound it to HTTP.sys with
`netsh http add sslcert`/`urlacl` — all of which required a UAC elevation
prompt. Root cause was `serve-openclerk.ps1`'s use of
`System.Net.HttpListener`, which sits on HTTP.sys and forces both the
machine-global bindings and a machine-store certificate.

**Fix:** `serve-openclerk.ps1` was rewritten to terminate TLS in-process
with `TcpListener` + `SslStream` instead of `HttpListener`, which removes
the HTTP.sys dependency entirely. This let `setup-local-server.ps1` move to
a fully per-user, non-elevated design:
- **Per-user cert, non-exportable key.** `New-SelfSignedCertificate` now
  targets `Cert:\CurrentUser\My` with `-KeyExportPolicy NonExportable` and
  an EKU restricted to Server Authentication only. The private key never
  enters a machine-wide store and can't be exported even by the installing
  user.
- **Per-user trust.** The public certificate is added to
  `Cert:\CurrentUser\Root` — no machine-wide trust anchor is created, so
  other users on the same machine are unaffected.
- **No elevation, no `netsh`.** All `Test-IsAdmin`/`-ElevatedCertStep`/`Start-Process
  -Verb RunAs` code was removed; the scheduled task's principal is
  explicitly `-RunLevel Limited`. The entire install now runs as the
  current user with no UAC prompt.
- **Stronger secret generation.** The per-install auth secret moved from a
  128-bit `Guid.NewGuid()` to a 256-bit value from
  `RandomNumberGenerator`, while keeping the existing ACL hardening
  (inherited permissions stripped, restricted to the installing user) on
  the file it's stored in.
- **`-Uninstall` switch added**, closing the previously-noted gap: it
  stops/unregisters the scheduled task, removes the cert from both
  `CurrentUser\My` and `CurrentUser\Root` by subject, and removes the
  install directory and WEF manifest.

TLS is fixed to 1.2 (not 1.3) because the scheduled task runs under
Windows PowerShell 5.1/.NET Framework, whose `SslStream` doesn't support
1.3. The `Test-Authorized` cookie/query-secret check, MIME table,
Content-Security-Policy header, and path-traversal guard (`GetFullPath`
must resolve inside `ContentRoot`) were all carried over unchanged from
the previous implementation.

## Documented only (round 2)

### E. Offline packaging vendors Microsoft CDN assets without integrity pinning
`scripts/package-release-offline.js` fetches `office.js` and Fabric's
`fabric.min.css` from Microsoft's CDNs at packaging time and bundles them
into the offline zip, with no Subresource-Integrity hash or checksum. Much
lower risk than round 1's `reporters-db` finding: it's a first-party
Microsoft CDN over HTTPS, it runs only in the maintainer's CI/dev
environment (not on end-user machines), and `office.js` is specifically
designed to be loaded as the current published build rather than pinned.
Flagged only so a maintainer producing offline packages is aware the
bundled copies are whatever the CDN served at build time; pinning
`fabric.min.css` by version+hash (it already has a version-stamped URL)
would be a cheap hardening if offline packages are distributed widely.
