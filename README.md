# WordClerk

[![CI](https://github.com/wbarnha/WordClerk/actions/workflows/ci.yml/badge.svg)](https://github.com/wbarnha/WordClerk/actions/workflows/ci.yml)
[![Latest release](https://img.shields.io/github/v/release/wbarnha/WordClerk?style=flat-square)](https://github.com/wbarnha/WordClerk/releases)
[![License: MIT](https://img.shields.io/github/license/wbarnha/WordClerk?style=flat-square)](LICENSE)
[![Last commit](https://img.shields.io/github/last-commit/wbarnha/WordClerk?style=flat-square)](https://github.com/wbarnha/WordClerk/commits/main)
[![Open issues](https://img.shields.io/github/issues/wbarnha/WordClerk?style=flat-square)](https://github.com/wbarnha/WordClerk/issues)
[![Open PRs](https://img.shields.io/github/issues-pr/wbarnha/WordClerk?style=flat-square)](https://github.com/wbarnha/WordClerk/pulls)
[![Stars](https://img.shields.io/github/stars/wbarnha/WordClerk?style=flat-square)](https://github.com/wbarnha/WordClerk/stargazers)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.4-3178C6?style=flat-square&logo=typescript&logoColor=white)](tsconfig.json)
[![Node.js](https://img.shields.io/badge/Node.js-16%2B-339933?style=flat-square&logo=node.js&logoColor=white)](README.md#development)
[![Code style: Prettier](https://img.shields.io/badge/code_style-prettier-ff69b4?style=flat-square&logo=prettier&logoColor=white)](https://www.npmjs.com/package/office-addin-prettier-config)
[![Platform: Word Add-in](https://img.shields.io/badge/platform-Word%20Add--in-2B579A?style=flat-square&logo=microsoftword&logoColor=white)](manifest.xml)
[![PRs welcome](https://img.shields.io/badge/PRs-welcome-brightgreen?style=flat-square)](https://github.com/wbarnha/WordClerk/pulls)

WordClerk is a Word add-in (task pane) for legal citation work: hyperlinking case-law and parenthetical citations, looking citations up live against public and enterprise legal databases, and checking citations for Bluebook formatting problems — all from one task pane, and entirely local by default.

> **Bringing this to IT for approval?** Start with [Security & IT review](#security--it-review) — it has a one-page summary, a data-flow table, a plain-English permissions breakdown, and a centralized-deployment path, written for a security reviewer rather than a developer.

## Features

WordClerk has four tabs, each a self-contained workflow:

| Tab | What it does | Network calls? |
| --- | --- | --- |
| **Case Law** | Apply or remove hyperlinks by copying them from a source `.docx` that already has case-law citations hyperlinked — the visible text never changes, only the links. | None — parsed entirely in-browser |
| **Non-patent Literature** | Scan the open document for parenthetical citations (e.g. `(Smith v. Jones, 2020)`), then assign a URL to each and apply/remove the resulting hyperlinks. | None |
| **Online Lookup** | An alternative to Case Law's file-based workflow: queries a citation lookup provider live and hyperlinks only the citations that resolve to exactly one case. Ships with free **CourtListener** support out of the box, plus a plugin architecture for enterprise providers (**LexisNexis**, **Westlaw**, **Bloomberg Law**) using your firm's own contracted credentials. See [Citation hyperlink providers](#citation-hyperlink-providers-plugin-architecture). | Opt-in, per-citation only |
| **Bluebook Check** | Scans the document's case citations for common Bluebook mechanical formatting problems (the `"v."` abbreviation, reporter series form, year/court presence, edition-specific case-name abbreviations) across three selectable editions (20th/21st/22nd). Click any flagged citation to jump straight to it in the document. See [Bluebook citation checking](#bluebook-citation-checking-plugin-architecture). | None |

A **Report an issue** link sits at the bottom of every tab, pointing straight to this repo's [GitHub Issues](https://github.com/wbarnha/WordClerk/issues).

Three tabs make zero network calls, ever — Online Lookup is the only feature that leaves the machine, and only when a user explicitly turns it on. See [Security & IT review](#security--it-review) for the full data-flow breakdown.

## Development

Prerequisites
- Node.js (16+ recommended)
- npm

Install dependencies:

```bash
cd "c:\Users\willi\WordClerk"
npm install
```

Generate PNG icons from the SVG (optional but recommended for manifest icons):

```bash
npm run convert-logos
```

Start the dev server and sideload the add-in into Word (desktop):

```bash
npm run start
```

This runs the webpack dev server and uses `office-addin-debugging` to sideload the manifest into Word for debugging.

## Scripts
- `npm run build:dev` — build development bundle
- `npm run build` — production build
- `npm run convert-logos` — convert `dist/assets/logo-filled.svg` into PNG variants (16/32/80px)
- `npm run start` — start dev server and sideload into Word
- `npm run recover:start` — clear stale Office dev settings, then stop and restart sideload debugging
- `npm run install:addin` — install the manifest into the default Office WEF manifest folder
- `npm run install:addin:dry-run` — validate installer behavior without writing files

## Notes
- If icons in the manifest are SVG and Word rejects the manifest, convert or reference PNGs instead.
- If `insertHyperlink` is unavailable in your Word environment, the add-in falls back to using `insertHtml` or plain text. This fallback also makes the add-in behave the same way on Windows (WebView2) and Mac (WKWebView) Word builds that support different Office.js requirement sets.

## Citation hyperlink providers (plugin architecture)

The **Online Lookup** tab is an alternative to the file-parsing workflow on the **Case Law** tab: instead of copying hyperlinks out of a source `.docx`, it queries a case-law lookup provider live and hyperlinks only the citations that resolve to exactly one case. For a citation like:

> Norfolk & W. Ry. Co. v. Liepelt, 444 U.S. 490 (U.S.Ill., 1980)

the selected provider is asked about that citation; if it immediately resolves to a single case, WordClerk hyperlinks it, otherwise that citation is left untouched and the scan moves on to the next one — no error, no interruption.

Providers are plugins implementing `CitationProvider` in [src/taskpane/providers/types.ts](src/taskpane/providers/types.ts) and self-registering with `citationProviderRegistry` in [src/taskpane/providers/index.ts](src/taskpane/providers/index.ts). To add a new one (a firm's internal search API, another public case-law database, etc.), implement the interface and register an instance — nothing else in the add-in needs to change.

**Citation scanner known limitations** (validated by hand against a real trial brief, not committed to this repo): the scanner in [citationParser.ts](src/taskpane/providers/citationParser.ts) handles multi-pincite lists (`503, 505, 508, 513`), pincite ranges (`705-06`), and reporters with embedded digits (`F. Supp. 3d`, `F.4th`) correctly, but does not currently handle footnote pincites (`567 n.1`), a stray typo like a doubled comma before the parenthetical, or parallel citations to a second reporter (`24 Misc. 2d 790, 200 N.Y.S.2d 126 (1960)`) — the last of these gets misparsed (matching the second reporter, treating the first as part of the case name) rather than cleanly skipped. See the "known limitation" tests in `tests/providers.test.ts` for exact behavior.

Built-in providers:
- **CourtListener** ([courtListenerProvider.ts](src/taskpane/providers/courtListenerProvider.ts)) — Free Law Project's free, public case-law search. Works with no credentials; get an optional API token at [courtlistener.com/profile/api-token](https://www.courtlistener.com/profile/api-token/) to raise the (fairly low) anonymous rate limit.
- **LexisNexis, Westlaw, Bloomberg Law** ([lexisNexisProvider.ts](src/taskpane/providers/lexisNexisProvider.ts), [westlawProvider.ts](src/taskpane/providers/westlawProvider.ts), [bloombergLawProvider.ts](src/taskpane/providers/bloombergLawProvider.ts)) — these are contract-gated enterprise APIs. Each vendor provisions its own base URL and OAuth2 client credentials per customer, so WordClerk doesn't (and can't) ship a fixed endpoint or key; you supply your API base URL, client ID, and client secret from your firm's contract in the Online Lookup tab. The bundled implementations use the standard OAuth2 client-credentials flow and a `POST <base>/search/cases`-shaped request as a starting point — confirm the exact token and search paths in your vendor's API documentation and adjust the `TOKEN_PATH`/`SEARCH_PATH` constants at the top of each file if they differ.
- **USPTO Patent Center** ([usptoPatentCenterProvider.ts](src/taskpane/providers/usptoPatentCenterProvider.ts)) — **TODO, not implemented.** Registered as a placeholder (shows up in the provider list, always reports "not found" so citations are skipped) so the plugin wiring can be exercised end-to-end before a real Patent Center / PEDS integration is built for the Non-patent Literature workflow.

### Getting credentials for each provider

**CourtListener** (free, optional token — self-service):
1. Create a free account at [courtlistener.com/sign-in](https://www.courtlistener.com/sign-in/) (there's a "Register" link if you don't have one).
2. Once signed in, go to [courtlistener.com/profile/api-token](https://www.courtlistener.com/profile/api-token/) and copy the token shown there.
3. Paste it into the "API token" field on the Online Lookup tab and click Connect. You can leave it blank and use CourtListener anonymously, just at a lower rate limit.

**LexisNexis** (requires an existing Lexis subscription):
1. Register for a developer account at the [LexisNexis Developer Portal](https://dev.lexisnexis.com/) (use your firm/business email) — see their [Getting Started guide](https://dev.lexisnexis.com/gettingStarted).
2. Request access to the case-law/citation API product for your organization; a LexisNexis data specialist reviews and provisions access (their published turnaround is about 24 hours for trial requests).
3. Once approved, the portal issues a Client ID and Client Secret and documents your tenant's API base URL — enter those three values into the LexisNexis fields on the Online Lookup tab.

**Westlaw** (requires an existing Westlaw/Thomson Reuters subscription):
1. Visit the [Thomson Reuters Developer Portal](https://developers.thomsonreuters.com/) and contact Thomson Reuters support to request API access under your firm's existing Westlaw agreement — this isn't self-service signup; TR provisions it per customer.
2. TR support provides a Client ID, Client Secret, the authentication (token) endpoint, and your API base URL.
3. Enter those into the Westlaw fields on the Online Lookup tab (adjust `TOKEN_PATH` in [westlawProvider.ts](src/taskpane/providers/westlawProvider.ts) if the endpoint TR gives you doesn't match the default `/oauth/token`).

**Bloomberg Law** (requires an existing Bloomberg Law subscription):
1. Contact your Bloomberg account representative, or Bloomberg Law's help desk (888.560.2529 / help@bloomberglaw.com), to request Web API access for your firm's account.
2. Bloomberg provisions credentials through its Enterprise Console (an admin on your account downloads a Client ID and Secret there — see Bloomberg's [Web API Connectivity Policy](https://www.bloomberg.com/professional/support/api-library/) for what the console shows). Store the downloaded secret securely; per Bloomberg's policy these credentials expire after 18 months and will need to be refreshed then.
3. Enter the Client ID, Secret, and the API base URL Bloomberg gives you into the Bloomberg Law fields on the Online Lookup tab.

Because the three enterprise integrations are contract-gated and provisioned per customer, treat the exact field names/paths above as a starting point, not a guarantee — confirm specifics against the documentation your vendor rep provides.

## Bluebook citation checking (plugin architecture)

The **Bluebook Check** tab scans the current document's case citations for common Bluebook formatting problems and lists any it finds, per citation. It's a mechanical checker, not the rulebook: it verifies conventions like

- `"v."` is used (not `"v"` or `"vs."`) between party names,
- the reporter series uses the ordinal abbreviations `"2d"`/`"3d"` rather than `"2nd"`/`"3rd"`,
- a decision year is present in the parenthetical, and
- a court abbreviation is present for any reporter other than U.S. Reports (`"U.S."`, where the Supreme Court is implied),

plus a small set of edition-specific case-name abbreviations (see below). **It does not check every Bluebook rule** (typeface/italics, signal usage, pinpoint-citation style for non-case authorities, etc.) and isn't a substitute for the actual rulebook — treat a clean result as "no obvious mechanical problems," not "Bluebook-perfect."

Each citation in the results list is clickable: clicking it searches the document for that exact citation text and selects it, which moves Word's view to show where it actually appears — no manual scrolling/searching needed to find a flagged citation.

Like the citation lookup providers, Bluebook editions are plugins implementing `BluebookRuleSet` in [src/taskpane/bluebook/types.ts](src/taskpane/bluebook/types.ts) and registering with `bluebookRuleSetRegistry` in [src/taskpane/bluebook/index.ts](src/taskpane/bluebook/index.ts). Pick an edition from the dropdown on the Bluebook Check tab; each is checked independently and adding a new edition (or a firm/journal-specific house style) means implementing the interface and registering an instance.

Built-in editions — **20th (2015)**, **21st (2020)**, **22nd (2025, current)**:
- All three run the same shared, edition-stable Rule 10 (case citation) checks listed above — see [commonRules.ts](src/taskpane/bluebook/commonRules.ts). Per the [University of Washington Law Library's Bluebook 101 guide](https://lib.law.uw.edu/bluebook101/editions) ([22nd edition page](https://lib.law.uw.edu/bluebook101/22nd)), the documented changes across these three editions are concentrated in statutory/online-source citation, typeface terminology, and new source types (audio/video, AI-generated content, state administrative materials) — not in core case-citation format, so there was no accurate edition-specific case-citation rule to invent for most of Rule 10.
- The one genuine, verified case-citation-relevant difference: the **21st edition merged Table T6** (case-name word abbreviations) **with the former Table T13.2** (periodical/institutional-author abbreviations), so several words — e.g. *Laboratory* → `Lab'y`, *Employment*/*Employee* → `Emp.`, and similarly *Environment* → `Env't`, *Research* → `Rsch.`, *Psychology* → `Psych.`, *Sociology* → `Socio.`, *Comparative* → `Compar.` — went from "spelled out in case names, abbreviated only in institutional-author citations" to "abbreviated everywhere." The 20th-edition rule-set doesn't flag these words in case names; the 21st and 22nd do. Sources: the UW guide above and Mary Whisner, [Bluebook Weight Loss Program, Part Two: The Merger of Tables T6 and T13.2](https://citeblog.access-to-law.com/?p=1074) (which specifically confirms the *Laboratory*/*Employment* mappings; the remaining word list is the same category of T6/T13.2-merger abbreviation but hasn't been independently verified word-for-word against the official table — see the sourcing note in [caseNameAbbreviations.ts](src/taskpane/bluebook/caseNameAbbreviations.ts)).

## Security & IT review

### One-page summary

- **What it is:** a Word task-pane add-in (static HTML/JS/CSS rendered inside Word's built-in browser control). It is **not** a hosted web app, has no backend server of its own, no user accounts, and no database.
- **What it touches:** only the Word document that is currently open, and only when the user clicks a button. It never scans other files, other Office apps, mail, or the file system beyond a single `.docx` the user explicitly picks from a file-open dialog.
- **What leaves the machine, by default:** nothing. Two of the add-in's three workflows (**Case Law** file-parsing and **Non-patent Literature** parentheticals) make zero network calls, ever.
- **What can leave the machine, opt-in only:** if a user turns on the **Online Lookup** tab, short citation strings (not document content) are sent over HTTPS to whichever single lookup provider that user selects and connects. See the data-flow table below for exactly what's sent where.
- **Source & license:** fully open source in this repository under the MIT license (see [package.json](package.json)); every claim in this section can be checked directly against [src/taskpane/](src/taskpane/) rather than taken on faith.
- **Dependencies:** `npm audit --omit=dev` currently reports **0 known vulnerabilities** in production dependencies — i.e., in the code that actually ships inside the add-in bundle loaded into Word. Running a plain `npm audit` (no flags) will additionally show advisories in **devDependencies only** (the local build/debugging toolchain: `webpack-dev-server`, `copy-webpack-plugin`, and Microsoft's own `office-addin-manifest`/Teams Toolkit tooling, none of which are packaged into `dist/`). Those are worth tracking but don't affect the shipped add-in; versions are pinned via `package-lock.json` and CI (`npm ci`) installs the exact locked tree on every build.
- **Distribution:** no auto-update mechanism reaches out to any WordClerk-controlled service. Installing or upgrading is something your organization does deliberately (see [Centralized deployment](#centralized-deployment-it-managed-rollout) below), not something that happens silently in the background.

### Data flow, by feature

| Feature (tab) | Network calls? | What's sent | Where | Trigger |
| --- | --- | --- | --- | --- |
| **Case Law** (apply hyperlinks from a source `.docx`) | None | Nothing — the source file is parsed entirely in-browser with [JSZip](https://www.npmjs.com/package/jszip) | N/A | User clicks "Apply hyperlinks" |
| **Case Law** (remove hyperlinks) | None | Nothing — edits the open document's OOXML in place | N/A | User clicks "Remove hyperlinks" |
| **Non-patent Literature** (scan/add/remove parentheticals) | None | Nothing — citations are extracted from the open document locally; URLs are typed in by the user | N/A | User clicks Scan / Add / Remove |
| **Online Lookup** — CourtListener | Yes (opt-in) | The **individual matched citation string** (e.g. `"Norfolk & W. Ry. Co. v. Liepelt, 444 U.S. 490 (U.S.Ill., 1980)"`) — never the surrounding document text — plus an API token if one was entered | `https://www.courtlistener.com` only | User clicks "Scan & hyperlink via API" with CourtListener selected |
| **Online Lookup** — LexisNexis / Westlaw / Bloomberg Law | Yes (opt-in) | The same per-citation string, plus an OAuth2 bearer token obtained from the credentials the user entered | The API base URL *that user typed in*, from their firm's existing vendor contract | User clicks "Connect", then "Scan & hyperlink via API" |

The extraction that produces those per-citation strings (`extractCaseCitations` in [citationParser.ts](src/taskpane/providers/citationParser.ts)) runs entirely inside the Word document object model, in-browser; only the short matched substrings — never `context.document.body.getText()`'s full output — are ever passed to `fetch()`. That's checkable directly: every network call in the add-in lives in [src/taskpane/providers/](src/taskpane/providers/) and nowhere else.

### Permissions requested

The manifest requests exactly one Office permission: `<Permissions>ReadWriteDocument</Permissions>` (see [manifest.xml](manifest.xml)). In plain English, that grants the add-in the ability to read and edit the currently open Word document through the Office JavaScript API — nothing else. It does **not** grant access to: other open documents, other Office apps (Outlook, Excel, Teams), the file system, the network in general (browser same-origin/CORS rules still apply to any `fetch()` call), or any data outside the current document. This is the lowest permission tier Word add-ins offer above read-only.

### Credential handling

- Credential inputs are rendered as password-masked fields (`<input type="password">`) for every secret-typed field (API tokens, client secrets) — see `renderProviderPanel` in [word.ts](src/taskpane/word.ts).
- Credentials are held in memory only, inside each provider instance (see `EnterpriseCitationProvider` in [base.ts](src/taskpane/providers/base.ts)). They are never written to `localStorage`, `sessionStorage`, cookies, or any WordClerk-controlled server, and are cleared by clicking "Disconnect" or by closing/reloading the task pane.
- Enterprise provider API base URLs are required to start with `https://` — WordClerk refuses to authenticate over plain HTTP, so a credential can't accidentally be sent unencrypted (see the check in `EnterpriseCitationProvider.authenticate`, [base.ts](src/taskpane/providers/base.ts)).
- **No telemetry or analytics** are collected by the add-in itself, successful or not.
- **`manifest.xml`'s `<AppDomains>`** declares every external domain WordClerk talks to out of the box (`courtlistener.com`); if your policy requires allow-listing every contacted domain at the network/firewall level, add your firm's specific LexisNexis/Westlaw/Bloomberg Law API host there too before deploying.

### Centralized deployment (IT-managed rollout)

Rather than each user sideloading the manifest individually (the default developer flow described above), IT can push WordClerk to specific users or groups — and control updates — through the **integrated apps** feature of the Microsoft 365 admin center:

1. In the [Microsoft 365 admin center](https://admin.microsoft.com), go to **Settings → Integrated apps**, then **Upload custom apps**.
2. Upload the `manifest.xml` produced by `npm run package` (or a release asset — see [Download and install from GitHub](#download-and-install-from-github) below) and choose which users/groups it's deployed to (`Everyone`, specific users/groups, or just yourself for a pilot).
3. When a new version is approved, IT re-uploads the updated manifest/package and clicks **Update** — end users never see an unreviewed, unapproved version.

See Microsoft's own docs for the full process and requirements: [Deploy add-ins in the Microsoft 365 admin center](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins) and [Requirements for centralized deployment](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/centralized-deployment-of-add-ins).

### Frequently asked (by IT reviewers)

- **Does it ever send a whole document anywhere?** No. The only network calls the add-in ever makes are the per-citation lookups described in the data-flow table above; full document content is never transmitted, by any workflow, under any configuration.
- **Does it phone home, track usage, or call any WordClerk-controlled service?** No such service exists. There is nothing to phone home to.
- **What if we don't want *any* outbound network calls at all?** Don't turn on the Online Lookup tab — the Case Law and Non-patent Literature workflows (the add-in's original functionality) are 100% local and unaffected by its presence.
- **Can we restrict which external domains it's allowed to reach?** Yes, at the network layer (firewall/proxy allow-list) using the domains named in the data-flow table, and/or at the manifest layer via `<AppDomains>`.
- **Who maintains this and where do we report a security concern?** It's maintained in the open at [github.com/wbarnha/WordClerk](https://github.com/wbarnha/WordClerk); file an issue or PR there, or read the source directly — there's no vendor support line to call.

## Compliance considerations

WordClerk itself is a thin, client-side Word add-in with no hosted service of its own, so heavyweight vendor attestations like **SOC 2 Type II** or **ISO/IEC 27001** — which certify how an organization *operates and hosts* a service over time — don't map cleanly onto "review this add-in's code." WordClerk does not hold, and does not claim, either certification. What's actually relevant to a law firm's review is more specific to how it's distributed and what it touches:

- If you're evaluating this add-in as a piece of software (rather than as a hosted vendor), Microsoft's [Microsoft 365 App Compliance Program](https://learn.microsoft.com/en-us/microsoft-365-app-certification/overview) is the applicable path for Office Add-ins: **Publisher Attestation** (a self-assessment of security/data-handling practices) or full **Microsoft 365 Certification** (a third-party audit against SOC 2/PCI-DSS/ISO 27001-aligned controls). Neither has been completed for WordClerk; pursue whichever your IT team requires before broad deployment.
- Firms increasingly vet vendors with standardized questionnaires (e.g., the Shared Assessments **SIG** questionnaire) rather than, or in addition to, certifications — the "Security & IT review" section above is meant to answer those questions directly from the source.
- Because the LexisNexis/Westlaw/Bloomberg Law integrations are opt-in and use *your firm's own* contracted API credentials, your firm's existing vendor agreements and compliance posture with those three vendors govern that data flow — WordClerk is just the client dialing out to an endpoint you already have a relationship with.

## Performance notes

- The file-parsing (**Case Law**) and parenthetical (**Non-patent Literature**) workflows are entirely local (no network calls) and are unaffected by the Online Lookup feature.
- Online Lookup makes one network round-trip per detected citation, **sequentially** (never in parallel), specifically to stay under each provider's rate limits — CourtListener's free anonymous tier, for example, allows only a few requests per minute. A document with many citations will take roughly `(number of citations) × (provider round-trip time)` to finish; expect low-to-mid tens of seconds for a document with a few dozen citations, and prefer the file-parsing workflow when you already have a hyperlinked source document, since it's effectively instant.
- No new bundle dependencies were added for this feature — it's built on the browser's native `fetch`, already available in both the Windows (WebView2) and Mac (WKWebView) Word task pane runtimes.

## Dependency security fixes

All of the following are **devDependency-only** issues — they live in the local build/debug toolchain (webpack, its dev server, and Microsoft's own Office Add-in CLI tooling) and were never part of `dist/`, the bundle Word actually loads. Still, `npm audit`/GitHub Dependabot flagged 10 findings (9 moderate, 1 high) before this branch; all 10 are fixed here (`npm audit` now reports **0 vulnerabilities**), with no breaking changes to the packages we depend on directly and no change to the project's minimum supported Node version.

| Package (path) | Advisory | Severity | Before | After | How it's reached / fixed |
| --- | --- | --- | --- | --- | --- |
| `serialize-javascript` (via `copy-webpack-plugin`) | [GHSA-5c6j-r48x-rmvq](https://github.com/advisories/GHSA-5c6j-r48x-rmvq) — RCE via `RegExp.flags`/`Date.prototype.toISOString()` | High | `<=7.0.4` | `^7.0.7` (forced via `overrides`) | `copy-webpack-plugin@12.0.2` (our pinned version) depends on a vulnerable `serialize-javascript` range. Bumping `copy-webpack-plugin` itself to the only fixed release (`14.0.0`) would've required Node ≥20.9.0 in CI/local dev, so instead we added a package.json [`overrides`](https://docs.npmjs.com/cli/v10/configuring-npm/package-json#overrides) entry that forces every consumer in the tree onto a patched `serialize-javascript`, regardless of what `copy-webpack-plugin` itself asks for. |
| `serialize-javascript` (same package) | [GHSA-qj8w-gfj5-8c6v](https://github.com/advisories/GHSA-qj8w-gfj5-8c6v) — CPU-exhaustion DoS via crafted array-like objects | High | `<=7.0.4` | `^7.0.7` | Same root cause and fix as above (one vulnerable package, two CVEs). |
| `uuid` (via `webpack-dev-server` → `sockjs`) | [GHSA-w5hq-g745-h8pq](https://github.com/advisories/GHSA-w5hq-g745-h8pq) — missing buffer bounds check in `uuid` v3/v5/v6 when a buffer is supplied | Moderate | `<11.1.1` (`sockjs`'s own dependency is pinned to `^8.3.2`, unpatched upstream as of this writing) | `^14.0.1` (forced via `overrides`) | `sockjs` (used by `webpack-dev-server` for its WebSocket fallback transport) hasn't updated its own `uuid` dependency, so no version of `webpack-dev-server` fixes this by itself. The `overrides` entry forces `uuid` to a patched version tree-wide instead. |
| `uuid` (via `office-addin-manifest` → `office-addin-usage-data`) | [GHSA-w5hq-g745-h8pq](https://github.com/advisories/GHSA-w5hq-g745-h8pq) | Moderate | `office-addin-manifest@2.1.5` → `office-addin-usage-data 1.6.2–2.0.8` → vulnerable `uuid` | `office-addin-manifest@2.1.6` → `office-addin-usage-data@2.1.1` (no longer depends on `uuid` at all) | Plain devDependency update — `office-addin-manifest` and `office-addin-cli` were bumped within their existing `^2.1.x`/`^2.0.x` package.json ranges; no version-range change needed. |
| `uuid` (via `office-addin-debugging` → `office-addin-dev-settings` → (optional) `@microsoft/m365agentstoolkit-cli` → `@microsoft/teamsfx-core` → `@azure/msal-node`, and a second nested copy of `office-addin-manifest@1.13.6` inside the same `teamsfx-core` tree) | [GHSA-w5hq-g745-h8pq](https://github.com/advisories/GHSA-w5hq-g745-h8pq) | Moderate (×5 distinct dependency paths in the original audit output) | `@azure/msal-node@<=5.1.4` and a pinned `office-addin-manifest@1.13.6` nested under `@microsoft/teamsfx-core`, both depending on vulnerable `uuid` | Same `^14.0.1` `overrides` entry | `office-addin-debugging@6.1.2` (updated from `6.1.1`, within our existing `^6.0.6` range) now pulls `office-addin-dev-settings@3.1.2`, but that package still lists `@microsoft/m365agentstoolkit-cli` as an *optional* dependency, and Microsoft hasn't yet updated the `teamsfx-core`/`msal-node` chain it drags in. This entire chain is dev-time-only tooling for an M365 Agents Toolkit CLI feature this project doesn't use; the `overrides` entry closes it without needing that upstream fix. |
| `webpack-dev-server` | [GHSA-mx8g-39q3-5c79](https://github.com/advisories/GHSA-mx8g-39q3-5c79) — HMR WebSocket interception via permissive user-configured proxies | Moderate | `5.2.4` (exact pin) | `5.2.6` (exact pin) | Plain patch-level version bump, same `5.x` line, same `node >= 18.12.0` engine requirement — no compatibility risk expected. |

**Why `overrides` instead of upgrading `copy-webpack-plugin`/`webpack-dev-server` to their latest majors?** `npm audit fix --force` offered `copy-webpack-plugin@14.0.0` and (for a full sockjs fix) `webpack-dev-server@6.0.0`, but those require Node ≥20.9.0 and ≥22.15.0 respectively — bumping this project's minimum Node version is a separate decision with its own blast radius (CI config, contributor tooling), not something to bundle silently into a vulnerability fix. The `overrides` field lets us force the two actually-vulnerable leaf packages (`uuid`, `serialize-javascript`) to patched versions everywhere in the tree, which is what actually eliminates the advisories, without forcing major-version churn on `copy-webpack-plugin`/`webpack-dev-server` or raising the Node floor. If a future contributor *does* want to move to `webpack-dev-server@6`/Node 22+, these overrides can simply be deleted at that point since they'd be redundant.

## Download and install from GitHub

WordClerk's add-in content is hosted on **GitHub Pages** (`https://wbarnha.github.io/WordClerk/`), so **end users do not need Node.js, npm, or any developer tooling** to install and use the add-in.

### Option 1: Install from GitHub Release asset (recommended — no Node.js required)
1. Go to the [GitHub Releases page](https://github.com/wbarnha/WordClerk/releases) for this repo.
2. Download the latest `wordclerk-addin.zip` release asset.
3. Extract the ZIP. It's intentionally small — it contains only:
   - `manifest.xml` — the add-in manifest (URLs already point to GitHub Pages)
   - `installer/install-wordclerk.ps1` — a standalone PowerShell installer (no Node.js required)

   Nothing else ships here: the manifest's URLs point at GitHub Pages, so Word fetches the taskpane, commands, and icons live over HTTPS rather than reading local files.
4. **Windows**: open PowerShell and run the installer:

```powershell
powershell -ExecutionPolicy Bypass -File installer\install-wordclerk.ps1
```

   This copies `manifest.xml` into Word's local add-in manifest folder (`%LOCALAPPDATA%\Microsoft\Office\16.0\WEF`).

5. Restart Word and the **WordClerk** button will appear on the **Home** ribbon.

> **No Node.js or npm is needed.** The add-in content is served from GitHub Pages; the installer is pure PowerShell.

### Verifying a release download

Every release also includes a `SHA256SUMS` file and a [build provenance attestation](https://github.com/wbarnha/WordClerk/attestations), so you can confirm a downloaded zip actually matches what this repo's CI built, rather than a modified copy from somewhere else.

**Checksum:**
```powershell
Get-FileHash wordclerk-addin.zip -Algorithm SHA256
# Compare the hash against the matching line in SHA256SUMS
```

**Provenance (requires the [GitHub CLI](https://cli.github.com/)):**
```bash
gh attestation verify wordclerk-addin.zip --repo wbarnha/WordClerk
```
This cryptographically proves the file was built by this repo's GitHub Actions workflow from a specific commit, not hand-uploaded by anyone with release-creation access.

### Option 2: Install from GitHub Actions workflow artifact
1. Open the **Actions** tab in this repo.
2. Select the latest successful **CI** workflow run.
3. Scroll to the **Artifacts** section and download `wordclerk-addin`.
4. Extract the downloaded artifact.
5. Follow the same steps as Option 1 (run `installer\install-wordclerk.ps1`).

### Option 3: Manual sideload via manifest
If you prefer to sideload the manifest manually in Word Desktop:
1. Extract `wordclerk-addin.zip` (from a release or CI artifact).
2. In Word, go to `Insert` → `My Add-ins` → `Upload My Add-in` → `Add from file`.
3. Select the `manifest.xml` from the extracted folder.

The manifest already points to the add-in content hosted on GitHub Pages — no local server needed.

### Option 4: Create a local install package
If you want to build and package locally (requires Node.js — for contributors only):

```bash
npm install
npm run package
```

This creates `wordclerk-addin.zip` at the repo root with the same contents as the release package.

### Option 5: Clone the repository and install locally (development)
See the [Development](#development) section below for instructions on running the add-in from a local dev server.

### Option 6: Run fully offline (no internet connection required)

Both Option 1 and Option 2 need internet access every time the add-in loads, since Word fetches its content live from GitHub Pages. If you need WordClerk to work with no network connection at all, use the offline package instead:

1. Download `wordclerk-addin-offline.zip` from the [GitHub Releases page](https://github.com/wbarnha/WordClerk/releases).
2. Extract it, then run the setup script:

```powershell
powershell -ExecutionPolicy Bypass -File installer\setup-local-server.ps1
```

3. Windows will prompt for administrator approval once (UAC) — this is only used to bind a self-signed HTTPS certificate to a single `127.0.0.1` port and is not required again after the first run.
4. Restart Word.

This installs a small local web server that:
- **Only binds to `127.0.0.1`** on one fixed port (`44399` by default) — it's unreachable from the network, and no other ports are opened.
- **Requires a per-install secret token** to serve any content, so other local processes can't casually request pages from it.
- **Runs hidden and starts automatically at login** (via a Scheduled Task named `WordClerkLocalServer`), so the add-in works immediately without you needing to start anything yourself.
- **Includes local copies of `office.js` and the Fabric CSS** (vendored from Microsoft's CDN at packaging time), so the taskpane doesn't silently require internet access just to load its own framework files.

See `scripts/local-server/serve-wordclerk.ps1` and `scripts/local-server/setup-local-server.ps1` for the implementation. To remove it later: `Unregister-ScheduledTask -TaskName WordClerkLocalServer`, then remove `%LOCALAPPDATA%\WordClerk\LocalServer`.

## Self-hosting

By default, production builds point the manifest at the project's GitHub Pages deployment (`https://wbarnha.github.io/WordClerk/`), so most users don't need to host anything themselves.

If you'd rather serve the add-in content from your own infrastructure (an internal HTTPS server, Azure Static Web Apps, S3+CloudFront, etc. — useful for IT-managed rollouts that don't want to depend on GitHub Pages), set `WORDCLERK_HOST_URL` before building or packaging:

```bash
WORDCLERK_HOST_URL=https://addins.example.com/wordclerk/ npm run build
WORDCLERK_HOST_URL=https://addins.example.com/wordclerk/ npm run package
```

This bakes your URL into both `dist/manifest.xml` and the packaged manifest, so the add-in fetches its taskpane, commands, and icons from your host instead of GitHub Pages. You're responsible for uploading the contents of `dist/` to that URL yourself — this repo's CI only deploys to GitHub Pages.

