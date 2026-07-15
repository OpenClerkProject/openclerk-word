# OpenClerk

[![CI](https://github.com/OpenClerkProject/openclerk-word/actions/workflows/ci.yml/badge.svg)](https://github.com/OpenClerkProject/openclerk-word/actions/workflows/ci.yml)
[![Latest release](https://img.shields.io/github/v/release/OpenClerkProject/openclerk-word?style=flat-square)](https://github.com/OpenClerkProject/openclerk-word/releases)
[![License: MIT](https://img.shields.io/github/license/OpenClerkProject/openclerk-word?style=flat-square)](LICENSE)
[![Last commit](https://img.shields.io/github/last-commit/OpenClerkProject/openclerk-word?style=flat-square)](https://github.com/OpenClerkProject/openclerk-word/commits/main)
[![Open issues](https://img.shields.io/github/issues/OpenClerkProject/openclerk-word?style=flat-square)](https://github.com/OpenClerkProject/openclerk-word/issues)
[![Open PRs](https://img.shields.io/github/issues-pr/OpenClerkProject/openclerk-word?style=flat-square)](https://github.com/OpenClerkProject/openclerk-word/pulls)
[![Stars](https://img.shields.io/github/stars/OpenClerkProject/openclerk-word?style=flat-square)](https://github.com/OpenClerkProject/openclerk-word/stargazers)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.4-3178C6?style=flat-square&logo=typescript&logoColor=white)](tsconfig.json)
[![Node.js](https://img.shields.io/badge/Node.js-16%2B-339933?style=flat-square&logo=node.js&logoColor=white)](README.md#development)
[![Code style: Prettier](https://img.shields.io/badge/code_style-prettier-ff69b4?style=flat-square&logo=prettier&logoColor=white)](https://www.npmjs.com/package/office-addin-prettier-config)
[![Platform: Word Add-in](https://img.shields.io/badge/platform-Word%20Add--in-2B579A?style=flat-square&logo=microsoftword&logoColor=white)](manifest.xml)
[![PRs welcome](https://img.shields.io/badge/PRs-welcome-brightgreen?style=flat-square)](https://github.com/OpenClerkProject/openclerk-word/pulls)

OpenClerk is a Word add-in (task pane) for legal citation work: hyperlinking case-law and parenthetical citations, looking citations up live against public and enterprise legal databases, and checking citations for Bluebook formatting problems — all from one task pane, and entirely local by default.

> **Bringing this to IT for approval?** Start with [Security & IT review](#security--it-review) — it has a one-page summary, a data-flow table, a plain-English permissions breakdown, and a centralized-deployment path, written for a security reviewer rather than a developer.

> **Spotted a wrong Bluebook citation rule?** No coding experience needed to fix it — see [CONTRIBUTING.md](CONTRIBUTING.md#contributing-a-bluebook-citation-correction-no-coding-experience-needed) for a form you can fill out, or a one-file edit you can make directly on GitHub.com.

## Download and install from GitHub

OpenClerk's add-in content is hosted on **GitHub Pages** (`https://openclerkproject.github.io/openclerk-word/`), so **end users do not need Node.js, npm, or any developer tooling** to install and use the add-in.

### Option 1: Install from GitHub Release asset (recommended — no Node.js required)
1. Go to the [GitHub Releases page](https://github.com/OpenClerkProject/openclerk-word/releases) for this repo.
2. Download the latest `openclerk-addin.zip` release asset.
3. Extract the ZIP. It's intentionally small — it contains only:
   - `manifest.xml` — the add-in manifest (URLs already point to GitHub Pages)
   - `installer/install-openclerk.ps1` — a standalone PowerShell installer for Windows (no Node.js required)
   - `installer/install-openclerk.cmd` — a double-clickable wrapper around the `.ps1` installer
   - `installer/install-openclerk.sh` — a standalone bash installer for macOS (no Node.js required)

   Nothing else ships here: the manifest's URLs point at GitHub Pages, so Word fetches the taskpane, commands, and icons live over HTTPS rather than reading local files.
4. Run the installer for your platform:

   **Windows** — double-click `installer\install-openclerk.cmd` (double-clicking the `.ps1` directly opens it in an editor instead of running it, since that's Windows' default action for `.ps1` files), or run the PowerShell script from a console yourself:

   ```powershell
   powershell -ExecutionPolicy Bypass -File installer\install-openclerk.ps1
   ```

   This copies `manifest.xml` into Word's local add-in manifest folder (`%LOCALAPPDATA%\Microsoft\Office\16.0\WEF`).

   **macOS** — open Terminal, `cd` into the extracted folder, and run:

   ```bash
   ./installer/install-openclerk.sh
   ```

   This copies `manifest.xml` into Word for Mac's shared sideloading folder (`~/Library/Containers/com.Microsoft.OsfWebHost/Data/documents/wef`) — the same folder every Office app on the Mac scans for sideloaded manifests, regardless of which one you use to install it. The script is already marked executable inside the ZIP; if your unzip tool strips that bit, run `chmod +x installer/install-openclerk.sh` first. Pass `--dry-run` to preview the install target without copying anything, or `--target <dir>` to install somewhere else.

5. Restart Word and the **OpenClerk** button will appear on the **Home** ribbon.

> **No Node.js or npm is needed.** The add-in content is served from GitHub Pages; the installers are pure PowerShell (Windows) and bash (macOS).

### Verifying a release download

Every release also includes a `SHA256SUMS` file and a [build provenance attestation](https://github.com/OpenClerkProject/openclerk-word/attestations), so you can confirm a downloaded zip actually matches what this repo's CI built, rather than a modified copy from somewhere else.

**Checksum:**
```powershell
Get-FileHash openclerk-addin.zip -Algorithm SHA256
# Compare the hash against the matching line in SHA256SUMS
```

**Provenance (requires the [GitHub CLI](https://cli.github.com/)):**
```bash
gh attestation verify openclerk-addin.zip --repo OpenClerkProject/openclerk-word
```
This cryptographically proves the file was built by this repo's GitHub Actions workflow from a specific commit, not hand-uploaded by anyone with release-creation access.

### Option 2: Install from GitHub Actions workflow artifact
1. Open the **Actions** tab in this repo.
2. Select the latest successful **CI** workflow run.
3. Scroll to the **Artifacts** section and download `openclerk-addin`.
4. Extract the downloaded artifact.
5. Follow the same steps as Option 1 (run `installer\install-openclerk.ps1` on Windows, or `./installer/install-openclerk.sh` on macOS).

### Option 3: Manual sideload via manifest
If you prefer to sideload the manifest manually in Word Desktop:
1. Extract `openclerk-addin.zip` (from a release or CI artifact).
2. In Word, go to `Insert` → `My Add-ins` → `Upload My Add-in` → `Add from file`.
3. Select the `manifest.xml` from the extracted folder.

The manifest already points to the add-in content hosted on GitHub Pages — no local server needed.

### Option 4: Create a local install package
If you want to build and package locally (requires Node.js — for contributors only):

```bash
npm install
npm run package
```

This creates `openclerk-addin.zip` at the repo root with the same contents as the release package.

### Option 5: Clone the repository and install locally (development)
See the [Development](#development) section below for instructions on running the add-in from a local dev server.

### Option 6: Run fully offline (no internet connection required, Windows only)

Both Option 1 and Option 2 need internet access every time the add-in loads, since Word fetches its content live from GitHub Pages. If you need OpenClerk to work with no network connection at all, use the offline package instead. **This option is currently Windows-only** — the setup script binds a self-signed HTTPS certificate and registers a scheduled task using Windows-specific APIs (`New-SelfSignedCertificate`, Task Scheduler) that don't have a macOS equivalent:

1. Download `openclerk-addin-offline.zip` from the [GitHub Releases page](https://github.com/OpenClerkProject/openclerk-word/releases).
2. Extract it, then run the setup script:

```powershell
powershell -ExecutionPolicy Bypass -File installer\setup-local-server.ps1
```

3. Restart Word.

No administrator approval or UAC prompt is needed — the certificate is created per-user (`Cert:\CurrentUser\My`, non-exportable, trusted via `Cert:\CurrentUser\Root`) and the local server terminates TLS itself, so nothing here touches machine-wide configuration.

This installs a small local web server that:
- **Only binds to `127.0.0.1`** on one fixed port (`44399` by default) — it's unreachable from the network, and no other ports are opened.
- **Requires a per-install secret token** to serve any content, so other local processes can't casually request pages from it.
- **Runs hidden and starts automatically at login** (via a Scheduled Task named `OpenClerkLocalServer`, running with standard — not elevated — rights), so the add-in works immediately without you needing to start anything yourself.
- **Includes local copies of `office.js` and the Fabric CSS** (vendored from Microsoft's CDN at packaging time), so the taskpane doesn't silently require internet access just to load its own framework files.

See `scripts/local-server/serve-openclerk.ps1` and `scripts/local-server/setup-local-server.ps1` for the implementation. To remove it later: `powershell -ExecutionPolicy Bypass -File installer\setup-local-server.ps1 -Uninstall`.

## Features

OpenClerk has four tabs, each a self-contained workflow:

| Tab | What it does | Network calls? |
| --- | --- | --- |
| **Case Law** | Apply or remove hyperlinks by copying them from a source `.docx` that already has case-law citations hyperlinked — the visible text never changes, only the links. | None — parsed entirely in-browser |
| **Non-patent Literature** | Scan the open document for parenthetical citations (e.g. `(Smith v. Jones, 2020)`), then assign a URL to each and apply/remove the resulting hyperlinks. | None |
| **Online Lookup** | An alternative to Case Law's file-based workflow: queries a citation lookup provider live and hyperlinks only the citations that resolve to exactly one case. Ships with free **CourtListener** support out of the box, plus a plugin architecture for enterprise providers (**LexisNexis**, **Westlaw**, **Bloomberg Law**) using your firm's own contracted credentials. See [Citation hyperlink providers](#citation-hyperlink-providers-plugin-architecture). | Opt-in, per-citation only |
| **Bluebook Check** | Scans the document's case citations for common Bluebook mechanical formatting problems (the `"v."` abbreviation, reporter series form, year/court presence, edition-specific case-name abbreviations) across three selectable editions (20th/21st/22nd). Click any flagged citation to jump straight to it in the document. See [Bluebook citation checking](#bluebook-citation-checking-plugin-architecture). | None |
| **Embed Cited Text** | Finds citations that pinpoint a page beyond the first page of the opinion (a single page, a list of pages, or a page range) and embeds the cited court opinion text for that pincite as a Word comment on the citation — comments are collapsed by default and expand on click. See [Embedding cited opinion text](#embedding-cited-opinion-text). | Opt-in, per-citation only |

A **Report an issue** link sits at the bottom of every tab, pointing straight to this repo's [GitHub Issues](https://github.com/OpenClerkProject/openclerk-word/issues).

Most of OpenClerk makes zero network calls, ever. Online Lookup, Find Hallucinations, and Embed Cited Text are the only features that leave the machine, and each one only does so when a user explicitly runs it. See [Security & IT review](#security--it-review) for the full data-flow breakdown.

## Development

Prerequisites
- Node.js (16+ recommended)
- npm

Install dependencies:

```bash
cd openclerk-word
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

the selected provider is asked about that citation; if it immediately resolves to a single case, OpenClerk hyperlinks it, otherwise that citation is left untouched and the scan moves on to the next one — no error, no interruption.

Providers are plugins implementing `CitationProvider` in [openclerk-core: src/providers/types.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/providers/types.ts) and self-registering with `citationProviderRegistry` in [openclerk-core: src/providers/index.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/providers/index.ts). To add a new one (a firm's internal search API, another public case-law database, etc.), implement the interface and register an instance — nothing else in the add-in needs to change.

**Citation scanner known limitations** (validated by hand against a real trial brief, not committed to this repo): the scanner in [citationParser.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/providers/citationParser.ts) handles multi-pincite lists (`503, 505, 508, 513`), pincite ranges (`705-06`), and reporters with embedded digits (`F. Supp. 3d`, `F.4th`) correctly, but does not currently handle footnote pincites (`567 n.1`), a stray typo like a doubled comma before the parenthetical, or parallel citations to a second reporter (`24 Misc. 2d 790, 200 N.Y.S.2d 126 (1960)`) — the last of these gets misparsed (matching the second reporter, treating the first as part of the case name) rather than cleanly skipped. See the "known limitation" tests in `tests/providers.test.ts` for exact behavior.

Built-in providers:
- **CourtListener** ([courtListenerProvider.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/providers/courtListenerProvider.ts)) — Free Law Project's free case-law search. Requires a free account and API token from [courtlistener.com/profile/api-token](https://www.courtlistener.com/profile/api-token/) — CourtListener's citation-lookup API returns `401` for unauthenticated requests, so there's no working anonymous mode despite what some of their documentation examples show.
- **LexisNexis, Westlaw, Bloomberg Law** ([lexisNexisProvider.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/providers/lexisNexisProvider.ts), [westlawProvider.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/providers/westlawProvider.ts), [bloombergLawProvider.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/providers/bloombergLawProvider.ts)) — these are contract-gated enterprise APIs. Each vendor provisions its own base URL and OAuth2 client credentials per customer, so OpenClerk doesn't (and can't) ship a fixed endpoint or key; you supply your API base URL, client ID, and client secret from your firm's contract in the Online Lookup tab. The bundled implementations use the standard OAuth2 client-credentials flow and a `POST <base>/search/cases`-shaped request as a starting point — confirm the exact token and search paths in your vendor's API documentation and adjust the `TOKEN_PATH`/`SEARCH_PATH` constants at the top of each file if they differ.
- **USPTO Patent Center** ([usptoPatentCenterProvider.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/providers/usptoPatentCenterProvider.ts)) — **TODO, not implemented.** Registered as a placeholder (shows up in the provider list, always reports "not found" so citations are skipped) so the plugin wiring can be exercised end-to-end before a real Patent Center / PEDS integration is built for the Non-patent Literature workflow.

### Getting credentials for each provider

**CourtListener** (free, self-service — token required):
1. Create a free account at [courtlistener.com/sign-in](https://www.courtlistener.com/sign-in/) (there's a "Register" link if you don't have one).
2. Once signed in, go to [courtlistener.com/profile/api-token](https://www.courtlistener.com/profile/api-token/) and copy the token shown there.
3. Paste it into the "API token" field on the Online Lookup tab and click Connect. A token is required — CourtListener's citation-lookup API rejects unauthenticated requests, so there's no anonymous fallback.

**LexisNexis** (requires an existing Lexis subscription):
1. Register for a developer account at the [LexisNexis Developer Portal](https://dev.lexisnexis.com/) (use your firm/business email) — see their [Getting Started guide](https://dev.lexisnexis.com/gettingStarted).
2. Request access to the case-law/citation API product for your organization; a LexisNexis data specialist reviews and provisions access (their published turnaround is about 24 hours for trial requests).
3. Once approved, the portal issues a Client ID and Client Secret and documents your tenant's API base URL — enter those three values into the LexisNexis fields on the Online Lookup tab.

**Westlaw** (requires an existing Westlaw/Thomson Reuters subscription):
1. Visit the [Thomson Reuters Developer Portal](https://developers.thomsonreuters.com/) and contact Thomson Reuters support to request API access under your firm's existing Westlaw agreement — this isn't self-service signup; TR provisions it per customer.
2. TR support provides a Client ID, Client Secret, the authentication (token) endpoint, and your API base URL.
3. Enter those into the Westlaw fields on the Online Lookup tab (adjust `TOKEN_PATH` in [westlawProvider.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/providers/westlawProvider.ts) if the endpoint TR gives you doesn't match the default `/oauth/token`).

**Bloomberg Law** (requires an existing Bloomberg Law subscription):
1. Contact your Bloomberg account representative, or Bloomberg Law's help desk (888.560.2529 / help@bloomberglaw.com), to request Web API access for your firm's account.
2. Bloomberg provisions credentials through its Enterprise Console (an admin on your account downloads a Client ID and Secret there — see Bloomberg's [Web API Connectivity Policy](https://www.bloomberg.com/professional/support/api-library/) for what the console shows). Store the downloaded secret securely; per Bloomberg's policy these credentials expire after 18 months and will need to be refreshed then.
3. Enter the Client ID, Secret, and the API base URL Bloomberg gives you into the Bloomberg Law fields on the Online Lookup tab.

Because the three enterprise integrations are contract-gated and provisioned per customer, treat the exact field names/paths above as a starting point, not a guarantee — confirm specifics against the documentation your vendor rep provides.

## Bluebook citation checking (plugin architecture)

The **Bluebook Check** tab scans the current document's case citations for common Bluebook formatting problems and lists any it finds, per citation. It's a mechanical checker, not the rulebook: it verifies conventions like

- `"v."` is used (not `"v"` or `"vs."`) between party names,
- the reporter abbreviation is a recognized, correctly-formatted form (see "Vendored reference data" below) — including the ordinal typo `"2nd"`/`"3rd"` instead of `"2d"`/`"3d"` and a Rule 6.1 spacing mistake (e.g. `"S.Ct."` instead of `"S. Ct."`),
- every full word in the case name that Table T6 abbreviates is actually abbreviated (e.g. `"Company"` → `"Co."`, `"Association"` → `"Ass'n"`),
- a decision year is present in the parenthetical,
- a court abbreviation is present for any reporter other than U.S. Reports (`"U.S."`, where the Supreme Court is implied),
- the court/jurisdiction parenthetical uses the Table T10 state abbreviation rather than the spelled-out state name (e.g. `"California"` → `"Cal."`), and
- a pinpoint page range drops repetitious digits per Rule 3.2 (e.g. `"705-706"` should be `"705-06"`),

plus a small set of edition-specific case-name abbreviations (see below). **It does not check every Bluebook rule** (typeface/italics, signal/authority ordering within a citation string, short-form/`id.` citations, parenthetical-phrase formatting, Lexis/Westlaw or internet-source citations, etc.) and isn't a substitute for the actual rulebook — treat a clean result as "no obvious mechanical problems," not "Bluebook-perfect."

Each citation in the results list is clickable: clicking it searches the document for that exact citation text and selects it, which moves Word's view to show where it actually appears — no manual scrolling/searching needed to find a flagged citation.

### Vendored reference data

The reporter (Table T1), case-name (Table T6), and state (Table T10) abbreviation checks are powered by [reporters-db](https://github.com/freelawproject/reporters-db) (BSD-2-Clause), Free Law Project's open, actively-maintained abbreviation database — the same one that powers CourtListener's own citation parser (`eyecite`). As of this writing it covers **1,342 valid reporter forms + 2,250 known malformed variations**, **231 case-name abbreviations**, and all 50 states.

This data is fetched and trimmed **once, at development time**, by [openclerk-core's scripts/generate-bluebook-data.js](https://github.com/OpenClerkProject/openclerk-core/blob/main/scripts/generate-bluebook-data.js) into `src/bluebook/generated/*.ts` in that repo — those generated files are committed there like any other source file, and this add-in picks them up transitively through its `openclerk-core` dependency. **The add-in itself makes no network call to fetch this data**; it's bundled into the JS just like the hand-written rule-sets, preserving the zero-network-calls-by-default architecture (see [Security & IT review](#security--it-review)). To pick up upstream updates, run `npm run bluebook:update-data` in [openclerk-core](https://github.com/OpenClerkProject/openclerk-core), commit the diff there, cut a new release tag, and bump this repo's `openclerk-core` dependency version in `package.json`.

**Reporter-format checking distinguishes two categories** to avoid false positives: a reporter's independently valid *edition* forms (e.g. `"A."`, `"A.2d"`, `"A.3d"` are all correct — different chronological editions of the same series, not errors) versus its known *variations* (malformed spellings like `"A2d"` or `"Atl.2d"`, which are flagged with the specific correct form). An ordinal typo (`"2nd"`/`"3rd"`) is caught generically for every reporter, not just the ones reporters-db happens to have recorded a variation entry for.

**Found a wrong or missing rule?** [`openclerk-core: src/bluebook/manualCorrections.ts`](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/bluebook/manualCorrections.ts) is a small, hand-maintained file — separate from the generated data above, so it's never overwritten by `bluebook:update-data` — where corrections and additions go. See [CONTRIBUTING.md](CONTRIBUTING.md#contributing-a-bluebook-citation-correction-no-coding-experience-needed) for a walkthrough that assumes no coding background.

Like the citation lookup providers, Bluebook editions are plugins implementing `BluebookRuleSet` in [openclerk-core: src/bluebook/types.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/bluebook/types.ts) and registering with `bluebookRuleSetRegistry` in [openclerk-core: src/bluebook/index.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/bluebook/index.ts). Pick an edition from the dropdown on the Bluebook Check tab; each is checked independently and adding a new edition (or a firm/journal-specific house style) means implementing the interface and registering an instance.

Built-in editions — **20th (2015)**, **21st (2020)**, **22nd (2025, current)**:
- All three run the same shared, edition-stable Rule 10 (case citation) checks listed above — see [commonRules.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/bluebook/commonRules.ts). Per the [University of Washington Law Library's Bluebook 101 guide](https://lib.law.uw.edu/bluebook101/editions) ([22nd edition page](https://lib.law.uw.edu/bluebook101/22nd)), the documented changes across these three editions are concentrated in statutory/online-source citation, typeface terminology, and new source types (audio/video, AI-generated content, state administrative materials) — not in core case-citation format, so there was no accurate edition-specific case-citation rule to invent for most of Rule 10.
- The one genuine, verified case-citation-relevant difference: the **21st edition merged Table T6** (case-name word abbreviations) **with the former Table T13.2** (periodical/institutional-author abbreviations), so several words — e.g. *Laboratory* → `Lab'y`, *Employment*/*Employee* → `Emp.`, and similarly *Environment* → `Env't`, *Research* → `Rsch.`, *Psychology* → `Psych.`, *Sociology* → `Socio.`, *Comparative* → `Compar.` — went from "spelled out in case names, abbreviated only in institutional-author citations" to "abbreviated everywhere." The 20th-edition rule-set doesn't flag these words in case names; the 21st and 22nd do. Sources: the UW guide above and Mary Whisner, [Bluebook Weight Loss Program, Part Two: The Merger of Tables T6 and T13.2](https://citeblog.access-to-law.com/?p=1074) (which specifically confirms the *Laboratory*/*Employment* mappings; the remaining word list is the same category of T6/T13.2-merger abbreviation but hasn't been independently verified word-for-word against the official table — see the sourcing note in [caseNameAbbreviations.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/bluebook/caseNameAbbreviations.ts)).

## Embedding cited opinion text

The **Embed Cited Text** tab finds case citations that pinpoint a specific page beyond the first page of the opinion — Bluebook calls this a pincite, and it shows up as a single page (`, 496`), a comma-separated list (`, 505, 508, 513`), or a page range (`, 705-06`) after the citation's starting page. For each one, it fetches the cited opinion's text at that exact page (or pages) and attaches it to the citation **as a Word comment**.

A Word comment is collapsed to a small icon in the margin by default and expands inline when clicked — that's deliberately reused as the "embedded, expandable/collapsible" mechanism here instead of building a custom widget, since it's a native Word feature that already does exactly this, and the user can review, reply to, resolve, or delete the comments with Word's own comment tools. Click **Remove embedded text** to delete only the comments OpenClerk added (it never touches comments you wrote yourself).

**This requires a provider that can supply full opinion text, not just a hyperlink.** Today only **CourtListener** implements this (see `OpinionTextCapableProvider` in [openclerk-core: src/providers/types.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/providers/types.ts)), and — unlike CourtListener's basic citation-to-hyperlink lookup — it requires an API token; CourtListener's opinion-text endpoints don't have a free anonymous tier. Connect a token first via the Manage Hyperlinks tab's Online Lookup source, then the Embed Cited Text tab will show "Ready" next to the provider once it's usable.

Extracting the exact cited page from an opinion's full text is a best-effort heuristic, not a guarantee: it looks for "star pagination" markers (e.g. `*705`) — the standard convention Westlaw and most public-domain legal text corpora use to mark where each print-reporter page begins — and an opinion's text isn't required to include them (it depends on the opinion's original source). If no markers are found, or none match the requested page, that citation is skipped and reported as such rather than guessing at an excerpt. See [opinionTextExtractor.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/providers/opinionTextExtractor.ts) and [pincitePages.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/providers/pincitePages.ts) for the extraction and page-range-expansion logic, both covered by unit tests against synthetic fixture text.

**CourtListener's default rate limit is modest** — 5 requests/minute, 50/hour, 125/day per [their API documentation](https://www.courtlistener.com/help/api/rest/) — and each pincite citation costs two requests here (resolving the citation to a case, then fetching that case's opinion text). On a document with several pincite citations, later ones can come back rate-limited before earlier ones finish. This is reported distinctly from "not found" (see `OpinionExcerptResult.rateLimited` in [types.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/providers/types.ts)) — wait a minute and click "Embed cited opinion text" again to pick up the rest; already-embedded citations aren't re-fetched or duplicated.

### Testing against the real CourtListener API locally

`npm test` never makes a network call — everything above is verified with mocked `fetch` responses. If you want to sanity-check the real integration (CourtListener's exact field names, HTML markup, and star-pagination conventions can change upstream independently of this repo), there's an opt-in live test suite that's skipped automatically unless you provide a token:

```bash
# macOS / Linux / Git Bash
COURTLISTENER_API_TOKEN=your-token npm run test:live

# Windows PowerShell
$env:COURTLISTENER_API_TOKEN="your-token"; npm run test:live
```

Get a free token from [CourtListener's API documentation](https://www.courtlistener.com/help/api/rest/#authentication). **Never commit a token or put it in a file** — pass it as an environment variable for the one command, the same way you would any other secret. Without `COURTLISTENER_API_TOKEN` set, `npm test` and `npm run test:live` both skip this file entirely (see [courtListener.live.test.ts](tests/courtListener.live.test.ts)), so CI and every other contributor's local run are completely unaffected.

## Security & IT review

### One-page summary

- **What it is:** a Word task-pane add-in (static HTML/JS/CSS rendered inside Word's built-in browser control). It is **not** a hosted web app, has no backend server of its own, no user accounts, and no database.
- **What it touches:** only the Word document that is currently open, and only when the user clicks a button. It never scans other files, other Office apps, mail, or the file system beyond a single `.docx` the user explicitly picks from a file-open dialog.
- **What leaves the machine, by default:** nothing. Two of the add-in's three workflows (**Case Law** file-parsing and **Non-patent Literature** parentheticals) make zero network calls, ever.
- **What can leave the machine, opt-in only:** if a user turns on the **Online Lookup** tab, short citation strings (not document content) are sent over HTTPS to whichever single lookup provider that user selects and connects. See the data-flow table below for exactly what's sent where.
- **Source & license:** fully open source in this repository under the MIT license (see [package.json](package.json)); every claim in this section can be checked directly against [src/taskpane/](src/taskpane/) rather than taken on faith.
- **Dependencies:** `npm audit --omit=dev` currently reports **0 known vulnerabilities** in production dependencies — i.e., in the code that actually ships inside the add-in bundle loaded into Word. Running a plain `npm audit` (no flags) will additionally show advisories in **devDependencies only** (the local build/debugging toolchain: `webpack-dev-server`, `copy-webpack-plugin`, and Microsoft's own `office-addin-manifest`/Teams Toolkit tooling, none of which are packaged into `dist/`). Those are worth tracking but don't affect the shipped add-in; versions are pinned via `package-lock.json` and CI (`npm ci`) installs the exact locked tree on every build.
- **Distribution:** no auto-update mechanism reaches out to any OpenClerk-controlled service. Installing or upgrading is something your organization does deliberately (see [Centralized deployment](#centralized-deployment-it-managed-rollout) below), not something that happens silently in the background.

### Data flow, by feature

| Feature (tab) | Network calls? | What's sent | Where | Trigger |
| --- | --- | --- | --- | --- |
| **Case Law** (apply hyperlinks from a source `.docx`) | None | Nothing — the source file is parsed entirely in-browser with [JSZip](https://www.npmjs.com/package/jszip) | N/A | User clicks "Apply hyperlinks" |
| **Case Law** (remove hyperlinks) | None | Nothing — edits the open document's OOXML in place | N/A | User clicks "Remove hyperlinks" |
| **Non-patent Literature** (scan/add/remove parentheticals) | None | Nothing — citations are extracted from the open document locally; URLs are typed in by the user | N/A | User clicks Scan / Add / Remove |
| **Online Lookup** — CourtListener | Yes (opt-in) | The **individual matched citation string** (e.g. `"Norfolk & W. Ry. Co. v. Liepelt, 444 U.S. 490 (U.S.Ill., 1980)"`) — never the surrounding document text — plus the required API token | `https://www.courtlistener.com` only | User clicks "Scan & hyperlink via API" with CourtListener selected |
| **Online Lookup** — LexisNexis / Westlaw / Bloomberg Law | Yes (opt-in) | The same per-citation string, plus an OAuth2 bearer token obtained from the credentials the user entered | The API base URL *that user typed in*, from their firm's existing vendor contract | User clicks "Connect", then "Scan & hyperlink via API" |
| **Embed Cited Text** — CourtListener | Yes (opt-in) | The individual matched citation string, plus a required API token, to resolve the citation and then fetch that opinion's text | `https://www.courtlistener.com` only | User clicks "Embed cited opinion text" with CourtListener selected |

The extraction that produces those per-citation strings (`extractCaseCitations` in [citationParser.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/providers/citationParser.ts)) runs entirely inside the Word document object model, in-browser; only the short matched substrings — never `context.document.body.getText()`'s full output — are ever passed to `fetch()`. That's checkable directly: every network call OpenClerk makes lives in [openclerk-core's `src/providers/`](https://github.com/OpenClerkProject/openclerk-core/tree/main/src/providers) (a public, open-source dependency this add-in consumes unmodified) and nowhere else in this repo — `word.ts` is the only place here that invokes it.

### Permissions requested

The manifest requests exactly one Office permission: `<Permissions>ReadWriteDocument</Permissions>` (see [manifest.xml](manifest.xml)). In plain English, that grants the add-in the ability to read and edit the currently open Word document through the Office JavaScript API — nothing else. It does **not** grant access to: other open documents, other Office apps (Outlook, Excel, Teams), the file system, the network in general (browser same-origin/CORS rules still apply to any `fetch()` call), or any data outside the current document. This is the lowest permission tier Word add-ins offer above read-only.

### Credential handling

- Credential inputs are rendered as password-masked fields (`<input type="password">`) for every secret-typed field (API tokens, client secrets) — see `renderProviderPanel` in [word.ts](src/taskpane/word.ts).
- Credentials are held in memory only, inside each provider instance (see `EnterpriseCitationProvider` in [base.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/providers/base.ts)). They are never written to `localStorage`, `sessionStorage`, cookies, or any OpenClerk-controlled server, and are cleared by clicking "Disconnect" or by closing/reloading the task pane.
- Enterprise provider API base URLs are required to start with `https://` — OpenClerk refuses to authenticate over plain HTTP, so a credential can't accidentally be sent unencrypted (see the check in `EnterpriseCitationProvider.authenticate`, [base.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/providers/base.ts)).
- **No telemetry or analytics** are collected by the add-in itself, successful or not.
- **`manifest.xml`'s `<AppDomains>`** declares every external domain OpenClerk talks to out of the box (`courtlistener.com`); if your policy requires allow-listing every contacted domain at the network/firewall level, add your firm's specific LexisNexis/Westlaw/Bloomberg Law API host there too before deploying.
- **⚠️ When enabling an enterprise provider, widen `connect-src` in [taskpane.html](src/taskpane/taskpane.html)'s CSP to your firm's *specific* API domain — never to a wildcard (e.g. `https:`).** The add-in only checks that a typed-in API base URL starts with `https://`; it does not allow-list which HTTPS host it's allowed to be. A wildcarded `connect-src` would let a mistyped or malicious base URL send that provider's OAuth credentials to an unintended host. A domain-specific `connect-src` closes this off at the browser level regardless of what URL a user types in.

### Centralized deployment (IT-managed rollout)

Rather than each user sideloading the manifest individually (the default developer flow described above), IT can push OpenClerk to specific users or groups — and control updates — through the **integrated apps** feature of the Microsoft 365 admin center:

1. In the [Microsoft 365 admin center](https://admin.microsoft.com), go to **Settings → Integrated apps**, then **Upload custom apps**.
2. Upload the `manifest.xml` produced by `npm run package` (or a release asset — see [Download and install from GitHub](#download-and-install-from-github) below) and choose which users/groups it's deployed to (`Everyone`, specific users/groups, or just yourself for a pilot).
3. When a new version is approved, IT re-uploads the updated manifest/package and clicks **Update** — end users never see an unreviewed, unapproved version.

See Microsoft's own docs for the full process and requirements: [Deploy add-ins in the Microsoft 365 admin center](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins) and [Requirements for centralized deployment](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/centralized-deployment-of-add-ins).

### Frequently asked (by IT reviewers)

- **Does it ever send a whole document anywhere?** No. The only network calls the add-in ever makes are the per-citation lookups described in the data-flow table above; full document content is never transmitted, by any workflow, under any configuration.
- **Does it phone home, track usage, or call any OpenClerk-controlled service?** No such service exists. There is nothing to phone home to.
- **What if we don't want *any* outbound network calls at all?** Don't turn on the Online Lookup tab — the Case Law and Non-patent Literature workflows (the add-in's original functionality) are 100% local and unaffected by its presence.
- **Can we restrict which external domains it's allowed to reach?** Yes, at the network layer (firewall/proxy allow-list) using the domains named in the data-flow table, and/or at the manifest layer via `<AppDomains>`.
- **Who maintains this and where do we report a security concern?** It's maintained in the open at [github.com/OpenClerkProject/openclerk-word](https://github.com/OpenClerkProject/openclerk-word); file an issue or PR there, or read the source directly — there's no vendor support line to call.

## Compliance considerations

OpenClerk itself is a thin, client-side Word add-in with no hosted service of its own, so heavyweight vendor attestations like **SOC 2 Type II** or **ISO/IEC 27001** — which certify how an organization *operates and hosts* a service over time — don't map cleanly onto "review this add-in's code." OpenClerk does not hold, and does not claim, either certification. What's actually relevant to a law firm's review is more specific to how it's distributed and what it touches:

- If you're evaluating this add-in as a piece of software (rather than as a hosted vendor), Microsoft's [Microsoft 365 App Compliance Program](https://learn.microsoft.com/en-us/microsoft-365-app-certification/overview) is the applicable path for Office Add-ins: **Publisher Attestation** (a self-assessment of security/data-handling practices) or full **Microsoft 365 Certification** (a third-party audit against SOC 2/PCI-DSS/ISO 27001-aligned controls). Neither has been completed for OpenClerk; pursue whichever your IT team requires before broad deployment.
- Firms increasingly vet vendors with standardized questionnaires (e.g., the Shared Assessments **SIG** questionnaire) rather than, or in addition to, certifications — the "Security & IT review" section above is meant to answer those questions directly from the source.
- Because the LexisNexis/Westlaw/Bloomberg Law integrations are opt-in and use *your firm's own* contracted API credentials, your firm's existing vendor agreements and compliance posture with those three vendors govern that data flow — OpenClerk is just the client dialing out to an endpoint you already have a relationship with.

## Performance notes

- The file-parsing (**Case Law**) and parenthetical (**Non-patent Literature**) workflows are entirely local (no network calls) and are unaffected by the Online Lookup feature.
- Online Lookup makes one network round-trip per detected citation, **sequentially** (never in parallel), specifically to stay under each provider's rate limits — CourtListener's [documented default](https://www.courtlistener.com/help/api/rest/), for example, is just 5 requests/minute (50/hour, 125/day) for a free account. A document with many citations will take roughly `(number of citations) × (provider round-trip time)` to finish, and a large document may hit CourtListener's rate limit partway through, leaving the rest unresolved until you wait and re-run; prefer the file-parsing workflow when you already have a hyperlinked source document, since it's effectively instant and has no rate limit at all.
- Being rate-limited is reported distinctly from "not found" wherever a provider supports it (see `RateLimitAwareProvider` in [types.ts](https://github.com/OpenClerkProject/openclerk-core/blob/main/src/providers/types.ts)) — Online Lookup's status line calls out how many citations were rate-limited versus genuinely unresolved, and Find Hallucinations won't call a citation a "possible hallucination" just because a request got throttled before it could check. Re-running Online Lookup after a partial run skips citations that already have a hyperlink instead of re-spending API quota verifying them again.
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

## Self-hosting

By default, production builds point the manifest at the project's GitHub Pages deployment (`https://openclerkproject.github.io/openclerk-word/`), so most users don't need to host anything themselves.

If you'd rather serve the add-in content from your own infrastructure (an internal HTTPS server, Azure Static Web Apps, S3+CloudFront, etc. — useful for IT-managed rollouts that don't want to depend on GitHub Pages), set `OPENCLERK_HOST_URL` before building or packaging:

```bash
OPENCLERK_HOST_URL=https://addins.example.com/openclerk/ npm run build
OPENCLERK_HOST_URL=https://addins.example.com/openclerk/ npm run package
```

This bakes your URL into both `dist/manifest.xml` and the packaged manifest, so the add-in fetches its taskpane, commands, and icons from your host instead of GitHub Pages. You're responsible for uploading the contents of `dist/` to that URL yourself — this repo's CI only deploys to GitHub Pages.

