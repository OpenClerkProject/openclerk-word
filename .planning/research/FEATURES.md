# Feature Research: Microsoft Partner Center / AppSource Submission

**Domain:** Microsoft 365 Marketplace (AppSource) submission requirements for an Office Word
task-pane add-in
**Researched:** 2026-07-15
**Confidence:** HIGH (all findings sourced directly from current learn.microsoft.com Partner
Center / Office Add-ins documentation, cross-checked against this repo's actual `manifest.xml`,
`webpack.config.js`, and `.github/workflows/ci.yml`; publish/update dates on the source pages
range 2025-09-25 through 2026-07-10, i.e. current as of this research date)

This file answers "what does OpenClerk's code/repo need so a Partner Center submission doesn't
get rejected?" — scoped to a **free, no-account, no-telemetry, entirely client-side** Word
add-in. It intentionally does not re-litigate what OpenClerk already does (see `PROJECT.md`);
it maps Microsoft's requirements onto this specific add-in's profile.

## Feature Landscape

### Table Stakes (Blocks Submission/Certification If Missing)

These are non-negotiable. Missing any one of them is a documented top cause of automatic
rejection or of the manifest failing automated validation before a human reviewer even looks at
the listing.

| Requirement | Why Required | Complexity | Notes |
|---|---|---|---|
| **Manifest `ProviderName` matches Partner Center Publisher name** | Checklist Step 3 + policy 1120.1: "the provider/developer name in the manifest must match the Publisher." Values don't need to be identical but must "clearly represent the same entity." | LOW | **Confirmed gap:** `manifest.xml` line 5 is `<ProviderName>Contoso</ProviderName>` — an unedited Yeoman-template placeholder. Must be changed to the real publisher/org name used when the Partner Center account is enrolled (business/ops decision, out of scope per `PROJECT.md`, but the manifest edit itself is in scope). |
| **Manifest `Description` is real, not placeholder text** | Policy 100.12 (functionality/UI must not "look unfinished") and 1100.1 (clear value proposition). The `<Description>` element is one of the fields Microsoft's own build/validation flow surfaces. | LOW | **Confirmed gap:** `manifest.xml` line 8 is `<Description DefaultValue="A template to get started."/>` — literal Yeoman boilerplate, never replaced. |
| **`GetStarted` ribbon-callout title/description are real, not template text** | Same 100.12 "unfinished UI" concern — this text is user-visible in Word on first add-in use. | LOW | **Confirmed gap:** lines 80/85 still read "Get started with your sample add-in!" / "Your sample add-in loaded successfully..." |
| **`SourceLocation`/task-pane/commands URLs are HTTPS, not `localhost`** | Policy 1100.5 (must be SSL-secured) and 1120.1 (source location must point to a valid web address). Marketplace validation and end-user Office clients must actually be able to load the add-in. | LOW (already solved) | **Not a gap** — confirmed `webpack.config.js` string-replaces `https://localhost:3000/` with the production GitHub Pages URL (`https://openclerkproject.github.io/openclerk-word/`, overridable via `OPENCLERK_HOST_URL`) at build time, and `ci.yml`'s release job greps the packaged manifest to fail the build if `localhost` survives. The *repo-root* `manifest.xml` (dev-mode, localhost) is correct as committed — only the **built/packaged** manifest that actually gets uploaded to Partner Center matters, and that one is already clean. Confirm the Partner Center CI job (`ci.yml` "Publish to Partner Center" step) uploads `openclerk-addin.zip` (the production-built package), not the raw repo manifest — it does. |
| **Unique manifest `Id` (GUID)** | Policy: "app ID in the manifest must be unique," reused across version updates, never regenerated. | LOW (already solved) | `manifest.xml` already has a GUID (`3e0d3ccf-cbc6-4a3c-a29a-75d96be5bf89`). Do not regenerate it on future submissions — reuse it for every version update. |
| **`HighResolutionIconUrl` is a genuine 64×64 asset (task pane/content add-in)** | Store-listing icon spec (`create-effective-office-store-listings`): task pane/content add-ins need a 32×32 regular icon **and a separate 64×64 high-DPI icon**. (Outlook add-ins use 64×64/128×128 instead — different tier, not applicable here.) | LOW | **Confirmed gap:** `manifest.xml`'s `HighResolutionIconUrl` points at `logo-filled-80.png` (80×80 — the ribbon large-icon asset, correct for `Icon.80x80` but wrong size for the store high-DPI field). Only 16×16/32×32/80×80 PNGs exist in `assets/`; no 64×64 variant. Need to export/add `logo-filled-64.png` and repoint `HighResolutionIconUrl` at it. |
| **≥1 store screenshot** | Submission checklist: "One screenshot is required" on the Marketplace listing. | MEDIUM | Not a manifest field — supplied directly in the Partner Center "Marketplace listings" step. Needs a real, redacted (no personal info) screenshot of the task pane in Word. Not currently produced by this repo; needs to be created as part of the submission, likely outside the code repo itself unless the team wants a `docs/store-assets/` folder for review/versioning. |
| **`SupportUrl` is a valid HTTPS URL, not an email address** | Explicitly called out: "This can't be an email address; it must be an https:// URL." Required by policy 1120.1. | LOW (already solved) | `manifest.xml` already points at `https://github.com/OpenClerkProject/openclerk-word/issues` — compliant. |
| **Privacy Policy URL (HTTPS, product-named, separate from ToU)** | Checklist Step 6 + policy 100.6/1100.5: must describe the *app* by name (not just the org/website), be a working link (no 404), and be a genuinely distinct document from Terms of Use — "A Terms of Use policy isn't considered a privacy policy." | already in flight | This is `PRIVACY.md` from PR #27 per `PROJECT.md` — confirm it names "OpenClerk" specifically and is reachable at a live HTTPS URL (e.g. GitHub Pages render of the markdown, not just the raw `.md` file, since Partner Center's link validator likely expects a normal web page, not a raw-content URL). **Verify the URL scheme used for PR #27's links before submission** — a raw `raw.githubusercontent.com` link or bare `.md` file may render but should be checked against how Microsoft's validator fetches/renders it. |
| **EULA / Terms of Use URL (HTTPS) — or accept Microsoft's standard EULA checkbox** | Checklist Step 6: an EULA link is required at submission; Partner Center offers a "Standard Contract" checkbox as a zero-effort alternative to a custom one. | already in flight | `TERMS.md` from PR #27 covers this. Note the **anti-feature** entry below — using Microsoft's standard EULA is a legitimate alternative to custom `TERMS.md` and would resolve the currently-open `§8 Governing Law` placeholder question by making it moot (Microsoft's own EULA governs). Worth raising to the repo owner as an option, though `PROJECT.md` treats resolving `TERMS.md §8` as the intended path. |
| **Certification/testing notes for the reviewer** | Checklist Step 5 + "top 5 rejection reasons" #3: reviewers must be able to fully exercise every feature; missing notes cause automatic rejection ("Applications that don't list clear instructions... will automatically fail"). | MEDIUM | Not a code change, but a **submission-time deliverable this milestone should produce a draft of**: since CourtListener lookup needs a free API token (per `README.md`) and the enterprise providers (Westlaw/LexisNexis/Bloomberg) need firm credentials the reviewer won't have, the notes must (a) give the reviewer CourtListener test credentials or clear self-service signup steps, and (b) explicitly say the enterprise providers are opt-in/BYO-credential and out of scope for a full functional test, pointing the reviewer at the CourtListener-only path and the Bluebook-check/hyperlink features that need zero external accounts. |
| **Disclosure of external-service dependency in the listing Description** | "Top 5 rejection reasons" #4/#5: any add-in requiring an external account or paid service must name it and explain how to get it, directly in the AppSource Description field (not just README). | LOW | OpenClerk's core hyperlinking/Bluebook-check features work with **zero** external accounts; only the "Online Lookup"/"Embed Cited Text" actions need a free CourtListener token. The Marketplace **Description text** (not code) needs a sentence disclosing this, e.g. "Optional live citation lookup uses the free CourtListener API and requires a free CourtListener account/API token; core hyperlinking and Bluebook checks work fully offline." |
| **`AppDomains` / CSP cover every domain the add-in actually calls** | Policy 1100.5 (no undisclosed data transmission) and general functionality-must-match-listing (100.12) — reviewers will exercise the CourtListener flow and expect it to work without unexpected cross-origin failures. | LOW (already solved) | `manifest.xml` already declares `https://www.courtlistener.com` in `AppDomains`, matching the `connect-src` CSP noted in `CONCERNS.md`. No change needed for the default (non-enterprise) submission profile. |
| **Manifest passes `office-addin-manifest validate -p`** (production-mode schema + Marketplace-specific checks) | The checklist explicitly recommends running the *same automated tool Microsoft's own review pipeline uses* before submitting, to "pass our automated testing before you submit." | LOW | `ci.yml` already runs `npx --no-install office-addin-manifest validate manifest.xml` (line 315) in some job — confirm it's run in **production mode** (`-p` flag) against the *built/packaged* manifest (with real URLs, not localhost) somewhere in the pipeline, since `-p` "allows developer information like localhost URLs" to be flagged that plain mode doesn't. |
| **Version number increments on every update** | Policy 1120.1: "version number in the app package updates must be incremented," and checklist: "version number on the submission form matches the version number in the app manifest." | LOW | `manifest.xml` is currently `1.0.0.0` — fine for a first submission; needs a process note (not code) that future releases bump this and keep the submission-form version in sync. |
| **Register for a Microsoft 365 and Copilot Partner Center account/program** | Prerequisite to all of the above. | N/A — explicitly out of scope | Confirmed out of scope per `PROJECT.md` ("business/ops task"). Listed here only so the code-facing checklist above isn't mistaken for the complete picture. |

### Differentiators (Strengthen the Listing, Not Required to Pass Certification)

| Feature | Value Proposition | Complexity | Notes |
|---|---|---|---|
| **Multiple/richer store screenshots + short demo video** | Only one screenshot is technically required; Microsoft's own "craft effective store images" guidance says richer, captioned images with real content (not an empty document) improve discoverability and conversion. | MEDIUM | Good candidate for a `docs/store-assets/` folder with 3-5 annotated screenshots (hyperlinking, Bluebook check results, hallucination-detection flag) captured against a realistic (but scrubbed) sample brief. |
| **Custom EULA tailored to legal-professional users, vs. Microsoft's standard EULA** | A custom `TERMS.md` can address legal-tool-specific disclaimers (no legal advice, accuracy-of-citation-check caveats) that Microsoft's generic standard EULA won't cover. | already chosen | `PROJECT.md` already commits to the custom `TERMS.md` path (PR #27); this is listed as a differentiator because it's the *better* choice for this domain, not because it's required — the "anti-feature" entry below is the fallback if `§8` jurisdiction stays blocked. |
| **Up to 3 Marketplace categories + up to 2 industries selected deliberately** | Submission guide Step 5: pick 1-3 categories and up to 2 industries "to help your customers find your product." No canonical "Legal" category was confirmed in this research pass (Microsoft's public category list — Productivity, Analytics, Security, etc. — is oriented around SaaS/cloud offers more than Office task panes); worth checking the actual category picker inside Partner Center at submission time rather than guessing here. | LOW | Not a code task — a submission-form decision. Flag as an open question for whoever fills out the form (likely "Productivity," possibly with an industry tag if a legal-specific one exists in the live picker). |
| **Localized listing / `DefaultLocale` + `Override` elements for non-English markets** | Policy 1100.7 requires *declaring* supported languages and keeping the experience "reasonably similar" across them, but a single-locale (`en-US`) submission is fully compliant — this is purely upside. | HIGH | Not needed for initial submission; flag as a future-milestone item only if OpenClerk targets non-US legal markets. |
| **Installation deep-link** (`ms-word:https://api.addins.store.office.com/...`) for the README/marketing site once listed | Publish-docs feature: after Marketplace approval, Microsoft supports a "click and run" install link that opens Word and installs the add-in directly, bypassing manual store search. | LOW (post-approval only) | Can't be built until the add-in has a live Marketplace `addInId` — a nice follow-up for `README.md` once approved, not part of this milestone's code changes. |
| **Additional-certification-info PDF for reviewers** | Submission guide: an optional PDF upload (persists across resubmissions) lets you include screenshots/diagrams in reviewer instructions beyond the plain-text Notes for Certification box. | LOW | Useful if the CourtListener signup flow or the "which features work with zero accounts" distinction is hard to describe in plain text; low cost, directly reduces re-submission risk. |

### Anti-Features (Do Not Build — Not Applicable to This Add-in's Profile)

Requirements that exist in Microsoft's policy for *other* kinds of Office-Marketplace offers but
that do not apply to a free, no-account-required, client-side, no-telemetry, desktop/web Word
add-in like OpenClerk. Building any of these for this submission would be wasted effort.

| Would-Be Feature | Why It Seems Relevant | Why It's Not Applicable Here | What To Do Instead |
|---|---|---|---|
| **"Additional purchases" disclosure / checkbox** | Checklist Step 5 explicitly walks through an "additional purchases" flow with license keys and paid-SaaS test credentials. | OpenClerk itself is free with no in-app purchases. The *optional* enterprise providers (Westlaw/LexisNexis/Bloomberg) are BYO-credential integrations with third parties the user already pays separately for — not something OpenClerk sells or charges for. | Leave the "Does your app require additional purchases?" box **unchecked** at submission; disclose the optional BYO-credential providers in the listing Description (a Table Stakes item above) instead, which is a different, lighter-weight requirement (1100.1's "extra charge" disclosure is about *the app itself* charging, and does not clearly apply when the third-party service is independently licensed by the customer). |
| **Sign-in / sign-up / sign-out UI in the add-in chrome** | Policy 1100.5 + "top 5 rejection reasons" #4 require "clear and simple sign in/sign out and sign-up" UI for add-ins "that depend on external accounts or services." | OpenClerk's core value (hyperlinking, Bluebook checks) requires **no account at all**; only the optional CourtListener/enterprise lookup needs a token, entered as a settings field, not a full OAuth-style account system. | Don't build a sign-in/sign-out flow. If a reviewer flags the CourtListener API-token field as needing "sign-in affordance," the fix is a one-line UI/copy clarification (e.g., a "Get a free API token →" link next to the token field), not an authentication system — flag as a possible **Moderate**-priority pitfall to watch for during review, not a feature to pre-build speculatively. |
| **Microsoft Entra ID / SSO integration + fallback-auth certification notes** | Checklist Step 3 and policy 1120.3 both have dedicated SSO sections requiring fallback-auth documentation. | OpenClerk has no SSO of any kind — it's entirely client-side with optional per-user API tokens, not Microsoft-identity-based auth. | Answer "No" to the Entra ID/SSO question in Partner Center's Product Setup step; no code or docs needed. |
| **Apple ID / iOS App Store compliance, Android support** | Policy 1120.2 has a full mobile-specific subsection (no in-app purchase UI, Apple Terms acceptance, Apple ID in account settings). | `PROJECT.md`/`README.md` describe OpenClerk as a Word desktop/web add-in for Windows/macOS; there's no indication of Outlook-on-Android/iOS support (mobile-add-in support is Outlook-only per Microsoft's own platform docs), and Word doesn't have the same Android/iOS add-in surface as Outlook. | Answer "No" to iOS/Android listing questions; skip Apple ID account setup entirely. |
| **PCI DSS / payment-processor compliance (policy 100.11)** | General marketplace security policy mentions credit-card handling requirements. | OpenClerk has no payment flow of any kind — free product, no in-app purchases. | N/A — not applicable, no action needed. |
| **Excel custom-functions `helpUrl` + custom-function certification notes (policy 1120.5)** | This policy section is explicitly numbered alongside the Office Add-in policies researched here. | OpenClerk is a **Word** task-pane add-in; it defines no Excel custom functions. | N/A — policy section doesn't apply to this Hosts declaration (`<Host Name="Document"/>` = Word only). |
| **Unified (JSON/Teams-style) manifest migration** | Microsoft's newer docs promote the unified manifest as "the modern approach" and it's what new Teams/Copilot-combined apps use. | Research found no current mandatory-migration deadline for the add-in-only XML manifest for a plain Word task-pane add-in submitted solely to Microsoft Marketplace/AppSource (not combined with a Teams app or Copilot agent) — the add-in-only manifest remains fully supported, and AppSource auto-derives one from a unified manifest anyway for legacy-platform compatibility. Confidence: MEDIUM (no explicit "will never deprecate for XML-only submissions" statement found, just absence of a current deadline). | Keep the existing add-in-only `manifest.xml` for this submission. Flag unified-manifest migration as a **possible future milestone** only if OpenClerk later wants to bundle a Teams/Copilot surface — not blocking for this one. |
| **Ad-supported-app guidelines compliance** | Marketplace listing guidance has a full section on ad placement/behavior rules. | OpenClerk has no ads. | N/A. |
| **International-accessibility-standard deep audit as a *new* body of work** | Checklist explicitly requires accessibility conformance. | This is real but is a **cross-cutting engineering concern**, not a Marketplace-listing-specific feature to scope here — treat as its own accessibility-audit backlog item if not already covered, not something invented fresh for this compliance milestone. | Out of scope for this FEATURES.md; note as a possible follow-up if no accessibility pass exists yet (not confirmed either way in this research pass — check `CONCERNS.md`/existing test coverage separately). |

## Feature Dependencies

```
Partner Center account enrollment (out of scope, business/ops)
    └──requires──> Manifest ProviderName fixed (Contoso → real publisher name)
                       └──blocks──> Any Partner Center submission attempt (hard validation gate)

Manifest Description + GetStarted text fixed
    └──independent of──> ProviderName fix (both are "finish the Yeoman template" cleanup,
                          can land in the same PR/commit)

HighResolutionIconUrl repointed to a real 64x64 asset
    └──requires──> New 64x64 PNG exported from existing logo-filled.svg/.png source
                       └──enhances──> Store listing "correctly sized icon" compliance (policy 1100.5)

PRIVACY.md / TERMS.md live HTTPS URLs (PR #27)
    └──requires──> TERMS.md §8 governing-law jurisdiction resolved (per PROJECT.md, deferred to
                    repo owner/counsel)
    └──blocks──> Partner Center listing submission (privacy + EULA URLs are hard-required fields)

Certification/testing notes (CourtListener test token + "enterprise providers are BYO/opt-in")
    └──requires──> A working free CourtListener account to hand the reviewer credentials for
    └──enhances──> Reduces risk of "insufficient testing instructions" rejection (top-5 reason)

Screenshot(s) for Marketplace listing
    └──independent──> can be produced any time before final submission, no code dependency

Unified-manifest migration (deferred, anti-feature for this milestone)
    └──conflicts with──> "ship this milestone quickly" — explicitly deferred, do not start
```

### Dependency Notes

- **Manifest ProviderName fix blocks submission entirely:** this is the single highest-priority
  code change in this research — Microsoft's own validation checks it, and it's currently a
  literal `Contoso` placeholder never touched since the Yeoman scaffold. It must land before any
  Partner Center submission attempt succeeds.
- **PRIVACY.md/TERMS.md block submission independently of the manifest fixes** — both must be
  live HTTPS URLs (per `PROJECT.md`, already drafted in PR #27) before the "Legal and support
  info" step of the submission form can be completed. The `§8` jurisdiction placeholder is the
  one open item blocking `TERMS.md` from being submission-ready.
- **Icon-size fix and certification-notes drafting are independent of each other and of the
  manifest text fixes** — all three can be done in parallel, none blocks the others.
- **Unified-manifest migration explicitly does not belong in this milestone** — flagged as an
  anti-feature/deferred item so it isn't accidentally pulled in during roadmap phase-splitting.

## MVP Definition

### Launch With (v1 — required for a submission attempt to have a realistic chance of passing)

- [ ] Fix `manifest.xml` `ProviderName` (Contoso → real publisher name matching Partner Center
      enrollment) — blocks submission outright if missed
- [ ] Fix `manifest.xml` `Description` (remove "A template to get started.")
- [ ] Fix `GetStarted.Title`/`GetStarted.Description` template text
- [ ] Add a real 64×64 icon asset and repoint `HighResolutionIconUrl` at it (currently mis-sized
      at 80×80)
- [ ] Land PR #27 (`PRIVACY.md`/`TERMS.md`) with `§8` governing-law jurisdiction resolved —
      required submission-form fields
- [ ] Confirm the Partner Center CI publish step uploads the production-built package (URLs
      already verified clean at build time) and that `office-addin-manifest validate -p` runs
      against it
- [ ] Draft certification/testing notes text (CourtListener free-token path + enterprise-provider
      opt-in disclosure) — even though final submission is out of scope, the milestone should
      leave this ready to paste into Partner Center
- [ ] Draft the AppSource listing Description text disclosing the optional CourtListener/
      enterprise-provider external-service dependency

### Add After Validation (v1.x — do once the above is confirmed sufficient, or if Microsoft
feedback on a real submission attempt asks for it)

- [ ] Produce 3-5 real store screenshots (beyond the one required minimum)
- [ ] Decide Marketplace category/industry tags by checking the live Partner Center picker
- [ ] Optional "Additional certification info" PDF for reviewers if plain-text notes prove
      insufficient on a first submission round

### Future Consideration (v2+ — explicitly not this milestone)

- [ ] Unified (JSON) manifest migration — only relevant if OpenClerk later adds a Teams/Copilot
      surface; no current deadline forcing this
- [ ] Non-English localized listing
- [ ] Post-approval installation deep-link in `README.md`/marketing (needs a live `addInId`
      first)

## Feature Prioritization Matrix

| Feature | User Value | Implementation Cost | Priority |
|---|---|---|---|
| Fix `ProviderName` placeholder | N/A (compliance gate) | LOW | P1 |
| Fix `Description`/`GetStarted` placeholder text | LOW (cosmetic, but policy-relevant) | LOW | P1 |
| Add real 64×64 icon, fix `HighResolutionIconUrl` | LOW (cosmetic, but policy-relevant) | LOW | P1 |
| Land PRIVACY.md/TERMS.md (resolve §8) | N/A (compliance gate) | MEDIUM (jurisdiction decision, not code) | P1 |
| Confirm production manifest is what CI publishes; run `-p` validation | N/A (compliance gate) | LOW | P1 |
| Draft certification/testing notes | N/A (reduces rejection risk) | LOW | P1 |
| Draft listing Description with dependency disclosure | N/A (reduces rejection risk) | LOW | P1 |
| Extra screenshots | MEDIUM (better conversion) | MEDIUM | P2 |
| Category/industry selection | LOW | LOW | P2 |
| Additional-info PDF for reviewers | LOW (contingency only) | LOW | P3 |
| Unified manifest migration | LOW (no current forcing function) | HIGH | P3 (deferred) |

**Priority key:**
- P1: Must have — blocks a Partner Center submission from passing or materially raises rejection risk
- P2: Should have, strengthens the listing once P1 items are done
- P3: Nice to have / explicitly deferred

## Sources

- [Microsoft 365 app publishing checklist](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/checklist) — updated 2025-09-25
- [Microsoft Marketplace step-by-step submission guide](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/add-in-submission-guide) — updated 2025-09-25
- [Create effective listings in Microsoft Marketplace and within Microsoft 365 app stores](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/create-effective-office-store-listings) — updated 2025-09-25 (icon sizes, description length limits)
- [Certification policies for Microsoft Marketplace (1100 Microsoft 365, 1120 Office Add-ins, 100.x general)](https://learn.microsoft.com/en-us/legal/marketplace/certification-policies)
- [Requirements for running Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/requirements-for-running-office-add-ins) — updated 2025-11-06
- [Office Add-ins manifest overview](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests) — updated 2026-03-23
- [Validate an Office Add-in's manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-manifest) — updated 2026-06-09
- [Publish your Office Add-in to Microsoft Marketplace](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-office-add-ins-to-appsource) — updated 2026-07-10
- [Office Add-ins with the unified app manifest for Microsoft 365](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/unified-manifest-overview)
- [Top 5 AppSource validation errors for Office Add-ins submissions – December 2023](https://devblogs.microsoft.com/microsoft365dev/top-5-appsource-validation-errors-for-office-add-ins-submissions-december-2023/) (recurring monthly series; content stable across 2023-2025 editions checked)
- This repo, read directly during research: `manifest.xml`, `webpack.config.js`, `.github/workflows/ci.yml`, `README.md`, `.planning/codebase/CONCERNS.md`, `.planning/PROJECT.md`

---
*Feature research for: Microsoft Partner Center / AppSource submission — OpenClerk (Word add-in)*
*Researched: 2026-07-15*
