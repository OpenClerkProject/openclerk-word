<!--
  Internal submission checklist / reviewer notes for publishing OpenClerk to Microsoft AppSource
  via Partner Center. This file is NOT shipped inside the add-in bundle; it exists so the person
  filling out the Partner Center submission form has the exact values and testing instructions in
  one place. Items marked  <FILL IN>  need a human decision before submitting.
-->

# OpenClerk — Microsoft AppSource / Partner Center submission notes

OpenClerk is a client-side Word task-pane add-in (no backend). This document collects everything the
Partner Center submission form asks for, plus the "Notes for Certification" a Microsoft validator
needs to test every feature. Fill in the `<FILL IN>` items before submitting.

## 1. Offer identity

| Field | Value |
|---|---|
| Add-in name (`DisplayName`) | **OpenClerk** |
| Publisher / `ProviderName` | **William Barnhart** — *must exactly match the Partner Center account's Publisher display name.* `<FILL IN: confirm>` |
| Add-in ID (`Id`) | `3e0d3ccf-cbc6-4a3c-a29a-75d96be5bf89` |
| Version | `1.0.0.0` |
| Host / platforms | Word — desktop (Windows & Mac), Word on the web. Requires **WordApi 1.4+** (declared in the manifest). |

## 2. Required hosted URLs

All are served from GitHub Pages (`https://openclerkproject.github.io/openclerk-word/`, overridable
via `OPENCLERK_HOST_URL` at build time). **Confirm each returns HTTP 200 over HTTPS before submitting.**

| Partner Center field | URL |
|---|---|
| Privacy policy URL | `https://openclerkproject.github.io/openclerk-word/privacy.html` |
| Terms of use URL | `https://openclerkproject.github.io/openclerk-word/terms.html` |
| Support URL | `https://github.com/OpenClerkProject/openclerk-word/issues` |
| Task pane source (`SourceLocation`) | `https://openclerkproject.github.io/openclerk-word/taskpane.html` |

> The privacy and terms pages are generated from `PRIVACY.md` / `TERMS.md` by `scripts/build-docs.js`
> and deployed with the rest of the site. They are separate documents (a Microsoft requirement) and
> each names the product explicitly.

## 3. Notes for Certification (reviewer testing instructions)

Paste this into the Partner Center **"Notes for certification"** box. Insufficient testing
instructions are one of the most common AppSource rejection reasons — so be explicit.

**No sign-in is required to test the core add-in.** Three workflows are fully local and need no
account, credentials, or network access — a reviewer can test them immediately:

1. **Manage Hyperlinks → from a source .docx** — Open the provided test document (`<FILL IN: attach a
   sample .docx with a few case citations>`). Open OpenClerk, choose the source file, click *Apply
   hyperlinks*. Case-law citations become hyperlinks.
2. **Bluebook Check** — On the Bluebook tab, pick an edition and click *Check citations*. Formatting
   issues in the open document are listed.
3. **Parenthetical hyperlinking** — On Manage Hyperlinks, scan and add parenthetical citation links.

**Two workflows are opt-in and query a third-party legal-research service** (see §4). To test them,
the reviewer must connect a provider:

4. **Online Lookup / Find Hallucinations** — On the provider panel, select **CourtListener**, enter
   the API token below, click *Connect*, then run *Find Hallucinations* (or online *Apply hyperlinks*).
   Real citations are marked verified; unverifiable ones are flagged.
5. **Embed Cited Text** — With CourtListener connected, run *Embed Cited Text* on a document
   containing a pinpoint citation; the cited opinion excerpt is inserted as a Word comment.

- **Test CourtListener API token (free, read-only):** `<FILL IN: provide a working token, or a free
  account the reviewer can use>` — obtainable at https://www.courtlistener.com/help/api/rest/.
- **Sign out:** click *Disconnect* on the provider panel. Credentials are held in memory only and are
  cleared on disconnect or when the task pane reloads (never written to disk/localStorage).

## 4. Dependency & purchase disclosures

Declare these on the submission form (undisclosed dependencies / paid access are rejection reasons):

- **CourtListener** (Free Law Project) — free third-party API used by the default lookup/hallucination
  features. Requires a free API token. Governed by CourtListener's own terms.
- **Westlaw / LexisNexis / Bloomberg Law** — optional enterprise providers. These require the
  **reviewer's or user's own paid, contracted credentials**; OpenClerk does not provide them and does
  not charge for them. Disclose as an external account dependency, not an in-app purchase.
- OpenClerk itself is **free and open source (MIT)** — no in-app purchases, no OpenClerk-operated
  backend, no telemetry.

## 5. Pre-submission checklist

- [ ] `ProviderName` in `manifest.xml` matches the Partner Center Publisher name exactly.
- [ ] `TERMS.md` § 8 governing-law jurisdiction filled in (currently a placeholder). `<FILL IN>`
- [ ] Privacy, Terms, Support, and Task pane URLs all return 200 over HTTPS on the live site.
- [ ] Production manifest contains **no** `localhost` URLs (CI "Verify production manifest URLs" gate).
- [ ] `npx office-addin-manifest validate manifest.xml` passes (run where the Microsoft validation
      endpoint is reachable — it is blocked in the sandboxed CI/agent network).
- [ ] Store logo (300×300+) and 1–5 screenshots prepared for the Partner Center listing.
- [ ] A sample test `.docx` and a working CourtListener token attached to the certification notes.
- [ ] Long description, category (Productivity / Legal), and search keywords drafted.
