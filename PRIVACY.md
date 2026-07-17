# OpenClerk Privacy Policy

**Last updated:** July 2026

This Privacy Policy describes what data OpenClerk, a Microsoft Word add-in, collects, uses, and
shares. It's separate from — and should be read alongside — [OpenClerk's Terms of Use](TERMS.md).

> **Short version:** OpenClerk has no backend, no user accounts, and collects no personal
> information or telemetry. The only data that ever leaves your computer is a short citation
> string sent to a legal-research service — and only when you explicitly turn that feature on.
> Everything below is checkable directly against the source code; nothing here is a claim you
> have to take on faith.

## Who this policy covers

This policy covers OpenClerk itself — the add-in code that runs inside Microsoft Word, published
and maintained by William Barnhart at
[github.com/OpenClerkProject/openclerk-word](https://github.com/OpenClerkProject/openclerk-word). OpenClerk is open source
under the [MIT License](LICENSE); the entire codebase is public, so every statement in this policy
can be verified by reading [src/taskpane/](src/taskpane/) rather than taken on trust.

## What OpenClerk is

OpenClerk is a task-pane add-in: static HTML/JavaScript/CSS that Word loads and runs inside its
own built-in browser control. It is **not** a hosted web application. There is no OpenClerk
server, no database, no user account system, and no way for OpenClerk's publisher to see what
documents you open, what you type, or how you use the add-in.

## Data OpenClerk itself collects

**None.** OpenClerk collects no personal information, no usage analytics, no crash reports, and no
telemetry of any kind — successful or failed. There is nothing to opt out of because there is
nothing being collected in the first place. See the "One-page summary" in
[README.md § Security & IT review](README.md#security--it-review) for the technical basis of this
claim.

## What OpenClerk touches on your computer

OpenClerk only reads and edits the Word document you currently have open, and only when you click
a button to trigger a specific action (e.g., "Apply hyperlinks," "Check citations," "Find
Hallucinations"). It does not scan other open documents, other Office applications, your email, or
any part of your file system beyond a single `.docx` file you explicitly choose from a file-picker
dialog for the "Case Law" (file-based) workflow. The Word permission OpenClerk requests is limited
to reading and editing the currently open document (`ReadWriteDocument` — see
[README.md § Permissions requested](README.md#permissions-requested)); it cannot access anything
broader than that.

## Data sent outside your computer — opt-in only

By default, OpenClerk makes **zero network calls**. Several of its workflows (Case Law
file-parsing, Non-patent Literature parenthetical hyperlinking, and Bluebook citation-format
checking) are entirely local and never transmit anything, under any configuration.

Two features are opt-in and, only once you explicitly enable them, send data to a third-party
legal-research service you choose:

- **Online Lookup** (part of Manage Hyperlinks) and **Find Hallucinations**: when you select a
  provider and run a scan, OpenClerk sends the **text of individual matched citations** (e.g.
  `"Norfolk & W. Ry. Co. v. Liepelt, 444 U.S. 490 (U.S.Ill., 1980)"`) — never the surrounding
  document text, and never the document as a whole — to that provider's API, over HTTPS. By
  default this is [CourtListener](https://www.courtlistener.com), a free legal-research service
  operated by the nonprofit Free Law Project. You can instead configure OpenClerk to send these
  same citation strings to LexisNexis, Westlaw, or Bloomberg Law using your own firm's contracted
  API credentials.
- If you supply an API token or other credentials for one of these providers, that credential is
  sent (over HTTPS only — OpenClerk refuses plain HTTP) to the provider you're authenticating
  with, to establish your session with their service.

See [README.md § Data flow, by feature](README.md#data-flow-by-feature) for the exact table of
what's sent where, and [citationParser.ts](src/taskpane/providers/citationParser.ts) /
[src/taskpane/providers/](src/taskpane/providers/) for the code that produces and sends it — every
network call OpenClerk ever makes lives in that one folder.

Because these lookups query **third-party services you choose**, that provider's own privacy
policy governs how they handle the citation text and any credentials you send them. For
CourtListener, see [Free Law Project's Terms of Service and
Policies](https://www.courtlistener.com/terms/). For LexisNexis, Westlaw, or Bloomberg Law, refer
to your firm's existing agreement with that vendor.

## How credentials are handled

- Credential fields (API tokens, client secrets) are rendered as password-masked inputs.
- Credentials are held **in memory only**, for the current task-pane session. They are never
  written to `localStorage`, `sessionStorage`, cookies, or any OpenClerk-controlled server, and are
  cleared when you click "Disconnect" or close/reload the task pane.
- See [README.md § Credential handling](README.md#credential-handling) for the exact code
  references backing each of these claims.

## Data retention

OpenClerk retains nothing, because it stores nothing — there is no OpenClerk-operated server or
database for data to be retained on. Any data sent to a third-party provider (CourtListener, or
your firm's enterprise legal-research vendor) is retained according to that provider's own privacy
policy and retention practices, not OpenClerk's.

## Children's privacy

OpenClerk is a professional legal-research tool intended for use by attorneys, paralegals, and
legal professionals. It is not directed at, and is not knowingly used by, children under 13.

## Changes to this policy

If this policy changes, the update will be reflected here with a new "Last updated" date and noted
in the project's release notes at
[github.com/OpenClerkProject/openclerk-word/releases](https://github.com/OpenClerkProject/openclerk-word/releases).

## Contact

Questions about this policy or OpenClerk's data handling can be raised as a
[GitHub issue](https://github.com/OpenClerkProject/openclerk-word/issues) — there is no separate support line or
vendor contact, since OpenClerk is maintained in the open as an individual open-source project.
