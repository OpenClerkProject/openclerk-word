# OpenClerk Terms of Use

**Last updated:** July 2026

These Terms of Use govern your use of OpenClerk, a Microsoft Word add-in published and maintained
by William Barnhart ("**we**," "**us**," or "**the publisher**"). By installing or using OpenClerk,
you agree to these terms. This document covers your rights and obligations in using the software;
for how OpenClerk handles data, see the separate [Privacy Policy](PRIVACY.md).

## 1. Not legal advice

**OpenClerk is a citation-formatting and hyperlinking tool, not a substitute for professional
judgment.** It does not provide legal advice, and nothing it outputs — a hyperlink, a Bluebook
formatting flag, a "verified"/"not verified" result from the Find Hallucinations feature, or an
embedded excerpt of opinion text — should be relied upon without independent verification.

**Every citation, hyperlink, formatting suggestion, and flagged result must be independently
verified by a licensed attorney before being relied upon in any filing, publication, or other
professional work product.** OpenClerk's checks are mechanical (regular-expression and database
lookups against publicly available data); they are not a comprehensive review against the actual
Bluebook rulebook or against the authoritativeness of any citation's underlying source, and they
can produce false positives, false negatives, or otherwise incomplete results. This disclaimer
appears directly inside the add-in's task pane as well, so it isn't fine print you'd only see here.

## 2. License to use OpenClerk

OpenClerk is open-source software, licensed to you under the [MIT License](LICENSE). The MIT
License already grants you broad rights to use, copy, modify, and redistribute OpenClerk; these
Terms of Use don't restrict those rights, but do describe additional expectations around how the
add-in and the third-party services it can connect to should be used.

## 3. No warranty

OpenClerk is provided **"as is," without warranty of any kind, express or implied**, including but
not limited to the warranties of merchantability, fitness for a particular purpose, and
non-infringement, as stated in the [MIT License](LICENSE). The publisher does not warrant that
OpenClerk will be error-free, uninterrupted, or that any citation-checking, hyperlinking, or
lookup feature will produce accurate or complete results.

## 4. Limitation of liability

To the maximum extent permitted by applicable law, in no event shall the publisher or copyright
holders be liable for any claim, damages, or other liability — whether in an action of contract,
tort, or otherwise — arising from, out of, or in connection with OpenClerk or the use or other
dealings in OpenClerk, including but not limited to reliance on any citation, hyperlink,
formatting check, or verification result it produces.

## 5. Your responsibilities

- **Independent verification.** As stated in Section 1, you are responsible for independently
  verifying every result OpenClerk produces before relying on it professionally.
- **Third-party credentials and services.** If you connect OpenClerk to CourtListener, LexisNexis,
  Westlaw, Bloomberg Law, or any other lookup provider, you are responsible for complying with that
  provider's own terms of service and for the security of any credentials (API tokens, client
  secrets) you enter. OpenClerk transmits those credentials only to the provider you're
  authenticating with, over HTTPS — see the [Privacy Policy](PRIVACY.md) for details — but does not
  control, and is not responsible for, that provider's own handling of your credentials or data
  once received.
- **Compliance with your own organization's policies.** If you use OpenClerk in a professional
  setting (e.g., a law firm), you're responsible for ensuring your use — including which lookup
  providers you connect and what data you send them — complies with your organization's IT,
  security, and client-confidentiality policies. See
  [README.md § Security & IT review](README.md#security--it-review) for the technical detail an IT
  reviewer would need to evaluate this.

## 6. Third-party services

OpenClerk's opt-in lookup features connect to third-party services (CourtListener by default, or
enterprise providers you configure). Those services are operated independently of OpenClerk and
governed by their own terms of service, not these Terms of Use. The publisher does not warrant the
availability, accuracy, or content of any third-party service OpenClerk connects to.

## 7. Changes to OpenClerk and these terms

OpenClerk has no auto-update mechanism that reaches out to any publisher-controlled service —
installing or upgrading it is something you or your organization does deliberately (see
[README.md § Distribution](README.md#one-page-summary)). These Terms of Use may be updated from
time to time; changes will be reflected here with a new "Last updated" date and noted in the
project's [release notes](https://github.com/OpenClerkProject/openclerk-word/releases).

## 8. Contact

Questions about these terms can be raised as a
[GitHub issue](https://github.com/OpenClerkProject/openclerk-word/issues) — there is no separate support line or
vendor contact, since OpenClerk is maintained in the open as an individual open-source project.
