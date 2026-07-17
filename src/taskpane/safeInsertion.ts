/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/// <reference types="office-js" />

/* global Word */

import { escapeHtml } from "openclerk-core";
import type { SafeHyperlinkUrl } from "openclerk-core";

/**
 * This file is the only place in openclerk-word allowed to call a raw Office.js insertion API
 * (insertHtml/insertHyperlink/insertComment/insertOoxml) -- enforced by this repo's ESLint
 * no-restricted-syntax guard (a later plan in this phase). insertSafeHyperlink takes the branded
 * SafeHyperlinkUrl for its url (scheme safety is compiler-enforced -- a caller cannot pass an
 * unvalidated URL in even if the ESLint rule were bypassed). Its displayText is a plain string that
 * this wrapper escapes at the point of use: the insertHtml branch HTML-escapes it, while the
 * insertHyperlink/insertText branches are plain-text sinks that must receive it raw -- feeding them
 * pre-escaped HTML would render literal &amp;/&#39; entities to the user (the same plain-text-sink
 * corruption documented on insertSafeComment below). Escaping at the sink, not at the call site,
 * keeps each of the three branches correct for its own context. insertSafeOoxml's ooxml parameter is
 * intentionally plain string (see its own doc comment) since its only current caller re-inserts
 * document-derived, not untrusted, OOXML.
 */
export async function insertSafeHyperlink(
  context: Word.RequestContext,
  item: Word.Range,
  url: SafeHyperlinkUrl,
  displayText: string
): Promise<void> {
  if (typeof (item as any).insertHyperlink === "function") {
    // insertHyperlink is a newer/preview Office.js API with no declaration in the installed
    // @types/office-js version -- this cast bridges that gap and is pre-existing, not new risk
    // introduced by this refactor (RESEARCH.md Pitfall 2). This is a plain-text sink, so
    // displayText is passed raw -- HTML-escaping it here would surface literal entities.
    (item as any).insertHyperlink(url, displayText, Word.InsertLocation.replace);
  } else if (typeof (item as any).insertHtml === "function") {
    // Both interpolations are escaped for THIS HTML sink. url is only branded for scheme safety
    // (SafeHyperlinkUrl certifies http/https/mailto, nothing about the HTML-attribute context it's
    // spliced into here), so escapeHtml on it stops a "><... payload in the URL breaking out of the
    // href attribute. displayText arrives raw (see header) and is HTML-escaped here so it can't
    // inject markup into the anchor's text content.
    const html = `<a href="${escapeHtml(url)}">${escapeHtml(displayText)}</a>`;
    (item as any).insertHtml(html, Word.InsertLocation.replace);
  } else {
    // Last-resort: replace with plain text (no hyperlink). Plain-text sink -- displayText raw.
    item.insertText(displayText, Word.InsertLocation.replace);
  }
  await context.sync();
}

/**
 * The only place in openclerk-word allowed to call Word.Range.insertComment directly -- see the
 * rationale on insertSafeHyperlink above. Unlike insertSafeHyperlink's insertHtml branch,
 * Word.Range.insertComment(commentText: string) is a plain-text-only API -- it does not render
 * HTML. text is intentionally a plain string here, not SafeHtml: passing HTML-escaped content
 * into a plain-text sink would corrupt any cited opinion excerpt or citation string containing
 * &/</>/"/' into literal HTML-entity text instead of the original characters.
 */
export async function insertSafeComment(
  context: Word.RequestContext,
  range: Word.Range,
  text: string
): Promise<void> {
  range.insertComment(text);
  await context.sync();
}

/**
 * The only place in openclerk-word allowed to call Body.insertOoxml directly -- same rationale
 * as insertSafeHyperlink/insertSafeComment above. Currently only used by removeAllHyperlinks
 * (word.ts) to re-insert OOXML derived from the document's own existing content (getOoxml(),
 * regex-stripped of hyperlink wrapper tags) -- not attacker-influenced input -- but the call is
 * centralized here regardless so no raw Office.js insertion API is ever called outside this file,
 * matching the header comment's "only place" claim.
 */
export async function insertSafeOoxml(
  context: Word.RequestContext,
  body: Word.Body,
  ooxml: string
): Promise<void> {
  body.insertOoxml(ooxml, Word.InsertLocation.replace);
  await context.sync();
}
