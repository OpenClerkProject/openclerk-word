/*
 * Feature-level insertion smoke test.
 *
 * This automates the security-relevant core of PR #38's manual Word smoke test (02-UAT.md) without a
 * real Word host. The two OpenClerk features that WRITE escaping-sensitive content into the document
 * -- Manage Hyperlinks and Embed Cited Text -- funnel every write through a single sink each:
 * insertSafeHyperlink and insertSafeComment (src/taskpane/safeInsertion.ts). This suite builds a
 * minimal fake Word document that records exactly what those sinks hand to Office.js, then drives the
 * REAL sink code (plus word.ts's REAL comment-content builder) through realistic legal-citation
 * inputs containing &, ', ", <, > -- the characters that CR-01 (href attribute breakout), CR-02, and
 * audit finding #1 (plain-text-sink over-escaping) are all about.
 *
 * What it does NOT cover (honest boundary): whether real Microsoft Word RENDERS the recorded strings
 * as intended. That relies on documented Office.js API behavior (insertHtml renders HTML;
 * insertHyperlink/insertText/insertComment are plain-text), needs a real Office host, and stays a
 * per-release manual check. Bluebook Check and Hallucination Check are analysis features whose logic
 * lives in openclerk-core (tested there) and which do not write escaped content into the body, so
 * they carry none of the rendering risk this suite guards.
 */

// safeInsertion.ts and word.ts reference Word.InsertLocation.replace at module scope; word.ts also
// calls Office.onReady(...) at load. Neither is a real global under Jest's node environment, so stub
// the tiny surface both touch BEFORE requiring the modules (require, not import, to keep this order).
(global as any).Word = { InsertLocation: { replace: "Replace" } };
(global as any).Office = { onReady: () => {}, HostType: { Word: "Word" } };

/* eslint-disable @typescript-eslint/no-var-requires */
import { toSafeHyperlinkUrl } from "openclerk-core";
import { insertSafeHyperlink, insertSafeComment } from "../src/taskpane/safeInsertion";
const { buildEmbeddedCommentContent, citationHasEmbeddedComment } = require("../src/taskpane/word");

// ---------------------------------------------------------------------------------------------
// Fake Word document: records what each insertion sink receives, and lets a test model a host that
// supports only a given subset of the duck-typed insertion methods (insertSafeHyperlink dispatches
// on which of insertHyperlink/insertHtml/insertText is a function).
// ---------------------------------------------------------------------------------------------
type Caps = Array<"insertHyperlink" | "insertHtml">;

interface Recorded {
  hyperlink?: { url: string; display: string; location: string };
  html?: { html: string; location: string };
  text?: { text: string; location: string };
  comment?: string;
}

function makeContext() {
  const state = { syncCount: 0 };
  const context = {
    sync: async () => {
      state.syncCount += 1;
    },
    _state: state,
  };
  return context as any;
}

// A Range that advertises the given optional capabilities. insertText and insertComment are always
// present (base Office.js Range APIs). Every method records its arguments into `rec`.
function makeRange(caps: Caps) {
  const rec: Recorded = {};
  const range: any = {
    insertText: (text: string, location: string) => {
      rec.text = { text, location };
    },
    insertComment: (content: string) => {
      rec.comment = content;
    },
    _rec: rec,
  };
  if (caps.includes("insertHyperlink")) {
    range.insertHyperlink = (url: string, display: string, location: string) => {
      rec.hyperlink = { url, display, location };
    };
  }
  if (caps.includes("insertHtml")) {
    range.insertHtml = (html: string, location: string) => {
      rec.html = { html, location };
    };
  }
  return range;
}

// A realistic case-law citation exercising every character the escaping fixes care about.
const CITATION = `Smith & Jones v. O'Brien, 5 U.S. 137 (1803)`;
const URL = toSafeHyperlinkUrl("https://www.courtlistener.com/opinion/5/smith-v-jones/")!;

describe("Manage Hyperlinks smoke — insertSafeHyperlink into a fake document", () => {
  test("preview host (insertHyperlink): display text reaches the plain-text sink raw", async () => {
    const range = makeRange(["insertHyperlink", "insertHtml"]);
    const context = makeContext();

    await insertSafeHyperlink(context, range, URL, CITATION);

    // insertHyperlink wins the dispatch; it is a plain-text sink, so the citation is passed verbatim.
    expect(range._rec.hyperlink).toEqual({
      url: URL,
      display: CITATION,
      location: "Replace",
    });
    expect(range._rec.html).toBeUndefined();
    expect(range._rec.text).toBeUndefined();
    expect(range._rec.hyperlink!.display).not.toContain("&amp;");
    expect(range._rec.hyperlink!.display).not.toContain("&#39;");
    expect(context._state.syncCount).toBe(1);
  });

  test("HTML fallback host (insertHtml): url attribute-escaped, display HTML-escaped in the anchor", async () => {
    const range = makeRange(["insertHtml"]);
    const context = makeContext();

    await insertSafeHyperlink(context, range, URL, CITATION);

    expect(range._rec.html!.html).toBe(
      `<a href="https://www.courtlistener.com/opinion/5/smith-v-jones/">Smith &amp; Jones v. O&#39;Brien, 5 U.S. 137 (1803)</a>`
    );
    expect(range._rec.hyperlink).toBeUndefined();
    expect(context._state.syncCount).toBe(1);
  });

  test("HTML fallback host: a URL carrying a breakout payload cannot escape the href attribute (CR-01)", async () => {
    const evilUrl = toSafeHyperlinkUrl('https://example.com/"><script>alert(1)</script>')!;
    const range = makeRange(["insertHtml"]);
    const context = makeContext();

    await insertSafeHyperlink(context, range, evilUrl, CITATION);

    // The closing quote / angle brackets from the URL must be entity-encoded, so no <script> tag or
    // stray attribute can materialize in the built HTML string.
    expect(range._rec.html!.html).not.toContain('"><script>');
    expect(range._rec.html!.html).not.toMatch(/<script/i);
    expect(range._rec.html!.html).toContain("&quot;&gt;&lt;script&gt;");
  });

  test("legacy host (insertText only): last-resort plain-text replacement gets the citation raw", async () => {
    const range = makeRange([]);
    const context = makeContext();

    await insertSafeHyperlink(context, range, URL, CITATION);

    expect(range._rec.text).toEqual({ text: CITATION, location: "Replace" });
    expect(range._rec.text!.text).not.toContain("&amp;");
    expect(range._rec.text!.text).not.toContain("&#39;");
    expect(context._state.syncCount).toBe(1);
  });

  test("document round trip: every matched citation range ends up correctly linked", async () => {
    const citations = [
      `Brown v. Board of Ed., 347 U.S. 483 (1954)`,
      `Katz & Co. v. United States, 389 U.S. 347 (1967)`,
      `People ex rel. O'Neil v. "The State"`,
    ];
    const context = makeContext();
    const ranges = citations.map(() => makeRange(["insertHyperlink"]));

    // Mirror the feature's insertion loop: one insertSafeHyperlink call per matched range.
    for (let i = 0; i < citations.length; i += 1) {
      await insertSafeHyperlink(context, ranges[i], URL, citations[i]);
    }

    ranges.forEach((range, i) => {
      expect(range._rec.hyperlink!.display).toBe(citations[i]);
      expect(range._rec.hyperlink!.display).not.toContain("&amp;");
      expect(range._rec.hyperlink!.display).not.toContain("&#39;");
      expect(range._rec.hyperlink!.display).not.toContain("&quot;");
    });
    expect(context._state.syncCount).toBe(citations.length);
  });
});

describe("Embed Cited Text smoke — real content builder into insertSafeComment", () => {
  const raw = CITATION;
  const excerpt = `The Court held "the writ" issues & the petitioner's claim <survives> review.`;

  test("cited opinion text reaches the plain-text comment API with its characters intact (CR-02 / #1)", async () => {
    const content = buildEmbeddedCommentContent(raw, excerpt);
    const range = makeRange([]);
    const context = makeContext();

    await insertSafeComment(context, range, content);

    // The comment stores exactly what the builder produced -- no HTML-entity corruption anywhere.
    expect(range._rec.comment).toBe(content);
    expect(range._rec.comment).toContain("Smith & Jones");
    expect(range._rec.comment).toContain("O'Brien");
    expect(range._rec.comment).toContain(`"the writ"`);
    expect(range._rec.comment).toContain("<survives>");
    expect(range._rec.comment).not.toContain("&amp;");
    expect(range._rec.comment).not.toContain("&#39;");
    expect(range._rec.comment).not.toContain("&quot;");
    expect(range._rec.comment).not.toContain("&lt;");
    expect(context._state.syncCount).toBe(1);
  });

  test("skip-existing detection recognizes the builder's own output (idempotent re-runs)", () => {
    // The feature skips citations that already carry an OpenClerk comment; that check must accept the
    // exact content this builder emits, or re-running would insert duplicate comments.
    const content = buildEmbeddedCommentContent(raw, excerpt);
    expect(citationHasEmbeddedComment(content, raw)).toBe(true);
    expect(citationHasEmbeddedComment(content, `Some Other Citation, 1 U.S. 1`)).toBe(false);
  });
});
