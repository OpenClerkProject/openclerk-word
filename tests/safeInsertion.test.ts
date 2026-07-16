import { toSafeHyperlinkUrl, toSafeHtml } from "openclerk-core";
import { insertSafeHyperlink, insertSafeComment } from "../src/taskpane/safeInsertion";

// Word.InsertLocation.replace is a real Office.js global normally supplied by the WebView host at
// runtime -- it has no presence under Jest's node testEnvironment. safeInsertion.ts's dispatch
// logic references it directly (matching the pre-existing word.ts code being ported), so tests
// exercising that dispatch need a minimal stand-in, not a mocking framework.
(global as any).Word = { InsertLocation: { replace: "Replace" } };

describe("toSafeHyperlinkUrl", () => {
  test("rejects an unsafe scheme", () => {
    expect(toSafeHyperlinkUrl("javascript:alert(1)")).toBeNull();
  });

  test("accepts an https URL", () => {
    expect(toSafeHyperlinkUrl("https://example.com")).toBe("https://example.com");
  });
});

describe("insertSafeHyperlink dispatch", () => {
  const url = toSafeHyperlinkUrl("https://example.com")!;
  const text = toSafeHtml("Smith v. Jones");

  test("calls only insertHyperlink when all three duck-typed methods are present", async () => {
    const insertHyperlink = jest.fn();
    const insertHtml = jest.fn();
    const insertText = jest.fn();
    const item = { insertHyperlink, insertHtml, insertText };
    const context = { sync: jest.fn().mockResolvedValue(undefined) };

    await insertSafeHyperlink(context as any, item as any, url, text);

    expect(insertHyperlink).toHaveBeenCalledTimes(1);
    expect(insertHtml).not.toHaveBeenCalled();
    expect(insertText).not.toHaveBeenCalled();
    expect(context.sync).toHaveBeenCalledTimes(1);
  });

  test("calls only insertHtml when insertHyperlink is unavailable", async () => {
    const insertHtml = jest.fn();
    const insertText = jest.fn();
    const item = { insertHtml, insertText };
    const context = { sync: jest.fn().mockResolvedValue(undefined) };

    await insertSafeHyperlink(context as any, item as any, url, text);

    expect(insertHtml).toHaveBeenCalledTimes(1);
    expect(insertHtml).toHaveBeenCalledWith(
      '<a href="https://example.com">Smith v. Jones</a>',
      "Replace"
    );
    expect(insertText).not.toHaveBeenCalled();
    expect(context.sync).toHaveBeenCalledTimes(1);
  });

  // Regression test for CR-01: a URL containing a double quote must not be able to break out of
  // the href="..." attribute and inject arbitrary markup into the built HTML string.
  test("escapes a quote character in the URL for the insertHtml attribute context", async () => {
    const unsafeUrl = toSafeHyperlinkUrl(
      'https://example.com/"><img src=x onerror=alert(1)>'
    )!;
    const insertHtml = jest.fn();
    const item = { insertHtml };
    const context = { sync: jest.fn().mockResolvedValue(undefined) };

    await insertSafeHyperlink(context as any, item as any, unsafeUrl, text);

    const [html] = insertHtml.mock.calls[0];
    expect(html).toBe(
      '<a href="https://example.com/&quot;&gt;&lt;img src=x onerror=alert(1)&gt;">Smith v. Jones</a>'
    );
  });

  test("calls only insertText when neither insertHyperlink nor insertHtml is available", async () => {
    const insertText = jest.fn();
    const item = { insertText };
    const context = { sync: jest.fn().mockResolvedValue(undefined) };

    await insertSafeHyperlink(context as any, item as any, url, text);

    expect(insertText).toHaveBeenCalledTimes(1);
    expect(context.sync).toHaveBeenCalledTimes(1);
  });

  // Regression tests for audit finding #1: insertSafeHyperlink now owns escaping, taking a plain
  // string displayText. The insertHyperlink and insertText branches are plain-text sinks, so they
  // must receive displayText raw -- feeding them pre-escaped HTML would render literal &amp;/&#39;
  // entities to the user (e.g. "Smith & Jones v. O'Brien" -> "Smith &amp; Jones v. O&#39;Brien").
  // Only the insertHtml branch escapes, and it must escape the raw displayText itself.
  const displayWithSpecials = `Smith & Jones v. O'Brien`;

  test("passes displayText raw to the plain-text insertText fallback (no HTML entities)", async () => {
    const insertText = jest.fn();
    const item = { insertText };
    const context = { sync: jest.fn().mockResolvedValue(undefined) };

    await insertSafeHyperlink(context as any, item as any, url, displayWithSpecials);

    const [storedText] = insertText.mock.calls[0];
    expect(storedText).toBe(displayWithSpecials);
    expect(storedText).toContain("&");
    expect(storedText).not.toContain("&amp;");
    expect(storedText).not.toContain("&#39;");
  });

  test("passes displayText raw to the plain-text insertHyperlink branch (no HTML entities)", async () => {
    const insertHyperlink = jest.fn();
    const item = { insertHyperlink };
    const context = { sync: jest.fn().mockResolvedValue(undefined) };

    await insertSafeHyperlink(context as any, item as any, url, displayWithSpecials);

    const [, storedText] = insertHyperlink.mock.calls[0];
    expect(storedText).toBe(displayWithSpecials);
    expect(storedText).not.toContain("&amp;");
    expect(storedText).not.toContain("&#39;");
  });

  test("HTML-escapes the raw displayText for the insertHtml sink", async () => {
    const insertHtml = jest.fn();
    const item = { insertHtml };
    const context = { sync: jest.fn().mockResolvedValue(undefined) };

    await insertSafeHyperlink(context as any, item as any, url, displayWithSpecials);

    const [html] = insertHtml.mock.calls[0];
    expect(html).toBe(
      `<a href="https://example.com">Smith &amp; Jones v. O&#39;Brien</a>`
    );
  });
});

describe("insertSafeComment", () => {
  test("calls insertComment with the plain text unchanged and syncs exactly once", async () => {
    const text = "An excerpt of the cited opinion.";
    const insertComment = jest.fn();
    const range = { insertComment };
    const context = { sync: jest.fn().mockResolvedValue(undefined) };

    await insertSafeComment(context as any, range as any, text);

    expect(insertComment).toHaveBeenCalledWith(text);
    expect(context.sync).toHaveBeenCalledTimes(1);
  });

  // Regression test for CR-02: insertComment is a plain-text-only Word API, so text containing
  // characters that would be HTML-escaped by escapeHtml/toSafeHtml must reach it completely
  // unchanged -- otherwise cited opinion/case-name text like "Smith & Wesson" or "the Court's
  // holding" would be corrupted into literal HTML-entity text in the user's document.
  test("does not HTML-escape ampersands, quotes, or apostrophes", async () => {
    const text = `Smith & Wesson, "the Court's holding" <emphasis>`;
    const insertComment = jest.fn();
    const range = { insertComment };
    const context = { sync: jest.fn().mockResolvedValue(undefined) };

    await insertSafeComment(context as any, range as any, text);

    const [storedText] = insertComment.mock.calls[0];
    expect(storedText).toBe(text);
    expect(storedText).not.toContain("&amp;");
    expect(storedText).not.toContain("&#39;");
    expect(storedText).not.toContain("&quot;");
  });
});
