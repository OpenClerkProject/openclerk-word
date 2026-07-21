import { assertSafeXmlPart, InspectableXmlDocument } from "../src/taskpane/xmlPartGuard";

// assertSafeXmlPart is duck-typed against the two DOM members it inspects (see xmlPartGuard.ts),
// so these fakes stand in for what the browser DOMParser returns without needing a real DOM. Each
// mirrors an actual DOMParser outcome: `doctype` is a DocumentType node or null; a malformed
// parse yields a document containing one or more <parsererror> elements.
function fakeDom(options: { hasDoctype?: boolean; parseErrorCount?: number }): InspectableXmlDocument {
  const parseErrorCount = options.parseErrorCount ?? 0;
  return {
    doctype: options.hasDoctype ? { name: "html" } : null,
    getElementsByTagName(tagName: string) {
      return { length: tagName === "parsererror" ? parseErrorCount : 0 };
    },
  };
}

describe("assertSafeXmlPart", () => {
  test("accepts a well-formed part with no DOCTYPE and no parse error", () => {
    expect(() => assertSafeXmlPart(fakeDom({}), "document body")).not.toThrow();
  });

  test("rejects a part that declares a DOCTYPE (inline-DTD entity-expansion vector)", () => {
    expect(() => assertSafeXmlPart(fakeDom({ hasDoctype: true }), "document body")).toThrow(/DOCTYPE/);
  });

  test("names the offending part in the DOCTYPE rejection message", () => {
    expect(() => assertSafeXmlPart(fakeDom({ hasDoctype: true }), "relationships part")).toThrow(
      /relationships part/
    );
  });

  test("rejects a part DOMParser flagged as malformed (<parsererror> present)", () => {
    expect(() => assertSafeXmlPart(fakeDom({ parseErrorCount: 1 }), "document body")).toThrow(
      /not a valid Word document/
    );
  });

  test("DOCTYPE is rejected even when the rest of the document parsed cleanly", () => {
    // A "billion laughs" payload parses without a <parsererror>; the DOCTYPE check is what stops
    // it, so this must throw before any caller walks the (potentially hugely expanded) document.
    expect(() => assertSafeXmlPart(fakeDom({ hasDoctype: true, parseErrorCount: 0 }), "document body")).toThrow(
      /DOCTYPE/
    );
  });
});
