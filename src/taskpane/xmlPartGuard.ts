// Guards for the one place OpenClerk parses untrusted XML: the individual parts of a source
// .docx a user imports on the Case Law tab (see parseSourceDocument / parseRelationships in
// word.ts). The .docx is a ZIP of XML parts; word/document.xml and word/_rels/document.xml.rels
// are read straight from a user-chosen file and handed to the browser DOMParser.
//
// This module holds the rejection rules that go beyond DOMParser's own behavior, kept separate
// so they can be unit-tested without a real DOMParser (the Jest suite runs under the node
// environment, which has no DOM). The shape below is duck-typed to just the two DOM members the
// checks touch, so a test can exercise the logic with plain fakes.

export interface InspectableXmlDocument {
  /** DOMParser sets this to a DocumentType when the input declared a DOCTYPE, otherwise null. */
  doctype: unknown;
  getElementsByTagName(tagName: string): { length: number };
}

/**
 * Rejects an imported .docx XML part that a legitimate Word document never produces. Throws a
 * user-facing Error (surfaced through parseSourceDocument's existing failure handling) when:
 *
 * - The part declares a DOCTYPE. The browser DOMParser never resolves *external* entities, so
 *   classic XXE (reading local files, SSRF via `SYSTEM` entities) is already out of reach here.
 *   But it does expand *internal* general entities declared in an inline DTD, and the expanded
 *   output is not bounded by the decompressed-size cap applied to the raw XML input -- a
 *   nested-entity ("billion laughs" / quadratic-blowup) part could still burn CPU and memory.
 *   Word never emits a DOCTYPE in an OOXML part, so refusing one outright costs nothing, removes
 *   that DoS vector, and keeps this safe even if the parser is ever swapped for a non-browser XML
 *   library that *does* resolve external entities.
 * - DOMParser flagged the input as malformed. parseFromString does not throw on invalid XML; it
 *   returns a document containing a <parsererror> element. Without this check a corrupt or
 *   deliberately malformed part would be read as simply "no citations found" rather than failing
 *   loudly, hiding the problem from the user.
 */
export function assertSafeXmlPart(dom: InspectableXmlDocument, partName: string): void {
  if (dom.doctype !== null) {
    throw new Error(
      `The selected file's ${partName} contains a DOCTYPE declaration, which Word documents never use, and was rejected.`
    );
  }
  if (dom.getElementsByTagName("parsererror").length > 0) {
    throw new Error("The selected file is not a valid Word document.");
  }
}
