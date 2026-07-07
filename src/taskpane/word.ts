/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/// <reference types="office-js" />

/* global document, Office, Word */

import JSZip from "jszip";
import {
  normalizeText,
  isLikelyCaseCitation,
  extractParentheticalCitations,
  escapeHtml,
  isSafeHyperlinkUrl,
} from "./utils";

type CitationMap = Map<string, string>;
type ParentheticalEntry = { citation: string; url: string; id: string };

// A source .docx is just a zip file; without limits, a small maliciously-crafted zip can
// decompress to a huge amount of data in memory ("zip bomb") and hang or crash the taskpane.
// Legitimate citation-source documents are always small, so these caps are generous in practice.
const MAX_SOURCE_FILE_BYTES = 20 * 1024 * 1024; // 20 MB compressed upload
const MAX_ZIP_ENTRY_COUNT = 500;
const MAX_DECOMPRESSED_XML_BYTES = 50 * 1024 * 1024; // 50 MB per extracted XML part

// Word's Range.search() rejects search strings over 255 characters. citationText here comes
// from arbitrary hyperlink display text in an untrusted source .docx, so it isn't guaranteed to
// respect that -- and because all searches for a batch are queued before the first sync, one
// oversized citation would otherwise fail the sync and abort every citation in the batch, not
// just the bad one.
const MAX_SEARCH_TEXT_LENGTH = 255;

let sourceCitationMap: CitationMap | null = null;
let parentheticalEntries: ParentheticalEntry[] = [];

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const sideloadMessage = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");
    const sourceFileInput = document.getElementById("source-file") as HTMLInputElement | null;
    const applyButton = document.getElementById("apply-hyperlinks") as HTMLButtonElement | null;
    const removeButton = document.getElementById("remove-hyperlinks") as HTMLButtonElement | null;
    const caseLawTab = document.getElementById("case-law-tab") as HTMLButtonElement | null;
    const nonPatentTab = document.getElementById("non-patent-tab") as HTMLButtonElement | null;
    const scanButton = document.getElementById("scan-parentheticals") as HTMLButtonElement | null;
    const addParentheticalButton = document.getElementById("add-parenthetical-hyperlinks") as HTMLButtonElement | null;
    const removeParentheticalButton = document.getElementById("remove-parenthetical-hyperlinks") as HTMLButtonElement | null;

    if (sideloadMessage) {
      sideloadMessage.style.display = "none";
    }
    if (appBody) {
      appBody.style.display = "flex";
    }
    if (sourceFileInput) {
      sourceFileInput.addEventListener("change", onSourceFileSelected);
    }
    if (applyButton) {
      applyButton.addEventListener("click", applyCaseLawHyperlinksFromSource);
    }
    if (removeButton) {
      removeButton.addEventListener("click", removeCaseLawHyperlinks);
    }
    if (caseLawTab) {
      caseLawTab.addEventListener("click", () => setActiveTab("case-law"));
    }
    if (nonPatentTab) {
      nonPatentTab.addEventListener("click", () => setActiveTab("non-patent"));
    }
    if (scanButton) {
      scanButton.addEventListener("click", scanParentheticalCitations);
    }
    if (addParentheticalButton) {
      addParentheticalButton.addEventListener("click", addParentheticalHyperlinks);
    }
    if (removeParentheticalButton) {
      removeParentheticalButton.addEventListener("click", removeParentheticalHyperlinks);
    }

    setActiveTab("case-law");
  }
});

async function onSourceFileSelected(event: Event) {
  const input = event.target as HTMLInputElement;
  const file = input.files && input.files[0];

  if (!file) {
    setStatus("No source file was selected.");
    return;
  }

  setStatus(`Reading ${file.name}...`);

  try {
    sourceCitationMap = await parseSourceDocument(file);
    if (sourceCitationMap.size === 0) {
      setStatus(`No hyperlink citations were found in ${file.name}.`);
      return;
    }

    setStatus(`Loaded ${sourceCitationMap.size} citation hyperlink(s) from ${file.name}.`);
  } catch (error) {
    setStatus(`Unable to read ${file.name}. ${error instanceof Error ? error.message : String(error)}`);
  }
}

async function applyCaseLawHyperlinksFromSource() {
  if (!sourceCitationMap || sourceCitationMap.size === 0) {
    setStatus("Choose a source .docx file that contains hyperlink citations first.");
    return;
  }

  setStatus("Applying hyperlinks...");

  try {
    await Word.run(async (context) => {
      // Batch all searches/loads/inserts into a handful of context.sync() calls instead of a
      // few per citation -- a document with 50-100+ citations previously meant 200-400+
      // sequential round-trips, which is enough to make the taskpane visibly freeze.
      const citationEntries = Array.from(sourceCitationMap!.entries()).filter(
        ([citationText, url]) =>
          citationText && citationText.length <= MAX_SEARCH_TEXT_LENGTH && url && isSafeHyperlinkUrl(url)
      );

      const searches = citationEntries.map(([citationText, url]) => ({
        url,
        results: context.document.body.search(citationText, { matchCase: false, matchWholeWord: false }),
      }));
      searches.forEach((entry) => entry.results.load("items"));
      await context.sync();

      const matchedItems: { item: Word.Range; url: string }[] = [];
      for (const entry of searches) {
        for (const item of entry.results.items) {
          item.load("text");
          matchedItems.push({ item, url: entry.url });
        }
      }
      await context.sync();

      const candidates = matchedItems.filter(({ item }) => {
        const normalizedText = normalizeText(item.text || "");
        return normalizedText && isLikelyCaseCitation(normalizedText);
      });
      candidates.forEach(({ item }) => item.hyperlinks.load("items"));
      await context.sync();

      let appliedCount = 0;
      for (const { item, url } of candidates) {
        if (item.hyperlinks.items.length > 0) {
          continue;
        }
        const normalizedText = normalizeText(item.text || "");
        if (typeof (item as any).insertHyperlink === "function") {
          (item as any).insertHyperlink(url, normalizedText, Word.InsertLocation.replace);
        } else if (typeof (item as any).insertHtml === "function") {
          const html = `<a href="${escapeHtml(url)}">${escapeHtml(normalizedText)}</a>`;
          (item as any).insertHtml(html, Word.InsertLocation.replace);
        } else {
          // Last-resort: replace with plain text (no hyperlink)
          item.insertText(normalizedText, Word.InsertLocation.replace);
        }
        appliedCount += 1;
      }
      await context.sync();

      setStatus(`Added ${appliedCount} hyperlink(s) to matching case-law citations.`);
    });
  } catch (error) {
    setStatus(`Unable to apply hyperlinks. ${error instanceof Error ? error.message : String(error)}`);
  }
}

async function removeCaseLawHyperlinks() {
  setStatus("Removing hyperlinks...");

  try {
    await removeAllHyperlinks();
  } catch (error) {
    setStatus(`Unable to remove hyperlinks. ${error instanceof Error ? error.message : String(error)}`);
  }
}

async function scanParentheticalCitations() {
  setStatus("Scanning for parenthetical citations...");

  try {
    await Word.run(async (context) => {
      context.document.body.load("text");
      await context.sync();
      const bodyText = context.document.body.text;
      const citations = extractParentheticalCitations(bodyText);
      parentheticalEntries = citations.map((citation, index) => ({
        citation,
        url: "",
        id: `${citation}-${index}`,
      }));
      renderParentheticalEntries();
      if (parentheticalEntries.length === 0) {
        setStatus("No parenthetical citations were found in the current document.");
      } else {
        setStatus(`Found ${parentheticalEntries.length} parenthetical citation(s).`);
      }
    });
  } catch (error) {
    setStatus(`Unable to scan the document. ${error instanceof Error ? error.message : String(error)}`);
  }
}

async function addParentheticalHyperlinks() {
  if (parentheticalEntries.length === 0) {
    setStatus("Scan the document for parenthetical citations first.");
    return;
  }

  setStatus("Adding hyperlinks...");

  try {
    await Word.run(async (context) => {
      const validEntries = parentheticalEntries
        .map((entry) => ({ ...entry, url: entry.url.trim() }))
        .filter(
          (entry) => entry.url && entry.citation.length <= MAX_SEARCH_TEXT_LENGTH && isSafeHyperlinkUrl(entry.url)
        );

      const searches = validEntries.map((entry) => ({
        entry,
        results: context.document.body.search(entry.citation, { matchCase: false, matchWholeWord: false }),
      }));
      searches.forEach((s) => s.results.load("items"));
      await context.sync();

      const matchedItems: { item: Word.Range; entry: ParentheticalEntry }[] = [];
      for (const s of searches) {
        for (const item of s.results.items) {
          matchedItems.push({ item, entry: s.entry });
        }
      }
      matchedItems.forEach(({ item }) => item.hyperlinks.load("items"));
      await context.sync();

      let addedCount = 0;
      for (const { item, entry } of matchedItems) {
        if (item.hyperlinks.items.length > 0) {
          continue;
        }
        if (typeof (item as any).insertHyperlink === "function") {
          (item as any).insertHyperlink(entry.url, entry.citation, Word.InsertLocation.replace);
        } else if (typeof (item as any).insertHtml === "function") {
          const html = `<a href="${escapeHtml(entry.url)}">${escapeHtml(entry.citation)}</a>`;
          (item as any).insertHtml(html, Word.InsertLocation.replace);
        } else {
          item.insertText(entry.citation, Word.InsertLocation.replace);
        }
        addedCount += 1;
      }
      await context.sync();

      setStatus(`Added ${addedCount} hyperlink(s) to parenthetical citations.`);
    });
  } catch (error) {
    setStatus(`Unable to add hyperlinks. ${error instanceof Error ? error.message : String(error)}`);
  }
}

async function removeParentheticalHyperlinks() {
  setStatus("Removing hyperlinks...");

  try {
    await removeAllHyperlinks();
  } catch (error) {
    setStatus(`Unable to remove hyperlinks. ${error instanceof Error ? error.message : String(error)}`);
  }
}

function setActiveTab(tabName: string) {
  const caseLawPanel = document.getElementById("case-law-panel");
  const nonPatentPanel = document.getElementById("non-patent-panel");
  const caseLawTab = document.getElementById("case-law-tab");
  const nonPatentTab = document.getElementById("non-patent-tab");

  if (caseLawPanel && nonPatentPanel && caseLawTab && nonPatentTab) {
    const isCaseLaw = tabName === "case-law";
    caseLawPanel.classList.toggle("active", isCaseLaw);
    nonPatentPanel.classList.toggle("active", !isCaseLaw);
    caseLawTab.classList.toggle("active", isCaseLaw);
    caseLawTab.setAttribute("aria-selected", isCaseLaw ? "true" : "false");
    nonPatentTab.classList.toggle("active", !isCaseLaw);
    nonPatentTab.setAttribute("aria-selected", isCaseLaw ? "false" : "true");
  }

  if (tabName === "non-patent" && parentheticalEntries.length === 0) {
    scanParentheticalCitations();
  }
}

function renderParentheticalEntries() {
  const container = document.getElementById("parenthetical-citation-list");
  if (!container) {
    return;
  }

  container.innerHTML = "";
  if (parentheticalEntries.length === 0) {
    container.innerHTML = '<p class="helper-text">No parenthetical citations found yet.</p>';
    return;
  }

  const fragment = document.createDocumentFragment();
  parentheticalEntries.forEach((entry) => {
    const row = document.createElement("div");
    row.className = "citation-row";

    const label = document.createElement("label");
    label.setAttribute("for", entry.id);
    label.textContent = entry.citation;

    const input = document.createElement("input");
    input.id = entry.id;
    input.className = "url-input";
    input.type = "text";
    input.placeholder = "https://example.com";
    input.value = entry.url;
    input.addEventListener("input", (event) => {
      const target = event.target as HTMLInputElement;
      const matchingEntry = parentheticalEntries.find((item) => item.id === entry.id);
      if (matchingEntry) {
        matchingEntry.url = target.value;
      }
    });

    row.appendChild(label);
    row.appendChild(input);
    fragment.appendChild(row);
  });

  container.appendChild(fragment);
}

function setStatus(message: string) {
  const statusElement = document.getElementById("status");
  if (statusElement) {
    statusElement.textContent = message;
  }
}

async function parseSourceDocument(file: File): Promise<CitationMap> {
  if (file.size > MAX_SOURCE_FILE_BYTES) {
    throw new Error(
      `File is too large (${Math.round(file.size / (1024 * 1024))} MB). The maximum supported size is ${
        MAX_SOURCE_FILE_BYTES / (1024 * 1024)
      } MB.`
    );
  }

  const zip = await JSZip.loadAsync(file);

  const entryCount = Object.keys(zip.files).length;
  if (entryCount > MAX_ZIP_ENTRY_COUNT) {
    throw new Error("The selected file contains an unexpectedly large number of entries and was rejected.");
  }

  const documentXml = await readZipEntryWithLimit(zip, "word/document.xml");
  const relationshipsXml = await readZipEntryWithLimit(zip, "word/_rels/document.xml.rels");

  if (!documentXml || !relationshipsXml) {
    throw new Error("The selected file is not a valid Word document.");
  }

  const relationships = parseRelationships(relationshipsXml);
  const parser = new DOMParser();
  const documentDom = parser.parseFromString(documentXml, "application/xml");
  const hyperlinks = documentDom.getElementsByTagNameNS(
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "hyperlink"
  );

  const citationMap: CitationMap = new Map<string, string>();
  for (let index = 0; index < hyperlinks.length; index += 1) {
    const hyperlink = hyperlinks[index];
    const relationshipId = hyperlink.getAttributeNS(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "id"
    );
    const url = relationshipId ? relationships.get(relationshipId) || "" : "";
    const text = normalizeText(getElementText(hyperlink));

    if (text && url && isLikelyCaseCitation(text) && isSafeHyperlinkUrl(url)) {
      citationMap.set(text, url);
    }
  }

  return citationMap;
}

// Reads a zip entry as text, aborting if the decompressed content is unexpectedly large --
// a defense-in-depth backstop against zip bombs on top of the whole-file and entry-count caps.
async function readZipEntryWithLimit(zip: JSZip, path: string): Promise<string | undefined> {
  const entry = zip.file(path);
  if (!entry) {
    return undefined;
  }
  const content = await entry.async("string");
  if (content.length > MAX_DECOMPRESSED_XML_BYTES) {
    throw new Error("The selected file's contents are unexpectedly large and were rejected.");
  }
  return content;
}

async function removeAllHyperlinks(): Promise<void> {
  setStatus("Removing all hyperlinks...");

  try {
    await Word.run(async (context: Word.RequestContext) => {
      const body = context.document.body as any;
      const ooxmlResult = body.getOoxml();
      await context.sync();

      const originalOoxml = ooxmlResult && ooxmlResult.value ? String(ooxmlResult.value) : "";
      if (!originalOoxml) {
        setStatus("No document OOXML available to process.");
        return;
      }

      // Count <w:hyperlink ...> occurrences to report how many links we'll remove
      const hyperlinkMatches = originalOoxml.match(/<w:hyperlink\b[^>]*>/g);
      const removedCount = hyperlinkMatches ? hyperlinkMatches.length : 0;

      if (removedCount === 0) {
        setStatus("No hyperlinks were found in the document.");
        return;
      }

      // Remove the hyperlink wrapper tags but keep inner runs (visible text)
      let strippedOoxml = originalOoxml.replace(/<w:hyperlink\b[^>]*>/g, "").replace(/<\/w:hyperlink>/g, "");

      // Remove the character style reference that causes blue/underline rendering
      strippedOoxml = strippedOoxml.replace(/<w:rStyle\s+w:val="Hyperlink"\s*\/\>/g, "");
      // Also remove non-self-closing form if present
      strippedOoxml = strippedOoxml.replace(/<w:rStyle\s+w:val="Hyperlink"\s*>\s*<\/w:rStyle>/g, "");

      // Replace entire body OOXML with cleaned version
      body.insertOoxml(strippedOoxml, Word.InsertLocation.replace);
      await context.sync();

      setStatus(`Removed ${removedCount} hyperlink(s) from the document.`);
    });
  } catch (err) {
    setStatus(`Unable to remove hyperlinks. ${err instanceof Error ? err.message : String(err)}`);
    throw err;
  }
}

function parseRelationships(relsXml: string): Map<string, string> {
  const parser = new DOMParser();
  const relsDom = parser.parseFromString(relsXml, "application/xml");
  const relationships = relsDom.getElementsByTagName("Relationship");
  const result = new Map<string, string>();

  for (let index = 0; index < relationships.length; index += 1) {
    const relationship = relationships[index];
    const id = relationship.getAttribute("Id");
    const target = relationship.getAttribute("Target");
    if (id && target) {
      result.set(id, target);
    }
  }

  return result;
}

function getElementText(node: Element): string {
  const textNodes = node.getElementsByTagNameNS(
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "t"
  );
  let result = "";

  for (let index = 0; index < textNodes.length; index += 1) {
    result += textNodes[index].textContent || "";
  }

  return result;
}
