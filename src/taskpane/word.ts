/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import JSZip from "jszip";

type CitationMap = Map<string, string>;
type ParentheticalEntry = { citation: string; url: string; id: string };

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
      let appliedCount = 0;
      const citationEntries = Array.from(sourceCitationMap!.entries());

      for (const [citationText, url] of citationEntries) {
        if (!citationText || !url) {
          continue;
        }

        const results = context.document.body.search(citationText, {
          matchCase: false,
          matchWholeWord: false,
        });
        results.load("items");
        await context.sync();

        for (const item of results.items) {
          item.load("text");
        }
        await context.sync();

        for (const item of results.items) {
          const normalizedText = normalizeText(item.text || "");
          if (!normalizedText || !isLikelyCaseCitation(normalizedText)) {
            continue;
          }

          const hyperlinks = item.hyperlinks;
          hyperlinks.load("items");
          await context.sync();

          if (hyperlinks.items.length > 0) {
            continue;
          }

          item.insertHyperlink(url, normalizedText, Word.InsertLocation.replace);
          await context.sync();
          appliedCount += 1;
        }
      }

      setStatus(`Added ${appliedCount} hyperlink(s) to matching case-law citations.`);
    });
  } catch (error) {
    setStatus(`Unable to apply hyperlinks. ${error instanceof Error ? error.message : String(error)}`);
  }
}

async function removeCaseLawHyperlinks() {
  setStatus("Removing hyperlinks...");

  try {
    await Word.run(async (context) => {
      const hyperlinks = context.document.body.hyperlinks;
      hyperlinks.load("items");
      await context.sync();

      let removedCount = 0;

      for (const hyperlink of hyperlinks.items) {
        hyperlink.load("text");
      }
      await context.sync();

      for (const hyperlink of hyperlinks.items) {
        const normalizedText = normalizeText(hyperlink.text || "");
        if (!normalizedText || !isLikelyCaseCitation(normalizedText)) {
          continue;
        }

        const range = hyperlink.range;
        range.load("text");
        await context.sync();
        const replacementRange = range.insertText(range.text, Word.InsertLocation.replace);
        replacementRange.font.color = "Auto";
        replacementRange.font.underline = "None";
        await context.sync();
        removedCount += 1;
      }

      setStatus(`Removed ${removedCount} hyperlink(s) from case-law citations.`);
    });
  } catch (error) {
    setStatus(`Unable to remove hyperlinks. ${error instanceof Error ? error.message : String(error)}`);
  }
}

async function scanParentheticalCitations() {
  setStatus("Scanning for parenthetical citations...");

  try {
    await Word.run(async (context) => {
      const bodyText = context.document.body.getText();
      await context.sync();
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
      let addedCount = 0;
      for (const entry of parentheticalEntries) {
        const url = entry.url.trim();
        if (!url) {
          continue;
        }

        const results = context.document.body.search(entry.citation, {
          matchCase: false,
          matchWholeWord: false,
        });
        results.load("items");
        await context.sync();

        for (const item of results.items) {
          const hyperlinks = item.hyperlinks;
          hyperlinks.load("items");
          await context.sync();

          if (hyperlinks.items.length > 0) {
            continue;
          }

          item.insertHyperlink(url, entry.citation, Word.InsertLocation.replace);
          await context.sync();
          addedCount += 1;
        }
      }

      setStatus(`Added ${addedCount} hyperlink(s) to parenthetical citations.`);
    });
  } catch (error) {
    setStatus(`Unable to add hyperlinks. ${error instanceof Error ? error.message : String(error)}`);
  }
}

async function removeParentheticalHyperlinks() {
  if (parentheticalEntries.length === 0) {
    setStatus("Scan the document for parenthetical citations first.");
    return;
  }

  setStatus("Removing hyperlinks...");

  try {
    await Word.run(async (context) => {
      const hyperlinks = context.document.body.hyperlinks;
      hyperlinks.load("items");
      await context.sync();

      let removedCount = 0;
      for (const hyperlink of hyperlinks.items) {
        hyperlink.load("text");
      }
      await context.sync();

      for (const hyperlink of hyperlinks.items) {
        const normalizedText = normalizeText(hyperlink.text || "");
        if (!normalizedText) {
          continue;
        }

        if (!parentheticalEntries.some((entry) => normalizeText(entry.citation) === normalizedText)) {
          continue;
        }

        const range = hyperlink.range;
        range.load("text");
        await context.sync();
        const replacementRange = range.insertText(range.text, Word.InsertLocation.replace);
        replacementRange.font.color = "Auto";
        replacementRange.font.underline = "None";
        await context.sync();
        removedCount += 1;
      }

      setStatus(`Removed ${removedCount} hyperlink(s) from parenthetical citations.`);
    });
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

function normalizeText(value: string): string {
  return value
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function isLikelyCaseCitation(value: string): boolean {
  const normalized = normalizeText(value);
  if (!normalized) {
    return false;
  }

  if (normalized.includes(" v. ") || normalized.includes(" v ")) {
    return true;
  }

  return /\b\d{4}\b/.test(normalized);
}

function extractParentheticalCitations(text: string): string[] {
  const matches = text.match(/\(([^()]{1,120})\)/g) || [];
  const uniqueMatches = new Set<string>();

  matches.forEach((match) => {
    const citation = normalizeText(match.slice(1, -1));
    if (!citation || !/[A-Za-z0-9]/.test(citation)) {
      return;
    }
    uniqueMatches.add(citation);
  });

  return Array.from(uniqueMatches);
}

async function parseSourceDocument(file: File): Promise<CitationMap> {
  const zip = await JSZip.loadAsync(file);
  const documentXml = await zip.file("word/document.xml")?.async("string");
  const relationshipsXml = await zip.file("word/_rels/document.xml.rels")?.async("string");

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

    if (text && url && isLikelyCaseCitation(text)) {
      citationMap.set(text, url);
    }
  }

  return citationMap;
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
