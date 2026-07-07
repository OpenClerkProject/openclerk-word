/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/// <reference types="office-js" />

/* global document, Office, Word */

import JSZip from "jszip";
import { normalizeText, isLikelyCaseCitation, extractParentheticalCitations, escapeHtml } from "./utils";
import { citationProviderRegistry, parseCaseCitation, extractCaseCitations, CitationProvider } from "./providers";

type CitationMap = Map<string, string>;
type ParentheticalEntry = { citation: string; url: string; id: string };
type TabId = "case-law" | "non-patent" | "online-lookup";

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
    const onlineLookupTab = document.getElementById("online-lookup-tab") as HTMLButtonElement | null;
    const providerSelect = document.getElementById("provider-select") as HTMLSelectElement | null;
    const providerConnectButton = document.getElementById("provider-connect") as HTMLButtonElement | null;
    const providerDisconnectButton = document.getElementById("provider-disconnect") as HTMLButtonElement | null;
    const applyOnlineHyperlinksButton = document.getElementById("apply-online-hyperlinks") as HTMLButtonElement | null;

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
    if (onlineLookupTab) {
      onlineLookupTab.addEventListener("click", () => setActiveTab("online-lookup"));
    }
    if (providerSelect) {
      providerSelect.addEventListener("change", renderProviderPanel);
    }
    if (providerConnectButton) {
      providerConnectButton.addEventListener("click", connectSelectedProvider);
    }
    if (providerDisconnectButton) {
      providerDisconnectButton.addEventListener("click", disconnectSelectedProvider);
    }
    if (applyOnlineHyperlinksButton) {
      applyOnlineHyperlinksButton.addEventListener("click", applyHyperlinksViaProvider);
    }

    populateProviderSelect();
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

async function applyHyperlinkToItem(
  context: Word.RequestContext,
  item: Word.Range,
  url: string,
  displayText: string
): Promise<void> {
  if (typeof (item as any).insertHyperlink === "function") {
    (item as any).insertHyperlink(url, displayText, Word.InsertLocation.replace);
  } else if (typeof (item as any).insertHtml === "function") {
    const html = `<a href="${escapeHtml(url)}">${escapeHtml(displayText)}</a>`;
    (item as any).insertHtml(html, Word.InsertLocation.replace);
  } else {
    // Last-resort: replace with plain text (no hyperlink)
    item.insertText(displayText, Word.InsertLocation.replace);
  }
  await context.sync();
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

          await applyHyperlinkToItem(context, item, url, normalizedText);
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
          await applyHyperlinkToItem(context, item, url, entry.citation);
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
  setStatus("Removing hyperlinks...");

  try {
    await removeAllHyperlinks();
  } catch (error) {
    setStatus(`Unable to remove hyperlinks. ${error instanceof Error ? error.message : String(error)}`);
  }
}

function populateProviderSelect() {
  const select = document.getElementById("provider-select") as HTMLSelectElement | null;
  if (!select) {
    return;
  }

  select.innerHTML = "";
  citationProviderRegistry.list().forEach((provider) => {
    const option = document.createElement("option");
    option.value = provider.id;
    option.textContent = provider.name;
    select.appendChild(option);
  });

  renderProviderPanel();
}

function getSelectedProvider(): CitationProvider | undefined {
  const select = document.getElementById("provider-select") as HTMLSelectElement | null;
  if (!select || !select.value) {
    return undefined;
  }
  return citationProviderRegistry.get(select.value);
}

function renderProviderPanel() {
  const provider = getSelectedProvider();
  const descriptionEl = document.getElementById("provider-description");
  const fieldsContainer = document.getElementById("provider-credential-fields");

  if (descriptionEl) {
    descriptionEl.textContent = provider?.description || "";
  }

  if (fieldsContainer) {
    fieldsContainer.innerHTML = "";
    provider?.credentialFields.forEach((field) => {
      const row = document.createElement("div");
      row.className = "citation-row";

      const label = document.createElement("label");
      label.setAttribute("for", `credential-${provider.id}-${field.key}`);
      label.textContent = field.label;

      const input = document.createElement("input");
      input.id = `credential-${provider.id}-${field.key}`;
      input.className = "url-input";
      input.type = field.type;
      if (field.placeholder) {
        input.placeholder = field.placeholder;
      }

      row.appendChild(label);
      row.appendChild(input);
      fieldsContainer.appendChild(row);
    });
  }

  updateProviderAuthStatus();
}

function updateProviderAuthStatus() {
  const provider = getSelectedProvider();
  const statusEl = document.getElementById("provider-auth-status");
  if (!statusEl) {
    return;
  }

  if (!provider) {
    statusEl.textContent = "";
  } else if (!provider.requiresAuth) {
    statusEl.textContent = "Ready to use (no sign-in required).";
  } else {
    statusEl.textContent = provider.isAuthenticated() ? "Connected." : "Not connected.";
  }
}

async function connectSelectedProvider() {
  const provider = getSelectedProvider();
  if (!provider) {
    return;
  }

  const credentials: Record<string, string> = {};
  provider.credentialFields.forEach((field) => {
    const input = document.getElementById(`credential-${provider.id}-${field.key}`) as HTMLInputElement | null;
    credentials[field.key] = input?.value ?? "";
  });

  setStatus(`Connecting to ${provider.name}...`);
  try {
    await provider.authenticate(credentials);
    setStatus(`Connected to ${provider.name}.`);
  } catch (error) {
    setStatus(`Unable to connect to ${provider.name}. ${error instanceof Error ? error.message : String(error)}`);
  }
  updateProviderAuthStatus();
}

function disconnectSelectedProvider() {
  const provider = getSelectedProvider();
  if (!provider) {
    return;
  }
  provider.signOut();
  setStatus(`Disconnected from ${provider.name}.`);
  updateProviderAuthStatus();
}

async function applyHyperlinksViaProvider() {
  const provider = getSelectedProvider();
  if (!provider) {
    setStatus("Choose a lookup provider first.");
    return;
  }
  if (provider.requiresAuth && !provider.isAuthenticated()) {
    setStatus(`Connect to ${provider.name} first.`);
    return;
  }

  setStatus(`Scanning the document for citations to look up via ${provider.name}...`);

  try {
    await Word.run(async (context) => {
      const bodyText = (context.document.body as any).getText();
      await context.sync();

      const candidates = extractCaseCitations(bodyText);
      if (candidates.length === 0) {
        setStatus("No case citations were found in the current document.");
        return;
      }

      let linkedCount = 0;
      let skippedCount = 0;

      // Looked up one at a time (not in parallel) to stay within each provider's rate limits.
      for (const raw of candidates) {
        const parsed = parseCaseCitation(raw) || { raw };
        const match = await provider.lookupCitation(parsed);
        if (!match) {
          skippedCount += 1;
          continue;
        }

        const results = context.document.body.search(raw, { matchCase: false, matchWholeWord: false });
        results.load("items");
        await context.sync();

        for (const item of results.items) {
          const hyperlinks = item.hyperlinks;
          hyperlinks.load("items");
          await context.sync();

          if (hyperlinks.items.length > 0) {
            continue;
          }
          await applyHyperlinkToItem(context, item, match.url, raw);
        }

        linkedCount += 1;
      }

      setStatus(
        `Linked ${linkedCount} of ${candidates.length} citation(s) via ${provider.name}. ` +
          `${skippedCount} could not be resolved and were left unchanged.`
      );
    });
  } catch (error) {
    setStatus(`Unable to complete the online lookup. ${error instanceof Error ? error.message : String(error)}`);
  }
}

const TAB_ELEMENT_IDS: Record<TabId, { tab: string; panel: string }> = {
  "case-law": { tab: "case-law-tab", panel: "case-law-panel" },
  "non-patent": { tab: "non-patent-tab", panel: "non-patent-panel" },
  "online-lookup": { tab: "online-lookup-tab", panel: "online-lookup-panel" },
};

function setActiveTab(tabName: TabId) {
  (Object.keys(TAB_ELEMENT_IDS) as TabId[]).forEach((id) => {
    const isActive = id === tabName;
    const tabEl = document.getElementById(TAB_ELEMENT_IDS[id].tab);
    const panelEl = document.getElementById(TAB_ELEMENT_IDS[id].panel);
    tabEl?.classList.toggle("active", isActive);
    tabEl?.setAttribute("aria-selected", isActive ? "true" : "false");
    panelEl?.classList.toggle("active", isActive);
  });

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
