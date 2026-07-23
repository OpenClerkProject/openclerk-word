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
  toSafeHyperlinkUrl,
  citationProviderRegistry,
  parseCaseCitation,
  extractCaseCitations,
  expandPincitePages,
  supportsOpinionText,
  supportsRateLimitAwareness,
  checkCitationsForHallucinations,
  CitationProvider,
  OpinionTextCapableProvider,
  ParsedCitation,
  HallucinationCheckResult,
  bluebookRuleSetRegistry,
  BluebookRuleSet,
  BluebookIssue,
  SafeHyperlinkUrl,
} from "openclerk-core";
import { insertSafeHyperlink, insertSafeComment, insertSafeOoxml } from "./safeInsertion";
import { assertSafeXmlPart } from "./xmlPartGuard";

type CitationMap = Map<string, string>;
type ParentheticalEntry = { citation: string; url: string; id: string };
type TabId = "manage-hyperlinks" | "bluebook-check" | "hallucination-check" | "embed-cited-text";
type HyperlinkScope = "case-law" | "non-case-law" | "both";
type CaseLawSource = "file" | "online";
type HallucinationProviderEntry = { id: string; checked: boolean };
type BluebookCheckedCitation = {
  raw: string;
  parsed: ParsedCitation | null;
  issues: BluebookIssue[];
};
type EmbedTextResult = { raw: string; embedded: boolean; reason: string | null };

// Prefix on every comment OpenClerk inserts via "Embed cited opinion text", so "Remove embedded
// text" can find-and-delete only its own comments without touching any the user added by hand.
// The citation's raw text is embedded right after the marker (see buildEmbeddedCommentContent)
// so a re-run can tell which citations already have a comment and skip them, instead of
// re-fetching (wasteful given CourtListener's rate limit) and inserting a duplicate.
const EMBEDDED_TEXT_COMMENT_MARKER = "[OpenClerk embedded citation text]";

// Exported for the feature-level insertion smoke test (tests/smoke.wordInsertion.test.ts), which
// drives this real builder into the real insertSafeComment sink to prove cited text with &/'/"/<>
// reaches Word's plain-text comment API uncorrupted (the CR-02 / audit-#1 plain-text-sink guarantee).
export function buildEmbeddedCommentContent(raw: string, excerpt: string): string {
  return `${EMBEDDED_TEXT_COMMENT_MARKER} ${raw}\n\n${excerpt}`;
}

export function citationHasEmbeddedComment(commentContent: string, raw: string): boolean {
  return (
    commentContent === `${EMBEDDED_TEXT_COMMENT_MARKER} ${raw}` ||
    commentContent.startsWith(`${EMBEDDED_TEXT_COMMENT_MARKER} ${raw}\n`)
  );
}

// A source .docx is just a zip file; without limits, a small maliciously-crafted zip can
// decompress to a huge amount of data in memory ("zip bomb") and hang or crash the taskpane.
// Legitimate citation-source documents are always small, so these caps are generous in practice.
const MAX_SOURCE_FILE_BYTES = 20 * 1024 * 1024; // 20 MB compressed upload
const MAX_ZIP_ENTRY_COUNT = 500;
const MAX_DECOMPRESSED_XML_BYTES = 50 * 1024 * 1024; // 50 MB per extracted XML part

// A single .docx XML part that claims to expand at more than this ratio is treated as a zip bomb
// and rejected *before* it is decompressed. Real document XML (markup plus prose) compresses at
// roughly 10-20:1; a part whose declared uncompressed:compressed ratio is above 200:1 is not a
// normal document, it's a run of repeated bytes crafted to blow up in memory once expanded. The
// floor below keeps this from tripping on a tiny, legitimately repetitive part whose ratio is high
// but whose absolute expanded size is trivial.
const MAX_SAFE_COMPRESSION_RATIO = 200;
const COMPRESSION_RATIO_FLOOR_BYTES = 1 * 1024 * 1024; // 1 MB

// Word's Range.search() rejects search strings over 255 characters. citationText here comes
// from arbitrary hyperlink display text in an untrusted source .docx, so it isn't guaranteed to
// respect that -- and because all searches for a batch are queued before the first sync, one
// oversized citation would otherwise fail the sync and abort every citation in the batch, not
// just the bad one.
const MAX_SEARCH_TEXT_LENGTH = 255;

let sourceCitationMap: CitationMap | null = null;
let parentheticalEntries: ParentheticalEntry[] = [];
let hallucinationProviderOrder: HallucinationProviderEntry[] = [];
let hyperlinkScope: HyperlinkScope = "case-law";
let caseLawSource: CaseLawSource = "file";
let lastBluebookResults: BluebookCheckedCitation[] | null = null;
let bluebookShowFlaggedOnly = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const sideloadMessage = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");
    const sourceFileInput = document.getElementById("source-file") as HTMLInputElement | null;
    const applyButton = document.getElementById("apply-hyperlinks") as HTMLButtonElement | null;
    const removeButton = document.getElementById("remove-hyperlinks") as HTMLButtonElement | null;
    const workflowSelect = document.getElementById("workflow-select") as HTMLSelectElement | null;
    const hyperlinkScopeSelect = document.getElementById(
      "hyperlink-scope-select"
    ) as HTMLSelectElement | null;
    const caseLawSourceSelect = document.getElementById(
      "case-law-source-select"
    ) as HTMLSelectElement | null;
    const scanButton = document.getElementById("scan-parentheticals") as HTMLButtonElement | null;
    const addParentheticalButton = document.getElementById(
      "add-parenthetical-hyperlinks"
    ) as HTMLButtonElement | null;
    const removeParentheticalButton = document.getElementById(
      "remove-parenthetical-hyperlinks"
    ) as HTMLButtonElement | null;
    const providerSelect = document.getElementById("provider-select") as HTMLSelectElement | null;
    const providerConnectButton = document.getElementById(
      "provider-connect"
    ) as HTMLButtonElement | null;
    const providerDisconnectButton = document.getElementById(
      "provider-disconnect"
    ) as HTMLButtonElement | null;
    const applyOnlineHyperlinksButton = document.getElementById(
      "apply-online-hyperlinks"
    ) as HTMLButtonElement | null;
    const bluebookEditionSelect = document.getElementById(
      "bluebook-edition-select"
    ) as HTMLSelectElement | null;
    const checkBluebookCitationsButton = document.getElementById(
      "check-bluebook-citations"
    ) as HTMLButtonElement | null;
    const bluebookShowFlaggedOnlyCheckbox = document.getElementById(
      "bluebook-show-flagged-only"
    ) as HTMLInputElement | null;
    const checkHallucinationsButton = document.getElementById(
      "check-hallucinations"
    ) as HTMLButtonElement | null;
    const embedTextProviderSelect = document.getElementById(
      "embed-text-provider-select"
    ) as HTMLSelectElement | null;
    const embedCitedTextButton = document.getElementById(
      "embed-cited-text"
    ) as HTMLButtonElement | null;
    const removeEmbeddedTextButton = document.getElementById(
      "remove-embedded-text"
    ) as HTMLButtonElement | null;

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
      applyButton.addEventListener("click", () =>
        withBusyButton(applyButton, applyCaseLawHyperlinksFromSource)
      );
    }
    if (removeButton) {
      removeButton.addEventListener("click", () =>
        withBusyButton(removeButton, removeCaseLawHyperlinks)
      );
    }
    if (workflowSelect) {
      workflowSelect.addEventListener("change", () => setActiveTab(workflowSelect.value as TabId));
    }
    if (hyperlinkScopeSelect) {
      hyperlinkScopeSelect.addEventListener("change", () => {
        hyperlinkScope = hyperlinkScopeSelect.value as HyperlinkScope;
        updateManageHyperlinksVisibility();
      });
    }
    if (caseLawSourceSelect) {
      caseLawSourceSelect.addEventListener("change", () => {
        caseLawSource = caseLawSourceSelect.value as CaseLawSource;
        updateManageHyperlinksVisibility();
      });
    }
    if (scanButton) {
      scanButton.addEventListener("click", () =>
        withBusyButton(scanButton, scanParentheticalCitations)
      );
    }
    if (addParentheticalButton) {
      addParentheticalButton.addEventListener("click", () =>
        withBusyButton(addParentheticalButton, addParentheticalHyperlinks)
      );
    }
    if (removeParentheticalButton) {
      removeParentheticalButton.addEventListener("click", () =>
        withBusyButton(removeParentheticalButton, removeParentheticalHyperlinks)
      );
    }
    if (providerSelect) {
      providerSelect.addEventListener("change", renderProviderPanel);
    }
    if (providerConnectButton) {
      providerConnectButton.addEventListener("click", () =>
        withBusyButton(providerConnectButton, connectSelectedProvider)
      );
    }
    if (providerDisconnectButton) {
      providerDisconnectButton.addEventListener("click", () =>
        withBusyButton(providerDisconnectButton, async () => disconnectSelectedProvider())
      );
    }
    if (applyOnlineHyperlinksButton) {
      applyOnlineHyperlinksButton.addEventListener("click", () =>
        withBusyButton(applyOnlineHyperlinksButton, applyHyperlinksViaProvider)
      );
    }
    if (bluebookEditionSelect) {
      bluebookEditionSelect.addEventListener("change", () => {
        renderBluebookEditionDescription();
        invalidateBluebookResults();
      });
    }
    if (checkBluebookCitationsButton) {
      checkBluebookCitationsButton.addEventListener("click", () =>
        withBusyButton(checkBluebookCitationsButton, checkBluebookCitations)
      );
    }
    if (bluebookShowFlaggedOnlyCheckbox) {
      bluebookShowFlaggedOnlyCheckbox.addEventListener("change", () => {
        bluebookShowFlaggedOnly = bluebookShowFlaggedOnlyCheckbox.checked;
        renderBluebookResults();
      });
    }
    if (checkHallucinationsButton) {
      checkHallucinationsButton.addEventListener("click", () =>
        withBusyButton(checkHallucinationsButton, checkForHallucinations)
      );
    }
    if (embedTextProviderSelect) {
      embedTextProviderSelect.addEventListener("change", renderEmbedTextProviderStatus);
    }
    if (embedCitedTextButton) {
      embedCitedTextButton.addEventListener("click", () =>
        withBusyButton(embedCitedTextButton, embedCitedOpinionText)
      );
    }
    if (removeEmbeddedTextButton) {
      removeEmbeddedTextButton.addEventListener("click", () =>
        withBusyButton(removeEmbeddedTextButton, removeEmbeddedCitationText)
      );
    }

    populateProviderSelect();
    populateBluebookEditionSelect();
    populateHallucinationProviderList();
    populateEmbedTextProviderSelect();
    updateManageHyperlinksVisibility();
    renderBluebookResults();
    setActiveTab("manage-hyperlinks");
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
    setStatus(
      `Unable to read ${file.name}. ${error instanceof Error ? error.message : String(error)}`
    );
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
      // The search/load/filter passes below are batched into three context.sync() calls
      // regardless of citation count. The actual insertion loop further down, however, calls
      // insertSafeHyperlink once per matched citation, and insertSafeHyperlink unconditionally
      // syncs internally (safeInsertion.ts) -- so a document with 50-100+ citations still means
      // roughly one context.sync() round-trip per citation for the insertion phase itself. Only
      // the search/load phase is batched today, not insertion.
      const citationEntries = Array.from(sourceCitationMap!.entries())
        .map(([citationText, rawUrl]) => ({ citationText, url: toSafeHyperlinkUrl(rawUrl) }))
        .filter(
          (entry) =>
            entry.citationText &&
            entry.citationText.length <= MAX_SEARCH_TEXT_LENGTH &&
            entry.url !== null
        );

      const searches = citationEntries.map(({ citationText, url }) => ({
        url: url as SafeHyperlinkUrl,
        results: context.document.body.search(citationText, {
          matchCase: false,
          matchWholeWord: false,
        }),
      }));
      searches.forEach((entry) => entry.results.load("items"));
      await context.sync();

      const matchedItems: { item: Word.Range; url: SafeHyperlinkUrl }[] = [];
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
        await insertSafeHyperlink(context, item, url, normalizedText);
        appliedCount += 1;
      }

      setStatus(`Added ${appliedCount} hyperlink(s) to matching case-law citations.`);
    });
  } catch (error) {
    setStatus(
      `Unable to apply hyperlinks. ${error instanceof Error ? error.message : String(error)}`
    );
  }
}

async function removeCaseLawHyperlinks() {
  setStatus("Removing hyperlinks...");

  try {
    await removeAllHyperlinks();
  } catch (error) {
    setStatus(
      `Unable to remove hyperlinks. ${error instanceof Error ? error.message : String(error)}`
    );
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
    setStatus(
      `Unable to scan the document. ${error instanceof Error ? error.message : String(error)}`
    );
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
        .map((entry) => ({ ...entry, url: toSafeHyperlinkUrl(entry.url.trim()) }))
        .filter(
          (entry): entry is ParentheticalEntry & { url: SafeHyperlinkUrl } =>
            entry.url !== null && entry.citation.length <= MAX_SEARCH_TEXT_LENGTH
        );

      const searches = validEntries.map((entry) => ({
        entry,
        results: context.document.body.search(entry.citation, {
          matchCase: false,
          matchWholeWord: false,
        }),
      }));
      searches.forEach((s) => s.results.load("items"));
      await context.sync();

      const matchedItems: {
        item: Word.Range;
        entry: ParentheticalEntry & { url: SafeHyperlinkUrl };
      }[] = [];
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
        await insertSafeHyperlink(context, item, entry.url, entry.citation);
        addedCount += 1;
      }

      setStatus(`Added ${addedCount} hyperlink(s) to parenthetical citations.`);
    });
  } catch (error) {
    setStatus(
      `Unable to add hyperlinks. ${error instanceof Error ? error.message : String(error)}`
    );
  }
}

async function removeParentheticalHyperlinks() {
  setStatus("Removing hyperlinks...");

  try {
    await removeAllHyperlinks();
  } catch (error) {
    setStatus(
      `Unable to remove hyperlinks. ${error instanceof Error ? error.message : String(error)}`
    );
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
    const input = document.getElementById(
      `credential-${provider.id}-${field.key}`
    ) as HTMLInputElement | null;
    credentials[field.key] = input?.value ?? "";
  });

  setStatus(`Connecting to ${provider.name}...`);
  try {
    await provider.authenticate(credentials);
    setStatus(`Connected to ${provider.name}.`);
  } catch (error) {
    setStatus(
      `Unable to connect to ${provider.name}. ${error instanceof Error ? error.message : String(error)}`
    );
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
      context.document.body.load("text");
      await context.sync();
      const bodyText = context.document.body.text;

      const candidates = extractCaseCitations(bodyText);
      if (candidates.length === 0) {
        setStatus("No case citations were found in the current document.");
        return;
      }

      let linkedCount = 0;
      let skippedCount = 0;
      let rateLimitedCount = 0;

      // Looked up one at a time (not in parallel) to stay within each provider's rate limits.
      for (const raw of candidates) {
        const searchResults = context.document.body.search(raw, {
          matchCase: false,
          matchWholeWord: false,
        });
        searchResults.load("items");
        await context.sync();

        if (searchResults.items.length === 0) {
          skippedCount += 1;
          continue;
        }

        // Load hyperlink status for every instance before deciding whether a lookup is even
        // needed -- if every instance of this citation already has a hyperlink, skip the API
        // call entirely. This matters more than it might seem: with CourtListener's tight
        // default rate limit, re-running this scan after a partial run should spend its limited
        // quota on the citations that still need it, not re-verify ones already done.
        searchResults.items.forEach((item) => item.hyperlinks.load("items"));
        await context.sync();

        const unlinkedItems = searchResults.items.filter(
          (item) => item.hyperlinks.items.length === 0
        );
        if (unlinkedItems.length === 0) {
          linkedCount += 1;
          continue;
        }

        const parsed = parseCaseCitation(raw) || { raw };
        const match = await provider.lookupCitation(parsed);
        const safeUrl = match ? toSafeHyperlinkUrl(match.url) : null;
        if (!safeUrl) {
          if (supportsRateLimitAwareness(provider) && provider.wasLastRequestRateLimited()) {
            rateLimitedCount += 1;
          } else {
            skippedCount += 1;
          }
          continue;
        }

        for (const item of unlinkedItems) {
          await insertSafeHyperlink(context, item, safeUrl, raw);
        }

        linkedCount += 1;
      }

      const rateLimitNote =
        rateLimitedCount > 0
          ? ` ${rateLimitedCount} were rate-limited by ${provider.name} -- wait a minute and click "Scan & hyperlink via API" again to pick up the rest (already-linked citations won't be re-checked).`
          : "";
      setStatus(
        `Linked ${linkedCount} of ${candidates.length} citation(s) via ${provider.name}. ` +
          `${skippedCount} could not be resolved and were left unchanged.${rateLimitNote}`
      );
    });
  } catch (error) {
    setStatus(
      `Unable to complete the online lookup. ${error instanceof Error ? error.message : String(error)}`
    );
  }
}

const TAB_PANEL_IDS: Record<TabId, string> = {
  "manage-hyperlinks": "manage-hyperlinks-panel",
  "bluebook-check": "bluebook-check-panel",
  "hallucination-check": "hallucination-check-panel",
  "embed-cited-text": "embed-cited-text-panel",
};

function setActiveTab(tabName: TabId) {
  (Object.keys(TAB_PANEL_IDS) as TabId[]).forEach((id) => {
    const panelEl = document.getElementById(TAB_PANEL_IDS[id]);
    panelEl?.classList.toggle("active", id === tabName);
  });

  const workflowSelect = document.getElementById("workflow-select") as HTMLSelectElement | null;
  if (workflowSelect && workflowSelect.value !== tabName) {
    workflowSelect.value = tabName;
  }

  if (tabName === "hallucination-check") {
    // Re-render (not re-populate) so any provider connected via the Manage Hyperlinks tab since
    // this list was first built shows up-to-date "(not connected)" status, without losing the
    // user's checked/order selections.
    renderHallucinationProviderList();
  }

  if (tabName === "embed-cited-text") {
    // Same reasoning as hallucination-check above: a provider connected on the Manage Hyperlinks
    // tab since this panel was first opened should be reflected without losing the selection.
    renderEmbedTextProviderStatus();
  }
}

// Toggles which sub-section(s) of the Manage Hyperlinks panel are visible based on the current
// scope (case law / non-case-law / both) and, within case-law scope, which source is selected
// (a file with pre-existing hyperlinks, vs. a live online lookup). Both scope sections can be
// visible at once when scope is "both".
function updateManageHyperlinksVisibility() {
  const caseLawSection = document.getElementById("case-law-scope-section");
  const nonCaseLawSection = document.getElementById("non-case-law-scope-section");
  const fileSource = document.getElementById("case-law-file-source");
  const onlineSource = document.getElementById("case-law-online-source");

  const showCaseLaw = hyperlinkScope === "case-law" || hyperlinkScope === "both";
  const showNonCaseLaw = hyperlinkScope === "non-case-law" || hyperlinkScope === "both";

  caseLawSection?.classList.toggle("hidden", !showCaseLaw);
  nonCaseLawSection?.classList.toggle("hidden", !showNonCaseLaw);
  fileSource?.classList.toggle("hidden", caseLawSource !== "file");
  onlineSource?.classList.toggle("hidden", caseLawSource !== "online");

  if (showNonCaseLaw && parentheticalEntries.length === 0) {
    scanParentheticalCitations();
  }
}

function populateBluebookEditionSelect() {
  const select = document.getElementById("bluebook-edition-select") as HTMLSelectElement | null;
  if (!select) {
    return;
  }

  select.innerHTML = "";
  bluebookRuleSetRegistry.list().forEach((ruleSet) => {
    const option = document.createElement("option");
    option.value = ruleSet.id;
    option.textContent = ruleSet.name;
    select.appendChild(option);
  });

  renderBluebookEditionDescription();
}

function getSelectedBluebookRuleSet(): BluebookRuleSet | undefined {
  const select = document.getElementById("bluebook-edition-select") as HTMLSelectElement | null;
  if (!select || !select.value) {
    return undefined;
  }
  return bluebookRuleSetRegistry.get(select.value);
}

function renderBluebookEditionDescription() {
  const descriptionEl = document.getElementById("bluebook-edition-description");
  if (descriptionEl) {
    descriptionEl.textContent = getSelectedBluebookRuleSet()?.description || "";
  }
}

async function checkBluebookCitations() {
  const ruleSet = getSelectedBluebookRuleSet();
  if (!ruleSet) {
    setStatus("Choose a Bluebook edition first.");
    return;
  }

  setStatus(`Checking citations against the ${ruleSet.name}...`);

  try {
    await Word.run(async (context) => {
      context.document.body.load("text");
      await context.sync();
      const bodyText = context.document.body.text;

      const candidates = extractCaseCitations(bodyText);
      lastBluebookResults = candidates.map((raw) => {
        const parsed = parseCaseCitation(raw);
        const issues = parsed ? ruleSet.checkCitation(parsed) : [];
        return { raw, parsed, issues };
      });
      renderBluebookResults();

      if (candidates.length === 0) {
        setStatus("No case citations were found in the current document.");
        return;
      }

      const errorCount = lastBluebookResults.reduce(
        (count, result) =>
          count + result.issues.filter((issue) => issue.severity === "error").length,
        0
      );
      const warningCount = lastBluebookResults.reduce(
        (count, result) =>
          count + result.issues.filter((issue) => issue.severity === "warning").length,
        0
      );
      const flaggedCount = lastBluebookResults.filter(
        (result) => !result.parsed || result.issues.length > 0
      ).length;

      setStatus(
        `Checked ${candidates.length} citation(s) against the ${ruleSet.name}; ${flaggedCount} flagged ` +
          `(${errorCount} error(s), ${warningCount} warning(s)).`
      );
    });
  } catch (error) {
    setStatus(
      `Unable to check citations. ${error instanceof Error ? error.message : String(error)}`
    );
  }
}

function invalidateBluebookResults() {
  if (lastBluebookResults === null) {
    return;
  }
  lastBluebookResults = null;
  const summary = document.getElementById("bluebook-results-summary");
  if (summary) {
    summary.textContent = "";
  }
  const container = document.getElementById("bluebook-issue-list");
  if (container) {
    container.innerHTML =
      '<p class="helper-text">The Bluebook edition changed -- click "Check citations" again to see results for the new edition.</p>';
  }
}

/**
 * Deep-links into openclerk-core's "Bluebook citation correction" GitHub Issue Form (see
 * openclerk-core's .github/ISSUE_TEMPLATE/bluebook-correction.yml and this repo's CONTRIBUTING.md)
 * with the flagged citation and rule already filled in, so reporting a possible false positive
 * doesn't require retyping it. Points at openclerk-core, not this repo, because the Bluebook rule
 * data and rule-checking logic it reports against lives there now.
 */
function buildBluebookCorrectionReportUrl(citationText: string, issue: BluebookIssue): string {
  const params = new URLSearchParams({
    template: "bluebook-correction.yml",
    title: `[Bluebook correction] Possible false positive: ${issue.ruleId}`,
    "citation-example": citationText,
    "whats-wrong": `OpenClerk flagged this citation with rule "${issue.ruleId}":\n\n${issue.message}\n\nI believe this flag is incorrect because: `,
  });
  return `https://github.com/OpenClerkProject/openclerk-core/issues/new?${params.toString()}`;
}

async function goToCitationInDocument(citationText: string) {
  if (!citationText || citationText.length > MAX_SEARCH_TEXT_LENGTH) {
    return;
  }

  try {
    await Word.run(async (context) => {
      const results = context.document.body.search(citationText, {
        matchCase: false,
        matchWholeWord: false,
      });
      results.load("items");
      await context.sync();

      if (results.items.length === 0) {
        setStatus(`Could not find "${citationText}" in the document.`);
        return;
      }

      // Selecting the range moves Word's view to show it -- this is what actually
      // "jumps to" the citation for the user.
      results.items[0].select();
      await context.sync();
    });
  } catch (error) {
    setStatus(
      `Unable to locate that citation. ${error instanceof Error ? error.message : String(error)}`
    );
  }
}

function renderBluebookResults() {
  const container = document.getElementById("bluebook-issue-list");
  const summaryEl = document.getElementById("bluebook-results-summary");
  if (!container) {
    return;
  }

  container.innerHTML = "";
  if (summaryEl) {
    summaryEl.innerHTML = "";
  }

  if (lastBluebookResults === null) {
    container.innerHTML =
      '<p class="helper-text">No case citations found yet. Click "Check citations".</p>';
    return;
  }

  if (lastBluebookResults.length === 0) {
    container.innerHTML =
      '<p class="helper-text">No case citations were found in the current document.</p>';
    return;
  }

  const errorCount = lastBluebookResults.reduce(
    (count, result) => count + result.issues.filter((issue) => issue.severity === "error").length,
    0
  );
  const warningCount = lastBluebookResults.reduce(
    (count, result) => count + result.issues.filter((issue) => issue.severity === "warning").length,
    0
  );
  const cleanCount = lastBluebookResults.filter(
    (result) => result.parsed && result.issues.length === 0
  ).length;

  if (summaryEl) {
    const parts: { text: string; className: string }[] = [];
    if (errorCount > 0) {
      parts.push({
        text: `${errorCount} error${errorCount === 1 ? "" : "s"}`,
        className: "summary-errors",
      });
    }
    if (warningCount > 0) {
      parts.push({
        text: `${warningCount} warning${warningCount === 1 ? "" : "s"}`,
        className: "summary-warnings",
      });
    }
    parts.push({ text: `${cleanCount} clean`, className: "summary-ok" });
    parts.forEach((part, index) => {
      const span = document.createElement("span");
      span.className = part.className;
      span.textContent = part.text;
      summaryEl.appendChild(span);
      if (index < parts.length - 1) {
        summaryEl.appendChild(document.createTextNode(" · "));
      }
    });
  }

  const visibleResults = bluebookShowFlaggedOnly
    ? lastBluebookResults.filter((result) => !result.parsed || result.issues.length > 0)
    : lastBluebookResults;

  if (visibleResults.length === 0) {
    container.innerHTML =
      '<p class="helper-text">No flagged citations -- everything checked out clean.</p>';
    return;
  }

  const fragment = document.createDocumentFragment();
  visibleResults.forEach(({ raw, parsed, issues }) => {
    const row = document.createElement("div");
    row.className = "bluebook-issue-row";

    const label = document.createElement("button");
    label.type = "button";
    label.className = "citation-link";
    label.title = "Click to find this citation in the document";
    label.textContent = raw;
    label.addEventListener("click", () => goToCitationInDocument(raw));
    row.appendChild(label);

    if (!parsed) {
      const result = document.createElement("p");
      result.className = "helper-text issue-flagged";
      result.textContent =
        "Could not parse this citation's structure -- this can mean a real formatting problem, or just a " +
        "citation shape this tool doesn't yet recognize (e.g. a parallel citation or a nominative reporter). " +
        "Verify it manually.";
      row.appendChild(result);
    } else if (issues.length === 0) {
      const result = document.createElement("p");
      result.className = "helper-text issue-ok";
      result.textContent = "No issues found.";
      row.appendChild(result);
    } else {
      const list = document.createElement("ul");
      list.className = "bluebook-issue-item-list";
      issues.forEach((issue) => {
        const item = document.createElement("li");
        item.className = `bluebook-issue-item severity-${issue.severity}`;

        const text = document.createElement("span");
        text.textContent = `${issue.severity === "error" ? "Error" : "Warning"}: ${issue.message}`;
        item.appendChild(text);

        const reportLink = document.createElement("a");
        reportLink.className = "report-correction-link";
        reportLink.href = buildBluebookCorrectionReportUrl(raw, issue);
        reportLink.target = "_blank";
        reportLink.rel = "noopener noreferrer";
        reportLink.textContent = "Report as wrong";
        item.appendChild(reportLink);

        list.appendChild(item);
      });
      row.appendChild(list);
    }

    fragment.appendChild(row);
  });

  container.appendChild(fragment);
}

function populateHallucinationProviderList() {
  // Registry order is the sensible initial default (roughly free/public first); nothing is
  // checked by default since the user must explicitly choose which platforms to trust.
  hallucinationProviderOrder = citationProviderRegistry.list().map((provider) => ({
    id: provider.id,
    checked: false,
  }));
  renderHallucinationProviderList();
}

function renderHallucinationProviderList() {
  const container = document.getElementById("hallucination-provider-list");
  if (!container) {
    return;
  }

  container.innerHTML = "";

  const fragment = document.createDocumentFragment();
  hallucinationProviderOrder.forEach((entry, index) => {
    const provider = citationProviderRegistry.get(entry.id);
    if (!provider) {
      return;
    }

    const row = document.createElement("div");
    row.className = "hallucination-provider-row";

    const label = document.createElement("label");
    label.className = "hallucination-provider-label";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = entry.checked;
    checkbox.addEventListener("change", () => {
      entry.checked = checkbox.checked;
    });

    const nameSpan = document.createElement("span");
    const authNote = provider.requiresAuth && !provider.isAuthenticated() ? " (not connected)" : "";
    nameSpan.textContent = `${index + 1}. ${provider.name}${authNote}`;

    label.appendChild(checkbox);
    label.appendChild(nameSpan);
    row.appendChild(label);

    const moveButtons = document.createElement("div");
    moveButtons.className = "hallucination-move-buttons";

    const upButton = document.createElement("button");
    upButton.type = "button";
    upButton.className = "ms-Button move-button";
    upButton.textContent = "↑";
    upButton.title = "Move up";
    upButton.disabled = index === 0;
    upButton.addEventListener("click", () => moveHallucinationProvider(index, -1));

    const downButton = document.createElement("button");
    downButton.type = "button";
    downButton.className = "ms-Button move-button";
    downButton.textContent = "↓";
    downButton.title = "Move down";
    downButton.disabled = index === hallucinationProviderOrder.length - 1;
    downButton.addEventListener("click", () => moveHallucinationProvider(index, 1));

    moveButtons.appendChild(upButton);
    moveButtons.appendChild(downButton);
    row.appendChild(moveButtons);

    fragment.appendChild(row);
  });

  container.appendChild(fragment);
}

function moveHallucinationProvider(index: number, delta: number) {
  const newIndex = index + delta;
  if (newIndex < 0 || newIndex >= hallucinationProviderOrder.length) {
    return;
  }
  const [entry] = hallucinationProviderOrder.splice(index, 1);
  hallucinationProviderOrder.splice(newIndex, 0, entry);
  renderHallucinationProviderList();
}

async function checkForHallucinations() {
  const selectedProviders = hallucinationProviderOrder
    .filter((entry) => entry.checked)
    .map((entry) => citationProviderRegistry.get(entry.id))
    .filter((provider): provider is CitationProvider => Boolean(provider));

  if (selectedProviders.length === 0) {
    setStatus("Select at least one platform to check citations against.");
    return;
  }

  setStatus("Scanning the document for citations to verify...");

  try {
    await Word.run(async (context) => {
      context.document.body.load("text");
      await context.sync();
      const bodyText = context.document.body.text;

      const candidates = extractCaseCitations(bodyText);
      if (candidates.length === 0) {
        renderHallucinationResults([]);
        setStatus("No case citations were found in the current document.");
        return;
      }

      // Looked up one citation, one provider, at a time (not in parallel) to stay within each
      // platform's rate limits -- same reasoning as the Online Lookup tab. Delegates to
      // openclerk-core's checkCitationsForHallucinations instead of re-deriving this loop here --
      // that shared implementation is also what verifies a provider's match actually names the
      // same case (see caseNamesMatch), not just that it resolved some citation locator.
      const results: HallucinationCheckResult[] = await checkCitationsForHallucinations(
        candidates,
        selectedProviders
      );

      renderHallucinationResults(results);

      // Rate-limited citations are kept separate from "flagged" -- a request that got throttled
      // before it could even check isn't evidence of anything, and lumping it in with genuine
      // non-matches would misleadingly call a real citation a "possible hallucination" just
      // because a platform's rate limit was hit partway through the document.
      const rateLimitedCount = results.filter(
        (result) => !result.verifiedVia && result.rateLimitedProviders.length > 0
      ).length;
      const flaggedCount = results.filter(
        (result) => !result.verifiedVia && result.rateLimitedProviders.length === 0
      ).length;
      const rateLimitNote =
        rateLimitedCount > 0
          ? ` ${rateLimitedCount} could not be checked because a platform rate-limited the request -- wait a minute and try again.`
          : "";
      setStatus(
        `Checked ${results.length} citation(s) against ${selectedProviders.length} platform(s); ` +
          `${flaggedCount} could not be verified on any selected platform.${rateLimitNote}`
      );
    });
  } catch (error) {
    setStatus(
      `Unable to check citations. ${error instanceof Error ? error.message : String(error)}`
    );
  }
}

function renderHallucinationResults(results: HallucinationCheckResult[]) {
  const container = document.getElementById("hallucination-results-list");
  if (!container) {
    return;
  }

  container.innerHTML = "";
  if (results.length === 0) {
    container.innerHTML = '<p class="helper-text">No results yet. Click "Find Hallucinations".</p>';
    return;
  }

  const fragment = document.createDocumentFragment();
  results.forEach((result) => {
    const row = document.createElement("div");
    row.className = "bluebook-issue-row";

    const label = document.createElement("button");
    label.type = "button";
    label.className = "citation-link";
    label.title = "Click to find this citation in the document";
    label.textContent = result.raw;
    label.addEventListener("click", () => goToCitationInDocument(result.raw));
    row.appendChild(label);

    const status = document.createElement("p");
    status.className = "helper-text";
    if (result.verifiedVia) {
      status.classList.add("issue-ok");
      status.textContent = `Verified via ${result.verifiedVia}.`;
    } else if (result.nameMismatch) {
      // A stronger fabrication signal than a plain "not found": the citation's locator
      // (reporter/volume/page) is real, but it belongs to a different case than the one named
      // here -- exactly the pattern of a citation with a real-looking citation attached to a
      // fabricated case name.
      status.classList.add("issue-flagged");
      status.textContent = `Possible hallucination -- ${result.nameMismatch.provider} resolves this citation to a different case: "${result.nameMismatch.foundCaseName}".`;
    } else if (result.rateLimitedProviders.length > 0) {
      // Deliberately not styled/worded like a flagged hallucination -- a throttled request isn't
      // evidence of anything, and this citation may well be genuine once re-checked.
      status.textContent = `Not checked -- rate-limited by ${result.rateLimitedProviders.join(", ")}. Not a confirmed hallucination; wait a minute and try again.`;
    } else {
      status.classList.add("issue-flagged");
      status.textContent =
        result.skippedProviders.length > 0
          ? `Not found on any connected platform. Not checked (not connected): ${result.skippedProviders.join(", ")}.`
          : "Not found on any selected platform — possible hallucination.";
    }
    row.appendChild(status);

    fragment.appendChild(row);
  });

  container.appendChild(fragment);
}

function populateEmbedTextProviderSelect() {
  const select = document.getElementById("embed-text-provider-select") as HTMLSelectElement | null;
  if (!select) {
    return;
  }

  select.innerHTML = "";
  const capableProviders = citationProviderRegistry.list().filter(supportsOpinionText);

  if (capableProviders.length === 0) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "No provider available";
    select.appendChild(option);
    select.disabled = true;
    renderEmbedTextProviderStatus();
    return;
  }

  select.disabled = false;
  capableProviders.forEach((provider) => {
    const option = document.createElement("option");
    option.value = provider.id;
    option.textContent = provider.name;
    select.appendChild(option);
  });

  renderEmbedTextProviderStatus();
}

function getSelectedEmbedTextProvider(): OpinionTextCapableProvider | undefined {
  const select = document.getElementById("embed-text-provider-select") as HTMLSelectElement | null;
  if (!select || !select.value) {
    return undefined;
  }
  const provider = citationProviderRegistry.get(select.value);
  return provider && supportsOpinionText(provider) ? provider : undefined;
}

// Shown next to the provider dropdown so the user knows *before* clicking "Embed cited opinion
// text" whether the run will actually be able to fetch anything -- this feature can't work
// anonymously (see CourtListenerProvider.isReadyForOpinionText), and discovering that only after
// every citation comes back "skipped" would be a frustrating way to find out.
function renderEmbedTextProviderStatus() {
  const statusEl = document.getElementById("embed-text-provider-status");
  if (!statusEl) {
    return;
  }

  const provider = getSelectedEmbedTextProvider();
  statusEl.classList.remove("issue-ok", "issue-flagged");
  if (!provider) {
    statusEl.textContent = "";
    return;
  }

  if (provider.isReadyForOpinionText()) {
    statusEl.classList.add("issue-ok");
    statusEl.textContent = `Ready -- connected to ${provider.name}.`;
  } else {
    statusEl.classList.add("issue-flagged");
    statusEl.textContent = `Not ready -- connect to ${provider.name} with an API token on the Manage Hyperlinks tab first.`;
  }
}

async function embedCitedOpinionText() {
  const provider = getSelectedEmbedTextProvider();
  if (!provider) {
    setStatus("Choose a provider that supports embedding opinion text first.");
    return;
  }

  setStatus(`Scanning the document for pincite citations to embed via ${provider.name}...`);

  try {
    await Word.run(async (context) => {
      context.document.body.load("text");
      await context.sync();
      const bodyText = context.document.body.text;

      const candidates = extractCaseCitations(bodyText);
      const pinciteCitations = candidates
        .map((raw) => ({ raw, parsed: parseCaseCitation(raw) }))
        .filter((item): item is { raw: string; parsed: ParsedCitation } =>
          Boolean(item.parsed?.pincite)
        );

      if (pinciteCitations.length === 0) {
        renderEmbedTextResults([]);
        setStatus(
          "No citations with a pincite (a page beyond the first page) were found in the document."
        );
        return;
      }

      // Citations that already have an OpenClerk comment are skipped rather than re-fetched --
      // both to avoid inserting a duplicate comment and to conserve CourtListener's rate limit
      // when re-running after a previous run got partially rate-limited.
      const existingComments = context.document.body.getComments();
      existingComments.load("items");
      await context.sync();
      existingComments.items.forEach((comment) => comment.load("content"));
      await context.sync();
      const existingCommentContents = existingComments.items.map((comment) => comment.content);

      const results: EmbedTextResult[] = [];

      // Looked up one citation at a time (not in parallel), same reasoning as the other
      // provider-backed workflows: stay within the provider's rate limits.
      for (const { raw, parsed } of pinciteCitations) {
        if (existingCommentContents.some((content) => citationHasEmbeddedComment(content, raw))) {
          results.push({ raw, embedded: true, reason: null });
          continue;
        }

        const targetPages = expandPincitePages(parsed.pincite as string);
        const { excerpt, rateLimited } = await provider.fetchOpinionExcerpt(parsed, targetPages);

        if (!excerpt) {
          let reason: string;
          if (rateLimited) {
            reason = `${provider.name} rate-limited this request -- wait a minute and try the remaining citations again.`;
          } else if (!provider.isReadyForOpinionText()) {
            reason = `Connect to ${provider.name} with an API token first.`;
          } else {
            reason = "Opinion text not found, or has no page markers matching this pincite.";
          }
          results.push({ raw, embedded: false, reason });
          continue;
        }

        const searchResults = context.document.body.search(raw, {
          matchCase: false,
          matchWholeWord: false,
        });
        searchResults.load("items");
        await context.sync();

        if (searchResults.items.length === 0) {
          results.push({
            raw,
            embedded: false,
            reason: "Could not find this citation's text in the document.",
          });
          continue;
        }

        // A Word comment is collapsed by default (just a margin icon) and expands on click --
        // exactly the "embedded, expandable/collapsible" behavior this feature is for, using
        // Word's own native UI instead of a custom widget.
        await insertSafeComment(
          context,
          searchResults.items[0],
          buildEmbeddedCommentContent(raw, excerpt)
        );
        results.push({ raw, embedded: true, reason: null });
      }

      renderEmbedTextResults(results);
      const embeddedCount = results.filter((result) => result.embedded).length;
      setStatus(
        `Embedded opinion text for ${embeddedCount} of ${pinciteCitations.length} pincite citation(s).`
      );
    });
  } catch (error) {
    setStatus(
      `Unable to embed cited text. ${error instanceof Error ? error.message : String(error)}`
    );
  }
}

async function removeEmbeddedCitationText() {
  setStatus("Removing OpenClerk-embedded citation text...");

  try {
    await Word.run(async (context) => {
      // document.body.getComments() (WordApi 1.4) is used instead of the desktop-only
      // document.comments property (WordApiDesktop 1.4) for broader cross-platform support.
      const comments = context.document.body.getComments();
      comments.load("items");
      await context.sync();

      comments.items.forEach((comment) => comment.load("content"));
      await context.sync();

      const ours = comments.items.filter((comment) =>
        comment.content.startsWith(EMBEDDED_TEXT_COMMENT_MARKER)
      );
      ours.forEach((comment) => comment.delete());
      await context.sync();

      setStatus(`Removed ${ours.length} embedded citation text comment(s).`);
    });
  } catch (error) {
    setStatus(
      `Unable to remove embedded text. ${error instanceof Error ? error.message : String(error)}`
    );
  }
}

function renderEmbedTextResults(results: EmbedTextResult[]) {
  const container = document.getElementById("embed-text-results-list");
  const summaryEl = document.getElementById("embed-text-results-summary");
  if (!container) {
    return;
  }

  container.innerHTML = "";
  if (summaryEl) {
    summaryEl.innerHTML = "";
  }

  if (results.length === 0) {
    container.innerHTML =
      '<p class="helper-text">No results yet. Click "Embed cited opinion text".</p>';
    return;
  }

  const embeddedCount = results.filter((result) => result.embedded).length;
  if (summaryEl) {
    const embeddedSpan = document.createElement("span");
    embeddedSpan.className = "summary-ok";
    embeddedSpan.textContent = `${embeddedCount} embedded`;
    summaryEl.appendChild(embeddedSpan);

    const skippedCount = results.length - embeddedCount;
    if (skippedCount > 0) {
      summaryEl.appendChild(document.createTextNode(" · "));
      const skippedSpan = document.createElement("span");
      skippedSpan.className = "summary-warnings";
      skippedSpan.textContent = `${skippedCount} skipped`;
      summaryEl.appendChild(skippedSpan);
    }
  }

  const fragment = document.createDocumentFragment();
  results.forEach((result) => {
    const row = document.createElement("div");
    row.className = "bluebook-issue-row";

    const label = document.createElement("button");
    label.type = "button";
    label.className = "citation-link";
    label.title = "Click to find this citation in the document";
    label.textContent = result.raw;
    label.addEventListener("click", () => goToCitationInDocument(result.raw));
    row.appendChild(label);

    const status = document.createElement("p");
    status.className = "helper-text";
    if (result.embedded) {
      status.classList.add("issue-ok");
      status.textContent =
        "Embedded as a comment -- click the comment icon in the margin to expand it.";
    } else {
      status.classList.add("issue-flagged");
      status.textContent = result.reason || "Not embedded.";
    }
    row.appendChild(status);

    fragment.appendChild(row);
  });

  container.appendChild(fragment);
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

// Disables `button` for the duration of `action` so a slow operation (a document scan, a batch of
// rate-limited API lookups) can't be re-triggered by an impatient second click while the first
// run is still in flight -- which would waste API quota and could apply/insert duplicate edits.
// Disabled buttons don't dispatch click events at all, so this alone is enough to prevent re-entrancy.
async function withBusyButton(
  button: HTMLButtonElement | null,
  action: () => Promise<void>
): Promise<void> {
  if (!button) {
    await action();
    return;
  }
  button.disabled = true;
  button.setAttribute("aria-busy", "true");
  try {
    await action();
  } finally {
    button.disabled = false;
    button.removeAttribute("aria-busy");
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
    throw new Error(
      "The selected file contains an unexpectedly large number of entries and was rejected."
    );
  }

  const documentXml = await readZipEntryWithLimit(zip, "word/document.xml");
  const relationshipsXml = await readZipEntryWithLimit(zip, "word/_rels/document.xml.rels");

  if (!documentXml || !relationshipsXml) {
    throw new Error("The selected file is not a valid Word document.");
  }

  const relationships = parseRelationships(relationshipsXml);
  const documentDom = parseXmlPartStrict(documentXml, "document body");
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

    if (text && url && isLikelyCaseCitation(text) && toSafeHyperlinkUrl(url) !== null) {
      citationMap.set(text, url);
    }
  }

  return citationMap;
}

// The zip central directory records each entry's declared compressed/uncompressed sizes. JSZip
// exposes them on an internal `_data` object rather than a public API, so this reads them
// defensively (typeof-guarded, undefined when absent) -- it lets us reject an oversized or
// absurdly-compressible entry from its declared header *before* calling entry.async(), which is
// what actually allocates the decompressed bytes. The values are attacker-controlled, but a zip
// bomb has to declare its true expanded size for a reader to allocate that much, so honoring these
// caps up front rejects the common case; the post-decompression length check stays as a backstop
// for a header that lies about being smaller than it is.
function readDeclaredEntrySizes(entry: JSZip.JSZipObject): {
  uncompressedSize?: number;
  compressedSize?: number;
} {
  const data = (entry as unknown as { _data?: unknown })._data;
  if (!data || typeof data !== "object") {
    return {};
  }
  const record = data as { uncompressedSize?: unknown; compressedSize?: unknown };
  return {
    uncompressedSize:
      typeof record.uncompressedSize === "number" ? record.uncompressedSize : undefined,
    compressedSize: typeof record.compressedSize === "number" ? record.compressedSize : undefined,
  };
}

// Reads a zip entry as text, aborting if the decompressed content is unexpectedly large -- a
// defense-in-depth backstop against zip bombs on top of the whole-file and entry-count caps.
// Exported for the zip-bomb regression test (tests/sourceDocumentZipGuard.test.ts).
export async function readZipEntryWithLimit(zip: JSZip, path: string): Promise<string | undefined> {
  const entry = zip.file(path);
  if (!entry) {
    return undefined;
  }

  // Preflight against the entry's declared sizes *before* decompressing: entry.async("string")
  // allocates the full uncompressed content in memory, so a content.length check afterwards is
  // already too late to stop a zip bomb -- the memory is spent by then. See readDeclaredEntrySizes.
  const { uncompressedSize, compressedSize } = readDeclaredEntrySizes(entry);

  if (uncompressedSize !== undefined && uncompressedSize > MAX_DECOMPRESSED_XML_BYTES) {
    throw new Error("The selected file's contents are unexpectedly large and were rejected.");
  }

  if (
    uncompressedSize !== undefined &&
    compressedSize !== undefined &&
    compressedSize > 0 &&
    uncompressedSize > COMPRESSION_RATIO_FLOOR_BYTES &&
    uncompressedSize / compressedSize > MAX_SAFE_COMPRESSION_RATIO
  ) {
    throw new Error("The selected file's contents are unexpectedly large and were rejected.");
  }

  // Backstop: if the declared header lied about being small, this still catches the real expanded
  // size after decompression.
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

      // getOoxml() returns an OfficeExtension.ClientResult<string>, not a proxy object; its .value
      // is populated automatically by the context.sync() above, no .load() call applies here. The
      // lint rule's getFunctions.json name-matching heuristic can't distinguish ClientResult-
      // returning methods from proxy-object-returning ones, so the next line is a documented false
      // positive for office-addins/load-object-before-read.
      // eslint-disable-next-line office-addins/load-object-before-read
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
      let strippedOoxml = originalOoxml
        .replace(/<w:hyperlink\b[^>]*>/g, "")
        .replace(/<\/w:hyperlink>/g, "");

      // Remove the character style reference that causes blue/underline rendering
      strippedOoxml = strippedOoxml.replace(/<w:rStyle\s+w:val="Hyperlink"\s*\/>/g, "");
      // Also remove non-self-closing form if present
      strippedOoxml = strippedOoxml.replace(
        /<w:rStyle\s+w:val="Hyperlink"\s*>\s*<\/w:rStyle>/g,
        ""
      );

      // Replace entire body OOXML with cleaned version
      await insertSafeOoxml(context, body, strippedOoxml);

      setStatus(`Removed ${removedCount} hyperlink(s) from the document.`);
    });
  } catch (err) {
    setStatus(`Unable to remove hyperlinks. ${err instanceof Error ? err.message : String(err)}`);
    throw err;
  }
}

// Parses one XML part of an imported (untrusted) .docx and runs it through assertSafeXmlPart
// before any caller inspects it (see xmlPartGuard.ts for the rejection rationale). Kept thin so
// the security-relevant checks live in the separately-unit-tested guard module.
function parseXmlPartStrict(xml: string, partName: string) {
  const dom = new DOMParser().parseFromString(xml, "application/xml");
  assertSafeXmlPart(dom, partName);
  return dom;
}

function parseRelationships(relsXml: string): Map<string, string> {
  const relsDom = parseXmlPartStrict(relsXml, "relationships part");
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
