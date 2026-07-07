export function normalizeText(value: string): string {
  return value
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

export function isLikelyCaseCitation(value: string): boolean {
  const normalized = normalizeText(value);
  if (!normalized) {
    return false;
  }

  if (normalized.includes(" v. ") || normalized.includes(" v ")) {
    return true;
  }

  return /\b\d{4}\b/.test(normalized);
}

export function extractParentheticalCitations(text: string): string[] {
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

const ALLOWED_HYPERLINK_SCHEMES = new Set(["http:", "https:", "mailto:"]);

/**
 * Only http(s)/mailto URLs are safe to write into a Word hyperlink here -- source .docx files
 * come from whoever the user chooses to import, and parenthetical URLs are free-form user input,
 * so schemes like javascript:/vbscript:/file: must be rejected before insertHyperlink/insertHtml.
 */
export function isSafeHyperlinkUrl(url: string): boolean {
  const trimmed = url.trim();
  if (!trimmed) {
    return false;
  }
  try {
    const parsed = new URL(trimmed, "https://placeholder.invalid/");
    return ALLOWED_HYPERLINK_SCHEMES.has(parsed.protocol);
  } catch {
    return false;
  }
}

export function escapeHtml(str: string): string {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

export function stripHtmlHyperlinks(html: string): string {
  if (!html) return "";

  // Replace anchor tags with their inner content
  let result = html.replace(/<a\b[^>]*>([\s\S]*?)<\/a>/gi, "$1");

  // Remove any remaining HTML tags
  result = result.replace(/<[^>]+>/g, "");

  // Decode a handful of common HTML entities
  result = result
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");

  return normalizeText(result);
}
