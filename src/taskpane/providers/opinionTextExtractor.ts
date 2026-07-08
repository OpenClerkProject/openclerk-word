/**
 * Extracts the portion(s) of a court opinion's full text corresponding to specific print-reporter
 * page numbers, using "star pagination" markers (e.g. "*705") -- the standard convention used by
 * Westlaw and most public-domain legal text corpora (including the sources CourtListener
 * aggregates) to mark where each page of the original print reporter begins.
 *
 * This is a best-effort heuristic, not a guarantee: an opinion's text isn't required to include
 * star-pagination markers at all (it depends on the opinion's original source), and if none are
 * found -- or none match a requested page -- this returns null rather than guessing.
 */

interface PageMarker {
  index: number;
  page: number;
}

function findPageMarkers(text: string): PageMarker[] {
  const markers: PageMarker[] = [];
  const markerRegex = /\*\s?(\d+)\b/g;
  let match: RegExpExecArray | null;
  while ((match = markerRegex.exec(text)) !== null) {
    markers.push({ index: match.index, page: parseInt(match[1], 10) });
  }
  return markers;
}

/**
 * Returns the text between each requested page's marker and the next marker (of any page),
 * joined in ascending page order -- or null if the text has no star-pagination markers, or none
 * of them match a requested page.
 */
export function extractPageExcerpt(fullText: string, targetPages: number[]): string | null {
  if (!fullText || targetPages.length === 0) {
    return null;
  }

  const markers = findPageMarkers(fullText);
  if (markers.length === 0) {
    return null;
  }

  const wanted = new Set(targetPages);
  const segments: string[] = [];
  for (let i = 0; i < markers.length; i++) {
    if (!wanted.has(markers[i].page)) {
      continue;
    }
    const start = markers[i].index;
    const end = i + 1 < markers.length ? markers[i + 1].index : fullText.length;
    const segment = fullText.slice(start, end).trim();
    if (segment) {
      segments.push(segment);
    }
  }

  return segments.length > 0 ? segments.join("\n\n[...]\n\n") : null;
}

/**
 * Strips HTML tags for opinions that only have an HTML text field (no plain_text). Blunt but
 * adequate here: the output only needs to be readable inside a Word comment, never re-rendered
 * as HTML, so there's no injection risk in being lossy about markup.
 */
export function stripHtmlTags(html: string): string {
  return html
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, " ")
    .trim();
}
