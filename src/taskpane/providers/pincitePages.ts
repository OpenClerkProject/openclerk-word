/**
 * Reconstructs a page-range end number that may have been written in Bluebook's "dropped
 * digits" short form (e.g. the "06" in "705-06") back to its full value, by borrowing the
 * corresponding leading digits from the range's start number -- exactly the assumption the
 * dropped notation itself relies on. If the written end is already the same length as (or
 * longer than) the start, it's returned unchanged.
 */
export function reconstructFullPageNumber(start: string, writtenEnd: string): string {
  return writtenEnd.length < start.length ? start.slice(0, start.length - writtenEnd.length) + writtenEnd : writtenEnd;
}

/**
 * Expands a citation's pincite string -- a single page ("496"), a comma-separated list
 * ("505, 508, 513"), and/or a Bluebook-dropped-digit range ("705-06") -- into the full,
 * deduplicated, ascending list of individual page numbers it refers to.
 */
export function expandPincitePages(pincite: string): number[] {
  const pages = new Set<number>();
  const segments = pincite.split(",").map((segment) => segment.trim());

  for (const segment of segments) {
    const rangeMatch = segment.match(/^(\d+)-(\d+)$/);
    if (rangeMatch) {
      const [, start, writtenEnd] = rangeMatch;
      const startNum = parseInt(start, 10);
      const endNum = parseInt(reconstructFullPageNumber(start, writtenEnd), 10);
      for (let page = startNum; page <= endNum; page++) {
        pages.add(page);
      }
      continue;
    }

    if (/^\d+$/.test(segment)) {
      pages.add(parseInt(segment, 10));
    }
  }

  return Array.from(pages).sort((a, b) => a - b);
}
