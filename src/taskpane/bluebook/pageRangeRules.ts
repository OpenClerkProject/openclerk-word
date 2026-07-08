import { ParsedCitation } from "../providers/types";
import { BluebookIssue } from "./types";

/**
 * Computes the Bluebook Rule 3.2 "drop repetitious digits" form of a page-range end number,
 * given the range's start number (both plain digit strings of the same length). Retains at
 * least the last two digits always, and never drops when the numbers have different digit
 * counts (e.g. 99-100 isn't shortened -- there's no shared leading digit to drop).
 */
function minimalDroppedForm(start: string, fullEnd: string): string {
  if (fullEnd.length !== start.length) {
    return fullEnd;
  }
  let i = 0;
  while (i < start.length && start[i] === fullEnd[i] && start.length - i > 2) {
    i++;
  }
  return fullEnd.slice(i);
}

/**
 * Checks each page range in a citation's pincite (e.g. the "705-06" in "705-06, 710") against
 * Bluebook Rule 3.2's digit-dropping convention: retain the last two digits of the end page and
 * drop the digits it shares with the start page (e.g. "705-706" should be "705-06"). Only
 * applies to case-citation page ranges -- this project's parser never captures statute/
 * regulation section citations, which Rule 3.3 explicitly exempts from this rule anyway.
 *
 * The end page as written might already be in dropped (shortened) form, so it's reconstructed
 * back to a same-length "full" number by borrowing the corresponding leading digits from the
 * start page -- exactly the assumption the dropped notation itself relies on -- before computing
 * what the correctly-dropped form should be. This also catches an under-dropped end page (fewer
 * than the required 2 digits), not just an over-full one.
 */
export function checkPincitePageRange(citation: ParsedCitation): BluebookIssue[] {
  const pincite = citation.pincite;
  if (!pincite) {
    return [];
  }

  const issues: BluebookIssue[] = [];
  const segments = pincite.split(",").map((segment) => segment.trim());

  for (const segment of segments) {
    const rangeMatch = segment.match(/^(\d+)-(\d+)$/);
    if (!rangeMatch) {
      continue;
    }
    const [, start, writtenEnd] = rangeMatch;

    const fullEnd =
      writtenEnd.length < start.length ? start.slice(0, start.length - writtenEnd.length) + writtenEnd : writtenEnd;
    const correctEnd = minimalDroppedForm(start, fullEnd);

    if (correctEnd !== writtenEnd) {
      issues.push({
        ruleId: "pincite-range-digits",
        message: `Page range "${start}-${writtenEnd}" should drop repetitious digits to "${start}-${correctEnd}" (Bluebook Rule 3.2) -- always keep at least the last two digits.`,
        severity: "error",
      });
    }
  }

  return issues;
}
