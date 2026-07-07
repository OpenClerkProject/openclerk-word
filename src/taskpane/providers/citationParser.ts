import { normalizeText } from "../utils";
import { ParsedCitation } from "./types";

// A case-name "word" is either a capitalized token (Norfolk, W., Ry., Co., State, York...), the
// literal "&", or one of a small set of lowercase connectors Bluebook case names commonly contain
// (State of New York, United States ex rel. Doe, In re Foo). Requiring every token to look like part
// of a proper name -- rather than matching any run of non-punctuation text -- keeps ordinary lowercase
// prose ("the court's holding in ...") from being swallowed into the captured case name, since periods
// alone (used heavily in reporter/party abbreviations) can't be relied on as a sentence boundary.
const NAME_START_TOKEN = "[A-Z][A-Za-z.'&-]*";
const NAME_CONT_TOKEN = "(?:[A-Z][A-Za-z.'&-]*|&|of|the|and|for|a|an|ex|rel\\.?)";
const CASE_NAME = `${NAME_START_TOKEN}(?:\\s+${NAME_CONT_TOKEN})*`;
// Optional Bluebook pincite immediately after the first page, e.g. ", 496" or ", 705-06".
const PINCITE = "(?:,\\s*\\d+(?:-\\d+)?)?";
const CASE_CITATION_REGEX = new RegExp(
  `${CASE_NAME}\\s+v\\.?\\s+${CASE_NAME},\\s*\\d+\\s+[A-Za-z0-9.&' ]+?\\s+\\d+${PINCITE}(?:\\s*\\([^)]*\\))?`,
  "g"
);

// Bluebook introductory signals (see Bluebook Rule 1.2) that commonly precede a citation with no
// intervening punctuation, e.g. "See generally Norfolk & W. Ry. Co. v. Liepelt, ...". Without
// stripping these, the greedy case-name capture above would swallow the signal into the "case name".
const LEADING_SIGNAL_REGEX =
  /^(?:see\s*,?\s*e\.g\.,?|see\s+also|see\s+generally|but\s+see|but\s+cf\.?|cf\.?|accord|contra|compare|see)\s+/i;

/**
 * Scans running document text for full Bluebook-style case citations, e.g.
 * "Norfolk & W. Ry. Co. v. Liepelt, 444 U.S. 490 (U.S.Ill., 1980)". This is
 * a best-effort heuristic scanner for the online-lookup workflow: a missed
 * or malformed match simply means that citation is skipped, since providers
 * are only ever asked about text this function judged citation-shaped.
 */
export function extractCaseCitations(text: string): string[] {
  const matches = text.match(CASE_CITATION_REGEX) || [];
  const unique = new Set<string>();

  matches.forEach((match) => {
    const citation = normalizeText(match).replace(LEADING_SIGNAL_REGEX, "");
    if (citation) {
      unique.add(citation);
    }
  });

  return Array.from(unique);
}

/**
 * Parses a Bluebook-style case citation, e.g.:
 *   "Norfolk & W. Ry. Co. v. Liepelt, 444 U.S. 490 (U.S.Ill., 1980)"
 * into its structural parts. Returns null when the text doesn't look like a
 * "<case name>, <volume> <reporter> <page> (<court info>)" citation.
 */
export function parseCaseCitation(text: string): ParsedCitation | null {
  const raw = normalizeText(text);
  if (!raw) {
    return null;
  }

  const match = raw.match(
    /^(.+?),\s*(\d+)\s+([A-Za-z0-9.&' ]+?)\s+(\d+)(?:,\s*(\d+(?:-\d+)?))?\s*(?:\(([^)]*)\))?\s*$/
  );

  if (!match) {
    return null;
  }

  const [, caseName, volume, reporter, page, pincite, parenthetical] = match;
  const parsed: ParsedCitation = {
    raw,
    caseName: caseName?.trim(),
    volume: volume?.trim(),
    reporter: reporter?.trim(),
    page: page?.trim(),
  };
  if (pincite) {
    parsed.pincite = pincite.trim();
  }

  if (parenthetical) {
    const yearMatch = parenthetical.match(/(\d{4})\s*$/);
    if (yearMatch) {
      parsed.year = yearMatch[1];
      const court = parenthetical.slice(0, yearMatch.index).replace(/,\s*$/, "").trim();
      if (court) {
        parsed.court = court;
      }
    } else {
      parsed.court = parenthetical.trim();
    }
  }

  return parsed;
}
