import { ParsedCitation } from "../providers/types";
import { BluebookIssue } from "./types";
import { STATE_ABBREVIATIONS } from "./generated/stateAbbreviations.generated";

const stateAbbreviations = STATE_ABBREVIATIONS as Record<string, string>;

/**
 * Flags a court/jurisdiction parenthetical that spells out a full state name instead of using
 * the Bluebook Table T10 abbreviation (e.g. "(California 1990)" instead of "(Cal. 1990)").
 * Edition-independent -- state abbreviations haven't changed across the editions this project
 * tracks. Uses Free Law Project's reporters-db state abbreviation table (vendored at dev time,
 * see generated/stateAbbreviations.generated.ts).
 */
export function checkCourtStateAbbreviation(citation: ParsedCitation): BluebookIssue[] {
  const court = citation.court;
  if (!court) {
    return [];
  }

  for (const [abbreviation, fullName] of Object.entries(stateAbbreviations)) {
    if (fullName === abbreviation) {
      // States whose Bluebook abbreviation is the full name itself (e.g. "Iowa", "Ohio") --
      // nothing to flag.
      continue;
    }
    if (new RegExp(`\\b${fullName}\\b`, "i").test(court)) {
      return [
        {
          ruleId: "court-state-not-abbreviated",
          message: `"${fullName}" should be abbreviated as "${abbreviation}" in the court/jurisdiction parenthetical (Table T10).`,
          severity: "warning",
        },
      ];
    }
  }

  return [];
}
