import { ParsedCitation } from "../providers/types";
import { BluebookIssue } from "./types";
import { T6_T13_MERGER_ABBREVIATIONS } from "./caseNameAbbreviations";

/**
 * Flags full words in a case name that the 21st-edition merger of Tables T6
 * and T13.2 requires to be abbreviated (see caseNameAbbreviations.ts for why
 * this is edition-gated). Only call this for rule-sets whose edition
 * actually applies the merged table (21st edition and later).
 */
export function checkCaseNameAbbreviations(citation: ParsedCitation): BluebookIssue[] {
  const caseName = citation.caseName || "";
  const issues: BluebookIssue[] = [];

  for (const entry of T6_T13_MERGER_ABBREVIATIONS) {
    const wordRegex = new RegExp(`\\b${entry.word}\\b`, "i");
    if (wordRegex.test(caseName)) {
      issues.push({
        ruleId: "t6-t13-merger-abbreviation",
        message: `"${entry.word}" should be abbreviated as "${entry.abbreviation}" in a case name (Table T6).`,
        severity: "warning",
      });
    }
  }

  return issues;
}
