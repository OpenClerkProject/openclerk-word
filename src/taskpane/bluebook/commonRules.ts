import { ParsedCitation } from "../providers/types";
import { BluebookIssue } from "./types";

/**
 * Checks for Bluebook Rule 10 (case citation) conventions that have stayed
 * stable across at least the last three editions (20th/2015, 21st/2020,
 * 22nd/2025) -- edition-specific rule-sets all run these, then layer their
 * own edition-specific checks on top (see caseNameAbbreviations.ts).
 */
export function checkCommonCaseCitationRules(citation: ParsedCitation): BluebookIssue[] {
  const issues: BluebookIssue[] = [];
  const caseName = citation.caseName || "";

  if (/\svs\.?\s/i.test(caseName)) {
    issues.push({
      ruleId: "v-abbreviation",
      message: 'Use "v." (not "vs." or "vs") to abbreviate "versus" between party names.',
      severity: "error",
    });
  } else if (/\sv\s/.test(caseName) && !/\sv\.\s/.test(caseName)) {
    issues.push({
      ruleId: "v-period",
      message: '"v." should include a period.',
      severity: "error",
    });
  }

  if (citation.reporter && /\b\d(?:nd|rd)\b/.test(citation.reporter)) {
    issues.push({
      ruleId: "reporter-ordinal",
      message: `Reporter series should use "2d"/"3d", not the ordinal form found in "${citation.reporter}".`,
      severity: "error",
    });
  }

  if (!citation.year) {
    issues.push({
      ruleId: "year-required",
      message: "No year found in the citation's parenthetical; Bluebook Rule 10.5 requires the decision year.",
      severity: "error",
    });
  }

  if (citation.reporter && citation.reporter !== "U.S." && !citation.court) {
    issues.push({
      ruleId: "court-abbreviation-required",
      message:
        'A court abbreviation is expected in the parenthetical for this reporter (Rule 10.4) -- only U.S. Reports ("U.S.") citations can omit it.',
      severity: "warning",
    });
  }

  return issues;
}
