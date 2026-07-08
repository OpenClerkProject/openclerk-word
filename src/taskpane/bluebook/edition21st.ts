import { ParsedCitation } from "../providers/types";
import { BluebookIssue, BluebookRuleSet } from "./types";
import { checkCommonCaseCitationRules } from "./commonRules";
import { checkCaseNameAbbreviations } from "./checkCaseNameAbbreviations";

/**
 * The Bluebook, 21st ed. (2020, incl. 2021+ printings). Merged Table T6
 * with the former Table T13.2, so the shared abbreviation list now applies
 * to case names too.
 */
export class Bluebook21stEdition implements BluebookRuleSet {
  readonly id = "bluebook-21st";
  readonly name = "21st Edition (2020)";
  readonly description = "Tables T6 and T13.2 were merged; more words are abbreviated in case names.";

  checkCitation(citation: ParsedCitation): BluebookIssue[] {
    return [...checkCommonCaseCitationRules(citation), ...checkCaseNameAbbreviations(citation)];
  }
}
