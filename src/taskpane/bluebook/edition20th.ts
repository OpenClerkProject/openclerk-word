import { ParsedCitation } from "../providers/types";
import { BluebookIssue, BluebookRuleSet } from "./types";
import { checkCommonCaseCitationRules } from "./commonRules";

/**
 * The Bluebook, 20th ed. (2015). Table T6 (case-name abbreviations) and
 * Table T13.2 (periodical/institutional-author abbreviations) were still
 * separate at this point, so several words that later editions abbreviate
 * in case names (see checkCaseNameAbbreviations.ts) were correctly spelled
 * out in full here -- this rule-set intentionally does not run that check.
 */
export class Bluebook20thEdition implements BluebookRuleSet {
  readonly id = "bluebook-20th";
  readonly name = "20th Edition (2015)";
  readonly description = "Table T6 and T13.2 were still separate; fewer case-name words are abbreviated.";

  checkCitation(citation: ParsedCitation): BluebookIssue[] {
    return checkCommonCaseCitationRules(citation);
  }
}
