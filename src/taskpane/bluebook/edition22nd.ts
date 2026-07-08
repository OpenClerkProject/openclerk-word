import { ParsedCitation } from "../providers/types";
import { BluebookIssue, BluebookRuleSet } from "./types";
import { checkCommonCaseCitationRules } from "./commonRules";
import { checkCaseNameAbbreviations } from "./checkCaseNameAbbreviations";

/**
 * The Bluebook, 22nd ed. (2025) -- the current edition as of this writing.
 * Its documented changes (typeface terminology renamed to "small capitals";
 * mandatory web-archiving for online sources; a new "contrast" signal; a
 * rewritten Rule 18 covering AI-generated content) don't affect case
 * citation (Rule 10) format, so this rule-set runs the same case-citation
 * checks as the 21st edition (the merged T6/T13.2 table carries forward
 * unchanged). See https://lib.law.uw.edu/bluebook101/22nd.
 */
export class Bluebook22ndEdition implements BluebookRuleSet {
  readonly id = "bluebook-22nd";
  readonly name = "22nd Edition (2025)";
  readonly description = "Current edition. Case-citation format is unchanged from the 21st edition.";

  checkCitation(citation: ParsedCitation): BluebookIssue[] {
    return [...checkCommonCaseCitationRules(citation), ...checkCaseNameAbbreviations(citation)];
  }
}
