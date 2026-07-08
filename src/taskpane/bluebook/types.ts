import { ParsedCitation } from "../providers/types";

/**
 * Plugin architecture for Bluebook-format checking, mirroring the citation
 * lookup provider plugins in ../providers/. A BluebookRuleSet represents one
 * edition/version of The Bluebook; different editions can register different
 * checks without the checking UI needing to know which edition is active.
 */

export type BluebookIssueSeverity = "error" | "warning";

export interface BluebookIssue {
  /** Short machine-readable identifier for the specific rule that flagged this, e.g. "v-period". */
  ruleId: string;
  /** One-sentence, human-readable description of the problem. */
  message: string;
  severity: BluebookIssueSeverity;
}

export interface BluebookRuleSet {
  readonly id: string;
  readonly name: string;
  /** One-sentence description shown next to the edition in the dropdown. */
  readonly description: string;
  /**
   * Checks a single parsed case citation and returns any formatting issues.
   * An empty array means the citation looks correct under this edition's
   * rules. Must never throw -- a rule-set that can't make sense of a
   * citation should simply not flag anything for it.
   */
  checkCitation(citation: ParsedCitation): BluebookIssue[];
}
