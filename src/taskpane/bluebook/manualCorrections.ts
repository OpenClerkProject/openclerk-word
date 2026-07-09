/**
 * Community-contributed Bluebook citation corrections and additions.
 *
 * Unlike everything in generated/*.ts, this file is NOT touched by
 * scripts/generate-bluebook-data.js and is never overwritten -- it's safe to hand-edit, and
 * entries here always take priority over the vendored reporters-db data. This is the one file
 * non-technical contributors are pointed at in CONTRIBUTING.md to fix or add a Bluebook rule
 * without touching any other code, tests, or build tooling.
 *
 * There are three lists below, matching the three ways the checker can be wrong. If you're not
 * sure which one applies, just open a GitHub issue describing the problem (see CONTRIBUTING.md)
 * and a maintainer will add the entry for you.
 */

export interface ManualReporterCorrection {
  /** The incorrect/non-standard form as it might appear in a document, e.g. "F.Supp.2d". */
  incorrectForm: string;
  /** The correct Table T1 form it should be changed to, e.g. "F. Supp. 2d". */
  correctForm: string;
  /** Full reporter series name, shown in the message, e.g. "Federal Supplement". */
  name: string;
  /** Where this correction came from: a GitHub issue URL, a law library guide, etc. */
  source: string;
}

export interface ManualValidReporterForm {
  /** A real, standard reporter abbreviation that reporters-db is missing, e.g. "N.M.". */
  form: string;
  /** Full reporter series name, shown in the message. */
  name: string;
  source: string;
}

export interface ManualCaseNameAbbreviation {
  /** Full word as it appears unabbreviated in running text (matched case-insensitively). */
  word: string;
  /** The Table T6 abbreviation it should become, e.g. "Ass'n". */
  abbreviation: string;
  source: string;
}

// --- Add new entries below. See CONTRIBUTING.md > "Contributing a Bluebook citation
// correction" for the exact format and a worked example. ---

/**
 * Reporter forms that should be flagged and corrected to a specific standard form
 * (the checker currently doesn't know about them or gets the correction wrong).
 *
 * Example:
 * { incorrectForm: "F.Supp.2d", correctForm: "F. Supp. 2d", name: "Federal Supplement",
 *   source: "https://github.com/OpenClerkProject/openclerk-word/issues/NN" },
 */
export const MANUAL_REPORTER_CORRECTIONS: ManualReporterCorrection[] = [];

/**
 * Real, standard reporter abbreviations that reporters-db is missing, so the checker wrongly
 * calls them "unrecognized" (a false positive).
 *
 * Example:
 * { form: "N.M.", name: "New Mexico Reports", source: "https://github.com/OpenClerkProject/openclerk-word/issues/NN" },
 */
export const MANUAL_VALID_REPORTER_FORMS: ManualValidReporterForm[] = [];

/**
 * Table T6 case-name word/abbreviation pairs that are missing or wrong in the vendored data.
 *
 * Example:
 * { word: "Corporation", abbreviation: "Corp.", source: "https://github.com/OpenClerkProject/openclerk-word/issues/NN" },
 */
export const MANUAL_CASE_NAME_ABBREVIATIONS: ManualCaseNameAbbreviation[] = [];
