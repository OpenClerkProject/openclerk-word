import { ParsedCitation } from "../providers/types";
import { BluebookIssue } from "./types";
import { REPORTER_ABBREVIATIONS } from "./generated/reporterAbbreviations.generated";
import {
  MANUAL_REPORTER_CORRECTIONS,
  MANUAL_VALID_REPORTER_FORMS,
  ManualReporterCorrection,
  ManualValidReporterForm,
} from "./manualCorrections";

type ReporterLookup = {
  validForms: Record<string, string>;
  corrections: Record<string, { correctForm: string; name: string }>;
};

/**
 * Layers community-contributed overrides (manualCorrections.ts) on top of the vendored
 * reporters-db lookup -- a manual valid-form entry always wins (removes any conflicting
 * generated "correction"), and a manual correction is skipped if something already accepts
 * that form as valid. Exported and kept pure so it's unit-testable with fixture data,
 * independent of the real (normally empty) manual-corrections file.
 */
export function applyManualReporterOverrides(
  generated: ReporterLookup,
  manualCorrections: ManualReporterCorrection[],
  manualValidForms: ManualValidReporterForm[]
): ReporterLookup {
  const validForms = { ...generated.validForms };
  const corrections = { ...generated.corrections };

  for (const entry of manualValidForms) {
    validForms[entry.form] = entry.name;
    delete corrections[entry.form];
  }
  for (const entry of manualCorrections) {
    if (entry.incorrectForm in validForms) {
      continue;
    }
    corrections[entry.incorrectForm] = { correctForm: entry.correctForm, name: entry.name };
  }

  return { validForms, corrections };
}

const { validForms, corrections } = applyManualReporterOverrides(
  REPORTER_ABBREVIATIONS as ReporterLookup,
  MANUAL_REPORTER_CORRECTIONS,
  MANUAL_VALID_REPORTER_FORMS
);

/**
 * Builds a "spacing-insensitive" index of validForms -- keyed by each valid form with internal
 * whitespace removed, mapping to the correctly-spaced form -- to catch Bluebook Rule 6.1 spacing
 * mistakes (e.g. "S.Ct." instead of "S. Ct.") generically. reporters-db's recorded "variations"
 * only cover the specific malformed spellings someone bothered to record (same limitation noted
 * for the ordinal-typo check above); a spot-check found 805 of the 920 valid forms that contain
 * internal spacing have no recorded spacing-variant correction at all. A form that multiple
 * distinct valid forms collapse to (rare, but possible) is excluded rather than guessed at.
 */
export function buildReporterSpacingLookup(forms: Record<string, string>): Record<string, string> {
  const strippedToForm: Record<string, string> = {};
  const ambiguous = new Set<string>();
  for (const form of Object.keys(forms)) {
    const stripped = form.replace(/\s+/g, "");
    if (stripped in strippedToForm && strippedToForm[stripped] !== form) {
      ambiguous.add(stripped);
      continue;
    }
    strippedToForm[stripped] = form;
  }
  Array.from(ambiguous).forEach((key) => {
    delete strippedToForm[key];
  });
  return strippedToForm;
}

const spacingLookup = buildReporterSpacingLookup(validForms);

/**
 * Checks a citation's reporter abbreviation against Free Law Project's reporters-db (Table T1
 * data, see generated/reporterAbbreviations.generated.ts) -- vendored at dev time from
 * https://github.com/freelawproject/reporters-db, not fetched at runtime.
 *
 * Edition-independent: reporter abbreviations themselves don't change between Bluebook
 * editions the way case-name/statutory abbreviations do, so this applies to every edition.
 */
export function checkReporterAbbreviation(citation: ParsedCitation): BluebookIssue[] {
  const reporter = citation.reporter;
  if (!reporter) {
    return [];
  }

  if (reporter in validForms) {
    return [];
  }

  // reporters-db's "variations" only lists the specific malformed forms someone bothered to
  // record (it's built from real-world parsing hits, not an exhaustive enumeration), so it only
  // catches the ordinal typo ("2nd" instead of "2d") for a handful of reporters. Normalize and
  // re-check generically so this common mistake is caught for every reporter, not just those.
  const ordinalNormalized = reporter.replace(/\b(\d)(?:nd|rd)\b/g, "$1d");
  if (ordinalNormalized !== reporter && ordinalNormalized in validForms) {
    return [
      {
        ruleId: "reporter-ordinal",
        message: `Reporter series should use "${ordinalNormalized}" (ordinal abbreviations like "2nd"/"3rd" should be "2d"/"3d") -- found "${reporter}".`,
        severity: "error",
      },
    ];
  }

  const correction = corrections[reporter];
  if (correction) {
    return [
      {
        ruleId: "reporter-nonstandard-form",
        message: `"${reporter}" is a known non-standard form of "${correction.correctForm}" (${correction.name}); use "${correction.correctForm}" (Table T1).`,
        severity: "error",
      },
    ];
  }

  // Generic Rule 6.1 spacing check -- only reached once the reporter didn't match anything
  // reporters-db already recorded a specific correction for above.
  const strippedReporter = reporter.replace(/\s+/g, "");
  const correctlySpaced = spacingLookup[strippedReporter];
  if (correctlySpaced && correctlySpaced !== reporter) {
    return [
      {
        ruleId: "reporter-spacing",
        message: `Reporter spacing should be "${correctlySpaced}" (Bluebook Rule 6.1: close up adjacent single capitals, but space out longer abbreviations) -- found "${reporter}".`,
        severity: "error",
      },
    ];
  }

  return [
    {
      ruleId: "reporter-unrecognized",
      message: `"${reporter}" was not found in the Table T1 reporter database -- verify this reporter abbreviation manually.`,
      severity: "warning",
    },
  ];
}
