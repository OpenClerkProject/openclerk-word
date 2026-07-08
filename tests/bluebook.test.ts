import { parseCaseCitation } from '../src/taskpane/providers/citationParser';
import { BluebookRuleSetRegistry } from '../src/taskpane/bluebook/registry';
import { Bluebook20thEdition } from '../src/taskpane/bluebook/edition20th';
import { Bluebook21stEdition } from '../src/taskpane/bluebook/edition21st';
import { Bluebook22ndEdition } from '../src/taskpane/bluebook/edition22nd';
import { checkCommonCaseCitationRules } from '../src/taskpane/bluebook/commonRules';
import { applyManualReporterOverrides } from '../src/taskpane/bluebook/reporterRules';
import { applyManualCaseNameOverrides } from '../src/taskpane/bluebook/checkCaseNameAbbreviations';
import { ParsedCitation } from '../src/taskpane/providers/types';

const EXAMPLE_CITATION = 'Norfolk & W. Ry. Co. v. Liepelt, 444 U.S. 490 (U.S.Ill., 1980)';

function parseOrThrow(raw: string): ParsedCitation {
  const parsed = parseCaseCitation(raw);
  if (!parsed) {
    throw new Error(`Test fixture citation failed to parse: ${raw}`);
  }
  return parsed;
}

describe('BluebookRuleSetRegistry', () => {
  test('registers and retrieves rule-sets by id', () => {
    const registry = new BluebookRuleSetRegistry();
    const ruleSet = new Bluebook22ndEdition();
    registry.register(ruleSet);

    expect(registry.get('bluebook-22nd')).toBe(ruleSet);
    expect(registry.list()).toContain(ruleSet);
  });

  test('returns undefined for an unknown id', () => {
    expect(new BluebookRuleSetRegistry().get('nope')).toBeUndefined();
  });
});

describe('checkCommonCaseCitationRules', () => {
  test('a well-formed citation has no issues', () => {
    expect(checkCommonCaseCitationRules(parseOrThrow(EXAMPLE_CITATION))).toEqual([]);
  });

  test('flags "vs." instead of "v."', () => {
    const issues = checkCommonCaseCitationRules(parseOrThrow('Norfolk & W. Ry. Co. vs. Liepelt, 444 U.S. 490 (1980)'));
    expect(issues.some((i) => i.ruleId === 'v-abbreviation')).toBe(true);
  });

  test('flags "v" missing its period', () => {
    const issues = checkCommonCaseCitationRules(parseOrThrow('Norfolk & W. Ry. Co. v Liepelt, 444 U.S. 490 (1980)'));
    expect(issues.some((i) => i.ruleId === 'v-period')).toBe(true);
  });

  test('flags "2nd"/"3rd" instead of "2d"/"3d" in the reporter', () => {
    const issues = checkCommonCaseCitationRules(parseOrThrow('Doe v. Roe, 955 So. 2nd 425 (Fla. 2007)'));
    expect(issues.some((i) => i.ruleId === 'reporter-ordinal')).toBe(true);
  });

  test('flags a missing year', () => {
    const issues = checkCommonCaseCitationRules(parseOrThrow('Smith v. Jones, 123 F.3d 456'));
    expect(issues.some((i) => i.ruleId === 'year-required')).toBe(true);
  });

  test('flags a missing court abbreviation for a non-U.S.-Reports reporter', () => {
    const issues = checkCommonCaseCitationRules(parseOrThrow('Smith v. Jones, 123 F.3d 456 (2010)'));
    expect(issues.some((i) => i.ruleId === 'court-abbreviation-required')).toBe(true);
  });

  test('does not require a court abbreviation for U.S. Reports citations', () => {
    const issues = checkCommonCaseCitationRules(parseOrThrow('Brown v. Board of Education, 347 U.S. 483 (1954)'));
    expect(issues.some((i) => i.ruleId === 'court-abbreviation-required')).toBe(false);
  });
});

describe('edition-specific case-name abbreviation checks', () => {
  const CASE_NAME_WITH_UNABBREVIATED_WORD =
    'Sierra Club v. Metro Environment Research Laboratory, 500 F.3d 100 (9th Cir. 2019)';

  test('20th edition does not flag pre-merger Table T6/T13.2 words', () => {
    const issues = new Bluebook20thEdition().checkCitation(parseOrThrow(CASE_NAME_WITH_UNABBREVIATED_WORD));
    expect(issues.some((i) => i.ruleId === 't6-t13-merger-abbreviation')).toBe(false);
  });

  test('21st edition flags the same words once T6/T13.2 are merged', () => {
    const issues = new Bluebook21stEdition().checkCitation(parseOrThrow(CASE_NAME_WITH_UNABBREVIATED_WORD));
    const flagged = issues.filter((i) => i.ruleId === 't6-t13-merger-abbreviation');
    expect(flagged.length).toBeGreaterThan(0);
    expect(flagged.some((i) => i.message.includes('Environment'))).toBe(true);
    expect(flagged.some((i) => i.message.includes('Laboratory'))).toBe(true);
  });

  test('22nd edition carries the merged table forward', () => {
    const issues = new Bluebook22ndEdition().checkCitation(parseOrThrow(CASE_NAME_WITH_UNABBREVIATED_WORD));
    expect(issues.some((i) => i.ruleId === 't6-t13-merger-abbreviation')).toBe(true);
  });

  test('all three editions still run the shared common-rule checks', () => {
    const badCitation = parseOrThrow('Norfolk & W. Ry. Co. vs. Liepelt, 444 U.S. 490 (1980)');
    for (const ruleSet of [new Bluebook20thEdition(), new Bluebook21stEdition(), new Bluebook22ndEdition()]) {
      expect(ruleSet.checkCitation(badCitation).some((i) => i.ruleId === 'v-abbreviation')).toBe(true);
    }
  });

  test('checkCitation never throws, even on a minimal/edge-case citation', () => {
    const minimal: ParsedCitation = { raw: 'X v. Y, 1 U.S. 1' };
    for (const ruleSet of [new Bluebook20thEdition(), new Bluebook21stEdition(), new Bluebook22ndEdition()]) {
      expect(() => ruleSet.checkCitation(minimal)).not.toThrow();
    }
  });
});

describe('reporter abbreviation checks (vendored reporters-db Table T1 data)', () => {
  test('a valid reporter edition form is not flagged', () => {
    const issues = checkCommonCaseCitationRules(parseOrThrow('Smith v. Jones, 123 A.2d 456 (Del. 2010)'));
    expect(issues.some((i) => i.ruleId.startsWith('reporter-'))).toBe(false);
  });

  test('flags a known non-standard reporter form with the correct suggestion', () => {
    const issues = checkCommonCaseCitationRules(parseOrThrow('Smith v. Jones, 123 F. 2d 456 (2d Cir. 2010)'));
    const flagged = issues.find((i) => i.ruleId === 'reporter-nonstandard-form');
    expect(flagged).toBeDefined();
    expect(flagged?.message).toContain('F.2d');
  });

  test('flags an ordinal-form reporter typo generically, not just the ~13 reporters reporters-db happens to list', () => {
    const issues = checkCommonCaseCitationRules(parseOrThrow('Smith v. Jones, 123 A.2nd 456 (Del. 2010)'));
    const flagged = issues.find((i) => i.ruleId === 'reporter-ordinal');
    expect(flagged).toBeDefined();
    expect(flagged?.message).toContain('A.2d');
  });

  test('flags a completely unrecognized reporter as a warning, not an error', () => {
    const issues = checkCommonCaseCitationRules(parseOrThrow('Smith v. Jones, 123 Not.A.Real.Reporter 456 (2010)'));
    const flagged = issues.find((i) => i.ruleId === 'reporter-unrecognized');
    expect(flagged).toBeDefined();
    expect(flagged?.severity).toBe('warning');
  });

  test('flags a reporter spacing mistake (Rule 6.1) not already recorded as a specific reporters-db correction', () => {
    // "Ohio App." is a valid form, but reporters-db has no recorded correction for the
    // run-together "OhioApp." -- only the generic spacing check catches this one.
    const issues = checkCommonCaseCitationRules(parseOrThrow('Smith v. Jones, 123 OhioApp. 456 (2010)'));
    const flagged = issues.find((i) => i.ruleId === 'reporter-spacing');
    expect(flagged).toBeDefined();
    expect(flagged?.message).toContain('Ohio App.');
    expect(flagged?.severity).toBe('error');
  });

  test('a specific reporters-db correction still wins over the generic spacing check', () => {
    // "F. 2d" already has a recorded reporters-db correction to "F.2d" -- confirms the new
    // generic spacing check doesn't shadow the existing, more specific message.
    const issues = checkCommonCaseCitationRules(parseOrThrow('Smith v. Jones, 123 F. 2d 456 (2d Cir. 2010)'));
    expect(issues.some((i) => i.ruleId === 'reporter-nonstandard-form')).toBe(true);
    expect(issues.some((i) => i.ruleId === 'reporter-spacing')).toBe(false);
  });
});

describe('page-range digit-dropping checks (Bluebook Rule 3.2)', () => {
  test('flags a page range written in full instead of with dropped digits', () => {
    const issues = checkCommonCaseCitationRules(
      parseOrThrow('United States v. Nixon, 418 U.S. 683, 705-706 (1974)')
    );
    const flagged = issues.find((i) => i.ruleId === 'pincite-range-digits');
    expect(flagged).toBeDefined();
    expect(flagged?.message).toContain('705-06');
  });

  test('does not flag an already-correctly-dropped page range', () => {
    const issues = checkCommonCaseCitationRules(
      parseOrThrow('United States v. Nixon, 418 U.S. 683, 705-06 (1974)')
    );
    expect(issues.some((i) => i.ruleId === 'pincite-range-digits')).toBe(false);
  });

  test('flags an under-dropped page range (fewer than the required last two digits)', () => {
    const issues = checkCommonCaseCitationRules(
      parseOrThrow('United States v. Nixon, 418 U.S. 683, 705-6 (1974)')
    );
    const flagged = issues.find((i) => i.ruleId === 'pincite-range-digits');
    expect(flagged).toBeDefined();
    expect(flagged?.message).toContain('705-06');
  });

  test('does not flag a single (non-range) pincite page', () => {
    const issues = checkCommonCaseCitationRules(
      parseOrThrow('Norfolk & W. Ry. Co. v. Liepelt, 444 U.S. 490, 496 (1980)')
    );
    expect(issues.some((i) => i.ruleId === 'pincite-range-digits')).toBe(false);
  });
});

describe('full case-name abbreviation table (vendored reporters-db data, all editions)', () => {
  const CASE_NAME_WITH_COMMON_ABBREVIATION = 'Smith v. Acme Company, 123 F.3d 456 (9th Cir. 2010)';

  test('flags a word from the full table for the 20th edition too (not just the T6/T13.2 merger words)', () => {
    const issues = new Bluebook20thEdition().checkCitation(parseOrThrow(CASE_NAME_WITH_COMMON_ABBREVIATION));
    const flagged = issues.find((i) => i.ruleId === 'case-name-abbreviation');
    expect(flagged).toBeDefined();
    expect(flagged?.message).toContain('Co.');
  });

  test('flags the same word for the 22nd edition', () => {
    const issues = new Bluebook22ndEdition().checkCitation(parseOrThrow(CASE_NAME_WITH_COMMON_ABBREVIATION));
    expect(issues.some((i) => i.ruleId === 'case-name-abbreviation')).toBe(true);
  });

  test('does not double-flag a T6/T13.2 merger word under both the full table and the merger-specific check', () => {
    const issues = new Bluebook21stEdition().checkCitation(
      parseOrThrow('Sierra Club v. Metro Environment Laboratory, 500 F.3d 100 (9th Cir. 2019)')
    );
    const fullTableFlags = issues.filter((i) => i.ruleId === 'case-name-abbreviation' && i.message.includes('Environment'));
    expect(fullTableFlags.length).toBe(0);
  });
});

describe('community-contributed manual corrections (manualCorrections.ts)', () => {
  test('a manual correction adds a new flagged form on top of the generated data', () => {
    const { corrections } = applyManualReporterOverrides(
      { validForms: { 'F. Supp. 2d': 'Federal Supplement' }, corrections: {} },
      [{ incorrectForm: 'F.Supp.2d', correctForm: 'F. Supp. 2d', name: 'Federal Supplement', source: 'test' }],
      []
    );
    expect(corrections['F.Supp.2d']).toEqual({ correctForm: 'F. Supp. 2d', name: 'Federal Supplement' });
  });

  test('a manual valid-form entry removes any conflicting generated "correction"', () => {
    const { validForms, corrections } = applyManualReporterOverrides(
      { validForms: {}, corrections: { 'N.M.': { correctForm: 'N. M.', name: 'New Mexico Reports' } } },
      [],
      [{ form: 'N.M.', name: 'New Mexico Reports', source: 'test' }]
    );
    expect(validForms['N.M.']).toBe('New Mexico Reports');
    expect(corrections['N.M.']).toBeUndefined();
  });

  test('a manual correction is skipped when the form is already accepted as valid', () => {
    const { corrections } = applyManualReporterOverrides(
      { validForms: { 'A.2d': 'Atlantic Reporter, Second Series' }, corrections: {} },
      [{ incorrectForm: 'A.2d', correctForm: 'A. 2d', name: 'Atlantic Reporter, Second Series', source: 'test' }],
      []
    );
    expect(corrections['A.2d']).toBeUndefined();
  });

  test('a manual case-name abbreviation is added (and can override a generated one)', () => {
    const table = applyManualCaseNameOverrides(
      { corporation: 'Corp.' },
      [
        { word: 'Corporation', abbreviation: 'Corp', source: 'test' },
        { word: 'Institute', abbreviation: "Inst.", source: 'test' },
      ]
    );
    expect(table['corporation']).toBe('Corp');
    expect(table['institute']).toBe('Inst.');
  });
});

describe('court/jurisdiction state abbreviation checks (vendored reporters-db Table T10 data)', () => {
  test('flags a spelled-out state name in the court parenthetical', () => {
    const issues = checkCommonCaseCitationRules(parseOrThrow('Smith v. Jones, 123 Cal. Rptr. 456 (California 2010)'));
    const flagged = issues.find((i) => i.ruleId === 'court-state-not-abbreviated');
    expect(flagged).toBeDefined();
    expect(flagged?.message).toContain('Cal.');
  });

  test('does not flag an already-correct state abbreviation', () => {
    const issues = checkCommonCaseCitationRules(parseOrThrow('Smith v. Jones, 123 Cal. Rptr. 456 (Cal. 2010)'));
    expect(issues.some((i) => i.ruleId === 'court-state-not-abbreviated')).toBe(false);
  });
});
