import { parseCaseCitation } from '../src/taskpane/providers/citationParser';
import { BluebookRuleSetRegistry } from '../src/taskpane/bluebook/registry';
import { Bluebook20thEdition } from '../src/taskpane/bluebook/edition20th';
import { Bluebook21stEdition } from '../src/taskpane/bluebook/edition21st';
import { Bluebook22ndEdition } from '../src/taskpane/bluebook/edition22nd';
import { checkCommonCaseCitationRules } from '../src/taskpane/bluebook/commonRules';
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
