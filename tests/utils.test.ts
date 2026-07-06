import { normalizeText, isLikelyCaseCitation, extractParentheticalCitations } from '../src/taskpane/utils';

describe('Citation helpers', () => {
  test('normalizeText removes NBSP and collapses spaces', () => {
    const input = 'This\u00A0is   a\t test';
    expect(normalizeText(input)).toBe('This is a test');
  });

  test('isLikelyCaseCitation detects v. and years', () => {
    expect(isLikelyCaseCitation('Smith v. Jones')).toBe(true);
    expect(isLikelyCaseCitation('Smith v Jones')).toBe(true);
    expect(isLikelyCaseCitation('In re Estate of Foo, 2001')).toBe(true);
    expect(isLikelyCaseCitation('Random text without case')).toBe(false);
  });

  test('extractParentheticalCitations finds parentheticals', () => {
    const text = 'See (Smith v. Jones), and also (2020). Ignore empty ().';
    const result = extractParentheticalCitations(text);
    expect(result).toContain('Smith v. Jones');
    expect(result).toContain('2020');
  });
});
