import { expandPincitePages, reconstructFullPageNumber } from '../src/taskpane/providers/pincitePages';
import { extractPageExcerpt, stripHtmlTags } from '../src/taskpane/providers/opinionTextExtractor';

describe('reconstructFullPageNumber', () => {
  test('borrows leading digits from the start page for a dropped end page', () => {
    expect(reconstructFullPageNumber('705', '06')).toBe('706');
    expect(reconstructFullPageNumber('1099', '101')).toBe('1101');
  });

  test('returns the end page unchanged when it is already full-length', () => {
    expect(reconstructFullPageNumber('705', '706')).toBe('706');
    expect(reconstructFullPageNumber('99', '100')).toBe('100');
  });
});

describe('expandPincitePages', () => {
  test('a single page', () => {
    expect(expandPincitePages('496')).toEqual([496]);
  });

  test('a comma-separated list of single pages', () => {
    expect(expandPincitePages('505, 508, 513')).toEqual([505, 508, 513]);
  });

  test('a Bluebook-dropped-digit range expands to every page in between, inclusive', () => {
    expect(expandPincitePages('705-06')).toEqual([705, 706]);
  });

  test('a wider range expands to every page in between, inclusive', () => {
    expect(expandPincitePages('1099-101')).toEqual([1099, 1100, 1101]);
  });

  test('a mix of a range and single pages, deduplicated and sorted', () => {
    expect(expandPincitePages('705-06, 705, 710')).toEqual([705, 706, 710]);
  });
});

describe('extractPageExcerpt', () => {
  const OPINION_TEXT =
    'Some introductory text.\n' +
    '*703 The court now turns to the merits of the claim.\n' +
    '*704 Continuing the analysis of the statute.\n' +
    '*705 Here is the holding on the first issue.\n' +
    '*706 And here is the holding on the second issue.\n' +
    '*707 Concluding remarks follow.';

  test('extracts the text for a single requested page', () => {
    const excerpt = extractPageExcerpt(OPINION_TEXT, [705]);
    expect(excerpt).toContain('Here is the holding on the first issue.');
    expect(excerpt).not.toContain('Continuing the analysis');
    expect(excerpt).not.toContain('second issue');
  });

  test('extracts and joins text for a range of requested pages, in order', () => {
    const excerpt = extractPageExcerpt(OPINION_TEXT, [705, 706]);
    expect(excerpt).toContain('first issue');
    expect(excerpt).toContain('second issue');
    expect(excerpt?.indexOf('first issue')).toBeLessThan(excerpt?.indexOf('second issue') ?? -1);
  });

  test('returns null when the text has no star-pagination markers at all', () => {
    expect(extractPageExcerpt('Plain opinion text with no page markers.', [705])).toBeNull();
  });

  test('returns null when markers exist but none match the requested pages', () => {
    expect(extractPageExcerpt(OPINION_TEXT, [900])).toBeNull();
  });

  test('returns null for an empty target page list', () => {
    expect(extractPageExcerpt(OPINION_TEXT, [])).toBeNull();
  });

  test('the last requested page runs to the end of the text', () => {
    const excerpt = extractPageExcerpt(OPINION_TEXT, [707]);
    expect(excerpt).toContain('Concluding remarks follow.');
  });
});

describe('stripHtmlTags', () => {
  test('removes tags and collapses whitespace', () => {
    expect(stripHtmlTags('<p>Hello <b>world</b></p>\n<p>Second   paragraph</p>')).toBe('Hello world Second paragraph');
  });

  test('decodes common HTML entities', () => {
    expect(stripHtmlTags('Smith &amp; Jones &lt;1980&gt; &quot;quoted&quot; &#39;text&#39;&nbsp;end')).toBe(
      'Smith & Jones <1980> "quoted" \'text\' end'
    );
  });

  test('returns an empty string for empty input', () => {
    expect(stripHtmlTags('')).toBe('');
  });
});
