import { stripHtmlHyperlinks } from '../src/taskpane/utils';

describe('HTML hyperlink stripping', () => {
  test('removes a single anchor and preserves surrounding text', () => {
    const html = '<p>This is a <a href="http://example.com">link</a>.</p>';
    expect(stripHtmlHyperlinks(html)).toBe('This is a link.');
  });

  test('removes multiple anchors', () => {
    const html = '<div><a href="#">One</a> and <a href="#">Two</a></div>';
    expect(stripHtmlHyperlinks(html)).toBe('One and Two');
  });

  test('preserves inner formatted text inside anchor', () => {
    const html = '<a href="#"><b>Bold</b> and <i>Italic</i></a>';
    expect(stripHtmlHyperlinks(html)).toBe('Bold and Italic');
  });

  test('decodes common HTML entities', () => {
    const html = '<p><a href="#">AT&amp;T &lt;Corp&gt;</a></p>';
    expect(stripHtmlHyperlinks(html)).toBe('AT&T <Corp>');
  });

  test('handles empty or null input gracefully', () => {
    expect(stripHtmlHyperlinks('')).toBe('');
    expect(stripHtmlHyperlinks(null as unknown as string)).toBe('');
  });
});
