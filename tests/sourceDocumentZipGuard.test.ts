/*
 * Regression tests for the zip-bomb preflight in src/taskpane/word.ts (readZipEntryWithLimit).
 *
 * A maliciously crafted source .docx is a small archive whose entries expand to enormous amounts
 * of memory once decompressed. The guard rejects such an entry from its declared central-directory
 * sizes *before* calling entry.async(), which is the call that would actually allocate the expanded
 * bytes -- the old code only checked content.length after decompression, i.e. after the bomb had
 * already gone off. See commit 362ee6a, which first added the whole-file / entry-count / post-hoc
 * caps; this suite covers the pre-decompression preflight that hardened it further.
 */

// word.ts references Word.InsertLocation.replace at module scope and calls Office.onReady(...) at
// load. Neither is a real global under Jest's node environment, so stub the tiny surface it touches
// BEFORE requiring the module (require, not import, to keep this order) -- mirrors
// tests/smoke.wordInsertion.test.ts.
(global as any).Word = { InsertLocation: { replace: 'Replace' } };
(global as any).Office = { onReady: () => {}, HostType: { Word: 'Word' } };

/* eslint-disable @typescript-eslint/no-var-requires */
import JSZip from 'jszip';
const { readZipEntryWithLimit } = require('../src/taskpane/word');

const OVER_LIMIT_MESSAGE = /unexpectedly large/;

// Reloads a generated archive through JSZip.loadAsync so the entry carries the declared
// compressed/uncompressed sizes read straight from the zip central directory -- the same code path
// a real uploaded .docx takes -- rather than the in-memory sizes of a freshly-authored entry.
async function reloadZip(zip: JSZip): Promise<JSZip> {
  const buffer = await zip.generateAsync({ type: 'arraybuffer' });
  return JSZip.loadAsync(buffer);
}

describe('word.ts zip-bomb defenses (readZipEntryWithLimit)', () => {
  it('rejects a single entry that declares an over-limit uncompressed size, before decompressing', async () => {
    // 60 MB of a single repeated byte: well above the 50 MB per-entry cap, but a few dozen KB
    // compressed -- exactly the shape of a zip bomb.
    const zip = new JSZip();
    zip.file('word/document.xml', 'A'.repeat(60 * 1024 * 1024), { compression: 'DEFLATE' });
    const reloaded = await reloadZip(zip);

    await expect(readZipEntryWithLimit(reloaded, 'word/document.xml')).rejects.toThrow(
      OVER_LIMIT_MESSAGE
    );
  });

  it('rejects a highly-compressible entry on its declared ratio even when it is under the size cap', async () => {
    // 5 MB of a repeated byte: under the 50 MB size cap (so only the ratio guard can catch it), but
    // compresses at roughly 1000:1 -- far above the 200:1 threshold a real document XML part ever
    // reaches, and above the 1 MB floor -- so it must still be rejected before decompression.
    const zip = new JSZip();
    zip.file('word/document.xml', 'A'.repeat(5 * 1024 * 1024), { compression: 'DEFLATE' });
    const reloaded = await reloadZip(zip);

    await expect(readZipEntryWithLimit(reloaded, 'word/document.xml')).rejects.toThrow(
      OVER_LIMIT_MESSAGE
    );
  });

  it('still reads a normal, legitimately-sized entry', async () => {
    const zip = new JSZip();
    zip.file('word/document.xml', '<w:document>Ashcroft v. Iqbal, 556 U.S. 662 (2009)</w:document>', {
      compression: 'DEFLATE',
    });
    const reloaded = await reloadZip(zip);

    await expect(readZipEntryWithLimit(reloaded, 'word/document.xml')).resolves.toContain(
      'Ashcroft'
    );
  });

  it('returns undefined for an absent entry (missing part is not an error)', async () => {
    const zip = new JSZip();
    zip.file('word/document.xml', '<w:document/>', { compression: 'DEFLATE' });
    const reloaded = await reloadZip(zip);

    await expect(readZipEntryWithLimit(reloaded, 'word/_rels/document.xml.rels')).resolves.toBeUndefined();
  });
});
