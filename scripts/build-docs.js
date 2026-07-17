const fs = require('fs');
const path = require('path');

const repoRoot = path.resolve(__dirname, '..');
const distDir = path.join(repoRoot, 'dist');
const repoRawBase = 'https://github.com/OpenClerkProject/openclerk-word/';

// The Markdown docs that are ALSO published as their own page on this site. A relative link to one
// of these (e.g. PRIVACY.md linking to TERMS.md, or to README.md#some-anchor) is rewritten to the
// sibling HTML page rather than off to GitHub, so the Privacy Policy / Terms of Use pages read as a
// self-contained site -- which is what Partner Center expects a Privacy/Terms URL to be. Every other
// relative link (source files, LICENSE) still points back to GitHub; see rewriteRelativeLinks.
const DOC_PAGES = {
  'README.md': 'index.html',
  'PRIVACY.md': 'privacy.html',
  'TERMS.md': 'terms.html',
};

// The set of pages to generate: source Markdown -> output file + <title>.
const PAGES = [
  { src: 'README.md', out: 'index.html', title: 'OpenClerk' },
  { src: 'PRIVACY.md', out: 'privacy.html', title: 'OpenClerk Privacy Policy' },
  { src: 'TERMS.md', out: 'terms.html', title: 'OpenClerk Terms of Use' },
];

// Matches GitHub's own heading-slug algorithm closely enough for this file: lowercase, strip
// anything that isn't a word character/space/hyphen, then turn spaces into hyphens. This is
// what makes in-page links like "(#security--it-review)" (already used throughout the README)
// actually land on the right heading once rendered as one long page.
//
// The exhaustive [^\w\- ] whitelist below is what actually makes this safe to drop into an
// HTML attribute (id="...") -- it strips every character that isn't alphanumeric/underscore/
// hyphen/space, full stop, so the output can never contain '<', '>', '"', or anything else
// HTML-meaningful, regardless of what the input looked like. (An earlier version of this
// function also ran a `.replace(/<[^>]+>/g, '')` tag-stripping pass first -- removed because a
// single-pass regex tag-stripper is a known-unreliable sanitization idiom on its own, and it
// added no actual protection here since this whitelist already subsumes it completely.)
function githubSlug(text) {
  return text
    .toLowerCase()
    .replace(/[^\w\- ]/g, '')
    .trim()
    .replace(/ /g, '-');
}

function buildRenderer(marked) {
  const renderer = new marked.Renderer();
  renderer.heading = function heading({ tokens, depth, text }) {
    const html = this.parser.parseInline(tokens);
    const id = githubSlug(text);
    return `<h${depth} id="${id}">${html}</h${depth}>\n`;
  };
  return renderer;
}

// README links to source files/folders with plain relative paths (e.g. "src/taskpane/word.ts",
// "src/taskpane/providers/"), which only resolve inside the repo. Rewritten to GitHub URLs so
// they work from the deployed docs site -- "tree" for directories (trailing slash), "blob" for
// files. Anchor-only links (#section) and already-absolute URLs are left untouched.
function rewriteRelativeLinks(html) {
  return html.replace(/href="([^"]+)"/g, (match, href) => {
    if (/^([a-z][a-z0-9+.-]*:|#)/i.test(href)) {
      return match;
    }
    // A link to a sibling published page (README/PRIVACY/TERMS, possibly with a #anchor) stays on
    // this site so the Privacy/Terms pages don't bounce a reviewer out to GitHub mid-document.
    const hashIndex = href.indexOf('#');
    const pathPart = hashIndex === -1 ? href : href.slice(0, hashIndex);
    const anchor = hashIndex === -1 ? '' : href.slice(hashIndex);
    if (Object.prototype.hasOwnProperty.call(DOC_PAGES, pathPart)) {
      return `href="${DOC_PAGES[pathPart]}${anchor}"`;
    }
    const kind = href.endsWith('/') ? 'tree' : 'blob';
    return `href="${repoRawBase}${kind}/main/${href}"`;
  });
}

function renderPage(bodyHtml, title) {
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>${title}</title>
<style>
  :root {
    color-scheme: light;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    font-family: "Segoe UI", -apple-system, Arial, sans-serif;
    color: #1f1f1f;
    background: #f7f7f7;
    line-height: 1.5;
  }
  header.site-header {
    background: #f0e6d6;
    border-bottom: 1px solid #c8b08a;
    padding: 20px 24px;
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
  }
  header.site-header h1 {
    margin: 0;
    font-size: 22px;
    color: #4a2f0f;
  }
  header.site-header nav a {
    color: #6f4518;
    text-decoration: none;
    font-weight: 600;
    margin-left: 16px;
    font-size: 14px;
  }
  header.site-header nav a:hover {
    text-decoration: underline;
  }
  main {
    max-width: 52rem;
    margin: 0 auto;
    padding: 24px;
  }
  main img { max-width: 100%; }
  main pre {
    background: #fffdf8;
    border: 1px solid #e2d1b0;
    border-radius: 6px;
    padding: 12px;
    overflow-x: auto;
  }
  main code {
    background: #fffdf8;
    border-radius: 3px;
    padding: 0 4px;
    font-size: 0.9em;
  }
  main pre code {
    background: none;
    padding: 0;
  }
  main table {
    border-collapse: collapse;
    width: 100%;
    margin: 1em 0;
    font-size: 14px;
  }
  main th, main td {
    border: 1px solid #e2d1b0;
    padding: 8px 10px;
    text-align: left;
    vertical-align: top;
  }
  main th {
    background: #f0e6d6;
  }
  main a { color: #6f4518; }
  main blockquote {
    border-left: 4px solid #b08a4b;
    margin: 1em 0;
    padding: 4px 16px;
    background: #fffdf8;
    color: #5f5344;
  }
  main h1, main h2, main h3 {
    color: #4a2f0f;
  }
  main h2 {
    border-bottom: 1px solid #e2d1b0;
    padding-bottom: 6px;
    margin-top: 2.2em;
  }
  footer.site-footer {
    text-align: center;
    color: #5f5344;
    font-size: 13px;
    padding: 24px;
  }
</style>
</head>
<body>
<header class="site-header">
  <h1>OpenClerk</h1>
  <nav>
    <a href="index.html">Home</a>
    <a href="privacy.html">Privacy</a>
    <a href="terms.html">Terms</a>
    <a href="https://github.com/OpenClerkProject/openclerk-word">Repository</a>
    <a href="https://github.com/OpenClerkProject/openclerk-word/releases">Releases</a>
  </nav>
</header>
<main>
${bodyHtml}
</main>
<footer class="site-footer">
  <a href="privacy.html">Privacy Policy</a> &middot; <a href="terms.html">Terms of Use</a><br />
  This site hosts OpenClerk's add-in content and documentation, rendered from the project's
  <a href="https://github.com/OpenClerkProject/openclerk-word">source repository</a>.
  OpenClerk is not a standalone web app &mdash; install it in Microsoft Word to use it.
</footer>
</body>
</html>
`;
}

async function main() {
  if (!fs.existsSync(distDir)) {
    throw new Error(`dist/ not found -- run the webpack build first: ${distDir}`);
  }

  // marked is published ESM-only; dynamic import() works from CommonJS on Node 18 (CI's build
  // job), unlike require(), which throws ERR_REQUIRE_ESM there.
  const { marked } = await import('marked');
  const renderer = buildRenderer(marked);

  for (const page of PAGES) {
    const srcPath = path.join(repoRoot, page.src);
    if (!fs.existsSync(srcPath)) {
      throw new Error(`Source doc not found: ${srcPath}`);
    }
    const markdown = fs.readFileSync(srcPath, 'utf8');
    const bodyHtml = rewriteRelativeLinks(marked.parse(markdown, { renderer }));
    const outPath = path.join(distDir, page.out);
    fs.writeFileSync(outPath, renderPage(bodyHtml, page.title), 'utf8');
    console.log('Wrote', outPath);
  }
}

main().catch((err) => {
  console.error(err instanceof Error ? err.message : err);
  process.exit(1);
});
