const fs = require('fs');
const path = require('path');
const https = require('https');
const crypto = require('crypto');
const { spawnSync } = require('child_process');
const archiver = require('archiver');

const repoRoot = path.resolve(__dirname, '..');
const distDir = path.join(repoRoot, 'dist');
const manifestPath = path.join(distDir, 'manifest.xml');
const outputPath = path.join(repoRoot, 'openclerk-addin-offline.zip');

// The taskpane/commands HTML load office.js and Fabric's CSS from Microsoft's CDN, which
// defeats the whole point of an offline package. Vendor local copies at packaging time (this
// runs in CI/dev, which has internet) and rewrite the HTML to reference them instead -- the
// GitHub Pages build is untouched, since CDN loading is the right call when a network is assumed.
// Each asset carries an optional `sha256` (lowercase hex of the raw downloaded bytes). When set,
// the download is verified against it and packaging fails hard on any mismatch, so a compromised,
// MITM'd, or silently-changed CDN response can never be vendored into the shipped offline package.
// When null, packaging still proceeds but logs the freshly-computed digest plus a warning, making
// it a one-line change to pin the asset. See vendorCdnAssets.
const VENDOR_ASSETS = [
  {
    // Evergreen production URL: Microsoft updates office.js in place at this path (there is no
    // Microsoft-published versioned production URL to pin instead), so a hardcoded digest would
    // break this build on every office.js update. Left unpinned deliberately.
    // TODO: pin `sha256` once a stable/versioned office.js URL exists, or wire this to a
    // maintainer-reviewed digest bumped deliberately alongside each office.js update.
    url: 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js',
    localName: 'office.js',
    sha256: null,
  },
  {
    // Fully version-pinned, immutable CDN path (fabric-cdn-prod_20230815.002/.../11.1.0/), so its
    // bytes never change and its digest can and should be pinned.
    // TODO: pin `sha256` with this exact asset's digest to enforce integrity. Capture it from a
    // trusted machine (`curl -sS <url> | sha256sum`); this build environment's proxy blocks the
    // Microsoft CDN, so it could not be captured in place.
    url:
      'https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css',
    localName: 'fabric.min.css',
    sha256: null,
  },
];

function fetchAsset(url, redirectsLeft = 5) {
  return new Promise((resolve, reject) => {
    https
      .get(url, (res) => {
        if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location && redirectsLeft > 0) {
          res.resume();
          resolve(fetchAsset(res.headers.location, redirectsLeft - 1));
          return;
        }
        if (res.statusCode !== 200) {
          reject(new Error(`Failed to fetch ${url}: HTTP ${res.statusCode}`));
          res.resume();
          return;
        }
        // Collect raw bytes (not a utf8 string) so the integrity digest below is computed over the
        // exact content served, and the vendored copy is written byte-for-byte identical to it.
        const chunks = [];
        res.on('data', (chunk) => chunks.push(chunk));
        res.on('end', () => resolve(Buffer.concat(chunks)));
      })
      .on('error', reject);
  });
}

async function vendorCdnAssets(vendorDir) {
  fs.mkdirSync(vendorDir, { recursive: true });
  for (const asset of VENDOR_ASSETS) {
    console.log(`Vendoring ${asset.url} -> vendor/${asset.localName}`);
    const content = await fetchAsset(asset.url);
    const digest = crypto.createHash('sha256').update(content).digest('hex');
    if (asset.sha256) {
      // Fail closed: never vendor an asset whose bytes don't match the pinned digest.
      if (digest !== asset.sha256.toLowerCase()) {
        throw new Error(
          `Integrity check failed for ${asset.url}: expected sha256 ${asset.sha256}, got ${digest}. ` +
            'Refusing to vendor an unexpected asset into the offline package.'
        );
      }
      console.log(`  sha256 OK (${digest})`);
    } else {
      console.warn(
        `  WARNING: ${asset.localName} is not integrity-pinned. Downloaded sha256 is ${digest}. ` +
          'Pin this value in VENDOR_ASSETS[].sha256 to enforce integrity (see the TODO there).'
      );
    }
    fs.writeFileSync(path.join(vendorDir, asset.localName), content);
  }
}

function rewriteCdnReferences(html) {
  let result = html;
  for (const asset of VENDOR_ASSETS) {
    result = result.split(asset.url).join(`vendor/${asset.localName}`);
  }
  // The CSP <meta> tag allowlists the CDN origins by hostname (CSP source lists are origins,
  // not full paths, so the string substitution above doesn't touch them) -- those become
  // unnecessary permissions once vendored locally, so strip them from the policy too rather
  // than leaving a stale allowlist entry for a domain this build never actually loads from.
  for (const asset of VENDOR_ASSETS) {
    const origin = new URL(asset.url).origin;
    result = result.split(` ${origin}`).join('');
  }
  // A directive that allowlisted only a CDN origin (e.g. font-src, which existed solely for
  // Fabric's icon fonts) is left with no source values at all after stripping it above --
  // that's invalid CSP syntax, so drop the whole directive rather than emit "font-src;".
  // default-src 'none' already blocks it the same way.
  result = result.replace(/\s*[a-z-]+-src;/g, '');
  return result;
}

function ensureExists(filePath, description) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`${description} not found: ${filePath}`);
  }
}

function buildWithLocalhostPlaceholders() {
  // {{PORT}}/{{SECRET}} are literal tokens -- webpack's manifest transform does a plain
  // string substitution, so they pass through untouched and get filled in for real by
  // setup-local-server.ps1 at install time (see scripts/local-server/).
  const env = { ...process.env, OPENCLERK_HOST_URL: 'https://localhost:{{PORT}}/' };
  const result = spawnSync('npx', ['webpack', '--mode', 'production'], {
    cwd: repoRoot,
    env,
    stdio: 'inherit',
    shell: true,
  });
  if (result.status !== 0) {
    throw new Error('webpack build (offline placeholders) failed');
  }
}

function addSecretToUrls(manifestXml) {
  // Append ?k={{SECRET}} to every URL that points at our own localhost placeholder host,
  // so the local server can authorize requests. See serve-openclerk.ps1 for the check.
  return manifestXml.replace(
    /(https:\/\/localhost:\{\{PORT\}\}\/[^"]*)"/g,
    (match, url) => `${url}?k={{SECRET}}"`
  );
}

function createArchive(vendorDir) {
  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(outputPath);
    const archive = archiver('zip', { zlib: { level: 9 } });

    output.on('close', () => resolve(archive.pointer()));
    archive.on('error', (err) => reject(err));

    archive.pipe(output);

    const manifestXml = addSecretToUrls(fs.readFileSync(manifestPath, 'utf8'));
    archive.append(manifestXml, { name: 'manifest.xml' });

    // The full static build -- this is what the local server serves from disk. taskpane.html
    // and commands.html get their CDN references rewritten to the vendored local copies below.
    for (const entry of fs.readdirSync(distDir, { withFileTypes: true })) {
      if (entry.name === 'taskpane.html' || entry.name === 'commands.html') {
        continue;
      }
      const fullPath = path.join(distDir, entry.name);
      if (entry.isDirectory()) {
        archive.directory(fullPath, `app/${entry.name}`);
      } else {
        archive.file(fullPath, { name: `app/${entry.name}` });
      }
    }
    for (const htmlName of ['taskpane.html', 'commands.html']) {
      const html = fs.readFileSync(path.join(distDir, htmlName), 'utf8');
      archive.append(rewriteCdnReferences(html), { name: `app/${htmlName}` });
    }
    archive.directory(vendorDir, 'app/vendor');

    archive.file(path.join(repoRoot, 'scripts', 'local-server', 'serve-openclerk.ps1'), {
      name: 'installer/serve-openclerk.ps1',
    });
    archive.file(path.join(repoRoot, 'scripts', 'local-server', 'setup-local-server.ps1'), {
      name: 'installer/setup-local-server.ps1',
    });

    archive.finalize();
  });
}

async function main() {
  const vendorDir = path.join(repoRoot, '.offline-vendor-cache');
  try {
    buildWithLocalhostPlaceholders();
    ensureExists(manifestPath, 'Production manifest (dist/manifest.xml)');

    fs.rmSync(vendorDir, { recursive: true, force: true });
    await vendorCdnAssets(vendorDir);

    if (fs.existsSync(outputPath)) {
      fs.unlinkSync(outputPath);
    }

    console.log('Creating offline add-in package:', outputPath);
    const bytes = await createArchive(vendorDir);
    console.log(`Created ${outputPath} (${bytes} bytes)`);
  } catch (err) {
    console.error(err instanceof Error ? err.message : err);
    process.exit(1);
  } finally {
    fs.rmSync(vendorDir, { recursive: true, force: true });
  }
}

main();
