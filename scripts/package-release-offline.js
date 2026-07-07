const fs = require('fs');
const path = require('path');
const https = require('https');
const { spawnSync } = require('child_process');
const archiver = require('archiver');

const repoRoot = path.resolve(__dirname, '..');
const distDir = path.join(repoRoot, 'dist');
const manifestPath = path.join(distDir, 'manifest.xml');
const outputPath = path.join(repoRoot, 'wordclerk-addin-offline.zip');

// The taskpane/commands HTML load office.js and Fabric's CSS from Microsoft's CDN, which
// defeats the whole point of an offline package. Vendor local copies at packaging time (this
// runs in CI/dev, which has internet) and rewrite the HTML to reference them instead -- the
// GitHub Pages build is untouched, since CDN loading is the right call when a network is assumed.
const VENDOR_ASSETS = [
  {
    url: 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js',
    localName: 'office.js',
  },
  {
    url:
      'https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css',
    localName: 'fabric.min.css',
  },
];

function fetchText(url, redirectsLeft = 5) {
  return new Promise((resolve, reject) => {
    https
      .get(url, (res) => {
        if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location && redirectsLeft > 0) {
          res.resume();
          resolve(fetchText(res.headers.location, redirectsLeft - 1));
          return;
        }
        if (res.statusCode !== 200) {
          reject(new Error(`Failed to fetch ${url}: HTTP ${res.statusCode}`));
          res.resume();
          return;
        }
        const chunks = [];
        res.on('data', (chunk) => chunks.push(chunk));
        res.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
      })
      .on('error', reject);
  });
}

async function vendorCdnAssets(vendorDir) {
  fs.mkdirSync(vendorDir, { recursive: true });
  for (const asset of VENDOR_ASSETS) {
    console.log(`Vendoring ${asset.url} -> vendor/${asset.localName}`);
    const content = await fetchText(asset.url);
    fs.writeFileSync(path.join(vendorDir, asset.localName), content, 'utf8');
  }
}

function rewriteCdnReferences(html) {
  let result = html;
  for (const asset of VENDOR_ASSETS) {
    result = result.split(asset.url).join(`vendor/${asset.localName}`);
  }
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
  const env = { ...process.env, WORDCLERK_HOST_URL: 'https://localhost:{{PORT}}/' };
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
  // so the local server can authorize requests. See serve-wordclerk.ps1 for the check.
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

    archive.file(path.join(repoRoot, 'scripts', 'local-server', 'serve-wordclerk.ps1'), {
      name: 'installer/serve-wordclerk.ps1',
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
