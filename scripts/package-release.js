const fs = require('fs');
const path = require('path');
const archiver = require('archiver');

const repoRoot = path.resolve(__dirname, '..');
const outputPath = path.join(repoRoot, 'openclerk-addin.zip');
// Use the manifest from the production build (URLs already replaced by webpack)
const manifestPath = path.join(repoRoot, 'dist', 'manifest.xml');

function ensureExists(filePath, description) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`${description} not found: ${filePath}`);
  }
}

function createArchive() {
  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(outputPath);
    const archive = archiver('zip', { zlib: { level: 9 } });

    output.on('close', () => resolve(archive.pointer()));
    archive.on('error', (err) => reject(err));

    archive.pipe(output);
    // manifest.xml at the package root, sourced from the production build. Nothing else needs to
    // ship here: the manifest's URLs point at GitHub Pages, so Word fetches taskpane/commands/icons
    // live over HTTPS instead of reading local files -- dist/ and assets/ would just be unused dead
    // weight in the end-user package (see the "Upload dist artifact for GitHub Pages" step in
    // ci.yml for the copy that actually gets deployed).
    archive.file(manifestPath, { name: 'manifest.xml' });
    // Standalone installers only — no Node.js required by end users
    archive.file(path.join(repoRoot, 'scripts', 'install-openclerk.ps1'), {
      name: 'installer/install-openclerk.ps1',
    });
    // Double-clickable wrapper around the .ps1 -- Windows opens .ps1 files in an editor by
    // default rather than running them, which makes the installer look like it does nothing.
    archive.file(path.join(repoRoot, 'scripts', 'install-openclerk.cmd'), {
      name: 'installer/install-openclerk.cmd',
    });
    // mode: 0o755 preserves the executable bit inside the zip so macOS/Linux unzip tools
    // don't require a manual `chmod +x` before the script can be run.
    archive.file(path.join(repoRoot, 'scripts', 'install-openclerk.sh'), {
      name: 'installer/install-openclerk.sh',
      mode: 0o755,
    });
    archive.finalize();
  });
}

async function main() {
  try {
    ensureExists(manifestPath, 'Production manifest (dist/manifest.xml) -- run `npm run build` first');

    if (fs.existsSync(outputPath)) {
      fs.unlinkSync(outputPath);
    }

    console.log('Creating add-in package:', outputPath);
    const bytes = await createArchive();
    console.log(`Created ${outputPath} (${bytes} bytes)`);
  } catch (err) {
    console.error(err instanceof Error ? err.message : err);
    process.exit(1);
  }
}

main();
