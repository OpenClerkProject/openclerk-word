const fs = require('fs');
const path = require('path');

function parseArgs(argv) {
  const args = {
    manifest: 'manifest.xml',
    target: process.platform === 'win32'
      ? path.join(process.env.LOCALAPPDATA || '', 'Microsoft', 'Office', '16.0', 'WEF')
      : path.join(process.cwd(), '.installer-test', 'wef'),
    dryRun: false,
  };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === '--manifest' && argv[i + 1]) {
      args.manifest = argv[i + 1];
      i += 1;
    } else if (arg === '--target' && argv[i + 1]) {
      args.target = argv[i + 1];
      i += 1;
    } else if (arg === '--dry-run') {
      args.dryRun = true;
    }
  }

  return args;
}

function installManifest(manifestPath, targetDir, dryRun) {
  if (!fs.existsSync(manifestPath)) {
    throw new Error(`Manifest not found: ${manifestPath}`);
  }

  const resolvedManifest = path.resolve(manifestPath);
  const resolvedTargetDir = path.resolve(targetDir);
  const targetManifest = path.join(resolvedTargetDir, 'openclerk-manifest.xml');

  console.log(`Source manifest: ${resolvedManifest}`);
  console.log(`Install target: ${targetManifest}`);

  if (dryRun) {
    console.log('Dry run enabled. No files were copied.');
    return targetManifest;
  }

  fs.mkdirSync(resolvedTargetDir, { recursive: true });
  fs.copyFileSync(resolvedManifest, targetManifest);

  console.log('Manifest installed successfully.');
  return targetManifest;
}

function main() {
  try {
    const args = parseArgs(process.argv.slice(2));
    installManifest(args.manifest, args.target, args.dryRun);
  } catch (err) {
    console.error(err instanceof Error ? err.message : String(err));
    process.exit(1);
  }
}

main();
