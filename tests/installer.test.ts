import { mkdtempSync, existsSync } from 'fs';
import { join } from 'path';
import { tmpdir } from 'os';
import { spawnSync } from 'child_process';

describe('Installer environment', () => {
  test('installer script copies manifest into target folder', () => {
    const tempRoot = mkdtempSync(join(tmpdir(), 'openclerk-installer-'));
    const target = join(tempRoot, 'wef');

    const result = spawnSync(
      process.execPath,
      ['scripts/install-openclerk.js', '--manifest', 'manifest.xml', '--target', target],
      { encoding: 'utf-8' }
    );

    expect(result.status).toBe(0);
    expect(existsSync(join(target, 'openclerk-manifest.xml'))).toBe(true);
  });

  test('installer supports dry-run mode', () => {
    const tempRoot = mkdtempSync(join(tmpdir(), 'openclerk-installer-dryrun-'));
    const target = join(tempRoot, 'wef');

    const result = spawnSync(
      process.execPath,
      ['scripts/install-openclerk.js', '--manifest', 'manifest.xml', '--target', target, '--dry-run'],
      { encoding: 'utf-8' }
    );

    expect(result.status).toBe(0);
    expect(existsSync(join(target, 'openclerk-manifest.xml'))).toBe(false);
    expect(result.stdout).toContain('Dry run enabled');
  });
});

// The macOS installer is a plain bash script, so it also runs (and is exercised here) on Linux,
// where `bash` is present but the platform-specific default target path doesn't apply -- these
// tests always pass an explicit --target, matching how the CI smoke tests exercise it too.
// Skipped only if bash itself is unavailable (e.g. a contributor running tests on plain Windows).
const describeIfBash = process.platform === 'win32' ? describe.skip : describe;

describeIfBash('macOS installer script (install-openclerk.sh)', () => {
  test('installer script copies manifest into target folder', () => {
    const tempRoot = mkdtempSync(join(tmpdir(), 'openclerk-installer-sh-'));
    const target = join(tempRoot, 'wef');

    const result = spawnSync(
      'bash',
      ['scripts/install-openclerk.sh', '--manifest', 'manifest.xml', '--target', target],
      { encoding: 'utf-8' }
    );

    expect(result.status).toBe(0);
    expect(existsSync(join(target, 'openclerk-manifest.xml'))).toBe(true);
  });

  test('installer supports dry-run mode', () => {
    const tempRoot = mkdtempSync(join(tmpdir(), 'openclerk-installer-sh-dryrun-'));
    const target = join(tempRoot, 'wef');

    const result = spawnSync(
      'bash',
      ['scripts/install-openclerk.sh', '--manifest', 'manifest.xml', '--target', target, '--dry-run'],
      { encoding: 'utf-8' }
    );

    expect(result.status).toBe(0);
    expect(existsSync(join(target, 'openclerk-manifest.xml'))).toBe(false);
    expect(result.stdout).toContain('Dry run enabled');
  });
});
