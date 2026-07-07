import { mkdtempSync, existsSync } from 'fs';
import { join } from 'path';
import { tmpdir } from 'os';
import { spawnSync } from 'child_process';

describe('Installer environment', () => {
  test('installer script copies manifest into target folder', () => {
    const tempRoot = mkdtempSync(join(tmpdir(), 'wordclerk-installer-'));
    const target = join(tempRoot, 'wef');

    const result = spawnSync(
      process.execPath,
      ['scripts/install-wordclerk.js', '--manifest', 'manifest.xml', '--target', target],
      { encoding: 'utf-8' }
    );

    expect(result.status).toBe(0);
    expect(existsSync(join(target, 'wordclerk-manifest.xml'))).toBe(true);
  });

  test('installer supports dry-run mode', () => {
    const tempRoot = mkdtempSync(join(tmpdir(), 'wordclerk-installer-dryrun-'));
    const target = join(tempRoot, 'wef');

    const result = spawnSync(
      process.execPath,
      ['scripts/install-wordclerk.js', '--manifest', 'manifest.xml', '--target', target, '--dry-run'],
      { encoding: 'utf-8' }
    );

    expect(result.status).toBe(0);
    expect(existsSync(join(target, 'wordclerk-manifest.xml'))).toBe(false);
    expect(result.stdout).toContain('Dry run enabled');
  });
});
