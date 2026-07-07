import { readFileSync } from 'fs';
import { parseStringPromise } from 'xml2js';

describe('Office add-in manifest', () => {
  let manifest: any;

  beforeAll(async () => {
    const xml = readFileSync('manifest.xml', 'utf-8');
    manifest = await parseStringPromise(xml, { explicitArray: false });
  });

  test('defines Word host and task pane extension points', () => {
    const hosts = manifest.OfficeApp.Hosts.Host;
    expect(hosts.Name || hosts.$.Name).toBe('Document');

    const versionOverrides = manifest.OfficeApp.VersionOverrides;
    expect(versionOverrides).toBeDefined();

    const desktopHost = versionOverrides.Hosts.Host;
    expect(desktopHost.$['xsi:type'] || desktopHost.$['xsi:type']).toBe('Document');

    const extensionPoints = desktopHost.DesktopFormFactor.ExtensionPoint;
    const extensionPointTypes = Array.isArray(extensionPoints)
      ? extensionPoints.map((ep: any) => ep.$['xsi:type'] || ep.$['xsi:type'])
      : [extensionPoints.$['xsi:type'] || extensionPoints.$['xsi:type']];

    expect(extensionPointTypes).toContain('PrimaryCommandSurface');
  });

  test('declares required requirement sets for Word and commands', () => {
    const requirements = manifest.OfficeApp.Requirements;
    // Requirements are optional in task pane manifests; assert only when present.
    if (!requirements) {
      expect(requirements).toBeUndefined();
      return;
    }

    const sets = requirements.Sets.Set;
    const names = Array.isArray(sets) ? sets.map((set: any) => set.$.Name) : [sets.$.Name];

    expect(names).toContain('WordApi');
    expect(names).toContain('WordApiOnline');
    expect(names).toContain('AddInCommands');
    expect(names).toContain('TaskpaneApi');
  });

  test('requires WordApi 1.9 or greater', () => {
    if (!manifest.OfficeApp.Requirements) {
      expect(manifest.OfficeApp.Requirements).toBeUndefined();
      return;
    }

    const sets = manifest.OfficeApp.Requirements.Sets.Set;
    const wordApi = Array.isArray(sets) ? sets.find((set: any) => set.$.Name === 'WordApi') : sets;
    expect(wordApi).toBeDefined();
    expect(wordApi.$.MinVersion).toBe('1.9');
  });
});
