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

  test('gates on the WordApi requirement set', () => {
    // The manifest declares a top-level Requirements block so the add-in only loads on hosts that
    // can run every advertised feature. It intentionally requires ONLY WordApi -- not AddInCommands
    // (gating on it would block hosts that support task panes but not ribbon commands, which degrade
    // gracefully) and not WordApiOnline/desktop-only sets (the code deliberately avoids those for
    // cross-platform reach; see the getComments-over-document.comments note in word.ts).
    const requirements = manifest.OfficeApp.Requirements;
    expect(requirements).toBeDefined();

    const sets = requirements.Sets.Set;
    const names = Array.isArray(sets) ? sets.map((set: any) => set.$.Name) : [sets.$.Name];

    expect(names).toContain('WordApi');
  });

  test('requires WordApi 1.4 -- the floor set by the comments API (Embed Cited Text)', () => {
    // 1.4 is the highest requirement-set version any used API needs: Range.insertComment /
    // Body.getComments are WordApi 1.4; everything else (search, insertHtml/insertText/insertOoxml,
    // Range.hyperlinks) is <= 1.4. Raising this would drop perpetual/volume-licensed Word (2019/
    // 2021/LTSC) for no functional gain, since no 1.5+ API is used.
    const requirements = manifest.OfficeApp.Requirements;
    expect(requirements).toBeDefined();

    const sets = requirements.Sets.Set;
    const wordApi = Array.isArray(sets) ? sets.find((set: any) => set.$.Name === 'WordApi') : sets;
    expect(wordApi).toBeDefined();
    expect(wordApi.$.MinVersion).toBe('1.4');
  });
});
