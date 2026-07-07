# WordClerk

[![CI](https://github.com/wbarnha/WordClerk/actions/workflows/ci.yml/badge.svg)](https://github.com/wbarnha/WordClerk/actions/workflows/ci.yml)

WordClerk is a Word add-in (task pane) for applying and removing hyperlinks to case-law and parenthetical citations.

## Development

Prerequisites
- Node.js (16+ recommended)
- npm

Install dependencies:

```bash
cd "c:\Users\willi\WordClerk"
npm install
```

Generate PNG icons from the SVG (optional but recommended for manifest icons):

```bash
npm run convert-logos
```

Start the dev server and sideload the add-in into Word (desktop):

```bash
npm run start
```

This runs the webpack dev server and uses `office-addin-debugging` to sideload the manifest into Word for debugging.

## Scripts
- `npm run build:dev` — build development bundle
- `npm run build` — production build
- `npm run convert-logos` — convert `dist/assets/logo-filled.svg` into PNG variants (16/32/80px)
- `npm run start` — start dev server and sideload into Word
- `npm run recover:start` — clear stale Office dev settings, then stop and restart sideload debugging
- `npm run install:addin` — install the manifest into the default Office WEF manifest folder
- `npm run install:addin:dry-run` — validate installer behavior without writing files

## Notes
- If icons in the manifest are SVG and Word rejects the manifest, convert or reference PNGs instead.
- If `insertHyperlink` is unavailable in your Word environment, the add-in falls back to using `insertHtml` or plain text.

## Download and install from GitHub
You can install the add-in into desktop Word from this repository by using the GitHub release package, workflow artifacts, or cloning the repo locally.

### Option 1: Install from GitHub Release asset
1. Go to the GitHub Releases page for this repo.
2. Download the latest `wordclerk-addin.zip` release asset.
3. Extract the ZIP. It contains `manifest.xml`, `dist`, and `assets`.
4. In a terminal, open the extracted folder and run:

```bash
npm install
npm run start
```

5. `npm run start` launches the local dev server and sideloads the add-in into Word Desktop.

### Option 2: Install from GitHub Actions workflow artifact
1. Open the Actions tab in this repo.
2. Select the latest successful workflow run for `CI`.
3. Scroll to the `Artifacts` section and download `wordclerk-addin`.
4. Extract the downloaded artifact.
5. Open the extracted folder and run:

```bash
npm install
npm run start
```

This gives you a packaged `wordclerk-addin.zip` plus the source contents.

### Option 3: Create a local install package
If you want a package that can be uploaded to Word manually, use the package script:

```bash
npm install
npm run package
```

This creates `wordclerk-addin.zip` at the repo root, which contains:
- `manifest.xml`
- `dist`
- `assets`
- `installer/install-wordclerk.js`
- `installer/install-wordclerk.ps1`

You can then upload the `manifest.xml` file to Word from `Insert` → `My Add-ins` → `Upload My Add-in` → `Add from file`.

### Option 4: Create a local test environment
For local QA, you can use the PowerShell helper script:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup-local-test.ps1
```

That script installs dependencies, builds the add-in, and packages `wordclerk-addin.zip`.
The generated package also includes installer scripts that can copy the manifest into Word's local WEF folder.

After the script completes, run:

```bash
npm run start
```

Then open Word and use the add-in from the sideloaded manifest.

### Option 5: Clone the repository and install locally
1. Clone the repo:

```bash
git clone https://github.com/wbarnha/WordClerk.git
cd WordClerk
```

2. Install dependencies:

```bash
npm install
```

3. Start the add-in locally and sideload into Word:

```bash
npm run start
```

This will launch the dev server on `https://localhost:3000` and load the add-in into Word Desktop using the manifest.

### Manual sideload via manifest
If you want to sideload the add-in manually in Word Desktop:
1. Run `npm install` and `npm run build:dev`.
2. Host the `dist` and `assets` folders on a local HTTPS server matching the URLs in `manifest.xml`.
3. In Word, go to `Insert` → `My Add-ins` → `Upload My Add-in` → `Add from file`.
4. Select the extracted or cloned `manifest.xml` file.

> Note: The manifest currently points to `https://localhost:3000/` for the add-in content, so the easiest install path is using `npm run start`.

