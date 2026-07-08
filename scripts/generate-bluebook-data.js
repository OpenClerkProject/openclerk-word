/*
 * Regenerates src/taskpane/bluebook/generated/*.ts from Free Law Project's
 * reporters-db (https://github.com/freelawproject/reporters-db, BSD-2-Clause),
 * the same reporter/case-name/state abbreviation database used by CourtListener's
 * own citation parser (eyecite). This is a manually-run, dev-time-only script --
 * it is NOT part of `npm run build` and the add-in makes no network calls to
 * fetch this data at runtime. The generated files are committed to the repo like
 * any other source file; re-run this script and commit the diff to pick up
 * upstream updates.
 *
 * Usage: node scripts/generate-bluebook-data.js
 */
const fs = require("fs");
const path = require("path");
const https = require("https");

const outDir = path.resolve(__dirname, "..", "src", "taskpane", "bluebook", "generated");

const SOURCES = {
  reporters: "https://raw.githubusercontent.com/freelawproject/reporters-db/main/reporters_db/data/reporters.json",
  caseNameAbbreviations:
    "https://raw.githubusercontent.com/freelawproject/reporters-db/main/reporters_db/data/case_name_abbreviations.json",
  stateAbbreviations:
    "https://raw.githubusercontent.com/freelawproject/reporters-db/main/reporters_db/data/state_abbreviations.json",
};

// The 21st-edition Table T6/T13.2 merger words (see ../caseNameAbbreviations.ts) are handled
// separately as an edition-gated overlay -- excluded here so the edition-independent table below
// doesn't incorrectly flag them for the 20th edition, where spelling them out was correct.
const EDITION_GATED_WORDS = new Set([
  "laboratory",
  "employment",
  "employee",
  "environment",
  "research",
  "psychology",
  "psychological",
  "sociology",
  "sociological",
  "comparative",
]);

function fetchJson(url) {
  return new Promise((resolve, reject) => {
    https
      .get(url, (res) => {
        if (res.statusCode !== 200) {
          reject(new Error(`Failed to fetch ${url}: HTTP ${res.statusCode}`));
          res.resume();
          return;
        }
        const chunks = [];
        res.on("data", (chunk) => chunks.push(chunk));
        res.on("end", () => {
          try {
            resolve(JSON.parse(Buffer.concat(chunks).toString("utf8")));
          } catch (err) {
            reject(err);
          }
        });
      })
      .on("error", reject);
  });
}

// Trims reporters.json down to what citation-format checking actually needs. Two distinct
// categories matter here, and conflating them causes false positives: a reporter's *edition*
// forms (e.g. "A.", "A.2d", "A.3d" are all independently valid -- different chronological
// editions of the same series, not "errors" to be corrected into each other) versus its known
// *variations* (malformed/non-standard spellings like "A2d" or "Atl.2d" that should be
// corrected to a specific valid form). Doesn't keep jurisdiction/date-range/publisher metadata
// -- this isn't a citator, just an "is this reporter abbreviation valid, and if not, what's the
// closest correct form" checker.
function buildReporterLookup(reportersJson) {
  const validForms = {};
  const corrections = {};
  for (const canonicalKey of Object.keys(reportersJson)) {
    for (const entry of reportersJson[canonicalKey]) {
      const name = entry.name || canonicalKey;
      const editionForms = new Set([canonicalKey, ...Object.keys(entry.editions || {})]);
      for (const form of editionForms) {
        validForms[form] = name;
      }
      for (const [variant, correctForm] of Object.entries(entry.variations || {})) {
        if (!(variant in validForms) && !(variant in corrections)) {
          corrections[variant] = { correctForm, name };
        }
      }
    }
  }
  // A string that's a known variation for one reporter series might legitimately be a valid
  // edition form of a *different* series -- valid-anywhere always wins over flagged-elsewhere.
  for (const form of Object.keys(validForms)) {
    delete corrections[form];
  }
  return { validForms, corrections };
}

// Inverts { abbreviation: [fullWord, ...] } into { fullWordLowercase: abbreviation }, since
// checking a case name means asking "does this full word appear that should be abbreviated?".
// Skips edition-gated words (handled separately) and any full word that maps to more than one
// distinct abbreviation upstream (ambiguous -- safer to not flag those automatically).
function buildCaseNameAbbreviationLookup(caseNameJson) {
  const wordToAbbreviation = {};
  const ambiguous = new Set();
  for (const [abbreviation, words] of Object.entries(caseNameJson)) {
    for (const word of words) {
      const key = word.toLowerCase();
      if (EDITION_GATED_WORDS.has(key)) {
        continue;
      }
      if (key in wordToAbbreviation && wordToAbbreviation[key] !== abbreviation) {
        ambiguous.add(key);
        continue;
      }
      wordToAbbreviation[key] = abbreviation;
    }
  }
  for (const key of ambiguous) {
    delete wordToAbbreviation[key];
  }
  return wordToAbbreviation;
}

function writeGeneratedFile(filename, exportName, data, sourceUrl, describeCount) {
  const header =
    `// GENERATED FILE -- do not edit by hand.\n` +
    `// Source: ${sourceUrl}\n` +
    `// Regenerate with: node scripts/generate-bluebook-data.js\n\n`;
  const body = `export const ${exportName} = ${JSON.stringify(data, null, 2)} as const;\n`;
  fs.writeFileSync(path.join(outDir, filename), header + body, "utf8");
  console.log(`Wrote ${filename} (${describeCount(data)})`);
}

async function main() {
  fs.mkdirSync(outDir, { recursive: true });

  console.log("Fetching reporters-db data...");
  const [reportersJson, caseNameJson, stateJson] = await Promise.all([
    fetchJson(SOURCES.reporters),
    fetchJson(SOURCES.caseNameAbbreviations),
    fetchJson(SOURCES.stateAbbreviations),
  ]);

  writeGeneratedFile(
    "reporterAbbreviations.generated.ts",
    "REPORTER_ABBREVIATIONS",
    buildReporterLookup(reportersJson),
    SOURCES.reporters,
    (data) =>
      `${Object.keys(data.validForms).length} valid forms, ${Object.keys(data.corrections).length} known corrections`
  );
  writeGeneratedFile(
    "caseNameAbbreviations.generated.ts",
    "CASE_NAME_ABBREVIATIONS",
    buildCaseNameAbbreviationLookup(caseNameJson),
    SOURCES.caseNameAbbreviations,
    (data) => `${Object.keys(data).length} entries`
  );
  writeGeneratedFile(
    "stateAbbreviations.generated.ts",
    "STATE_ABBREVIATIONS",
    stateJson,
    SOURCES.stateAbbreviations,
    (data) => `${Object.keys(data).length} entries`
  );
}

main().catch((err) => {
  console.error(err instanceof Error ? err.message : err);
  process.exit(1);
});
