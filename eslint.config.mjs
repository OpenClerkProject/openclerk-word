// Source: eslint.org/docs/latest/rules/no-restricted-syntax (selector syntax); exemption-via-
// second-config-object pattern confirmed against eslint/eslint#19239 (closed, "works as intended"
// -- a later flat-config entry fully replaces an earlier no-restricted-syntax value for files it
// matches, which is exactly the desired "off" override for safeInsertion.ts).
//
// This file is REQUIRED for ESCAPE-03's bypass-guard to have any effect at all: office-addin-lint's
// `npm run lint` invocation resolves this exact filename (project-root eslint.config.mjs) via
// fs.existsSync and passes it to ESLint via an explicit `-c` flag, bypassing ESLint's automatic
// config discovery entirely -- .eslintrc.json (the legacy config previously in this repo) is never
// read by `npm run lint`, regardless of ESLint version (see node_modules/office-addin-lint/lib/lint.js).
import officeAddins from "eslint-plugin-office-addins";
import tsParser from "@typescript-eslint/parser";

// eslint-plugin-office-addins's recommended flat config declares no browser globals at all --
// word.ts/taskpane.ts rely on inline `/* global document, Office, Word */` comments for the
// handful of globals they reference, but a few DOM constructor globals used across the taskpane
// (source-document/.docx parsing, form-field handling) are not covered by those comments. This is
// a pre-existing gap (present even under office-addin-lint's own bundled default config, verified
// during Plan 04 execution -- .eslintrc.json never actually enforced this project's lint rules per
// RESEARCH.md Pitfall 1, so no-undef on these globals was never previously surfaced). Declared here
// at the config level (rather than sprinkled across per-file `/* global */` comments) so any future
// file automatically inherits the same browser-global set.
const BROWSER_GLOBALS = {
  DOMParser: "readonly",
  Element: "readonly",
  Event: "readonly",
  File: "readonly",
  HTMLButtonElement: "readonly",
  HTMLInputElement: "readonly",
  HTMLSelectElement: "readonly",
  URLSearchParams: "readonly",
};

// All three selectors intentionally live in ONE rule-array entry, not split across separate config
// objects -- flat config's no-restricted-syntax does not merge across matching entries (eslint/eslint#19239).
const RAW_INSERTION_SELECTORS = [
  {
    selector: "CallExpression[callee.property.name='insertHtml']",
    message: "Raw insertHtml calls must go through src/taskpane/safeInsertion.ts.",
  },
  {
    selector: "CallExpression[callee.property.name='insertHyperlink']",
    message: "Raw insertHyperlink calls must go through src/taskpane/safeInsertion.ts.",
  },
  {
    selector: "CallExpression[callee.property.name='insertComment']",
    message: "Raw insertComment calls must go through src/taskpane/safeInsertion.ts.",
  },
];

export default [
  ...officeAddins.configs.recommended,
  {
    plugins: { "office-addins": officeAddins },
    languageOptions: { parser: tsParser, globals: BROWSER_GLOBALS },
    rules: {
      "no-restricted-syntax": ["error", ...RAW_INSERTION_SELECTORS],
    },
  },
  {
    // Path-anchored (not basename-only) glob -- matches only src/taskpane/safeInsertion.ts, the
    // sole file allowed to call the raw Office.js insertion APIs. A "**/safeInsertion.ts" glob
    // would match any same-named file at any depth under src/, silently disabling the guard
    // there too if a future contributor ever added one.
    files: ["src/taskpane/safeInsertion.ts"],
    rules: {
      "no-restricted-syntax": "off",
    },
  },
];
