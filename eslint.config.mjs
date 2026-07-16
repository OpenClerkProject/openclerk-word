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
// Each sink is guarded in two AST forms: the direct member call `x.insertHtml(...)`
// (callee.property.name, an Identifier) and the computed form `x["insertHtml"](...)`
// (callee.property.value, a string Literal). Without the computed selector, `x["insertHtml"](raw)`
// slips past the guard while doing exactly what the guard exists to forbid. Residual gap: a call
// routed through an alias (`const f = x.insertHtml.bind(x); f(...)`), `x.insertHtml.apply/call(...)`,
// or a dynamically-built method name is not statically matchable by a syntax selector -- the
// branded SafeHyperlinkUrl parameter on insertSafeHyperlink is the compile-time backstop for the
// url, and code review covers the rest. This lint is defense-in-depth, not a hard sandbox.
const RAW_INSERTION_SINKS = ["insertHtml", "insertHyperlink", "insertComment", "insertOoxml"];

const RAW_INSERTION_SELECTORS = RAW_INSERTION_SINKS.flatMap((name) => {
  const message = `Raw ${name} calls must go through src/taskpane/safeInsertion.ts.`;
  return [
    { selector: `CallExpression[callee.property.name='${name}']`, message },
    { selector: `CallExpression[callee.property.value='${name}']`, message },
  ];
});

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
