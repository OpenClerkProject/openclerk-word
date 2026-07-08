# Contributing to WordClerk

This project welcomes corrections and improvements from anyone -- including people who don't
write code. This guide has a section specifically for that: **fixing or adding a Bluebook
citation rule**. If you're a developer looking to contribute code more generally, the
[README](README.md#development) covers local setup.

## Contributing a Bluebook citation correction (no coding experience needed)

The **Bluebook Check** feature flags citation-formatting problems -- things like a reporter
abbreviation that isn't quite right, or a word in a case name that should be abbreviated. Most of
that data comes from [reporters-db](https://github.com/freelawproject/reporters-db), a large,
well-maintained legal-citation database, but it isn't perfect: it can be missing an abbreviation,
or WordClerk's rules on top of it can get something wrong. If you notice one of these while using
the add-in, here are two ways to report or fix it. Neither one requires installing anything,
knowing what Git is, or writing code.

### Option A: Fill out a form (easiest, recommended for most people)

1. Go to [github.com/wbarnha/WordClerk/issues/new/choose](https://github.com/wbarnha/WordClerk/issues/new/choose).
2. Pick **"Bluebook citation correction."**
3. Fill in the form: what citation you were checking, what WordClerk said versus what it should
   have said, and a source backing up the correction (a page of the Bluebook, a law library
   citation guide, or similar -- this is a legal tool, so we try to cite our sources rather than
   guess).
4. Submit it. A maintainer will review it and make the fix -- you're done.

That's genuinely the whole process. You never have to open a file, install anything, or touch
code.

### Option B: Edit the file directly on GitHub.com (a little more hands-on, still no install)

If you're comfortable with the idea of "editing a small text file in a web browser," you can
propose the fix yourself, and GitHub will create the pull request for you automatically:

1. Open [`src/taskpane/bluebook/manualCorrections.ts`](src/taskpane/bluebook/manualCorrections.ts)
   on GitHub. This is the **one file** set aside specifically for community corrections -- it's
   small, in plain English, and separate from the large auto-generated data files (which get
   overwritten whenever that data is refreshed from reporters-db, so edits there wouldn't stick).
2. Click the pencil icon (✏️) near the top-right of the file to start editing, in your browser --
   no download or install needed.
3. The file has three lists, one per kind of correction, each with a commented-out example showing
   the exact format:
   - **`MANUAL_REPORTER_CORRECTIONS`** -- "this reporter abbreviation is wrong, here's the correct
     one." Example: `F.Supp.2d` should be `F. Supp. 2d`.
   - **`MANUAL_VALID_REPORTER_FORMS`** -- "this reporter abbreviation is actually correct, stop
     flagging it as unrecognized."
   - **`MANUAL_CASE_NAME_ABBREVIATIONS`** -- "this word in a case name should be abbreviated this
     way."
   Add a new line to the relevant list, following the example format exactly (matching commas and
   curly braces matters -- copy the example, then change the values). Always fill in the `source`
   field with a link or citation backing up the correction.
4. Scroll down, add a short description of what you changed, and choose **"Create a new branch
   for this commit and start a pull request."**
5. Click **"Propose changes."** GitHub opens a pull request for you. A maintainer will review it,
   and our automated checks will confirm the file still loads correctly before it's merged.

If step 3 feels intimidating, that's completely fine -- use Option A instead and let a maintainer
make the edit.

### What makes a good correction report

- **A concrete example.** The exact citation text you were looking at, not just a description.
- **A source.** Bluebook citation rules are precise and sometimes counter-intuitive; a link to an
  authoritative source (the Bluebook itself, a law school citation guide, reporters-db) lets a
  maintainer verify the correction quickly instead of taking it on faith.
- **One correction per report.** Easier to review and merge than a batch of unrelated fixes.

### Why corrections live in a separate file

`src/taskpane/bluebook/manualCorrections.ts` exists specifically so community contributions are
safe from being silently overwritten. The bulk of the citation data
([`src/taskpane/bluebook/generated/`](src/taskpane/bluebook/generated/)) is regenerated
periodically straight from reporters-db (see [`npm run
bluebook:update-data`](README.md#vendored-reference-data)) -- any hand-edits there would be lost
the next time that happens. `manualCorrections.ts` is never touched by that process, so it's the
one place additions and corrections are guaranteed to stick around.

## Contributing code

For code contributions (bug fixes, new features, new citation-lookup providers, etc.), see
[README.md > Development](README.md#development) for how to set up a local environment, and open
a pull request as usual. `npm test` and `npm run build` should both pass before you submit.
