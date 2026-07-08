/**
 * Table T6 (case name / institutional-author word abbreviations).
 *
 * Before the 21st edition, case names used Table T6 while periodical/
 * institutional-author names used the separate Table T13.2, and several
 * words were only abbreviated under T13.2 -- so they were spelled out in
 * full in case names. The 21st edition merged T6 and T13.2 into one table
 * applied everywhere, which is what makes this list edition-dependent: a
 * pre-21st-edition case name spelling out one of these words was correct;
 * from the 21st edition on, Bluebook style abbreviates it.
 *
 * Sources: University of Washington Law Library, "Major Changes in the 21st
 * Edition" (https://lib.law.uw.edu/bluebook101/editions); Mary Whisner,
 * "Bluebook Weight Loss Program - Part Two: The Merger of Tables T6 and
 * T13.2" (https://citeblog.access-to-law.com/?p=1074), which specifically
 * confirms "Laboratory" -> "Lab'y" and "Employ(ee/ment)" -> "Emp." moved
 * from institution-only to case-name use in the merge. The remaining
 * entries below (Env't, Rsch., Psych., Socio., Compar.) are the same kind
 * of T6/T13.2-merger abbreviation and are included on that basis, but this
 * project has not independently verified each one's exact pre-merger scope
 * word-for-word against the official table -- treat this list as a
 * reasonable approximation, not a verbatim reproduction of Table T6.
 */
export interface CaseNameAbbreviationEntry {
  /** Full word as it would appear unabbreviated in running text. */
  word: string;
  /** The Table T6 abbreviation Bluebook style expects instead. */
  abbreviation: string;
}

export const T6_T13_MERGER_ABBREVIATIONS: CaseNameAbbreviationEntry[] = [
  { word: "Laboratory", abbreviation: "Lab'y" },
  { word: "Employment", abbreviation: "Emp." },
  { word: "Employee", abbreviation: "Emp." },
  { word: "Environment", abbreviation: "Env't" },
  { word: "Research", abbreviation: "Rsch." },
  { word: "Psychology", abbreviation: "Psych." },
  { word: "Psychological", abbreviation: "Psych." },
  { word: "Sociology", abbreviation: "Socio." },
  { word: "Sociological", abbreviation: "Socio." },
  { word: "Comparative", abbreviation: "Compar." },
];
