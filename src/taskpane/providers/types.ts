/**
 * Core contract for the citation hyperlink plugin architecture.
 *
 * Any online case-law lookup source (a free public API, or a paid research
 * platform like LexisNexis/Westlaw/Bloomberg Law) implements CitationProvider
 * and registers itself with the registry in registry.ts. Nothing else in the
 * add-in needs to know which provider is active.
 */

export interface ParsedCitation {
  /** The raw citation text as it appeared in the document. */
  raw: string;
  caseName?: string;
  volume?: string;
  reporter?: string;
  page?: string;
  /** Pinpoint page/range, e.g. "496" or "705-06" in "..., 490, 496 (1980)". */
  pincite?: string;
  court?: string;
  year?: string;
}

export interface CitationMatch {
  /** Direct URL to the case on the provider's site. */
  url: string;
  caseName?: string;
  /** Normalized citation string returned by the provider, if any. */
  citation?: string;
}

export interface ProviderCredentialField {
  key: string;
  label: string;
  type: "text" | "password";
  placeholder?: string;
  /** Defaults to true. Set false for genuinely optional fields (e.g. an API token that only raises rate limits). */
  required?: boolean;
}

export interface CitationProvider {
  readonly id: string;
  readonly name: string;
  readonly description: string;
  /** Whether this provider needs authenticate() before lookupCitation() can succeed. */
  readonly requiresAuth: boolean;
  readonly credentialFields: ProviderCredentialField[];

  isAuthenticated(): boolean;

  /** Validates and stores credentials in memory only (never persisted to disk). */
  authenticate(credentials: Record<string, string>): Promise<void>;

  /** Clears any in-memory credentials/tokens. */
  signOut(): void;

  /**
   * Looks up a single citation. Must resolve to null (never throw) when the
   * citation isn't found, the provider isn't authenticated, or the request
   * fails -- callers treat a null result as "move on to the next citation".
   */
  lookupCitation(citation: ParsedCitation): Promise<CitationMatch | null>;
}

export interface OpinionExcerptResult {
  /** The extracted text, or null if unavailable for any reason (see rateLimited for one specific reason). */
  excerpt: string | null;
  /**
   * True when `excerpt` is null specifically because the provider's API rate-limited the
   * request (HTTP 429) rather than because the citation/page genuinely couldn't be found --
   * callers should tell the user to wait and retry, not report this the same as "not found".
   */
  rateLimited?: boolean;
}

/**
 * Optional capability for providers that can supply a cited case's full opinion text, not just
 * a hyperlink -- used by the "Embed Cited Text" workflow to pull the text at a specific pincite
 * page (or pages) and attach it to the citation in the document. Not every CitationProvider can
 * do this (LexisNexis/Westlaw/Bloomberg Law are hyperlink-only integrations today), so this is
 * kept as a separate, optional interface rather than added to CitationProvider itself.
 */
export interface OpinionTextCapableProvider extends CitationProvider {
  /**
   * Fetches the opinion text covering the given (already-expanded, see providers/pincitePages.ts)
   * pincite page numbers for a citation. Must never throw -- an unavailable excerpt (citation not
   * found, provider not ready, request failure, or no page markers matching any requested page)
   * is reported via `excerpt: null` in the result, not an exception.
   */
  fetchOpinionExcerpt(citation: ParsedCitation, targetPages: number[]): Promise<OpinionExcerptResult>;
  /** Whether this provider currently has what it needs (e.g. an API token) to fetch opinion text. */
  isReadyForOpinionText(): boolean;
}

export function supportsOpinionText(provider: CitationProvider): provider is OpinionTextCapableProvider {
  const candidate = provider as Partial<OpinionTextCapableProvider>;
  return typeof candidate.fetchOpinionExcerpt === "function" && typeof candidate.isReadyForOpinionText === "function";
}

/**
 * Optional capability for providers whose API rate limit is tight enough that hitting it in
 * normal use is expected, not a rare edge case (e.g. CourtListener's default of 5 requests/
 * minute). Lets callers distinguish "rate-limited, try again later" from "genuinely not found"
 * -- conflating the two is misleading (e.g. Find Hallucinations would otherwise report a real,
 * valid citation as a possible hallucination just because the request got throttled).
 */
export interface RateLimitAwareProvider extends CitationProvider {
  /** Whether the most recent lookupCitation() call returned null specifically because of an HTTP 429, not a genuine miss. */
  wasLastRequestRateLimited(): boolean;
}

export function supportsRateLimitAwareness(provider: CitationProvider): provider is RateLimitAwareProvider {
  return typeof (provider as Partial<RateLimitAwareProvider>).wasLastRequestRateLimited === "function";
}
