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
