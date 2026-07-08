import { CitationMatch, ParsedCitation, ProviderCredentialField, RateLimitAwareProvider } from "./types";

const API_BASE = "https://www.courtlistener.com/api/rest/v4";
const SITE_ORIGIN = "https://www.courtlistener.com";

interface CourtListenerCluster {
  case_name?: string;
  absolute_url?: string;
}

interface CourtListenerCitationResult {
  citation?: string;
  status?: number;
  clusters?: CourtListenerCluster[];
}

/**
 * Free Law Project's CourtListener (courtlistener.com) citation-lookup API.
 * Free to use, but requires an API token -- as of 2026, CourtListener's v4 API returns 401 for
 * an unauthenticated citation-lookup request (confirmed live; there is no anonymous tier for
 * this endpoint, despite what some documentation examples might suggest). See:
 * https://www.courtlistener.com/help/api/rest/v4/citation-lookup/
 */
export class CourtListenerProvider implements RateLimitAwareProvider {
  readonly id = "courtlistener";
  readonly name = "CourtListener";
  readonly description = "Free Law Project's case-law search. Requires a free CourtListener account and API token.";
  readonly requiresAuth = true;
  readonly credentialFields: ProviderCredentialField[] = [
    {
      key: "apiToken",
      label: "API token",
      type: "password",
      placeholder: "Paste your CourtListener API token",
      required: true,
    },
  ];

  private apiToken: string | null = null;
  // CourtListener's documented default rate limit is a modest 5 requests/minute (50/hour,
  // 125/day) -- https://www.courtlistener.com/help/api/rest/ -- so hitting it partway through
  // scanning a document with several dozen citations is expected in normal use, not an edge
  // case. This tracks whether the most recent lookupCitation() call's null result was actually
  // a 429, so callers can tell "rate-limited, retry later" apart from "genuinely not found".
  private lastRequestWasRateLimited = false;

  isAuthenticated(): boolean {
    return this.apiToken !== null;
  }

  wasLastRequestRateLimited(): boolean {
    return this.lastRequestWasRateLimited;
  }

  async authenticate(credentials: Record<string, string>): Promise<void> {
    const token = (credentials.apiToken || "").trim();
    if (!token) {
      this.apiToken = null;
      throw new Error("CourtListener requires an API token.");
    }

    const response = await this.request(token, "1 U.S. 1");
    if (response.status === 401 || response.status === 403) {
      throw new Error("CourtListener rejected the supplied API token.");
    }

    this.apiToken = token;
  }

  signOut(): void {
    this.apiToken = null;
  }

  async lookupCitation(citation: ParsedCitation): Promise<CitationMatch | null> {
    this.lastRequestWasRateLimited = false;

    if (!this.apiToken) {
      return null;
    }

    const text = citation.raw.trim();
    if (!text) {
      return null;
    }

    let response: Response;
    try {
      response = await this.request(this.apiToken, text);
    } catch {
      return null;
    }

    if (response.status === 429) {
      this.lastRequestWasRateLimited = true;
      return null;
    }

    if (!response.ok) {
      return null;
    }

    let results: CourtListenerCitationResult[];
    try {
      results = await response.json();
    } catch {
      return null;
    }

    if (!Array.isArray(results)) {
      return null;
    }

    for (const result of results) {
      if (result.status !== 200 || !Array.isArray(result.clusters) || result.clusters.length === 0) {
        continue;
      }

      const cluster = result.clusters[0];
      if (!cluster.absolute_url) {
        continue;
      }

      return {
        url: `${SITE_ORIGIN}${cluster.absolute_url}`,
        caseName: cluster.case_name,
        citation: result.citation,
      };
    }

    return null;
  }

  private request(token: string | null, text: string): Promise<Response> {
    const headers: Record<string, string> = {
      "Content-Type": "application/x-www-form-urlencoded",
    };
    if (token) {
      headers["Authorization"] = `Token ${token}`;
    }

    return fetch(`${API_BASE}/citation-lookup/`, {
      method: "POST",
      headers,
      body: new URLSearchParams({ text }),
    });
  }
}
