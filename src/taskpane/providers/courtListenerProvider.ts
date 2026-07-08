import { CitationProvider, CitationMatch, ParsedCitation, ProviderCredentialField } from "./types";

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
 * Public and free to use; an API token is optional but raises the (fairly
 * low) anonymous rate limit. See:
 * https://www.courtlistener.com/help/api/rest/v4/citation-lookup/
 */
export class CourtListenerProvider implements CitationProvider {
  readonly id = "courtlistener";
  readonly name = "CourtListener";
  readonly description =
    "Free Law Project's free, public case-law search. Works without credentials; an optional API token raises the rate limit.";
  readonly requiresAuth = false;
  readonly credentialFields: ProviderCredentialField[] = [
    {
      key: "apiToken",
      label: "API token (optional)",
      type: "password",
      placeholder: "Paste your CourtListener API token",
      required: false,
    },
  ];

  private apiToken: string | null = null;

  isAuthenticated(): boolean {
    return true;
  }

  async authenticate(credentials: Record<string, string>): Promise<void> {
    const token = (credentials.apiToken || "").trim();
    if (!token) {
      this.apiToken = null;
      return;
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
      // 200 = resolved to exactly one case. 300 ("Multiple Choices") means the citation is
      // still real and valid -- it just matches more than one case -- so it must not be
      // treated the same as a genuine miss (404 not found, 400 unrecognized reporter, etc.).
      if (result.status !== 200 && result.status !== 300) {
        continue;
      }
      if (!Array.isArray(result.clusters) || result.clusters.length === 0) {
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
        ...(result.status === 300 ? { ambiguous: true } : {}),
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
