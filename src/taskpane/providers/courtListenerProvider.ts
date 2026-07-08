import { CitationProvider, CitationMatch, OpinionTextCapableProvider, ParsedCitation, ProviderCredentialField } from "./types";
import { extractPageExcerpt, stripHtmlTags } from "./opinionTextExtractor";

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

interface CourtListenerOpinion {
  plain_text?: string;
  html_with_citations?: string;
  html?: string;
  html_lawbox?: string;
  html_columbia?: string;
}

interface CourtListenerOpinionsResponse {
  results?: CourtListenerOpinion[];
}

/**
 * Free Law Project's CourtListener (courtlistener.com) citation-lookup API.
 * Public and free to use; an API token is optional but raises the (fairly
 * low) anonymous rate limit. See:
 * https://www.courtlistener.com/help/api/rest/v4/citation-lookup/
 */
export class CourtListenerProvider implements OpinionTextCapableProvider {
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

  /**
   * Opinion text fetching requires an API token even though basic citation lookup doesn't --
   * there's no free anonymous route to full opinion text on CourtListener's v4 API.
   */
  isReadyForOpinionText(): boolean {
    return this.apiToken !== null;
  }

  async fetchOpinionExcerpt(citation: ParsedCitation, targetPages: number[]): Promise<string | null> {
    if (!this.apiToken || targetPages.length === 0) {
      return null;
    }

    const text = citation.raw.trim();
    if (!text) {
      return null;
    }

    const clusterId = await this.resolveClusterId(text);
    if (!clusterId) {
      return null;
    }

    let opinions: CourtListenerOpinion[];
    try {
      const response = await fetch(`${API_BASE}/opinions/?cluster=${encodeURIComponent(clusterId)}`, {
        headers: { Authorization: `Token ${this.apiToken}` },
      });
      if (!response.ok) {
        return null;
      }
      const data: CourtListenerOpinionsResponse = await response.json();
      opinions = Array.isArray(data.results) ? data.results : [];
    } catch {
      return null;
    }

    for (const opinion of opinions) {
      const source =
        opinion.plain_text ||
        stripHtmlTags(opinion.html_with_citations || opinion.html || opinion.html_lawbox || opinion.html_columbia || "");
      if (!source) {
        continue;
      }
      const excerpt = extractPageExcerpt(source, targetPages);
      if (excerpt) {
        return excerpt;
      }
    }

    return null;
  }

  /** Resolves a citation to its CourtListener cluster ID by parsing it out of the citation-lookup result's absolute_url. */
  private async resolveClusterId(text: string): Promise<string | null> {
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
      if (result.status !== 200 || !Array.isArray(result.clusters) || result.clusters.length === 0) {
        continue;
      }
      const absoluteUrl = result.clusters[0].absolute_url;
      const idMatch = absoluteUrl && absoluteUrl.match(/^\/opinion\/(\d+)\//);
      if (idMatch) {
        return idMatch[1];
      }
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
