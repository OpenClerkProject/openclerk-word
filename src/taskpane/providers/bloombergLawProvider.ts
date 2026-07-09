import { CitationMatch, ParsedCitation, ProviderCredentialField } from "./types";
import { EnterpriseCitationProvider, fetchClientCredentialsToken, trimTrailingSlash } from "./base";

/**
 * Bloomberg Law's API is contract-gated and invite-only: access and the
 * exact token/search paths are provisioned per customer, so there is no
 * single public endpoint OpenClerk can ship. This provider implements the
 * common OAuth2 client-credentials shape and a configurable base URL;
 * confirm the exact paths in your firm's Bloomberg Law API documentation
 * and adjust TOKEN_PATH/SEARCH_PATH below if they differ.
 */
const TOKEN_PATH = "/oauth/token";
const SEARCH_PATH = "/api/v1/search/cases";

export class BloombergLawProvider extends EnterpriseCitationProvider {
  readonly id = "bloomberglaw";
  readonly name = "Bloomberg Law";
  readonly description =
    "Looks up citations through your organization's Bloomberg Law API subscription. Requires the API base URL and client credentials issued under your firm's Bloomberg Law contract.";
  readonly credentialFields: ProviderCredentialField[] = [
    { key: "apiBaseUrl", label: "API base URL (from your Bloomberg Law contract)", type: "text", placeholder: "https://your-tenant.api.bloomberglaw.com" },
    { key: "clientId", label: "Client ID", type: "text" },
    { key: "clientSecret", label: "Client secret", type: "password" },
  ];

  private accessToken: string | null = null;

  protected async verifyCredentials(credentials: Record<string, string>): Promise<void> {
    const baseUrl = trimTrailingSlash(credentials.apiBaseUrl);
    this.accessToken = await fetchClientCredentialsToken(`${baseUrl}${TOKEN_PATH}`, credentials.clientId, credentials.clientSecret);
  }

  signOut(): void {
    super.signOut();
    this.accessToken = null;
  }

  async lookupCitation(citation: ParsedCitation): Promise<CitationMatch | null> {
    if (!this.credentials || !this.accessToken) {
      return null;
    }

    try {
      const baseUrl = trimTrailingSlash(this.credentials.apiBaseUrl);
      const response = await fetch(`${baseUrl}${SEARCH_PATH}`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${this.accessToken}`,
        },
        body: JSON.stringify({ citation: citation.raw }),
      });

      if (!response.ok) {
        return null;
      }

      const payload = await response.json();
      const match = payload && Array.isArray(payload.results) ? payload.results[0] : null;
      if (!match || !match.url) {
        return null;
      }

      return { url: match.url, caseName: match.caseName || match.title, citation: citation.raw };
    } catch {
      return null;
    }
  }
}
