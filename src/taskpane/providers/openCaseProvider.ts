import { CitationMatch, ParsedCitation, ProviderCredentialField } from "./types";
import { EnterpriseCitationProvider, fetchClientCredentialsToken, trimTrailingSlash } from "./base";

/**
 * OpenCase (opencase.com) is an AI-powered legal research assistant built as
 * a Word add-in for U.S. legal professionals. As of this writing it has no
 * publicly documented developer API -- there is no equivalent to a
 * dev.opencase.com portal the way LexisNexis or Thomson Reuters publish one.
 * This provider follows the same shape as the other enterprise providers in
 * this directory (a configurable base URL plus the common OAuth2
 * client-credentials handshake) as a reasonable starting point, but every
 * detail here -- the token/search paths, the request/response shape -- is a
 * placeholder. Confirm actual integration details with OpenCase directly
 * (see the README's "Getting credentials" section) before relying on this.
 */
const TOKEN_PATH = "/oauth/token";
const SEARCH_PATH = "/search/cases";

export class OpenCaseProvider extends EnterpriseCitationProvider {
  readonly id = "opencase";
  readonly name = "OpenCase";
  readonly description =
    "Looks up citations through your organization's OpenCase subscription. No public API is documented as of this writing, so this uses the same configurable-endpoint shape as the other enterprise providers -- confirm the real integration details with OpenCase before relying on it.";
  readonly credentialFields: ProviderCredentialField[] = [
    { key: "apiBaseUrl", label: "API base URL (confirm with OpenCase)", type: "text", placeholder: "https://your-tenant.api.opencase.com" },
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
