import { CitationProvider, CitationMatch, ParsedCitation, ProviderCredentialField } from "./types";

/**
 * Base class for providers backed by a paid/contract-gated research API
 * (LexisNexis, Westlaw, Bloomberg Law, or a firm's own internal API).
 *
 * OpenClerk never ships a fixed endpoint or a bundled key for these
 * platforms: each vendor provisions API access per customer contract, so the
 * base URL and credentials are always supplied by the user at runtime and
 * held in memory only for the current session. Nothing is written to disk,
 * localStorage, or any OpenClerk-controlled server.
 */
export abstract class EnterpriseCitationProvider implements CitationProvider {
  abstract readonly id: string;
  abstract readonly name: string;
  abstract readonly description: string;
  abstract readonly credentialFields: ProviderCredentialField[];
  readonly requiresAuth = true;

  protected credentials: Record<string, string> | null = null;

  isAuthenticated(): boolean {
    return this.credentials !== null;
  }

  async authenticate(credentials: Record<string, string>): Promise<void> {
    const missing = this.credentialFields.filter(
      (field) => field.required !== false && !(credentials[field.key] || "").trim()
    );
    if (missing.length > 0) {
      throw new Error(`Missing required field(s): ${missing.map((field) => field.label).join(", ")}`);
    }

    const apiBaseUrl = credentials.apiBaseUrl;
    if (apiBaseUrl && !/^https:\/\//i.test(apiBaseUrl.trim())) {
      throw new Error("The API base URL must start with https:// so credentials and citations are never sent unencrypted.");
    }

    await this.verifyCredentials(credentials);
    this.credentials = credentials;
  }

  signOut(): void {
    this.credentials = null;
  }

  /** Subclasses perform the vendor-specific handshake (e.g. OAuth2 client-credentials) and throw on failure. */
  protected abstract verifyCredentials(credentials: Record<string, string>): Promise<void>;

  abstract lookupCitation(citation: ParsedCitation): Promise<CitationMatch | null>;
}

export function trimTrailingSlash(value: string): string {
  return value.replace(/\/+$/, "");
}

/**
 * Performs an OAuth2 client-credentials handshake, the pattern most
 * enterprise legal research APIs use for machine-to-machine access. Returns
 * the bearer token on success; throws on any non-2xx response or missing
 * access_token so callers can surface a clear "authentication failed" error.
 */
export async function fetchClientCredentialsToken(
  tokenUrl: string,
  clientId: string,
  clientSecret: string
): Promise<string> {
  const response = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "client_credentials",
      client_id: clientId,
      client_secret: clientSecret,
    }),
  });

  if (!response.ok) {
    throw new Error(`Authentication failed (HTTP ${response.status}). Verify the API base URL and credentials.`);
  }

  const payload = await response.json();
  const accessToken = payload && payload.access_token;
  if (!accessToken || typeof accessToken !== "string") {
    throw new Error("Authentication succeeded but no access token was returned.");
  }

  return accessToken;
}
