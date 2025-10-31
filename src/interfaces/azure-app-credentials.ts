/**
 * Represents credentials for a Microsoft Azure AD App.
 */
export interface AzureAppCredentials {
  /** The Client ID of the Azure AD App */
  clientId: string;
  /** The Client Secret of the Azure AD App */
  clientSecret: string;
}