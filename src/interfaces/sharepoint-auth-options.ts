import { AzureAppCredentials } from './azure-app-credentials';

/**
 * Represents credentials and settings to authenticate with SharePoint.
 */
export interface SharePointAuthOptions {
  /** Azure AD App credentials */
  appCredentials: AzureAppCredentials;
  /** The hostname of the SharePoint site (e.g., 'yourtenant.sharepoint.com') */
  siteHostname: string;
  /** The Azure tenant ID */
  tenantId: string;
  /** Refresh token for obtaining an access token */
  refreshToken: string;
  /** Optional OAuth2 scope (default: SharePoint online) */
  scope?: string;
  /** Optional grant type (default: 'refresh_token') */
  grantType?: string;
}
