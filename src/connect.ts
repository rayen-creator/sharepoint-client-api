import axios from "axios";
import qs from "qs";
import { MicrosoftSharePointWrapper } from "./client";
import { SharePointAuthOptions } from "./interfaces/sharepoint-auth-options";

/**
 * Retrieves an access token from Azure AD for SharePoint API.
 * @param options SharePoint authentication options
 * @returns Access token string
 */
async function fetchSharePointAccessToken(
  options: SharePointAuthOptions
): Promise<string | undefined> {
  const { tenantId, refreshToken, scope, grantType, appCredentials } = options;
  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const body = qs.stringify({
    client_id: appCredentials.clientId,
    client_secret: appCredentials.clientSecret,
    refresh_token: refreshToken,
    grant_type: grantType ?? "refresh_token",
    scope: scope ?? "https://microsoft.sharepoint.com/.default",
  });

  try {
    const { data } = await axios.post(tokenUrl, body, {
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
    });

    return data.access_token;
  } catch (error: any) {
    const status = error.response?.status;
    const message =
      error.response?.data?.error_description ||
      error.response?.data?.error ||
      error.message;

    console.error(
      `[SharePoint Token Error] ${status ? `(${status})` : ""} ${message || ""}`
    );
  }
}

/**
 * Connects to a SharePoint site using an Azure AD app and returns a wrapper instance.
 * @param options SharePoint authentication options
 * @returns MicrosoftSharePointWrapper instance
 */
export async function connectWithSharePoint(
  options: SharePointAuthOptions
): Promise<MicrosoftSharePointWrapper | undefined> {
  const accessToken = await fetchSharePointAccessToken(options);
  if (!accessToken) throw new Error("Failed to obtain access token for SharePoint.");

  const sp = new MicrosoftSharePointWrapper(options.siteHostname, {
    Accept: "application/json;odata=verbose",
    "Content-Type": "application/json;odata=verbose",
    Authorization: `Bearer ${accessToken}`,
  });

  return sp;
}

/**
 * Example function to connect if the token is already known.
 * @param siteHostname SharePoint site hostname
 * @param accessToken Access token string
 * @returns MicrosoftSharePointWrapper instance
 */
export function connectWithToken(
  siteHostname: string,
  accessToken: string
): MicrosoftSharePointWrapper {
  return new MicrosoftSharePointWrapper(siteHostname, {
    Accept: "application/json;odata=verbose",
    "Content-Type": "application/json;odata=verbose",
    Authorization: `Bearer ${accessToken}`,
  });
}
