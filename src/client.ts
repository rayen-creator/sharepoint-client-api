import axios from "axios";
import { ConfiguredStage } from "./interfaces/configured-stage";
import { HttpMethod } from "./enums/http-method";
import { InitialStage } from "./interfaces/initial-stage";

/**
 * Main SharePoint API wrapper for interacting with site or admin endpoints.
 * Implements a fluent API with chaining for query building and request execution.
 */
export class MicrosoftSharePointWrapper
  implements InitialStage, ConfiguredStage
{
  private headers: any;
  private siteHostname: string;
  private endpoint: string;
  private siteName: string;
  private isAdmin: boolean;
  private tempHeaders: Record<string, any>;
  private shouldIgnoreErrors: boolean;
  private initialized: boolean;
  private queryParams: Record<string, string>;

  /**
   * Creates a new SharePoint wrapper instance.
   * @param siteHostname The hostname of the SharePoint site (e.g., "mytenant.sharepoint.com")
   * @param headers Default headers to use for requests
   */
  constructor(siteHostname: string, headers: any) {
    this.siteHostname = siteHostname;
    this.headers = headers;
    this.endpoint = "";
    this.siteName = "";
    this.isAdmin = false;
    this.tempHeaders = {};
    this.shouldIgnoreErrors = false;
    this.initialized = false;
    this.queryParams = {};
  }

  // -----------------------
  //      Initial stage
  // -----------------------

  /**
   * Initialize a request for a specific SharePoint site endpoint.
   * @param siteName The site name to target
   * @param endpoint API endpoint path (relative)
   */
  api(siteName: string, endpoint: string): ConfiguredStage {
    this.siteName = siteName;
    this.endpoint = endpoint;
    this.isAdmin = false;
    this.initialized = true;
    this.queryParams = {};
    return this;
  }

  /**
   * Initialize a request for SharePoint admin endpoints.
   * @param endpoint Admin endpoint path
   */
  adminApi(endpoint: string): ConfiguredStage {
    this.endpoint = endpoint;
    this.siteName = "";
    this.isAdmin = true;
    this.initialized = true;
    this.queryParams = {};
    return this;
  }

  /**
   * Override or add temporary headers for this request.
   * @param overrides Headers to merge with default headers
   */
  setHeaders(overrides: Record<string, any>): ConfiguredStage {
    this.tempHeaders = overrides;
    return this;
  }

  /**
   * Ignore errors for this request and return null instead of throwing.
   */
  ignore(): ConfiguredStage {
    this.shouldIgnoreErrors = true;
    return this;
  }

  // -----------------------
  //   Query builder helpers
  // -----------------------

  /**
   * Select specific fields from the response.
   * @param fields Field names to select
   */
  select(fields: string | string[]): ConfiguredStage {
    const value = Array.isArray(fields) ? fields.join(",") : fields;
    this.queryParams["$select"] = value;
    return this;
  }

  /**
   * Filter results using an OData condition.
   * @param condition OData filter string
   */
  filter(condition: string): ConfiguredStage {
    this.queryParams["$filter"] = condition;
    return this;
  }

  /**
   * Expand related entities in the response.
   * @param fields Fields to expand
   */
  expand(fields: string | string[]): ConfiguredStage {
    const value = Array.isArray(fields) ? fields.join(",") : fields;
    this.queryParams["$expand"] = value;
    return this;
  }

  /**
   * Order results by a specific field.
   * @param field Field name to order by
   * @param ascending Whether to order ascending (default: true)
   */
  orderBy(field: string, ascending = true): ConfiguredStage {
    this.queryParams["$orderby"] = `${field} ${ascending ? "asc" : "desc"}`;
    return this;
  }

  /**
   * Limit the number of results returned.
   * @param count Number of items to return
   */
  top(count: number): ConfiguredStage {
    this.queryParams["$top"] = count.toString();
    return this;
  }

  /**
   * Skip a number of results.
   * @param count Number of items to skip
   */
  skip(count: number): ConfiguredStage {
    this.queryParams["$skip"] = count.toString();
    return this;
  }

  /**
   * Add raw query parameters.
   * @param params Key-value pairs to append to the query string
   */
  rawQuery(params: Record<string, any>): ConfiguredStage {
    Object.entries(params).forEach(([key, value]) => {
      this.queryParams[key] = String(value);
    });
    return this;
  }

  // -----------------------
  // Internal helpers
  // -----------------------

  /** Get SharePoint request digest for POST/PUT/PATCH requests */
  private async getRequestDigest(): Promise<string> {
    const response = await axios.post(
      `https://${this.siteHostname}/_api/contextinfo`,
      null,
      { headers: this.headers }
    );
    return response.data.d.GetContextWebInformation.FormDigestValue;
  }

  /** Build query string from query parameters */
  private buildQueryString(): string {
    const entries = Object.entries(this.queryParams);
    if (entries.length === 0) return "";
    const query = entries.map(([k, v]) => `${k}=${v}`).join("&");
    return `?${query}`;
  }

  /**
   * Execute a request with the given HTTP method and optional data.
   * @param method HTTP method
   * @param data Optional request body
   */
  private async makeRequest(method: HttpMethod, data?: any) {
    if (!this.initialized) {
      throw new Error(
        "❌ You must call .api() or .adminApi() before making a request."
      );
    }

    const baseUrl = this.isAdmin
      ? `https://${this.siteHostname.replace(
          ".sharepoint.com",
          ""
        )}-admin.sharepoint.com/_api/${this.endpoint}`
      : `https://${this.siteHostname}/sites/${this.siteName}/_api/${this.endpoint}`;

    const fullUrl = `${baseUrl}${this.buildQueryString()}`;
    const headers = { ...this.headers, ...this.tempHeaders };

    if (method !== HttpMethod.GET) {
      headers["X-RequestDigest"] = await this.getRequestDigest();
    }

    try {
      const response = await axios({
        url: fullUrl,
        method,
        headers,
        data: method !== HttpMethod.GET ? JSON.stringify(data) : undefined,
      });
      return response.data;
    } catch (error: any) {
      const status = error.response?.status;
      const spError = error.response?.data?.error;
      const message =
        spError?.message?.value || error.message || "Unknown SharePoint API error";

      if (this.shouldIgnoreErrors) {
        console.warn(
          `Ignoring SP error for: ${method} ${this.endpoint}` +
            (spError?.code ? ` (Code: ${spError.code})` : "") +
            ` - ${message}`
        );
        return null;
      }

      const cleanMessage =
        `Microsoft SharePoint API Error at: ${method} ${this.endpoint}` +
        (status ? ` (Status: ${status})` : "") +
        (spError?.code ? ` (Code: ${spError.code})` : "") +
        ` - ${message}`;

      throw new Error(cleanMessage);
    } finally {
      // Reset temporary state
      this.tempHeaders = {};
      this.isAdmin = false;
      this.shouldIgnoreErrors = false;
      this.initialized = false;
      this.queryParams = {};
    }
  }

  // -----------------------
  //   HTTP request methods
  // -----------------------

  /** Execute a GET request */
  async get() {
    return await this.makeRequest(HttpMethod.GET);
  }

  /** Execute a POST request */
  async post(data: any) {
    return await this.makeRequest(HttpMethod.POST, data);
  }

  /** Execute a PUT request */
  async put(data: Record<string, any>) {
    return await this.makeRequest(HttpMethod.PUT, data);
  }

  /** Execute a PATCH request */
  async patch(data: Record<string, any>) {
    return await this.makeRequest(HttpMethod.PATCH, data);
  }

  /** Execute a DELETE request */
  async delete() {
    return await this.makeRequest(HttpMethod.DELETE);
  }
}
