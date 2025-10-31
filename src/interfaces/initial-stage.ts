import { ConfiguredStage } from "./configured-stage";

/**
 * Initial stage to start an API request.
 * Provides entry points for regular or admin SharePoint APIs.
 */
export interface InitialStage {
  /**
   * Start a request for a specific site and endpoint
   * @param siteName SharePoint site name
   * @param endpoint Endpoint path (relative)
   */
  api(siteName: string, endpoint: string): ConfiguredStage;

  /**
   * Start a request for SharePoint admin endpoints
   * @param endpoint Admin endpoint path
   */
  adminApi(endpoint: string): ConfiguredStage;
}
