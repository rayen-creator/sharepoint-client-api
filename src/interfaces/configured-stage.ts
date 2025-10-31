/**
 * Fluent stage of an API request after endpoint selection.
 * Allows chaining query modifiers, headers, and performing HTTP requests.
 */
export interface ConfiguredStage {
  /**
   * Override headers for this request.
   * @param overrides Additional headers to merge
   */
  setHeaders(overrides: Record<string, any>): ConfiguredStage;

  /** Skip this request when executing (for conditional chains) */
  ignore(): ConfiguredStage;

  /** Select specific fields from the response */
  select(fields: string | string[]): ConfiguredStage;

  /** Filter the results using an OData condition */
  filter(condition: string): ConfiguredStage;

  /** Expand related entities (OData expand) */
  expand(fields: string | string[]): ConfiguredStage;

  /** Order results by field, ascending by default */
  orderBy(field: string, ascending?: boolean): ConfiguredStage;

  /** Limit the number of results returned */
  top(count: number): ConfiguredStage;

  /** Skip a number of results */
  skip(count: number): ConfiguredStage;

  /** Add raw query parameters */
  rawQuery(params: Record<string, any>): ConfiguredStage;

  /** Execute a GET request */
  get(): Promise<any>;

  /** Execute a POST request */
  post(data: any): Promise<any>;

  /** Execute a PUT request */
  put(data: Record<string, any>): Promise<any>;

  /** Execute a PATCH request */
  patch(data: Record<string, any>): Promise<any>;

  /** Execute a DELETE request */
  delete(): Promise<any>;
}
