import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';
import { GraphError } from '../types/graphTypes.template';

export abstract class BaseGraphService {
  protected graphClient: Client;
  protected accessToken: string;

  constructor(accessToken: string) {
    this.accessToken = accessToken;

    // Log the token for debugging purposes
    console.log('🔑 Microsoft Graph Access Token:', this.accessToken);
    console.log('🔑 Token length:', this.accessToken.length);
    console.log('🔑 Token starts with:', this.accessToken.substring(0, 20) + '...');

    // Initialize the Graph client with a custom auth provider
    const authProvider = (done: (error: any, accessToken: string | null) => void) => {
      try {
        console.log('🔐 AuthProvider called - providing token for Graph API call');
        done(null, this.accessToken);
      } catch (error) {
        console.error('🚨 AuthProvider error:', error);
        done(error, null);
      }
    };

    this.graphClient = Client.init({
      authProvider: authProvider
    });
  }

  /**
   * Handle common Graph API errors with user-friendly messages
   */
  protected handleGraphError(error: any, context: string): Error {
    console.error(`Error in ${context}:`, error);

    if (error.code === 'Forbidden' || error.status === 403) {
      return new Error(`Access denied. Please ensure you have the required permissions for ${context}.`);
    } else if (error.code === 'NotFound' || error.status === 404) {
      return new Error(`Resource not found. Please check the ID provided for ${context}.`);
    } else if (error.code === 'BadRequest' || error.status === 400) {
      return new Error(`Invalid request. Please check the data provided for ${context}.`);
    } else if (error.code === 'Unauthorized' || error.status === 401) {
      return new Error(`Authentication required. Please sign in to perform ${context}.`);
    } else if (error.code === 'TooManyRequests' || error.status === 429) {
      return new Error(`Rate limit exceeded. Please try again later for ${context}.`);
    }

    return new Error(`Failed to ${context}: ${error.message || 'Unknown error'}`);
  }

  /**
   * Get current user information
   */
  async getCurrentUser(): Promise<any> {
    try {
      const response = await this.graphClient.api('/me').get();
      return response;
    } catch (error: any) {
      throw this.handleGraphError(error, 'get current user information');
    }
  }

  /**
   * Get user's groups (useful for many workloads that require group membership)
   */
  async getUserGroups(): Promise<any[]> {
    try {
      const response = await this.graphClient.api('/me/memberOf').get();
      return response.value?.filter((group: any) => group['@odata.type'] === '#microsoft.graph.group') || [];
    } catch (error: any) {
      throw this.handleGraphError(error, 'get user groups');
    }
  }

  /**
   * Generic method to get paginated results from Graph API
   */
  protected async getPaginatedResults<T>(
    endpoint: string,
    queryParams: Record<string, any> = {}
  ): Promise<T[]> {
    try {
      let request = this.graphClient.api(endpoint);

      // Apply query parameters
      Object.entries(queryParams).forEach(([key, value]) => {
        if (value !== undefined && value !== null) {
          request = request.query(key, value.toString());
        }
      });

      const response = await request.get();
      let results = response.value || [];

      // Handle pagination
      let nextLink = response['@odata.nextLink'];
      while (nextLink) {
        const nextResponse = await this.graphClient.api(nextLink).get();
        results = results.concat(nextResponse.value || []);
        nextLink = nextResponse['@odata.nextLink'];
      }

      return results;
    } catch (error: any) {
      throw this.handleGraphError(error, `get data from ${endpoint}`);
    }
  }

  /**
   * Generic method to create a resource
   */
  protected async createResource<T>(endpoint: string, data: any): Promise<T> {
    try {
      const response = await this.graphClient.api(endpoint).post(data);
      return response;
    } catch (error: any) {
      throw this.handleGraphError(error, `create resource at ${endpoint}`);
    }
  }

  /**
   * Generic method to update a resource
   */
  protected async updateResource<T>(endpoint: string, data: any): Promise<T> {
    try {
      const response = await this.graphClient.api(endpoint).patch(data);
      return response;
    } catch (error: any) {
      throw this.handleGraphError(error, `update resource at ${endpoint}`);
    }
  }

  /**
   * Generic method to delete a resource
   */
  protected async deleteResource(endpoint: string): Promise<void> {
    try {
      await this.graphClient.api(endpoint).delete();
    } catch (error: any) {
      throw this.handleGraphError(error, `delete resource at ${endpoint}`);
    }
  }
}