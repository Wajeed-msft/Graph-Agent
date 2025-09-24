// Template types for Graph API workloads
// These interfaces should be replaced with workload-specific types

export interface {WORKLOAD_NAME}Item {
  id: string;
  // TODO: Add workload-specific properties
  // Example for Calendar: title, start, end, attendees
  // Example for Teams: name, description, members
  // Example for SharePoint: name, url, contentType
  [key: string]: any;
}

export interface {WORKLOAD_NAME}Collection {
  value: {WORKLOAD_NAME}Item[];
  '@odata.nextLink'?: string;
  '@odata.count'?: number;
}

export interface {WORKLOAD_NAME}CreateRequest {
  // TODO: Add required fields for creation
  // Example for Calendar: subject, start, end
  // Example for Teams: displayName, description
  // Example for SharePoint: name, template
  [key: string]: any;
}

export interface {WORKLOAD_NAME}UpdateRequest {
  // TODO: Add updatable fields
  // Example for Calendar: subject, start, end, body
  // Example for Teams: displayName, description
  // Example for SharePoint: name, description
  [key: string]: any;
}

export interface {WORKLOAD_NAME}QueryOptions {
  select?: string[];
  filter?: string;
  orderBy?: string;
  top?: number;
  skip?: number;
  expand?: string[];
}

// Common Graph API response patterns
export interface GraphResponse<T> {
  value: T[];
  '@odata.nextLink'?: string;
  '@odata.count'?: number;
}

export interface GraphError {
  error: {
    code: string;
    message: string;
    innerError?: {
      code: string;
      message: string;
    };
  };
}

// Workload-specific permissions mapping
export const {WORKLOAD_NAME_UPPER}_PERMISSIONS = {
  READ: '{WORKLOAD_READ_PERMISSION}',
  WRITE: '{WORKLOAD_WRITE_PERMISSION}',
  // TODO: Add specific permissions for workload
  // Example for Calendar: 'Calendars.Read', 'Calendars.ReadWrite'
  // Example for Teams: 'Team.ReadBasic.All', 'Team.Create'
  // Example for SharePoint: 'Sites.Read.All', 'Sites.ReadWrite.All'
} as const;

// Common Graph API endpoints for workload
export const {WORKLOAD_NAME_UPPER}_ENDPOINTS = {
  BASE: '/{WORKLOAD_BASE_ENDPOINT}',
  // TODO: Add workload-specific endpoints
  // Example for Calendar: '/me/events', '/me/calendars'
  // Example for Teams: '/teams', '/me/joinedTeams'
  // Example for SharePoint: '/sites', '/me/drive'
} as const;