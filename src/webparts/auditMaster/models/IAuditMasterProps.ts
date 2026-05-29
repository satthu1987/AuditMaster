import { WebPartContext } from '@microsoft/sp-webpart-base';

/**
 * Props passed from the webpart to the top-level React component.
 */
export interface IAuditMasterProps {
  description: string;
  context: WebPartContext;
  siteUrl: string;
}

/**
 * Validation error map – field name → error message.
 */
export interface IValidationErrors {
  [fieldName: string]: string;
}

/**
 * Navigation menu item shape.
 */
export interface IMenuItem {
  key: string;
  label: string;
  icon: string;
}

/**
 * Current user information enriched with role.
 */
export interface ICurrentUser {
  id: number;
  loginName: string;
  displayName: string;
  email: string;
  role: string;      // Admin | PIC | Verifier | Viewer
  segmentServiceIds?: number[];  // for Verifier scoping
}
