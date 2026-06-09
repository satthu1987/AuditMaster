import * as React from 'react';
import { RoleWorkspace, IRoleViewProps, IRoleWorkspaceConfig } from './RoleWorkspace';

/**
 * ViewerView – the fallback experience for users without a specific role
 * (e.g. an explicit "Viewer", an "Auditor", or any user with no UserRole record).
 *
 * Capabilities:
 *  • Read-only access to all audit findings (viewMode = 'all').
 *  • Cannot create items (allowCreate = false).
 *  • Opening an item shows it through AuditItemForm in read-only mode, because
 *    RoleService.canUpdate returns false for any non Admin/PIC/Verifier role.
 *
 * This guarantees a safe, non-destructive default for unknown roles.
 */
const VIEWER_CONFIG: IRoleWorkspaceConfig = {
  viewMode: 'all',
  allowCreate: false,
  roleLabel: 'Viewer',
  roleIcon: 'View',
  description: 'Read-only access – browse audit findings without editing.',
  infoBanner:
    'You have read-only access. You can browse and search all audit findings, but cannot make changes.'
};

const ViewerView: React.FC<IRoleViewProps> = (props) => {
  return <RoleWorkspace {...props} config={VIEWER_CONFIG} />;
};

export default ViewerView;
