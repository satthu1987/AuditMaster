import * as React from 'react';
import { RoleWorkspace, IRoleViewProps, IRoleWorkspaceConfig } from './RoleWorkspace';

/**
 * AdminView – the experience for users with the **Admin** role.
 *
 * Capabilities:
 *  • Full access to every audit item (viewMode = 'all').
 *  • May create brand-new audit items (allowCreate = true).
 *  • The underlying AuditItemForm renders every section as editable because
 *    RoleService.canCreate / canUpdate return true for Admins.
 *
 * This is the only role with unrestricted create + edit rights.
 */
const ADMIN_CONFIG: IRoleWorkspaceConfig = {
  viewMode: 'all',
  allowCreate: true,
  roleLabel: 'Administrator',
  roleIcon: 'Admin',
  description: 'Full access – create, edit and manage every audit finding.'
};

const AdminView: React.FC<IRoleViewProps> = (props) => {
  return <RoleWorkspace {...props} config={ADMIN_CONFIG} />;
};

export default AdminView;
