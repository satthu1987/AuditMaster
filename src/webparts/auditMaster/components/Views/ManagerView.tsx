import * as React from 'react';
import { RoleWorkspace, IRoleViewProps, IRoleWorkspaceConfig } from './RoleWorkspace';

/**
 * ManagerView – the experience for users with the **Quality Manager** role.
 *
 * Capabilities:
 *  • Oversight across all audit findings (viewMode = 'all') so the manager can
 *    monitor progress, status distribution and overdue items via the list view's
 *    statistics bar and filters.
 *  • Does not create findings (allowCreate = false) – their job is review and
 *    oversight rather than authoring.
 *  • Opening an item shows it through AuditItemForm; RoleService.canUpdate keeps
 *    a Quality Manager in read-only mode, giving a safe review experience while
 *    they assess root-cause quality, corrective actions and verification status.
 */
const MANAGER_CONFIG: IRoleWorkspaceConfig = {
  viewMode: 'all',
  allowCreate: false,
  roleLabel: 'Quality Manager',
  roleIcon: 'ComplianceAudit',
  description: 'Review and oversee all audit findings across the organisation.',
  infoBanner:
    'Oversight view: monitor every audit finding, track status and review root-cause quality. Items open in review (read-only) mode.'
};

const ManagerView: React.FC<IRoleViewProps> = (props) => {
  return <RoleWorkspace {...props} config={MANAGER_CONFIG} />;
};

export default ManagerView;
