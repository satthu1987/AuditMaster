import * as React from 'react';
import { RoleWorkspace, IRoleViewProps, IRoleWorkspaceConfig } from './RoleWorkspace';
import { PIC_DetailsView } from '../Forms';

/**
 * PICView – the experience for users with the **PIC** (Person In Charge) role.
 *
 * Capabilities:
 *  • Sees only the audits assigned to them (viewMode = 'myAssigned', which the
 *    list view backs with SharePointService.getAuditItemsByPIC).
 *  • Cannot create new audit items – findings are raised by Admins/Auditors.
 *  • Focuses on working their findings: filling in root-cause, action taken and
 *    responses. The AuditItemForm only lets a PIC edit items where PICId matches
 *    their own id (enforced by RoleService.canUpdate) and hides the Verification
 *    section until the item reaches a verification stage.
 */
const PIC_CONFIG: IRoleWorkspaceConfig = {
  viewMode: 'myAssigned',
  allowCreate: false,
  roleLabel: 'PIC',
  roleIcon: 'Contact',
  description: 'Manage the audits assigned to you – add responses, root-cause and corrective actions.',
  infoBanner:
    'These are the audit findings assigned to you. Open an item to record your root-cause analysis, corrective actions and responses.',
    formComponent: PIC_DetailsView
};

const PICView: React.FC<IRoleViewProps> = (props) => {
  return <RoleWorkspace {...props} config={PIC_CONFIG} />;
};

export default PICView;
