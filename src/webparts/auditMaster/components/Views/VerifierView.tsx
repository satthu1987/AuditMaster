import * as React from 'react';
import { RoleWorkspace, IRoleViewProps, IRoleWorkspaceConfig } from './RoleWorkspace';
import { Verifier_DetailsView } from '../Forms';

/**
 * VerifierView – the experience for users with the **Verifier** role.
 *
 * Capabilities:
 *  • Sees the items awaiting their verification (viewMode = 'myPending', backed
 *    by SharePointService.getAuditItemsForVerifier which filters to the
 *    verifier's segment/services and "Pending Verification" status).
 *  • Cannot create new findings (allowCreate = false).
 *  • Focused verification experience: the AuditItemForm locks the descriptive
 *    sections for Verifiers and keeps Status + Verification & Links editable, so
 *    they can record the verification result and approve/close the item.
 */
const VERIFIER_CONFIG: IRoleWorkspaceConfig = {
  viewMode: 'myPending',
  allowCreate: false,
  roleLabel: 'Verifier',
  roleIcon: 'CheckList',
  description: 'Verify and approve audit findings that are pending verification.',
  infoBanner:
    'These items are awaiting your verification. Open an item to review the corrective action, record the verification result and close it out.',

    formComponent: Verifier_DetailsView
};

const VerifierView: React.FC<IRoleViewProps> = (props) => {
  return <RoleWorkspace {...props} config={VERIFIER_CONFIG} />;
};

export default VerifierView;
