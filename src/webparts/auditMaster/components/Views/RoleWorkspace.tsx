import * as React from 'react';
import { Icon, MessageBar, MessageBarType, PrimaryButton } from '@fluentui/react';
import { IAuditMasterItem, ICurrentUser } from '../../models';
import { SharePointService, RoleService } from '../../services';
import AuditListView from './AuditListView';
import { AuditItemForm,IAuditItemFormProps } from '../Forms';

/**
 * Configuration that each role-specific view supplies to the shared workspace.
 * This keeps every role view tiny and declarative while the common
 * list ⇄ form navigation logic lives in one place.
 */
export interface IRoleWorkspaceConfig {
  /** Which server-side query the list view should run. */
  viewMode: 'all' | 'myAssigned' | 'myPending';
  /** Whether this role may create brand-new audit items. */
  allowCreate: boolean;
  /** Friendly role label shown in the workspace header. */
  roleLabel: string;
  /** Icon (Fluent UI icon name) shown beside the role label. */
  roleIcon: string;
  /** Short sentence describing what this role can do here. */
  description: string;
  /** Optional info banner shown above the list. */
  infoBanner?: string;

  formComponent?: React.FC<IAuditItemFormProps>; 
}

/**
 * Props common to every role view. AuditMaster passes these down after it has
 * resolved the current user via RoleService.
 */
export interface IRoleViewProps {
  spService: SharePointService;
  roleService: RoleService;
  currentUser: ICurrentUser;
  /** The sidebar menu key currently selected (e.g. 'create', 'all'). */
  selectedMenu: string;
}

type WorkspaceMode = 'list' | 'create' | 'edit';

interface IWorkspaceState {
  mode: WorkspaceMode;
  editItem: IAuditMasterItem | null;
}

/**
 * RoleWorkspace – the shared engine behind every role view.
 *
 * Responsibilities:
 *  • Owns the internal navigation between the list view and the item form
 *    (so each role view is self-contained, as required by the architecture).
 *  • Reacts to sidebar menu changes: selecting "create" opens a blank form
 *    (when the role is allowed to create), any other menu returns to the list.
 *  • Delegates the actual rendering to the existing, battle-tested
 *    `AuditListView` and `AuditItemForm` components – which already apply
 *    role-based field visibility / read-only rules via RoleService.
 *
 * The role views (AdminView, PICView, …) are thin wrappers that only supply an
 * `IRoleWorkspaceConfig`. This gives us five clearly-named entry points while
 * avoiding duplicated navigation code.
 */
export const RoleWorkspace: React.FC<IRoleViewProps & { config: IRoleWorkspaceConfig }> = (
  props
) => {
  const { spService, roleService, currentUser, selectedMenu, config } = props;

  const [state, setState] = React.useState<IWorkspaceState>({
    mode: 'list',
    editItem: null
  });

  // ── React to sidebar menu changes ─────────────────────────────────────────
  // When the user clicks "Create New Audit Item" we switch to create mode
  // (only if the role is permitted to create). Any other menu selection takes
  // the user back to the list for that role.
  React.useEffect(() => {
    if (selectedMenu === 'create' && config.allowCreate) {
      setState({ mode: 'create', editItem: null });
    } else {
      setState({ mode: 'list', editItem: null });
    }
  }, [selectedMenu, config.allowCreate]);

  // ── Navigation handlers ────────────────────────────────────────────────────
  const handleEditItem = (item: IAuditMasterItem): void => {
    // Open the selected item in the form. The form itself decides whether the
    // fields are editable based on RoleService.canUpdate – so even a Viewer can
    // safely open an item and see it in read-only mode.
    setState({ mode: 'edit', editItem: item });
  };

  const handleBackToList = (): void => {
    setState({ mode: 'list', editItem: null });
  };

  // ── Render: Create / Edit form ─────────────────────────────────────────────
  if (state.mode === 'create' || (state.mode === 'edit' && state.editItem)) {
    const FormComponent = config.formComponent || AuditItemForm; // Default to generic form if no custom form supplied
    return (
      <div>
        <FormComponent
          spService={spService}
          roleService={roleService}
          currentUser={currentUser}
          editItem={state.mode === 'edit' ? state.editItem! : undefined}
          onSaved={handleBackToList}
          onCancel={handleBackToList}
        />
      </div>
    );
  }

  // ── Render: List (default) ─────────────────────────────────────────────────
  return (
    <div>
      {/* Role workspace header – clarifies the active role context. */}
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          gap: 12,
          padding: '16px 20px 0 20px'
        }}
      >
        <Icon iconName={config.roleIcon} style={{ fontSize: 28, color: '#4facfe' }} />
        <div>
          <div style={{ fontSize: 18, fontWeight: 600, color: '#1a1a2e' }}>
            {config.roleLabel} Workspace
          </div>
          <div style={{ fontSize: 13, color: '#605e5c' }}>{config.description}</div>
        </div>
        {config.allowCreate && (
          <div style={{ marginLeft: 'auto' }}>
            <PrimaryButton
              iconProps={{ iconName: 'Add' }}
              text="New Audit Item"
              onClick={() => setState({ mode: 'create', editItem: null })}
            />
          </div>
        )}
      </div>

      {config.infoBanner && (
        <div style={{ padding: '12px 20px 0 20px' }}>
          <MessageBar messageBarType={MessageBarType.info} isMultiline={false}>
            {config.infoBanner}
          </MessageBar>
        </div>
      )}

      <AuditListView
        spService={spService}
        roleService={roleService}
        currentUser={currentUser}
        viewMode={config.viewMode}
        onEditItem={handleEditItem}
      />
    </div>
  );
};

export default RoleWorkspace;
