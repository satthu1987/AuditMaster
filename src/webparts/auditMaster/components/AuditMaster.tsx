import * as React from 'react';
import { Spinner, SpinnerSize, Icon, PrimaryButton } from '@fluentui/react';
import styles from './AuditMaster.module.scss';
import { IAuditMasterProps, ICurrentUser, IAuditMasterItem, IMenuItem } from '../models';
import { SharePointService, ListProvisioningService, RoleService } from '../services';
import { Sidebar } from './Sidebar';
import { AuditItemForm } from './Forms';
import { AuditListView } from './Views';

enum AppPhase {
  Initializing = 'Initializing',
  Provisioning = 'Provisioning',
  Ready = 'Ready',
  Error = 'Error'
}

interface IAppState {
  phase: AppPhase;
  statusMessage: string;
  errorMessage: string;
  currentUser: ICurrentUser | null;
  menuItems: IMenuItem[];
  selectedMenu: string;
  editItem: IAuditMasterItem | null;
  createFormKey: number;
}

/**
 * Top-level container component for the Audit Master webpart.
 *
 * Lifecycle:
 *  1. Provision lists (create if missing)
 *  2. Resolve current user + role
 *  3. Render sidebar + content area
 */
const AuditMaster: React.FC<IAuditMasterProps> = (props) => {
  const { context, siteUrl } = props;

  // Services (stable refs)
  const spService = React.useMemo(
    () => new SharePointService(context, siteUrl),
    [context, siteUrl]
  );
  const provisioningService = React.useMemo(
    () => new ListProvisioningService(context, siteUrl),
    [context, siteUrl]
  );
  const roleService = React.useMemo(
    () => new RoleService(spService),
    [spService]
  );

  const [state, setState] = React.useState<IAppState>({
    phase: AppPhase.Initializing,
    statusMessage: 'Checking environment...',
    errorMessage: '',
    currentUser: null,
    menuItems: [],
    selectedMenu: '',
    editItem: null,
    createFormKey: 0
  });

  // ── Initialization ────────────────────────────────────────────────────────
  React.useEffect(() => {
    initialize();
  }, []);

  const initialize = async (): Promise<void> => {
    try {
      // Phase 1: Provision lists
      setState(prev => ({
        ...prev,
        phase: AppPhase.Provisioning,
        statusMessage: 'Provisioning SharePoint lists...'
      }));
      await provisioningService.provisionAllLists();

      // Phase 2: Resolve user
      setState(prev => ({ ...prev, statusMessage: 'Resolving user permissions...' }));
      const user = await roleService.resolveCurrentUser();
      const menuItems = roleService.getMenuItems(user);
      const defaultMenu = menuItems.length > 0 ? menuItems[0].key : '';
      console.log('[AuditMaster] Initialization successful. Current user:', menuItems);
      setState({
        phase: AppPhase.Ready,
        statusMessage: '',
        errorMessage: '',
        currentUser: user,
        menuItems,
        selectedMenu: defaultMenu,
        editItem: null,
        createFormKey: 0
      });
    } catch (err) {
      console.error('[AuditMaster] Initialization failed:', err);
      setState(prev => ({
        ...prev,
        phase: AppPhase.Error,
        errorMessage: `Initialization failed: ${err instanceof Error ? err.message : 'Unknown error'}`
      }));

    }
  };

  // ── Navigation handlers ───────────────────────────────────────────────────
  const handleMenuSelect = (key: string): void => {
    setState(prev => ({
      ...prev,
      selectedMenu: key,
      editItem: null,
      createFormKey: key === 'create' ? prev.createFormKey + 1 : prev.createFormKey
    }));
  };

  const handleEditItem = (item: IAuditMasterItem): void => {
    setState(prev => ({ ...prev, selectedMenu: 'edit', editItem: item }));
  };

  const handleSaved = (): void => {
    // Go back to the appropriate list view
    const defaultMenu = state.menuItems.length > 0 ? state.menuItems[0].key : 'all';
    const listMenu = state.menuItems.find(m => m.key !== 'create')?.key || defaultMenu;
    setState(prev => ({ ...prev, selectedMenu: listMenu, editItem: null }));
  };

  const handleCancel = (): void => {
    const defaultMenu = state.menuItems.length > 0 ? state.menuItems[0].key : 'all';
    const listMenu = state.menuItems.find(m => m.key !== 'create')?.key || defaultMenu;
    setState(prev => ({ ...prev, selectedMenu: listMenu, editItem: null }));
  };

  // ── Render content area based on selected menu ────────────────────────────
  const renderContent = (): React.ReactElement => {
    if (!state.currentUser) return <div />;

    // Edit mode
    if (state.selectedMenu === 'edit' && state.editItem) {
      return (
        <AuditItemForm
          spService={spService}
          roleService={roleService}
          currentUser={state.currentUser}
          editItem={state.editItem}
          onSaved={handleSaved}
          onCancel={handleCancel}
        />
      );
    }

    switch (state.selectedMenu) {
      case 'create':
        return (
          <AuditItemForm
            key={`create-${state.createFormKey}`}
            spService={spService}
            roleService={roleService}
            currentUser={state.currentUser}
            onSaved={handleSaved}
            onCancel={handleCancel}
          />
        );

      case 'all':
        return (
          <AuditListView
            spService={spService}
            roleService={roleService}
            currentUser={state.currentUser}
            viewMode="all"
            onEditItem={handleEditItem}
          />
        );

      case 'myAssigned':
        return (
          <AuditListView
            spService={spService}
            roleService={roleService}
            currentUser={state.currentUser}
            viewMode="myAssigned"
            onEditItem={handleEditItem}
          />
        );

      case 'myPending':
        return (
          <AuditListView
            spService={spService}
            roleService={roleService}
            currentUser={state.currentUser}
            viewMode="myPending"
            onEditItem={handleEditItem}
          />
        );

      default:
        return (
          <div style={{ padding: 40, textAlign: 'center', color: '#605e5c' }}>
            <Icon iconName="Shield" style={{ fontSize: 64, color: '#4facfe', marginBottom: 16 }} />
            <h3 style={{ color: '#1a1a2e' }}>Welcome to Audit Master</h3>
            <p>Select an option from the sidebar to get started.</p>
          </div>
        );
    }
  };

  // ── Phase: Initializing / Provisioning ────────────────────────────────────
  if (state.phase === AppPhase.Initializing || state.phase === AppPhase.Provisioning) {
    return (
      <div className={styles.initOverlay}>
        <Icon iconName="Shield" style={{ fontSize: 56, color: '#4facfe' }} />
        <div className={styles.initTitle}>Audit Master</div>
        <Spinner size={SpinnerSize.large} />
        <div className={styles.initSubtitle}>{state.statusMessage}</div>
      </div>
    );
  }

  // ── Phase: Error ──────────────────────────────────────────────────────────
  if (state.phase === AppPhase.Error) {
    return (
      <div className={styles.errorContainer}>
        <Icon iconName="ErrorBadge" className={styles.errorIcon} />
        <div className={styles.errorText}>{state.errorMessage}</div>
        <PrimaryButton text="Retry" onClick={initialize} />
      </div>
    );
  }

  // ── Phase: Ready ──────────────────────────────────────────────────────────
  return (
    <div className={styles.auditMaster}>
      <Sidebar
        currentUser={state.currentUser!}
        menuItems={state.menuItems}
        selectedMenu={state.selectedMenu}
        onMenuSelect={handleMenuSelect}
      />
      <div className={styles.mainContent}>
        {renderContent()}
      </div>
    </div>
  );
};

export default AuditMaster;
