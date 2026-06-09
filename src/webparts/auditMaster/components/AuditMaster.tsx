import * as React from 'react';
import { Spinner, SpinnerSize, Icon, PrimaryButton } from '@fluentui/react';
import styles from './AuditMaster.module.scss';
import { IAuditMasterProps, ICurrentUser, IMenuItem, UserRole } from '../models';
import { SharePointService, ListProvisioningService, RoleService } from '../services';
import { Sidebar } from './Sidebar';
import {
  AdminView,
  PICView,
  ManagerView,
  VerifierView,
  ViewerView,
  IRoleViewProps
} from './Views';

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
}

/**
 * Top-level container component for the Audit Master webpart.
 *
 * Lifecycle:
 *  1. Provision lists (create if missing)
 *  2. Resolve current user + role (via RoleService)
 *  3. Render sidebar + the role-specific view component
 *
 * ────────────────────────────────────────────────────────────────────────────
 *  ROLE-BASED VIEW ROUTING
 * ────────────────────────────────────────────────────────────────────────────
 *  Rather than a single form that shows/hides sections, the app routes each
 *  user to a dedicated, self-contained view component based on their resolved
 *  role. Every view internally handles its own list ⇄ form navigation and
 *  reuses the shared SharePointService / RoleService instances.
 *
 *    Role            → View
 *    ----------------------------------
 *    Admin           → AdminView      (full create + edit, all items)
 *    PIC             → PICView        (my assigned audits, responses)
 *    Quality Manager → ManagerView    (oversight / review, all items)
 *    Verifier        → VerifierView   (verification & approval queue)
 *    Auditor/Viewer/ → ViewerView     (read-only fallback)
 *    unknown
 * ────────────────────────────────────────────────────────────────────────────
 */
const AuditMaster: React.FC<IAuditMasterProps> = (props) => {
  const { context, siteUrl } = props;

  // Services (stable refs) – created once and shared across every view.
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
    selectedMenu: ''
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

      // Phase 2: Resolve user + role, then build the role-specific menu.
      setState(prev => ({ ...prev, statusMessage: 'Resolving user permissions...' }));
      const user = await roleService.resolveCurrentUser();
      const menuItems = roleService.getMenuItems(user);
      const defaultMenu = menuItems.length > 0 ? menuItems[0].key : '';

      setState({
        phase: AppPhase.Ready,
        statusMessage: '',
        errorMessage: '',
        currentUser: user,
        menuItems,
        selectedMenu: defaultMenu
      });
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Failed to initialize audit item.';
      console.error('[AuditMaster] Initialization failed:', err);
      setState(prev => ({
        ...prev,
        phase: AppPhase.Error,
        errorMessage: `Initialization failed: ${errorMessage}`
      }));
    }
  };

  // ── Navigation ──────────────────────────────────────────────────────────--
  // AuditMaster only tracks which sidebar menu is selected. Each role view
  // interprets that selection (e.g. 'create' → open form) and manages its own
  // internal list ⇄ form navigation.
  const handleMenuSelect = (key: string): void => {
    setState(prev => ({ ...prev, selectedMenu: key }));
  };

  // ── Role → View routing ─────────────────────────────────────────────────--
  const renderRoleView = (): React.ReactElement => {
    const user = state.currentUser!;

    // Props shared by every role view.
    const viewProps: IRoleViewProps = {
      spService,
      roleService,
      currentUser: user,
      selectedMenu: state.selectedMenu
    };

    // Route to the dedicated view for the resolved role. Anything that is not
    // an explicit Admin/PIC/Quality Manager/Verifier falls back to the
    // read-only ViewerView (Auditor, Viewer, or no role record).
    switch (user.role) {
      case UserRole.Admin:
        return <AdminView {...viewProps} />;
      case UserRole.PIC:
        return <PICView {...viewProps} />;
      case UserRole.Verifier:
        return <VerifierView {...viewProps} />;
      default:
        return <ViewerView {...viewProps} />;
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
        {renderRoleView()}
      </div>
    </div>
  );
};

export default AuditMaster;
