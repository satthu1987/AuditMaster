import * as React from 'react';
import {
  SearchBox,
  Dropdown,
  IDropdownOption,
  Spinner,
  SpinnerSize,
  Icon,
  IconButton,
  MessageBar,
  MessageBarType
} from '@fluentui/react';
import styles from './AuditListView.module.scss';
import {
  IAuditMasterItem,
  ICurrentUser,
  AuditStatus
} from '../../models';
import { SharePointService, RoleService } from '../../services';

export interface IAuditListViewProps {
  spService: SharePointService;
  roleService: RoleService;
  currentUser: ICurrentUser;
  viewMode: 'all' | 'myAssigned' | 'myPending';
  onEditItem: (item: IAuditMasterItem) => void;
}

interface IViewState {
  items: IAuditMasterItem[];
  filteredItems: IAuditMasterItem[];
  isLoading: boolean;
  error: string;
  searchText: string;
  statusFilter: string;
  priorityFilter: string;
  currentPage: number;
}

const PAGE_SIZE = 15;

/**
 * Audit item list view – displays a filtered, searchable table of audit items.
 * Supports three view modes:
 *  • all         – Admin view: shows all items
 *  • myAssigned  – PIC view: shows items assigned to current user
 *  • myPending   – Verifier view: shows pending verification items for user's segments
 */
const AuditListView: React.FC<IAuditListViewProps> = (props) => {
  const { spService, currentUser, viewMode, onEditItem } = props;

  const [state, setState] = React.useState<IViewState>({
    items: [],
    filteredItems: [],
    isLoading: true,
    error: '',
    searchText: '',
    statusFilter: '',
    priorityFilter: '',
    currentPage: 1
  });

  // ── Load data ─────────────────────────────────────────────────────────────
  React.useEffect(() => {
    loadItems();
  }, [viewMode]);

  const loadItems = async (): Promise<void> => {
    setState(prev => ({ ...prev, isLoading: true, error: '' }));
    try {
      let items: IAuditMasterItem[];
      switch (viewMode) {
        case 'myAssigned':
          items = await spService.getAuditItemsByPIC(currentUser.id);
          break;
        case 'myPending':
          items = await spService.getAuditItemsForVerifier(currentUser.segmentServiceIds || []);
          break;
        default:
          items = await spService.getAllAuditItems();
          break;
      }
      setState(prev => ({
        ...prev,
        items,
        filteredItems: items,
        isLoading: false,
        currentPage: 1
      }));
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Failed to load audit items.';
      setState(prev => ({ ...prev, isLoading: false, error: errorMessage }));
    }
  };

  // ── Filtering ─────────────────────────────────────────────────────────────
  React.useEffect(() => {
    let result = [...state.items];
    if (state.searchText) {
      const q = state.searchText.toLowerCase();
      result = result.filter(i =>
        (i.Title && i.Title.toLowerCase().indexOf(q) >= 0) ||
        (i.Segment && i.Segment.toLowerCase().indexOf(q) >= 0)
      );
    }
    if (state.statusFilter) {
      result = result.filter(i => i.Status === state.statusFilter);
    }
    setState(prev => ({ ...prev, filteredItems: result, currentPage: 1 }));
  }, [state.searchText, state.statusFilter, state.priorityFilter, state.items]);

  // ── Pagination ────────────────────────────────────────────────────────────
  const totalPages = Math.ceil(state.filteredItems.length / PAGE_SIZE);
  const pageItems = state.filteredItems.slice(
    (state.currentPage - 1) * PAGE_SIZE,
    state.currentPage * PAGE_SIZE
  );

  // ── Stats ─────────────────────────────────────────────────────────────────
  const countByStatus = (status: string): number =>
    state.items.filter(i => i.Status === status).length;

  // ── Helpers ───────────────────────────────────────────────────────────────
  const statusClass = (status: string): string => {
    switch (status) {
      case AuditStatus.Open: return styles.open;
      case AuditStatus.InProgress: return styles.inProgress;
      case AuditStatus.PendingVerification: return styles.pendingVerification;
      case AuditStatus.Closed: return styles.closed;
      case AuditStatus.Overdue: return styles.overdue;
      case AuditStatus.Cancelled: return styles.cancelled;
      default: return '';
    }
  };



  const viewTitle = (): string => {
    switch (viewMode) {
      case 'myAssigned': return 'My Assigned Audits';
      case 'myPending': return 'My Pending Audits';
      default: return 'All Audit Items';
    }
  };

  const formatDate = (dateStr: string): string => {
    if (!dateStr) return '–';
    try {
      return new Date(dateStr).toLocaleDateString('en-GB', {
        day: '2-digit', month: 'short', year: 'numeric'
      });
    } catch {
      return dateStr;
    }
  };

  // ── Render ────────────────────────────────────────────────────────────────

  if (state.isLoading) {
    return (
      <div className={styles.listViewContainer}>
        <Spinner size={SpinnerSize.large} label="Loading audit items..." />
      </div>
    );
  }

  return (
    <div className={styles.listViewContainer}>
      {/* ── Header ─────────────────────────────────────────────────────── */}
      <div className={styles.viewHeader}>
        <h3>{viewTitle()}</h3>
        <IconButton
          iconProps={{ iconName: 'Refresh' }}
          title="Refresh"
          onClick={loadItems}
        />
      </div>

      {state.error && (
        <MessageBar messageBarType={MessageBarType.error}>{state.error}</MessageBar>
      )}

      {/* ── Stats Bar ──────────────────────────────────────────────────── */}
      <div className={styles.statsBar}>
        <div className={`${styles.statCard} ${styles.open}`}>
          <div className={styles.statValue}>{countByStatus(AuditStatus.Open)}</div>
          <div className={styles.statLabel}>Open</div>
        </div>
        <div className={`${styles.statCard} ${styles.inProgress}`}>
          <div className={styles.statValue}>{countByStatus(AuditStatus.InProgress)}</div>
          <div className={styles.statLabel}>In Progress</div>
        </div>
        <div className={`${styles.statCard} ${styles.pending}`}>
          <div className={styles.statValue}>{countByStatus(AuditStatus.PendingVerification)}</div>
          <div className={styles.statLabel}>Pending</div>
        </div>
        <div className={`${styles.statCard} ${styles.closed}`}>
          <div className={styles.statValue}>{countByStatus(AuditStatus.Closed)}</div>
          <div className={styles.statLabel}>Closed</div>
        </div>
        <div className={`${styles.statCard} ${styles.overdue}`}>
          <div className={styles.statValue}>{countByStatus(AuditStatus.Overdue)}</div>
          <div className={styles.statLabel}>Overdue</div>
        </div>
      </div>

      {/* ── Filter Bar ─────────────────────────────────────────────────── */}
      <div className={styles.filterBar}>
        <div className={styles.searchBox}>
          <SearchBox
            placeholder="Search by title, ID, finding, or department..."
            value={state.searchText}
            onChange={(_, val) => setState(prev => ({ ...prev, searchText: val || '' }))}
            onClear={() => setState(prev => ({ ...prev, searchText: '' }))}
          />
        </div>
        <Dropdown
          placeholder="Filter by Status"
          selectedKey={state.statusFilter || undefined}
          options={[
            { key: '', text: 'All Statuses' },
            ...Object.values(AuditStatus).map(s => ({ key: s as string, text: s as string }))
          ]}
          onChange={(_, opt) => setState(prev => ({
            ...prev,
            statusFilter: (opt?.key as string) || ''
          }))}
          styles={{ dropdown: { width: 180 } }}
        />

      </div>

      {/* ── Table ──────────────────────────────────────────────────────── */}
      {state.filteredItems.length === 0 ? (
        <div className={styles.emptyState}>
          <div className={styles.emptyIcon}>
            <Icon iconName="SearchIssue" />
          </div>
          <h4>No Audit Items Found</h4>
          <p>
            {state.items.length === 0
              ? 'There are no audit items to display.'
              : 'Try adjusting your search or filters.'}
          </p>
        </div>
      ) : (
        <div className={styles.tableWrapper}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th>Audit ID</th>
                <th>Title</th>
                <th>Status</th>
                <th>PIC</th>
                <th>Service</th>
                <th>Segment</th>
                <th>Audit Date</th>
                <th>Due Date</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {pageItems.map(item => (
                <tr key={item.AuditId} onClick={() => onEditItem(item)}>
                  <td><strong>{item.AuditId}</strong></td>
                  <td title={item.Title}>{item.Title}</td>
                  <td>
                    <span className={`${styles.statusBadge} ${statusClass(item.Status)}`}>
                      {item.Status}
                    </span>
                  </td>
                  <td>{item.PIC ? item.PIC.Title : '–'}</td>
                  <td>{item.Service || '–'}</td>
                  <td>{item.Segment || '–'}</td>
                  <td>{formatDate(item.AuditDate)}</td>
                  <td>{formatDate(item.DueDate)}</td>
                  <td>
                    <IconButton
                      iconProps={{ iconName: 'Edit' }}
                      title="View / Edit"
                      onClick={(e) => {
                        e.stopPropagation();
                        onEditItem(item);
                      }}
                    />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>

          {/* ── Pagination ──────────────────────────────────────────────── */}
          <div className={styles.pagination}>
            <span>
              Showing {(state.currentPage - 1) * PAGE_SIZE + 1}–
              {Math.min(state.currentPage * PAGE_SIZE, state.filteredItems.length)} of{' '}
              {state.filteredItems.length} items
            </span>
            <div>
              <IconButton
                iconProps={{ iconName: 'ChevronLeft' }}
                disabled={state.currentPage <= 1}
                onClick={() => setState(prev => ({ ...prev, currentPage: prev.currentPage - 1 }))}
              />
              <span style={{ margin: '0 8px' }}>
                Page {state.currentPage} of {totalPages}
              </span>
              <IconButton
                iconProps={{ iconName: 'ChevronRight' }}
                disabled={state.currentPage >= totalPages}
                onClick={() => setState(prev => ({ ...prev, currentPage: prev.currentPage + 1 }))}
              />
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default AuditListView;
