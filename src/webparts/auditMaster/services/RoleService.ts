import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService } from './SharePointService';
import { ICurrentUser, UserRole, IAuditMasterItem } from '../models';

/**
 * Service that resolves the current user's role from the UserRole list and
 * enforces permission checks for CRUD operations on Audit Master items.
 *
 * Role hierarchy:
 *  • Admin   – full create + update on all items
 *  • PIC     – update only on items where they are assigned as PIC
 *  • Verifier – update only on items belonging to their segment/department
 *  • Viewer  – (fallback) read-only access when no role record is found
 */
export class RoleService {
  private _spService: SharePointService;

  constructor(spService: SharePointService) {
    this._spService = spService;
  }

  // ─── Resolve current user + role ───────────────────────────────────────────

  /**
   * Fetch the current user from SharePoint, look up their role in the
   * UserRole list, and return an enriched ICurrentUser object.
   */
  public async resolveCurrentUser(): Promise<ICurrentUser> {
    const spUser = await this._spService.getCurrentUser();

    // Look up role assignment
    const roleItem = await this._spService.getUserRoleForUser(spUser.Id);

    const currentUser: ICurrentUser = {
      id: spUser.Id,
      loginName: spUser.LoginName,
      displayName: spUser.Title,
      email: spUser.Email,
      role: roleItem ? roleItem.Role : 'Viewer',
      segmentServiceIds: []
    };

    // If Verifier, resolve which segment/services they belong to
    // Convention: the UserRole.Title stores a comma-separated list of
    // SegmentService IDs for verifiers, or we fall back to all segments.
    if (currentUser.role === UserRole.Verifier && roleItem) {
      currentUser.segmentServiceIds = this._parseSegmentIds(roleItem.Title);
    }
    return currentUser;
  }

  // ─── Permission checks ────────────────────────────────────────────────────

  /**
   * Can the current user create new Audit Master items?
   */
  public canCreate(user: ICurrentUser): boolean {
    return user.role === UserRole.Admin;
  }

  /**
   * Can the current user update the given Audit Master item?
   */
  public canUpdate(user: ICurrentUser, item: IAuditMasterItem): boolean {
    switch (user.role) {
      case UserRole.Admin:
        return true;

      case UserRole.PIC:
        // PIC can only update items assigned to them
        return item.PICId === user.id;

      case UserRole.Verifier:
        // Verifier can update items belonging to their segment/department
        if (user.segmentServiceIds && user.segmentServiceIds.length > 0) {
          return user.segmentServiceIds.indexOf(item.ServiceId) !== -1;
        }
        // Fallback: verifier with no explicit segment restriction can verify
        // items in "Pending Verification" status
        return item.Status === 'Verification';

      default:
        return false; // Viewer – read only
    }
  }

  /**
   * Can the current user delete (or cancel) the given item?
   * Only Admins can delete/cancel.
   */
  public canDelete(user: ICurrentUser): boolean {
    return user.role === UserRole.Admin;
  }

  // ─── Menu items per role ───────────────────────────────────────────────────

  /**
   * Return the sidebar menu items appropriate for the user's role.
   */
  public getMenuItems(user: ICurrentUser): Array<{ key: string; label: string; icon: string }> {
    const items: Array<{ key: string; label: string; icon: string }> = [];

    switch (user.role) {
      case UserRole.Admin:
        items.push(
          { key: 'create', label: 'Create New Audit Item', icon: 'Add' },
          { key: 'all', label: 'All Audit Items', icon: 'BulletedList2' }
        );
        break;

      case UserRole.PIC:
        items.push(
          { key: 'myAssigned', label: 'My Assigned Audits', icon: 'Contact' }
        );
        break;

      case UserRole.Verifier:
        items.push(
          { key: 'myPending', label: 'My Pending Audits', icon: 'CheckList' }
        );
        break;

      default: // Viewer
        items.push(
          { key: 'all', label: 'All Audit Items', icon: 'BulletedList2' }
        );
        break;
    }

    return items;
  }

  // ─── Private helpers ───────────────────────────────────────────────────────

  private _parseSegmentIds(title: string): number[] {
    if (!title) return [];
    return title
      .split(',')
      .map(s => parseInt(s.trim(), 10))
      .filter(n => !isNaN(n));
  }
}
