import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './Sidebar.module.scss';
import { ICurrentUser, IMenuItem } from '../../models';

export interface ISidebarProps {
  currentUser: ICurrentUser;
  menuItems: IMenuItem[];
  selectedMenu: string;
  onMenuSelect: (key: string) => void;
}

/**
 * Collapsible sidebar navigation component.
 *
 * Shows role-specific menu items passed via props, the current user's name
 * and avatar, and a toggle button to collapse/expand the panel.
 */
const Sidebar: React.FC<ISidebarProps> = (props) => {
  const { currentUser, menuItems, selectedMenu, onMenuSelect } = props;
  const [collapsed, setCollapsed] = React.useState<boolean>(false);

  const initials = currentUser.displayName
    ? currentUser.displayName
        .split(' ')
        .map((n) => n[0])
        .join('')
        .substring(0, 2)
        .toUpperCase()
    : '?';

  return (
    <div className={`${styles.sidebar} ${collapsed ? styles.collapsed : styles.expanded}`}>
      {/* ── Header ─────────────────────────────────────────────────────── */}
      <div className={styles.header}>
        {!collapsed && (
          <div className={styles.logo}>
            <Icon iconName="Shield" style={{ fontSize: 22, color: '#4facfe' }} />
            <h3>Audit Master</h3>
          </div>
        )}
        <button
          className={styles.toggleButton}
          onClick={() => setCollapsed(!collapsed)}
          title={collapsed ? 'Expand sidebar' : 'Collapse sidebar'}
        >
          <Icon iconName={collapsed ? 'DoubleChevronRight' : 'DoubleChevronLeft'} />
        </button>
      </div>

      {/* ── Role Badge ─────────────────────────────────────────────────── */}
      {!collapsed && (
        <div className={styles.roleTag}>{currentUser.role}</div>
      )}

      {/* ── Menu Items ─────────────────────────────────────────────────── */}
      <ul className={styles.menuList}>
        {menuItems.map((item) => (
          <li
            key={item.key}
            className={`${styles.menuItem} ${selectedMenu === item.key ? styles.active : ''}`}
            onClick={() => onMenuSelect(item.key)}
            title={item.label}
          >
            <span className={styles.menuIcon}>
              <Icon iconName={item.icon} />
            </span>
            {!collapsed && <span className={styles.menuLabel}>{item.label}</span>}
          </li>
        ))}
      </ul>

      {/* ── User Info ──────────────────────────────────────────────────── */}
      <div className={styles.userInfo}>
        <div className={styles.userAvatar}>{initials}</div>
        {!collapsed && (
          <div className={styles.userName}>
            <div>{currentUser.displayName}</div>
            <div className={styles.email}>{currentUser.email}</div>
          </div>
        )}
      </div>
    </div>
  );
};

export default Sidebar;
