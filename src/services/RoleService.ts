/**
 * Role-based access – My requests vs Approver / Admin views
 * Uses SharePoint groups when available; falls back to email match.
 */

import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/site-users/web';
import '@pnp/sp/site-groups/web';

export type AppRole = 'Requester' | 'Approver' | 'Admin';

export interface ICurrentUserContext {
  displayName: string;
  email: string;
  loginName?: string;
  roles: AppRole[];
  isApprover: boolean;
  isAdmin: boolean;
}

/** Group titles that grant Approver / Admin (customize per tenant) */
const APPROVER_GROUPS = ['Admin Forms Approvers', 'Approvers', 'Managers'];
const ADMIN_GROUPS = ['Admin Forms Admins', 'Site Owners', 'Owners'];

export class RoleService {
  private static cached: ICurrentUserContext | null = null;

  public static async getCurrentUser(
    displayName?: string,
    email?: string
  ): Promise<ICurrentUserContext> {
    if (this.cached && (!email || this.cached.email === email)) {
      return this.cached;
    }

    let roles: AppRole[] = ['Requester'];
    let resolvedName = displayName || 'User';
    let resolvedEmail = email || '';
    let loginName: string | undefined;

    try {
      const me = await sp.web.currentUser.select('Id', 'Title', 'Email', 'LoginName').get();
      resolvedName = me.Title || resolvedName;
      resolvedEmail = me.Email || resolvedEmail;
      loginName = me.LoginName;

      const groups = await sp.web.currentUser.groups.select('Title').get();
      const titles = (groups || []).map((g: any) => (g.Title || '').toLowerCase());

      if (titles.some((t: string) => ADMIN_GROUPS.some((a) => t.includes(a.toLowerCase())))) {
        roles = ['Requester', 'Approver', 'Admin'];
      } else if (
        titles.some((t: string) => APPROVER_GROUPS.some((a) => t.includes(a.toLowerCase())))
      ) {
        roles = ['Requester', 'Approver'];
      }
    } catch (e) {
      console.warn('RoleService: could not resolve groups, using defaults', e);
      if (displayName) resolvedName = displayName;
      if (email) resolvedEmail = email;
    }

    this.cached = {
      displayName: resolvedName,
      email: resolvedEmail,
      loginName,
      roles,
      isApprover: roles.includes('Approver') || roles.includes('Admin'),
      isAdmin: roles.includes('Admin')
    };
    return this.cached;
  }

  public static clearCache(): void {
    this.cached = null;
  }

  /** Filter helper: requesters only see own; approvers/admins see all */
  public static scopeFilter(
    user: ICurrentUserContext,
    viewMode: 'my' | 'all' | 'pendingApproval'
  ): { requesterEmail?: string; statusPendingOnly?: boolean } {
    if (viewMode === 'my' || !user.isApprover) {
      return { requesterEmail: user.email || undefined };
    }
    if (viewMode === 'pendingApproval') {
      return { statusPendingOnly: true };
    }
    return {};
  }
}
