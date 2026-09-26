/**
 * Audit / history – surfaces SharePoint item versions when versioning is enabled.
 */

import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';

export interface IAuditEntry {
  versionLabel: string;
  modified: Date;
  editor?: string;
  status?: string;
  priority?: string;
  title?: string;
  changesSummary: string;
}

export class AuditService {
  private static readonly LIST = 'Admin Requests';

  /**
   * Requires list versioning enabled on Admin Requests.
   * Falls back to empty array if versions cannot be read.
   */
  public static async getRequestHistory(requestId: string): Promise<IAuditEntry[]> {
    try {
      const versions = await sp.web.lists
        .getByTitle(this.LIST)
        .items.getById(parseInt(requestId, 10))
        .versions.select(
          'VersionLabel',
          'Modified',
          'Editor',
          'Status',
          'Priority',
          'Title'
        )
        .get();

      if (!Array.isArray(versions)) {
        return [];
      }

      const entries: IAuditEntry[] = versions.map((v: any, index: number, arr: any[]) => {
        const prev = arr[index + 1];
        const parts: string[] = [];
        if (prev) {
          if (v.Status !== prev.Status) parts.push(`Status: ${prev.Status} → ${v.Status}`);
          if (v.Priority !== prev.Priority) parts.push(`Priority: ${prev.Priority} → ${v.Priority}`);
          if (v.Title !== prev.Title) parts.push('Title updated');
        } else {
          parts.push('Created / earliest version');
        }
        return {
          versionLabel: v.VersionLabel || String(v.VersionId || index),
          modified: v.Modified ? new Date(v.Modified) : new Date(),
          editor: v.Editor?.Title || v.Editor || undefined,
          status: v.Status,
          priority: v.Priority,
          title: v.Title,
          changesSummary: parts.join('; ') || 'Updated'
        };
      });

      return entries;
    } catch (e) {
      console.warn(
        'AuditService: versions unavailable (enable versioning on Admin Requests list)',
        e
      );
      return [];
    }
  }
}
