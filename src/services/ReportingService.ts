/**
 * Reporting – CSV/Excel-friendly export and simple aggregates for charts
 */

import { IAdminRequest, RequestStatus, RequestType } from '../models/IAdminRequest';
import { RequestService } from './RequestService';

export interface IReportSummary {
  total: number;
  byStatus: Record<string, number>;
  byType: Record<string, number>;
  byPriority: Record<string, number>;
  totalBudget: number;
  approvedBudget: number;
  pendingCount: number;
  overdueCount: number;
}

export class ReportingService {
  public static buildSummary(requests: IAdminRequest[]): IReportSummary {
    const byStatus: Record<string, number> = {};
    const byType: Record<string, number> = {};
    const byPriority: Record<string, number> = {};
    let totalBudget = 0;
    let approvedBudget = 0;
    const now = Date.now();

    requests.forEach((r) => {
      byStatus[r.status] = (byStatus[r.status] || 0) + 1;
      byType[r.requestType] = (byType[r.requestType] || 0) + 1;
      byPriority[r.priority] = (byPriority[r.priority] || 0) + 1;
      const b = r.totalBudget || 0;
      totalBudget += b;
      if (r.status === RequestStatus.Approved || r.status === RequestStatus.Completed) {
        approvedBudget += b;
      }
    });

    const pendingCount =
      (byStatus[RequestStatus.Pending] || 0) + (byStatus[RequestStatus.Draft] || 0);
    const overdueCount = requests.filter(
      (r) =>
        r.targetDeliveryDate &&
        r.targetDeliveryDate.getTime() < now &&
        r.status !== RequestStatus.Completed &&
        r.status !== RequestStatus.Cancelled &&
        r.status !== RequestStatus.Rejected
    ).length;

    return {
      total: requests.length,
      byStatus,
      byType,
      byPriority,
      totalBudget,
      approvedBudget,
      pendingCount,
      overdueCount
    };
  }

  public static async getSummary(): Promise<IReportSummary> {
    const all = await RequestService.getRequests();
    return this.buildSummary(all);
  }

  /** Download CSV of current request list */
  public static exportToCsv(requests: IAdminRequest[], fileName = 'admin-requests.csv'): void {
    const headers = [
      'ID',
      'Title',
      'Type',
      'Status',
      'Priority',
      'Requester',
      'Email',
      'Department',
      'Budget',
      'TargetDelivery',
      'Created',
      'ApprovedBy'
    ];
    const rows = requests.map((r) =>
      [
        r.id || '',
        csvEscape(r.title),
        r.requestType,
        r.status,
        r.priority,
        csvEscape(r.requesterName || ''),
        csvEscape(r.requesterEmail || ''),
        csvEscape(r.department || ''),
        r.totalBudget != null ? String(r.totalBudget) : '',
        r.targetDeliveryDate ? r.targetDeliveryDate.toISOString().slice(0, 10) : '',
        r.created ? r.created.toISOString().slice(0, 10) : '',
        csvEscape(r.approvedBy || '')
      ].join(',')
    );
    const csv = [headers.join(','), ...rows].join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    a.click();
    URL.revokeObjectURL(url);
  }

  public static chartDataByStatus(summary: IReportSummary): { key: string; count: number }[] {
    return Object.keys(summary.byStatus).map((key) => ({
      key,
      count: summary.byStatus[key]
    }));
  }

  public static chartDataByType(summary: IReportSummary): { key: string; count: number }[] {
    return Object.keys(summary.byType).map((key) => ({
      key,
      count: summary.byType[key]
    }));
  }
}

function csvEscape(value: string): string {
  if (value.includes(',') || value.includes('"') || value.includes('\n')) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}

// Silence unused import if tree-shaken
void RequestType;
