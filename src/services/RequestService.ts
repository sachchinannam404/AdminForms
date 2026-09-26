/**
 * High-level service for Admin Requests and child items across all request types
 */

import { SharePointRepository } from './SharePointRepository';
import {
  IAdminRequest,
  IAdminRequestFilter,
  RequestStatus,
  PriorityLevel,
  RequestType
} from '../models/IAdminRequest';
import { IChildItem } from '../models/IChildItem';
import { ItemStatus } from '../models/common/enums';
import { getRequestTypeConfig } from '../config/requestTypeRegistry';
import { NotificationService } from './NotificationService';

const PARENT_SELECT = [
  'ID',
  'Title',
  'Description',
  'RequestType',
  'RequesterName',
  'RequesterEmail',
  'Department',
  'Status',
  'Priority',
  'TargetDeliveryDate',
  'ApprovedBy',
  'ApprovedDate',
  'RejectionReason',
  'TotalBudget',
  'Comments',
  'DetailsJson',
  'Created',
  'Modified'
];

const CHILD_SELECT = [
  'ID',
  'RequestId',
  'ItemName',
  'Category',
  'Quantity',
  'Unit',
  'UnitPrice',
  'TotalPrice',
  'Description',
  'Status',
  'VendorName',
  'VendorEmail',
  'ExpectedDeliveryDate',
  'ActualDeliveryDate',
  'Remarks',
  'ExtraJson',
  'Created',
  'Modified'
];

function mapAdminRequest(item: any): IAdminRequest {
  let details: Record<string, any> | undefined;
  if (item.DetailsJson) {
    try {
      details =
        typeof item.DetailsJson === 'string' ? JSON.parse(item.DetailsJson) : item.DetailsJson;
    } catch {
      details = undefined;
    }
  }
  return {
    id: item.ID?.toString(),
    title: item.Title || '',
    description: item.Description || '',
    requestType: (item.RequestType as RequestType) || RequestType.General,
    requesterName: item.RequesterName,
    requesterEmail: item.RequesterEmail,
    department: item.Department,
    status: (item.Status as RequestStatus) || RequestStatus.Pending,
    priority: (item.Priority as PriorityLevel) || PriorityLevel.Medium,
    targetDeliveryDate: item.TargetDeliveryDate ? new Date(item.TargetDeliveryDate) : undefined,
    approvedBy: item.ApprovedBy,
    approvedDate: item.ApprovedDate ? new Date(item.ApprovedDate) : undefined,
    rejectionReason: item.RejectionReason,
    totalBudget: item.TotalBudget,
    comments: item.Comments,
    details,
    created: item.Created ? new Date(item.Created) : undefined,
    modified: item.Modified ? new Date(item.Modified) : undefined
  };
}

function mapToSharePointRequest(entity: Partial<IAdminRequest>): Record<string, any> {
  const data: Record<string, any> = {};
  if (entity.title !== undefined) data.Title = entity.title;
  if (entity.description !== undefined) data.Description = entity.description;
  if (entity.requestType !== undefined) data.RequestType = entity.requestType;
  if (entity.requesterName !== undefined) data.RequesterName = entity.requesterName;
  if (entity.requesterEmail !== undefined) data.RequesterEmail = entity.requesterEmail;
  if (entity.department !== undefined) data.Department = entity.department;
  if (entity.status !== undefined) data.Status = entity.status;
  if (entity.priority !== undefined) data.Priority = entity.priority;
  if (entity.targetDeliveryDate !== undefined) data.TargetDeliveryDate = entity.targetDeliveryDate;
  if (entity.approvedBy !== undefined) data.ApprovedBy = entity.approvedBy;
  if (entity.approvedDate !== undefined) data.ApprovedDate = entity.approvedDate;
  if (entity.rejectionReason !== undefined) data.RejectionReason = entity.rejectionReason;
  if (entity.totalBudget !== undefined) data.TotalBudget = entity.totalBudget;
  if (entity.comments !== undefined) data.Comments = entity.comments;
  if (entity.details !== undefined) data.DetailsJson = JSON.stringify(entity.details || {});
  return data;
}

function mapChildItem(item: any): IChildItem {
  let extra: Record<string, any> | undefined;
  if (item.ExtraJson) {
    try {
      extra = typeof item.ExtraJson === 'string' ? JSON.parse(item.ExtraJson) : item.ExtraJson;
    } catch {
      extra = undefined;
    }
  }
  // Promote known extra keys from form (serial, warranty, actualAmount) stored in ExtraJson
  return {
    id: item.ID?.toString(),
    requestId: item.RequestId,
    itemName: item.ItemName || '',
    category: item.Category,
    quantity: item.Quantity,
    unit: item.Unit || 'pcs',
    unitPrice: item.UnitPrice,
    totalPrice: item.TotalPrice,
    description: item.Description,
    status: item.Status || ItemStatus.Pending,
    vendorName: item.VendorName,
    vendorEmail: item.VendorEmail,
    expectedDeliveryDate: item.ExpectedDeliveryDate
      ? new Date(item.ExpectedDeliveryDate)
      : undefined,
    actualDeliveryDate: item.ActualDeliveryDate ? new Date(item.ActualDeliveryDate) : undefined,
    remarks: item.Remarks,
    extra,
    created: item.Created ? new Date(item.Created) : undefined,
    modified: item.Modified ? new Date(item.Modified) : undefined
  };
}

function mapToSharePointChild(entity: Partial<IChildItem> & Record<string, any>): Record<string, any> {
  const data: Record<string, any> = {};
  if (entity.requestId !== undefined) data.RequestId = entity.requestId;
  if (entity.itemName !== undefined) data.ItemName = entity.itemName;
  if (entity.category !== undefined) data.Category = entity.category;
  if (entity.quantity !== undefined) data.Quantity = entity.quantity;
  if (entity.unit !== undefined) data.Unit = entity.unit;
  if (entity.unitPrice !== undefined) data.UnitPrice = entity.unitPrice;
  if (entity.totalPrice !== undefined) data.TotalPrice = entity.totalPrice;
  if (entity.description !== undefined) data.Description = entity.description;
  if (entity.status !== undefined) data.Status = entity.status;
  if (entity.vendorName !== undefined) data.VendorName = entity.vendorName;
  if (entity.vendorEmail !== undefined) data.VendorEmail = entity.vendorEmail;
  if (entity.expectedDeliveryDate !== undefined)
    data.ExpectedDeliveryDate = entity.expectedDeliveryDate;
  if (entity.actualDeliveryDate !== undefined) data.ActualDeliveryDate = entity.actualDeliveryDate;
  if (entity.remarks !== undefined) data.Remarks = entity.remarks;

  // Fold non-standard form fields into ExtraJson
  const known =
    'requestId|itemName|category|quantity|unit|unitPrice|totalPrice|description|status|vendorName|vendorEmail|expectedDeliveryDate|actualDeliveryDate|remarks|extra|id|attachments|created|modified'.split(
      '|'
    );
  const extra: Record<string, any> = { ...(entity.extra || {}) };
  Object.keys(entity).forEach((k) => {
    if (!known.includes(k) && entity[k] !== undefined) {
      extra[k] = entity[k];
    }
  });
  if (Object.keys(extra).length) {
    data.ExtraJson = JSON.stringify(extra);
  } else if (entity.extra !== undefined) {
    data.ExtraJson = JSON.stringify(entity.extra || {});
  }
  return data;
}

export class RequestService {
  private static parentRepo = new SharePointRepository<IAdminRequest>(
    'Admin Requests',
    mapAdminRequest,
    mapToSharePointRequest,
    PARENT_SELECT
  );

  private static childRepos = new Map<string, SharePointRepository<IChildItem>>();

  public static initialize(context: any): void {
    SharePointRepository.initialize(context);
  }

  private static getChildRepo(listTitle: string): SharePointRepository<IChildItem> {
    if (!this.childRepos.has(listTitle)) {
      this.childRepos.set(
        listTitle,
        new SharePointRepository<IChildItem>(
          listTitle,
          mapChildItem,
          mapToSharePointChild,
          CHILD_SELECT
        )
      );
    }
    return this.childRepos.get(listTitle)!;
  }

  public static async createRequest(request: Partial<IAdminRequest>): Promise<IAdminRequest> {
    const config = getRequestTypeConfig(request.requestType || RequestType.General);
    const payload: Partial<IAdminRequest> = {
      ...request,
      status: request.status || config.defaultStatus || RequestStatus.Pending,
      priority: request.priority || config.defaultPriority || PriorityLevel.Medium,
      requestType: request.requestType || RequestType.General
    };
    const result = await this.parentRepo.create(payload);
    return {
      ...payload,
      id: result.ID?.toString(),
      title: payload.title || '',
      description: payload.description || '',
      status: payload.status!,
      priority: payload.priority!,
      requestType: payload.requestType!
    } as IAdminRequest;
  }

  public static async getRequest(id: string): Promise<IAdminRequest> {
    return this.parentRepo.getById(id);
  }

  public static async updateRequest(id: string, request: Partial<IAdminRequest>): Promise<void> {
    await this.parentRepo.update(id, request);
  }

  public static async deleteRequest(id: string): Promise<void> {
    await this.parentRepo.delete(id);
  }

  public static async getRequests(filter?: IAdminRequestFilter): Promise<IAdminRequest[]> {
    const parts: string[] = [];
    if (filter?.status) {
      const statuses = Array.isArray(filter.status) ? filter.status : [filter.status];
      if (statuses.length === 1) {
        parts.push(`Status eq '${statuses[0]}'`);
      } else if (statuses.length > 1) {
        parts.push(`(${statuses.map((s) => `Status eq '${s}'`).join(' or ')})`);
      }
    }
    if (filter?.requestType) {
      const types = Array.isArray(filter.requestType) ? filter.requestType : [filter.requestType];
      if (types.length === 1) {
        parts.push(`RequestType eq '${types[0]}'`);
      } else if (types.length > 1) {
        parts.push(`(${types.map((t) => `RequestType eq '${t}'`).join(' or ')})`);
      }
    }
    if (filter?.department) {
      parts.push(`Department eq '${filter.department.replace(/'/g, "''")}'`);
    }
    if (filter?.priority) {
      parts.push(`Priority eq '${filter.priority}'`);
    }
    if (filter?.requesterEmail) {
      parts.push(`RequesterEmail eq '${filter.requesterEmail.replace(/'/g, "''")}'`);
    }
    const odataFilter = parts.length ? parts.join(' and ') : undefined;
    let items = await this.parentRepo.getAll(odataFilter, 'Created', false, 500);

    if (filter?.searchText) {
      const q = filter.searchText.toLowerCase();
      items = items.filter(
        (r) =>
          (r.title && r.title.toLowerCase().includes(q)) ||
          (r.description && r.description.toLowerCase().includes(q)) ||
          (r.requesterName && r.requesterName.toLowerCase().includes(q)) ||
          (r.department && r.department.toLowerCase().includes(q))
      );
    }
    if (filter?.fromDate) {
      const from = filter.fromDate.getTime();
      items = items.filter((r) => r.created && r.created.getTime() >= from);
    }
    if (filter?.toDate) {
      const to = filter.toDate.getTime();
      items = items.filter((r) => r.created && r.created.getTime() <= to);
    }
    return items;
  }

  public static async approveRequest(
    id: string,
    approvedBy: string,
    comments?: string
  ): Promise<void> {
    const existing = await this.getRequest(id);
    const prev = existing.status;
    await this.parentRepo.update(id, {
      status: RequestStatus.Approved,
      approvedBy,
      approvedDate: new Date(),
      comments: comments || undefined
    });
    await NotificationService.notifyStatusChange(
      { ...existing, status: RequestStatus.Approved, approvedBy, approvedDate: new Date() },
      prev,
      approvedBy
    );
  }

  public static async rejectRequest(
    id: string,
    approvedBy: string,
    rejectionReason: string
  ): Promise<void> {
    const existing = await this.getRequest(id);
    const prev = existing.status;
    await this.parentRepo.update(id, {
      status: RequestStatus.Rejected,
      approvedBy,
      approvedDate: new Date(),
      rejectionReason
    });
    await NotificationService.notifyStatusChange(
      {
        ...existing,
        status: RequestStatus.Rejected,
        approvedBy,
        rejectionReason,
        approvedDate: new Date()
      },
      prev,
      approvedBy
    );
  }

  /** Bulk status update for multi-select dashboard actions */
  public static async bulkUpdateStatus(
    ids: string[],
    status: RequestStatus,
    actorName: string
  ): Promise<{ success: number; failed: number }> {
    let success = 0;
    let failed = 0;
    for (const id of ids) {
      try {
        await this.parentRepo.update(id, { status });
        success++;
      } catch {
        failed++;
      }
    }
    if (success > 0) {
      await NotificationService.notifyBulkStatusChange(ids, status, actorName);
    }
    return { success, failed };
  }

  public static async getChildItems(
    requestType: RequestType,
    requestId: string
  ): Promise<IChildItem[]> {
    const config = getRequestTypeConfig(requestType);
    if (!config.supportsChildren || !config.childListTitle) {
      return [];
    }
    const repo = this.getChildRepo(config.childListTitle);
    return repo.getAll(`RequestId eq '${requestId}'`, 'Created', false);
  }

  public static async createChildItem(
    requestType: RequestType,
    item: Partial<IChildItem> & Record<string, any>
  ): Promise<IChildItem> {
    const config = getRequestTypeConfig(requestType);
    if (!config.childListTitle) {
      throw new Error(`Request type ${requestType} does not support child items`);
    }
    const qty = item.quantity || 1;
    const unitPrice = item.unitPrice || 0;
    const payload: Partial<IChildItem> & Record<string, any> = {
      ...item,
      totalPrice: item.totalPrice !== undefined ? item.totalPrice : qty * unitPrice,
      status: item.status || ItemStatus.Pending,
      unit: item.unit || 'pcs'
    };
    const repo = this.getChildRepo(config.childListTitle);
    const result = await repo.create(payload);
    return {
      ...payload,
      id: result.ID?.toString(),
      requestId: item.requestId || '',
      itemName: item.itemName || ''
    } as IChildItem;
  }

  public static async updateChildItem(
    requestType: RequestType,
    itemId: string,
    item: Partial<IChildItem> & Record<string, any>
  ): Promise<void> {
    const config = getRequestTypeConfig(requestType);
    if (!config.childListTitle) {
      throw new Error(`Request type ${requestType} does not support child items`);
    }
    if (item.quantity !== undefined && item.unitPrice !== undefined) {
      item.totalPrice = item.quantity * item.unitPrice;
    }
    const repo = this.getChildRepo(config.childListTitle);
    await repo.update(itemId, item);
  }

  public static async deleteChildItem(requestType: RequestType, itemId: string): Promise<void> {
    const config = getRequestTypeConfig(requestType);
    if (!config.childListTitle) {
      throw new Error(`Request type ${requestType} does not support child items`);
    }
    const repo = this.getChildRepo(config.childListTitle);
    await repo.delete(itemId);
  }

  public static async getDashboardStats(): Promise<{
    total: number;
    pending: number;
    approved: number;
    rejected: number;
    inProgress: number;
    overdue: number;
    totalBudget: number;
    byType: Record<string, number>;
  }> {
    const all = await this.getRequests();
    const byType: Record<string, number> = {};
    let totalBudget = 0;
    const now = Date.now();
    let overdue = 0;
    all.forEach((r) => {
      byType[r.requestType] = (byType[r.requestType] || 0) + 1;
      totalBudget += r.totalBudget || 0;
      if (
        r.targetDeliveryDate &&
        r.targetDeliveryDate.getTime() < now &&
        r.status !== RequestStatus.Completed &&
        r.status !== RequestStatus.Cancelled &&
        r.status !== RequestStatus.Rejected
      ) {
        overdue++;
      }
    });
    return {
      total: all.length,
      pending: all.filter(
        (r) => r.status === RequestStatus.Pending || r.status === RequestStatus.Draft
      ).length,
      approved: all.filter((r) => r.status === RequestStatus.Approved).length,
      rejected: all.filter((r) => r.status === RequestStatus.Rejected).length,
      inProgress: all.filter((r) => r.status === RequestStatus.InProgress).length,
      overdue,
      totalBudget,
      byType
    };
  }
}
