/**
 * SharePoint Service - Handles SharePoint operations
 */

import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { IAdminRequest, RequestStatus, PriorityLevel } from '../models/IAdminRequest';
import { IStationeryItem, ItemCategory, ItemStatus } from '../models/IStationeryItem';

export class SharePointService {
  private static readonly ADMIN_REQUESTS_LIST = 'Admin Requests';
  private static readonly STATIONERY_ITEMS_LIST = 'Stationery Items';

  /**
   * Initialize SharePoint context
   */
  public static initialize(context: any): void {
    sp.setup({
      spfxContext: context
    });
  }

  // ============ ADMIN REQUEST OPERATIONS ============

  /**
   * Create new admin request
   */
  public static async createAdminRequest(request: IAdminRequest): Promise<any> {
    try {
      const result = await sp.web.lists
        .getByTitle(this.ADMIN_REQUESTS_LIST)
        .items.add({
          Title: request.title,
          Description: request.description,
          RequesterName: request.requesterName,
          RequesterEmail: request.requesterEmail,
          Department: request.department,
          Status: request.status || RequestStatus.Pending,
          Priority: request.priority || PriorityLevel.Medium,
          TargetDeliveryDate: request.targetDeliveryDate,
          TotalBudget: request.totalBudget,
          Comments: request.comments
        });

      return result.data;
    } catch (error) {
      console.error('Error creating admin request:', error);
      throw error;
    }
  }

  /**
   * Get admin request by ID
   */
  public static async getAdminRequest(requestId: string): Promise<IAdminRequest> {
    try {
      const item = await sp.web.lists
        .getByTitle(this.ADMIN_REQUESTS_LIST)
        .items.getById(parseInt(requestId))
        .select(
          'ID',
          'Title',
          'Description',
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
          'Created',
          'Modified'
        )
        .get();

      return this.mapAdminRequest(item);
    } catch (error) {
      console.error(`Error getting admin request ${requestId}:`, error);
      throw error;
    }
  }

  /**
   * Get all admin requests
   */
  public static async getAllAdminRequests(): Promise<IAdminRequest[]> {
    try {
      const items = await sp.web.lists
        .getByTitle(this.ADMIN_REQUESTS_LIST)
        .items.select(
          'ID',
          'Title',
          'Description',
          'RequesterName',
          'RequesterEmail',
          'Department',
          'Status',
          'Priority',
          'TargetDeliveryDate',
          'ApprovedBy',
          'ApprovedDate',
          'TotalBudget',
          'Created',
          'Modified'
        )
        .orderBy('Created', false)
        .get();

      return items.map((item: any) => this.mapAdminRequest(item));
    } catch (error) {
      console.error('Error getting all admin requests:', error);
      return [];
    }
  }

  /**
   * Update admin request
   */
  public static async updateAdminRequest(
    requestId: string,
    request: Partial<IAdminRequest>
  ): Promise<void> {
    try {
      const updateData: any = {};

      if (request.title) updateData.Title = request.title;
      if (request.description) updateData.Description = request.description;
      if (request.status) updateData.Status = request.status;
      if (request.priority) updateData.Priority = request.priority;
      if (request.targetDeliveryDate) updateData.TargetDeliveryDate = request.targetDeliveryDate;
      if (request.approvedBy) updateData.ApprovedBy = request.approvedBy;
      if (request.approvedDate) updateData.ApprovedDate = request.approvedDate;
      if (request.rejectionReason) updateData.RejectionReason = request.rejectionReason;
      if (request.totalBudget !== undefined) updateData.TotalBudget = request.totalBudget;
      if (request.comments) updateData.Comments = request.comments;

      await sp.web.lists
        .getByTitle(this.ADMIN_REQUESTS_LIST)
        .items.getById(parseInt(requestId))
        .update(updateData);
    } catch (error) {
      console.error(`Error updating admin request ${requestId}:`, error);
      throw error;
    }
  }

  /**
   * Delete admin request
   */
  public static async deleteAdminRequest(requestId: string): Promise<void> {
    try {
      await sp.web.lists
        .getByTitle(this.ADMIN_REQUESTS_LIST)
        .items.getById(parseInt(requestId))
        .delete();
    } catch (error) {
      console.error(`Error deleting admin request ${requestId}:`, error);
      throw error;
    }
  }

  // ============ STATIONERY ITEM OPERATIONS ============

  /**
   * Create stationery item
   */
  public static async createStationeryItem(item: IStationeryItem): Promise<any> {
    try {
      const result = await sp.web.lists
        .getByTitle(this.STATIONERY_ITEMS_LIST)
        .items.add({
          RequestId: item.requestId,
          ItemName: item.itemName,
          Category: item.category,
          Quantity: item.quantity,
          Unit: item.unit || 'pcs',
          UnitPrice: item.unitPrice,
          TotalPrice: item.totalPrice,
          Description: item.description,
          Status: item.status || ItemStatus.Pending,
          VendorName: item.vendorName,
          VendorEmail: item.vendorEmail,
          ExpectedDeliveryDate: item.expectedDeliveryDate,
          Remarks: item.remarks
        });

      return result.data;
    } catch (error) {
      console.error('Error creating stationery item:', error);
      throw error;
    }
  }

  /**
   * Get stationery items by request ID
   */
  public static async getStationeryItemsByRequest(requestId: string): Promise<IStationeryItem[]> {
    try {
      const items = await sp.web.lists
        .getByTitle(this.STATIONERY_ITEMS_LIST)
        .items.filter(`RequestId eq '${requestId}'`)
        .select(
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
          'Created',
          'Modified'
        )
        .orderBy('Created', false)
        .get();

      return items.map((item: any) => this.mapStationeryItem(item));
    } catch (error) {
      console.error(`Error getting stationery items for request ${requestId}:`, error);
      return [];
    }
  }

  /**
   * Get stationery item by ID
   */
  public static async getStationeryItem(itemId: string): Promise<IStationeryItem> {
    try {
      const item = await sp.web.lists
        .getByTitle(this.STATIONERY_ITEMS_LIST)
        .items.getById(parseInt(itemId))
        .select(
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
          'Created',
          'Modified'
        )
        .get();

      return this.mapStationeryItem(item);
    } catch (error) {
      console.error(`Error getting stationery item ${itemId}:`, error);
      throw error;
    }
  }

  /**
   * Update stationery item
   */
  public static async updateStationeryItem(
    itemId: string,
    item: Partial<IStationeryItem>
  ): Promise<void> {
    try {
      const updateData: any = {};

      if (item.itemName) updateData.ItemName = item.itemName;
      if (item.category) updateData.Category = item.category;
      if (item.quantity !== undefined) updateData.Quantity = item.quantity;
      if (item.unit) updateData.Unit = item.unit;
      if (item.unitPrice !== undefined) updateData.UnitPrice = item.unitPrice;
      if (item.totalPrice !== undefined) updateData.TotalPrice = item.totalPrice;
      if (item.description) updateData.Description = item.description;
      if (item.status) updateData.Status = item.status;
      if (item.vendorName) updateData.VendorName = item.vendorName;
      if (item.vendorEmail) updateData.VendorEmail = item.vendorEmail;
      if (item.expectedDeliveryDate) updateData.ExpectedDeliveryDate = item.expectedDeliveryDate;
      if (item.actualDeliveryDate) updateData.ActualDeliveryDate = item.actualDeliveryDate;
      if (item.remarks) updateData.Remarks = item.remarks;

      await sp.web.lists
        .getByTitle(this.STATIONERY_ITEMS_LIST)
        .items.getById(parseInt(itemId))
        .update(updateData);
    } catch (error) {
      console.error(`Error updating stationery item ${itemId}:`, error);
      throw error;
    }
  }

  /**
   * Delete stationery item
   */
  public static async deleteStationeryItem(itemId: string): Promise<void> {
    try {
      await sp.web.lists
        .getByTitle(this.STATIONERY_ITEMS_LIST)
        .items.getById(parseInt(itemId))
        .delete();
    } catch (error) {
      console.error(`Error deleting stationery item ${itemId}:`, error);
      throw error;
    }
  }

  // ============ HELPER METHODS ============

  private static mapAdminRequest(item: any): IAdminRequest {
    return {
      id: item.ID?.toString(),
      title: item.Title,
      description: item.Description,
      requesterName: item.RequesterName,
      requesterEmail: item.RequesterEmail,
      department: item.Department,
      status: item.Status || RequestStatus.Pending,
      priority: item.Priority || PriorityLevel.Medium,
      targetDeliveryDate: item.TargetDeliveryDate ? new Date(item.TargetDeliveryDate) : undefined,
      approvedBy: item.ApprovedBy,
      approvedDate: item.ApprovedDate ? new Date(item.ApprovedDate) : undefined,
      rejectionReason: item.RejectionReason,
      totalBudget: item.TotalBudget,
      comments: item.Comments,
      created: item.Created ? new Date(item.Created) : undefined,
      modified: item.Modified ? new Date(item.Modified) : undefined
    };
  }

  private static mapStationeryItem(item: any): IStationeryItem {
    return {
      id: item.ID?.toString(),
      requestId: item.RequestId,
      itemName: item.ItemName,
      category: item.Category || ItemCategory.Other,
      quantity: item.Quantity,
      unit: item.Unit || 'pcs',
      unitPrice: item.UnitPrice,
      totalPrice: item.TotalPrice,
      description: item.Description,
      status: item.Status || ItemStatus.Pending,
      vendorName: item.VendorName,
      vendorEmail: item.VendorEmail,
      orderDate: item.OrderDate ? new Date(item.OrderDate) : undefined,
      expectedDeliveryDate: item.ExpectedDeliveryDate ? new Date(item.ExpectedDeliveryDate) : undefined,
      actualDeliveryDate: item.ActualDeliveryDate ? new Date(item.ActualDeliveryDate) : undefined,
      remarks: item.Remarks,
      created: item.Created ? new Date(item.Created) : undefined,
      modified: item.Modified ? new Date(item.Modified) : undefined
    };
  }
}
