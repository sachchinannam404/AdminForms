/**
 * Legacy SharePoint Service – kept for backward compatibility.
 * New code should use RequestService + SharePointRepository.
 */

import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { IAdminRequest, RequestStatus, PriorityLevel, RequestType } from '../models/IAdminRequest';
import { IStationeryItem, ItemCategory, ItemStatus } from '../models/IStationeryItem';
import { RequestService } from './RequestService';

export class SharePointService {
  private static readonly ADMIN_REQUESTS_LIST = 'Admin Requests';
  private static readonly STATIONERY_ITEMS_LIST = 'Stationery Items';

  public static initialize(context: any): void {
    RequestService.initialize(context);
    sp.setup({ spfxContext: context });
  }

  public static async createAdminRequest(request: IAdminRequest): Promise<any> {
    const created = await RequestService.createRequest({
      ...request,
      requestType: request.requestType || RequestType.Stationery
    });
    return { ID: created.id ? parseInt(created.id, 10) : undefined, ...created };
  }

  public static async getAdminRequest(requestId: string): Promise<IAdminRequest> {
    return RequestService.getRequest(requestId);
  }

  public static async getAllAdminRequests(): Promise<IAdminRequest[]> {
    return RequestService.getRequests();
  }

  public static async updateAdminRequest(
    requestId: string,
    request: Partial<IAdminRequest>
  ): Promise<void> {
    await RequestService.updateRequest(requestId, request);
  }

  public static async deleteAdminRequest(requestId: string): Promise<void> {
    await RequestService.deleteRequest(requestId);
  }

  public static async createStationeryItem(item: IStationeryItem): Promise<any> {
    const created = await RequestService.createChildItem(RequestType.Stationery, {
      requestId: item.requestId,
      itemName: item.itemName,
      category: item.category,
      quantity: item.quantity,
      unit: item.unit,
      unitPrice: item.unitPrice,
      totalPrice: item.totalPrice,
      description: item.description,
      status: item.status,
      vendorName: item.vendorName,
      vendorEmail: item.vendorEmail,
      expectedDeliveryDate: item.expectedDeliveryDate,
      remarks: item.remarks
    });
    return { ID: created.id ? parseInt(created.id, 10) : undefined, ...created };
  }

  public static async getStationeryItemsByRequest(requestId: string): Promise<IStationeryItem[]> {
    const items = await RequestService.getChildItems(RequestType.Stationery, requestId);
    return items.map((i) => ({
      id: i.id,
      requestId: i.requestId,
      itemName: i.itemName,
      category: (i.category as ItemCategory) || ItemCategory.Other,
      quantity: i.quantity || 0,
      unit: i.unit,
      unitPrice: i.unitPrice,
      totalPrice: i.totalPrice,
      description: i.description,
      status: (i.status as ItemStatus) || ItemStatus.Pending,
      vendorName: i.vendorName,
      vendorEmail: i.vendorEmail,
      expectedDeliveryDate: i.expectedDeliveryDate,
      actualDeliveryDate: i.actualDeliveryDate,
      remarks: i.remarks,
      created: i.created,
      modified: i.modified
    }));
  }

  public static async getStationeryItem(itemId: string): Promise<IStationeryItem> {
    // Fallback: use list directly for single item by id
    const item = await sp.web.lists
      .getByTitle(this.STATIONERY_ITEMS_LIST)
      .items.getById(parseInt(itemId, 10))
      .get();
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
      expectedDeliveryDate: item.ExpectedDeliveryDate
        ? new Date(item.ExpectedDeliveryDate)
        : undefined,
      actualDeliveryDate: item.ActualDeliveryDate
        ? new Date(item.ActualDeliveryDate)
        : undefined,
      remarks: item.Remarks,
      created: item.Created ? new Date(item.Created) : undefined,
      modified: item.Modified ? new Date(item.Modified) : undefined
    };
  }

  public static async updateStationeryItem(
    itemId: string,
    item: Partial<IStationeryItem>
  ): Promise<void> {
    await RequestService.updateChildItem(RequestType.Stationery, itemId, item as any);
  }

  public static async deleteStationeryItem(itemId: string): Promise<void> {
    await RequestService.deleteChildItem(RequestType.Stationery, itemId);
  }
}
