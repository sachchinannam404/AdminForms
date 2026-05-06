/**
 * Attachment Service - Handles file attachment operations
 */

import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/attachments';
import { IAttachment } from '../models/IAdminRequest';

export class AttachmentService {
  private static readonly ADMIN_REQUESTS_LIST = 'Admin Requests';
  private static readonly STATIONERY_ITEMS_LIST = 'Stationery Items';

  /**
   * Upload attachment to admin request
   */
  public static async uploadRequestAttachment(
    requestId: string,
    file: File
  ): Promise<IAttachment> {
    try {
      const fileBuffer = await this.fileToArrayBuffer(file);
      const result = await sp.web.lists
        .getByTitle(this.ADMIN_REQUESTS_LIST)
        .items.getById(parseInt(requestId))
        .attachmentFiles.addUsingPath(file.name, fileBuffer);

      return {
        id: file.name,
        fileName: file.name,
        fileUrl: result.data.AbsoluteUrl,
        fileSize: file.size,
        uploadedBy: 'Current User',
        uploadedDate: new Date()
      };
    } catch (error) {
      console.error(
        `Error uploading attachment to request ${requestId}:`,
        error
      );
      throw error;
    }
  }

  /**
   * Upload attachment to stationery item
   */
  public static async uploadItemAttachment(
    itemId: string,
    file: File
  ): Promise<IAttachment> {
    try {
      const fileBuffer = await this.fileToArrayBuffer(file);
      const result = await sp.web.lists
        .getByTitle(this.STATIONERY_ITEMS_LIST)
        .items.getById(parseInt(itemId))
        .attachmentFiles.addUsingPath(file.name, fileBuffer);

      return {
        id: file.name,
        fileName: file.name,
        fileUrl: result.data.AbsoluteUrl,
        fileSize: file.size,
        uploadedBy: 'Current User',
        uploadedDate: new Date()
      };
    } catch (error) {
      console.error(`Error uploading attachment to item ${itemId}:`, error);
      throw error;
    }
  }

  /**
   * Get attachments for request
   */
  public static async getRequestAttachments(requestId: string): Promise<IAttachment[]> {
    try {
      const attachments = await sp.web.lists
        .getByTitle(this.ADMIN_REQUESTS_LIST)
        .items.getById(parseInt(requestId))
        .attachmentFiles.get();

      return attachments.map((att: any) => ({
        id: att.FileName,
        fileName: att.FileName,
        fileUrl: att.AbsoluteUrl,
        fileSize: att.FileSize || 0,
        uploadedBy: 'Unknown',
        uploadedDate: new Date(att.TimeCreated)
      }));
    } catch (error) {
      console.error(
        `Error fetching attachments for request ${requestId}:`,
        error
      );
      return [];
    }
  }

  /**
   * Get attachments for stationery item
   */
  public static async getItemAttachments(itemId: string): Promise<IAttachment[]> {
    try {
      const attachments = await sp.web.lists
        .getByTitle(this.STATIONERY_ITEMS_LIST)
        .items.getById(parseInt(itemId))
        .attachmentFiles.get();

      return attachments.map((att: any) => ({
        id: att.FileName,
        fileName: att.FileName,
        fileUrl: att.AbsoluteUrl,
        fileSize: att.FileSize || 0,
        uploadedBy: 'Unknown',
        uploadedDate: new Date(att.TimeCreated)
      }));
    } catch (error) {
      console.error(
        `Error fetching attachments for item ${itemId}:`,
        error
      );
      return [];
    }
  }

  /**
   * Delete attachment from request
   */
  public static async deleteRequestAttachment(
    requestId: string,
    fileName: string
  ): Promise<void> {
    try {
      await sp.web.lists
        .getByTitle(this.ADMIN_REQUESTS_LIST)
        .items.getById(parseInt(requestId))
        .attachmentFiles.getByName(fileName)
        .delete();
    } catch (error) {
      console.error(
        `Error deleting attachment ${fileName} from request ${requestId}:`,
        error
      );
      throw error;
    }
  }

  /**
   * Delete attachment from stationery item
   */
  public static async deleteItemAttachment(
    itemId: string,
    fileName: string
  ): Promise<void> {
    try {
      await sp.web.lists
        .getByTitle(this.STATIONERY_ITEMS_LIST)
        .items.getById(parseInt(itemId))
        .attachmentFiles.getByName(fileName)
        .delete();
    } catch (error) {
      console.error(
        `Error deleting attachment ${fileName} from item ${itemId}:`,
        error
      );
      throw error;
    }
  }

  /**
   * Convert File to ArrayBuffer
   */
  private static fileToArrayBuffer(file: File): Promise<ArrayBuffer> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        resolve(reader.result as ArrayBuffer);
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }
}
