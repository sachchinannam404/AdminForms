/**
 * Attachment Service – works with any list (parent or child)
 */

import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/attachments';
import { IAttachment } from '../models/common/IAttachment';

export class AttachmentService {
  public static async upload(
    listTitle: string,
    itemId: string,
    file: File
  ): Promise<IAttachment> {
    try {
      const fileBuffer = await this.fileToArrayBuffer(file);
      const result = await sp.web.lists
        .getByTitle(listTitle)
        .items.getById(parseInt(itemId, 10))
        .attachmentFiles.addUsingPath(file.name, fileBuffer);

      return {
        id: file.name,
        fileName: file.name,
        fileUrl: result.data.AbsoluteUrl || result.data.ServerRelativeUrl,
        fileSize: file.size,
        uploadedBy: 'Current User',
        uploadedDate: new Date()
      };
    } catch (error) {
      console.error(`Error uploading attachment to ${listTitle}/${itemId}:`, error);
      throw error;
    }
  }

  public static async getAll(listTitle: string, itemId: string): Promise<IAttachment[]> {
    try {
      const attachments = await sp.web.lists
        .getByTitle(listTitle)
        .items.getById(parseInt(itemId, 10))
        .attachmentFiles.get();

      return attachments.map((att: any) => ({
        id: att.FileName,
        fileName: att.FileName,
        fileUrl: att.AbsoluteUrl || att.ServerRelativeUrl,
        fileSize: att.FileSize || 0,
        uploadedBy: 'Unknown',
        uploadedDate: att.TimeCreated ? new Date(att.TimeCreated) : new Date()
      }));
    } catch (error) {
      console.error(`Error fetching attachments for ${listTitle}/${itemId}:`, error);
      return [];
    }
  }

  public static async delete(listTitle: string, itemId: string, fileName: string): Promise<void> {
    try {
      await sp.web.lists
        .getByTitle(listTitle)
        .items.getById(parseInt(itemId, 10))
        .attachmentFiles.getByName(fileName)
        .delete();
    } catch (error) {
      console.error(`Error deleting attachment ${fileName} from ${listTitle}/${itemId}:`, error);
      throw error;
    }
  }

  /** Backward-compatible helpers */
  public static async uploadRequestAttachment(requestId: string, file: File): Promise<IAttachment> {
    return this.upload('Admin Requests', requestId, file);
  }

  public static async uploadItemAttachment(
    itemId: string,
    file: File,
    listTitle: string = 'Stationery Items'
  ): Promise<IAttachment> {
    return this.upload(listTitle, itemId, file);
  }

  public static async getRequestAttachments(requestId: string): Promise<IAttachment[]> {
    return this.getAll('Admin Requests', requestId);
  }

  public static async getItemAttachments(
    itemId: string,
    listTitle: string = 'Stationery Items'
  ): Promise<IAttachment[]> {
    return this.getAll(listTitle, itemId);
  }

  public static async deleteRequestAttachment(requestId: string, fileName: string): Promise<void> {
    return this.delete('Admin Requests', requestId, fileName);
  }

  public static async deleteItemAttachment(
    itemId: string,
    fileName: string,
    listTitle: string = 'Stationery Items'
  ): Promise<void> {
    return this.delete(listTitle, itemId, fileName);
  }

  private static fileToArrayBuffer(file: File): Promise<ArrayBuffer> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }
}
