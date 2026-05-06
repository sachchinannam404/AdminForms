/**
 * Stationery Item Model - Child Form
 */

export enum ItemCategory {
  Paper = 'Paper',
  Pens = 'Pens',
  Notebooks = 'Notebooks',
  Folders = 'Folders',
  Tape = 'Tape',
  Ink = 'Ink',
  Other = 'Other'
}

export enum ItemStatus {
  Pending = 'Pending',
  Ordered = 'Ordered',
  InStock = 'InStock',
  Delivered = 'Delivered'
}

export interface IAttachment {
  id: string;
  fileName: string;
  fileUrl: string;
  fileSize: number;
  uploadedBy: string;
  uploadedDate: Date;
}

export interface IStationeryItem {
  id?: string;
  requestId: string;
  itemName: string;
  category: ItemCategory;
  quantity: number;
  unit?: string;
  unitPrice?: number;
  totalPrice?: number;
  description?: string;
  status?: ItemStatus;
  vendorName?: string;
  vendorEmail?: string;
  orderDate?: Date;
  expectedDeliveryDate?: Date;
  actualDeliveryDate?: Date;
  remarks?: string;
  attachments?: IAttachment[];
  created?: Date;
  modified?: Date;
}
