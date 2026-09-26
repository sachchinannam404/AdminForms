/**
 * Stationery Item Model - Child Form (backward compatible)
 * Prefer IChildItem + RequestType.Stationery for new code.
 */

import { IAttachment } from './common/IAttachment';
import { ItemStatus } from './common/enums';

export { ItemStatus };

export enum ItemCategory {
  Paper = 'Paper',
  Pens = 'Pens',
  Notebooks = 'Notebooks',
  Folders = 'Folders',
  Tape = 'Tape',
  Ink = 'Ink',
  Other = 'Other'
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
