/**
 * Generic child item linked to an admin request (stationery line, IT asset, travel leg, etc.)
 */

import { IAttachment } from './common/IAttachment';
import { ItemStatus } from './common/enums';

export interface IChildItem {
  id?: string;
  requestId: string;
  itemName: string;
  category?: string;
  quantity?: number;
  unit?: string;
  unitPrice?: number;
  totalPrice?: number;
  description?: string;
  status?: ItemStatus | string;
  vendorName?: string;
  vendorEmail?: string;
  expectedDeliveryDate?: Date;
  actualDeliveryDate?: Date;
  remarks?: string;
  attachments?: IAttachment[];
  /** Extra type-specific fields */
  extra?: Record<string, any>;
  created?: Date;
  modified?: Date;
}
