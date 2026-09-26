/**
 * Base Admin Request model – supports multiple request types via requestType + details
 */

import { IAttachment } from './common/IAttachment';
import { RequestStatus, PriorityLevel, RequestType } from './common/enums';

export { RequestStatus, PriorityLevel, RequestType };
export type { IAttachment };

export interface IAdminRequest {
  id?: string;
  title: string;
  description: string;
  requestType: RequestType;
  requesterId?: string;
  requesterName?: string;
  requesterEmail?: string;
  department?: string;
  status: RequestStatus;
  priority: PriorityLevel;
  requestDate?: Date;
  targetDeliveryDate?: Date;
  approvedBy?: string;
  approvedDate?: Date;
  rejectionReason?: string;
  totalBudget?: number;
  comments?: string;
  attachments?: IAttachment[];
  /** Type-specific payload (stationery lines, travel legs, leave dates, etc.) */
  details?: Record<string, any>;
  created?: Date;
  modified?: Date;
}

export interface IAdminRequestFilter {
  status?: RequestStatus | RequestStatus[];
  requestType?: RequestType | RequestType[];
  department?: string;
  priority?: PriorityLevel;
  requesterEmail?: string;
  fromDate?: Date;
  toDate?: Date;
  searchText?: string;
}
