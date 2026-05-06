/**
 * Admin Request Model - Parent Form
 */

export enum RequestStatus {
  Pending = 'Pending',
  Approved = 'Approved',
  Rejected = 'Rejected',
  InProgress = 'InProgress',
  Completed = 'Completed'
}

export enum PriorityLevel {
  Low = 'Low',
  Medium = 'Medium',
  High = 'High',
  Urgent = 'Urgent'
}

export interface IAttachment {
  id: string;
  fileName: string;
  fileUrl: string;
  fileSize: number;
  uploadedBy: string;
  uploadedDate: Date;
}

export interface IAdminRequest {
  id?: string;
  title: string;
  description: string;
  requesterId?: string;
  requesterName?: string;
  requesterEmail?: string;
  department?: string;
  status?: RequestStatus;
  priority?: PriorityLevel;
  requestDate?: Date;
  targetDeliveryDate?: Date;
  approvedBy?: string;
  approvedDate?: Date;
  rejectionReason?: string;
  totalBudget?: number;
  attachments?: IAttachment[];
  comments?: string;
  created?: Date;
  modified?: Date;
}
