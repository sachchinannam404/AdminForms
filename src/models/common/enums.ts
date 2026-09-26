/**
 * Shared enums for Admin Forms
 */

export enum RequestStatus {
  Draft = 'Draft',
  Pending = 'Pending',
  Approved = 'Approved',
  Rejected = 'Rejected',
  InProgress = 'InProgress',
  Completed = 'Completed',
  Cancelled = 'Cancelled'
}

export enum PriorityLevel {
  Low = 'Low',
  Medium = 'Medium',
  High = 'High',
  Urgent = 'Urgent'
}

export enum RequestType {
  Stationery = 'Stationery',
  ITEquipment = 'ITEquipment',
  Travel = 'Travel',
  Leave = 'Leave',
  Facilities = 'Facilities',
  Procurement = 'Procurement',
  General = 'General'
}

export enum ItemStatus {
  Pending = 'Pending',
  Ordered = 'Ordered',
  InStock = 'InStock',
  Delivered = 'Delivered',
  Cancelled = 'Cancelled'
}
