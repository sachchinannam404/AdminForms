/**
 * Central registry of request types, SharePoint lists, and form field schemas.
 * Add a new office admin functionality by registering a new entry here.
 */

import { RequestType, PriorityLevel, RequestStatus } from '../models/common/enums';

export type FieldType =
  | 'text'
  | 'multiline'
  | 'number'
  | 'currency'
  | 'date'
  | 'dropdown'
  | 'email'
  | 'checkbox';

export interface IFormField {
  key: string;
  label: string;
  type: FieldType;
  required?: boolean;
  placeholder?: string;
  options?: { key: string; text: string }[];
  /** Store under details.* vs top-level request */
  inDetails?: boolean;
}

export interface IRequestTypeConfig {
  key: RequestType;
  displayName: string;
  description: string;
  parentListTitle: string;
  childListTitle?: string;
  supportsChildren: boolean;
  iconName?: string;
  formFields: IFormField[];
  childFormFields?: IFormField[];
  defaultPriority?: PriorityLevel;
  defaultStatus?: RequestStatus;
  /** Optional budget threshold above which dual approval is recommended */
  approvalThreshold?: number;
}

const PARENT_LIST = 'Admin Requests';
const priorityOpts = Object.values(PriorityLevel).map((p) => ({ key: p, text: p }));

export const RequestTypeRegistry: Record<RequestType, IRequestTypeConfig> = {
  [RequestType.Stationery]: {
    key: RequestType.Stationery,
    displayName: 'Stationery / Office Supplies',
    description: 'Request pens, paper, notebooks and other consumables',
    parentListTitle: PARENT_LIST,
    childListTitle: 'Stationery Items',
    supportsChildren: true,
    iconName: 'EditNote',
    defaultPriority: PriorityLevel.Medium,
    defaultStatus: RequestStatus.Pending,
    formFields: [
      { key: 'title', label: 'Request Title', type: 'text', required: true },
      { key: 'description', label: 'Description', type: 'multiline', required: true },
      { key: 'department', label: 'Department', type: 'text' },
      { key: 'priority', label: 'Priority', type: 'dropdown', options: priorityOpts },
      { key: 'totalBudget', label: 'Total Budget', type: 'currency' },
      { key: 'targetDeliveryDate', label: 'Target Delivery Date', type: 'date' },
      { key: 'comments', label: 'Comments', type: 'multiline' }
    ],
    childFormFields: [
      { key: 'itemName', label: 'Item Name', type: 'text', required: true },
      {
        key: 'category',
        label: 'Category',
        type: 'dropdown',
        options: [
          { key: 'Paper', text: 'Paper' },
          { key: 'Pens', text: 'Pens' },
          { key: 'Notebooks', text: 'Notebooks' },
          { key: 'Folders', text: 'Folders' },
          { key: 'Tape', text: 'Tape' },
          { key: 'Ink', text: 'Ink' },
          { key: 'Other', text: 'Other' }
        ]
      },
      { key: 'quantity', label: 'Quantity', type: 'number', required: true },
      { key: 'unit', label: 'Unit', type: 'text', placeholder: 'pcs, ream, box' },
      { key: 'unitPrice', label: 'Unit Price', type: 'currency' },
      { key: 'vendorName', label: 'Vendor Name', type: 'text' },
      { key: 'vendorEmail', label: 'Vendor Email', type: 'email' },
      { key: 'expectedDeliveryDate', label: 'Expected Delivery', type: 'date' },
      { key: 'remarks', label: 'Remarks', type: 'multiline' }
    ]
  },

  [RequestType.ITEquipment]: {
    key: RequestType.ITEquipment,
    displayName: 'IT Equipment',
    description: 'Laptops, monitors, phones, peripherals – asset type, specs, serial, warranty',
    parentListTitle: PARENT_LIST,
    childListTitle: 'IT Equipment Items',
    supportsChildren: true,
    iconName: 'Devices3',
    defaultPriority: PriorityLevel.Medium,
    defaultStatus: RequestStatus.Pending,
    approvalThreshold: 2000,
    formFields: [
      { key: 'title', label: 'Request Title', type: 'text', required: true },
      { key: 'description', label: 'Business Justification', type: 'multiline', required: true },
      { key: 'department', label: 'Department', type: 'text' },
      { key: 'priority', label: 'Priority', type: 'dropdown', options: priorityOpts },
      { key: 'totalBudget', label: 'Estimated Budget', type: 'currency' },
      { key: 'targetDeliveryDate', label: 'Needed By', type: 'date' },
      { key: 'costCenter', label: 'Cost Center', type: 'text', inDetails: true },
      { key: 'comments', label: 'Comments', type: 'multiline' }
    ],
    childFormFields: [
      { key: 'itemName', label: 'Asset Name', type: 'text', required: true },
      {
        key: 'category',
        label: 'Asset Type',
        type: 'dropdown',
        options: [
          { key: 'Laptop', text: 'Laptop' },
          { key: 'Desktop', text: 'Desktop' },
          { key: 'Monitor', text: 'Monitor' },
          { key: 'Phone', text: 'Phone' },
          { key: 'Headset', text: 'Headset' },
          { key: 'Keyboard', text: 'Keyboard / Mouse' },
          { key: 'Other', text: 'Other' }
        ]
      },
      { key: 'quantity', label: 'Quantity', type: 'number', required: true },
      { key: 'unitPrice', label: 'Unit Price', type: 'currency' },
      { key: 'description', label: 'Specifications', type: 'multiline', placeholder: 'CPU, RAM, storage…' },
      { key: 'serialNumber', label: 'Serial Number', type: 'text', placeholder: 'If already assigned' },
      { key: 'warrantyMonths', label: 'Warranty (months)', type: 'number' },
      { key: 'warrantyExpiry', label: 'Warranty Expiry', type: 'date' },
      { key: 'vendorName', label: 'Preferred Vendor', type: 'text' },
      { key: 'expectedDeliveryDate', label: 'Expected Delivery', type: 'date' },
      { key: 'remarks', label: 'Remarks', type: 'multiline' }
    ]
  },

  [RequestType.Travel]: {
    key: RequestType.Travel,
    displayName: 'Travel & Expense',
    description: 'Flights, hotels, per-diem – budget vs actual',
    parentListTitle: PARENT_LIST,
    childListTitle: 'Travel Items',
    supportsChildren: true,
    iconName: 'Airplane',
    defaultPriority: PriorityLevel.Medium,
    defaultStatus: RequestStatus.Pending,
    approvalThreshold: 1500,
    formFields: [
      { key: 'title', label: 'Trip Title', type: 'text', required: true },
      { key: 'description', label: 'Purpose of Travel', type: 'multiline', required: true },
      { key: 'department', label: 'Department', type: 'text' },
      { key: 'priority', label: 'Priority', type: 'dropdown', options: priorityOpts },
      { key: 'destination', label: 'Destination', type: 'text', inDetails: true, required: true },
      { key: 'departureDate', label: 'Departure Date', type: 'date', inDetails: true },
      { key: 'returnDate', label: 'Return Date', type: 'date', inDetails: true },
      { key: 'totalBudget', label: 'Estimated Budget', type: 'currency' },
      { key: 'actualCost', label: 'Actual Cost (post-trip)', type: 'currency', inDetails: true },
      { key: 'comments', label: 'Comments', type: 'multiline' }
    ],
    childFormFields: [
      { key: 'itemName', label: 'Expense Item', type: 'text', required: true },
      {
        key: 'category',
        label: 'Category',
        type: 'dropdown',
        options: [
          { key: 'Flight', text: 'Flight' },
          { key: 'Hotel', text: 'Hotel' },
          { key: 'Train', text: 'Train' },
          { key: 'Taxi', text: 'Taxi / Transfer' },
          { key: 'PerDiem', text: 'Per Diem' },
          { key: 'Other', text: 'Other' }
        ]
      },
      { key: 'unitPrice', label: 'Estimated Amount', type: 'currency', required: true },
      { key: 'actualAmount', label: 'Actual Amount', type: 'currency' },
      { key: 'quantity', label: 'Qty / Nights', type: 'number' },
      { key: 'description', label: 'Details', type: 'multiline' },
      { key: 'remarks', label: 'Remarks', type: 'multiline' }
    ]
  },

  [RequestType.Leave]: {
    key: RequestType.Leave,
    displayName: 'Leave / Time Off',
    description: 'Annual, sick, personal – balance field ready for Graph later',
    parentListTitle: PARENT_LIST,
    supportsChildren: false,
    iconName: 'Calendar',
    defaultPriority: PriorityLevel.Medium,
    defaultStatus: RequestStatus.Pending,
    formFields: [
      {
        key: 'title',
        label: 'Leave Title',
        type: 'text',
        required: true,
        placeholder: 'e.g. Annual Leave – March'
      },
      { key: 'description', label: 'Reason / Notes', type: 'multiline' },
      {
        key: 'leaveType',
        label: 'Leave Type',
        type: 'dropdown',
        inDetails: true,
        required: true,
        options: [
          { key: 'Annual', text: 'Annual Leave' },
          { key: 'Sick', text: 'Sick Leave' },
          { key: 'Personal', text: 'Personal' },
          { key: 'Unpaid', text: 'Unpaid' },
          { key: 'Other', text: 'Other' }
        ]
      },
      { key: 'startDate', label: 'Start Date', type: 'date', inDetails: true, required: true },
      { key: 'endDate', label: 'End Date', type: 'date', inDetails: true, required: true },
      { key: 'halfDay', label: 'Half Day', type: 'checkbox', inDetails: true },
      {
        key: 'leaveBalanceDays',
        label: 'Leave Balance (days)',
        type: 'number',
        inDetails: true,
        placeholder: 'Manual for now; Graph later'
      },
      { key: 'daysRequested', label: 'Days Requested', type: 'number', inDetails: true },
      { key: 'comments', label: 'Comments', type: 'multiline' }
    ]
  },

  [RequestType.Facilities]: {
    key: RequestType.Facilities,
    displayName: 'Facilities / Maintenance',
    description: 'Location, urgency, assigned technician – attach photos on request',
    parentListTitle: PARENT_LIST,
    childListTitle: 'Facilities Items',
    supportsChildren: true,
    iconName: 'BuildingHome',
    defaultPriority: PriorityLevel.Medium,
    defaultStatus: RequestStatus.Pending,
    formFields: [
      { key: 'title', label: 'Issue Title', type: 'text', required: true },
      { key: 'description', label: 'Description', type: 'multiline', required: true },
      {
        key: 'location',
        label: 'Location / Floor / Desk',
        type: 'text',
        inDetails: true,
        required: true
      },
      {
        key: 'issueCategory',
        label: 'Category',
        type: 'dropdown',
        inDetails: true,
        options: [
          { key: 'HVAC', text: 'HVAC' },
          { key: 'Electrical', text: 'Electrical' },
          { key: 'Plumbing', text: 'Plumbing' },
          { key: 'Furniture', text: 'Furniture' },
          { key: 'Cleaning', text: 'Cleaning' },
          { key: 'Access', text: 'Access / Security' },
          { key: 'Other', text: 'Other' }
        ]
      },
      { key: 'priority', label: 'Urgency / Priority', type: 'dropdown', options: priorityOpts },
      {
        key: 'assignedTechnician',
        label: 'Assigned Technician',
        type: 'text',
        inDetails: true,
        placeholder: 'Name or email'
      },
      { key: 'targetDeliveryDate', label: 'Needed By', type: 'date' },
      { key: 'comments', label: 'Comments', type: 'multiline' }
    ],
    childFormFields: [
      { key: 'itemName', label: 'Task / Part', type: 'text', required: true },
      { key: 'quantity', label: 'Quantity', type: 'number' },
      { key: 'unitPrice', label: 'Cost', type: 'currency' },
      { key: 'description', label: 'Details', type: 'multiline' },
      { key: 'remarks', label: 'Remarks', type: 'multiline' }
    ]
  },

  [RequestType.Procurement]: {
    key: RequestType.Procurement,
    displayName: 'Procurement / Purchase Order',
    description: 'Vendor, PO number, approval threshold for high-value buys',
    parentListTitle: PARENT_LIST,
    childListTitle: 'Procurement Items',
    supportsChildren: true,
    iconName: 'ShoppingCart',
    defaultPriority: PriorityLevel.Medium,
    defaultStatus: RequestStatus.Pending,
    approvalThreshold: 5000,
    formFields: [
      { key: 'title', label: 'PO / Request Title', type: 'text', required: true },
      { key: 'description', label: 'Justification', type: 'multiline', required: true },
      { key: 'department', label: 'Department', type: 'text' },
      { key: 'vendorName', label: 'Preferred Vendor', type: 'text', inDetails: true },
      { key: 'poNumber', label: 'PO Number', type: 'text', inDetails: true },
      { key: 'contractRef', label: 'Contract Reference', type: 'text', inDetails: true },
      { key: 'totalBudget', label: 'Total Amount', type: 'currency', required: true },
      {
        key: 'requiresDualApproval',
        label: 'Requires dual approval',
        type: 'checkbox',
        inDetails: true
      },
      { key: 'priority', label: 'Priority', type: 'dropdown', options: priorityOpts },
      { key: 'targetDeliveryDate', label: 'Required By', type: 'date' },
      { key: 'comments', label: 'Comments', type: 'multiline' }
    ],
    childFormFields: [
      { key: 'itemName', label: 'Line Item', type: 'text', required: true },
      { key: 'quantity', label: 'Quantity', type: 'number', required: true },
      { key: 'unit', label: 'Unit', type: 'text' },
      { key: 'unitPrice', label: 'Unit Price', type: 'currency', required: true },
      { key: 'description', label: 'Description', type: 'multiline' },
      { key: 'vendorName', label: 'Vendor', type: 'text' },
      { key: 'expectedDeliveryDate', label: 'Expected Delivery', type: 'date' },
      { key: 'remarks', label: 'Remarks', type: 'multiline' }
    ]
  },

  [RequestType.General]: {
    key: RequestType.General,
    displayName: 'General Admin Request',
    description: 'Catch-all for other administrative requests',
    parentListTitle: PARENT_LIST,
    supportsChildren: false,
    iconName: 'PageList',
    defaultPriority: PriorityLevel.Medium,
    defaultStatus: RequestStatus.Pending,
    formFields: [
      { key: 'title', label: 'Title', type: 'text', required: true },
      { key: 'description', label: 'Description', type: 'multiline', required: true },
      { key: 'department', label: 'Department', type: 'text' },
      { key: 'priority', label: 'Priority', type: 'dropdown', options: priorityOpts },
      { key: 'totalBudget', label: 'Budget (if any)', type: 'currency' },
      { key: 'targetDeliveryDate', label: 'Needed By', type: 'date' },
      { key: 'comments', label: 'Comments', type: 'multiline' }
    ]
  }
};

export function getRequestTypeConfig(type: RequestType): IRequestTypeConfig {
  return RequestTypeRegistry[type];
}

export function getAllRequestTypes(): IRequestTypeConfig[] {
  return Object.values(RequestTypeRegistry);
}

/** Whether amount exceeds configured dual-approval threshold for the type */
export function exceedsApprovalThreshold(type: RequestType, amount?: number): boolean {
  const cfg = RequestTypeRegistry[type];
  if (!cfg.approvalThreshold || amount == null) return false;
  return amount >= cfg.approvalThreshold;
}
