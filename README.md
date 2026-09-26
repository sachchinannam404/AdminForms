# Admin Forms SPFx Solution (v2 – Full office admin suite)

Scalable SharePoint Framework solution for office administration: multi-type requests, dashboard, roles, bulk actions, approvals, audit, notifications, and reporting.

## Feature matrix

| Functionality | Implementation |
| ----------------------------- | ------------------------------------------------------------------------------------------------------------------------------- |
| **Dashboard** | `RequestsDashboard` – DetailsList, filters (status, type, department, date range, search), KPIs (Pending / Overdue / Budget) |
| **IT Equipment** | Registry + child list – asset type, specs, **serial**, **warranty** months/expiry |
| **Travel / Expense** | Child items (flights, hotels, per-diem); **estimated vs actual** cost on parent and lines |
| **Leave / Time-off** | Date range, leave type, **leave balance** + days requested (Graph-ready) |
| **Facilities / Maintenance** | Location, urgency/priority, **assigned technician**; photos via request attachments |
| **Procurement / PO** | Vendor, PO number, contract ref, **approval threshold** + dual-approval flag |
| **Approval workflow UI** | `ApprovalPanel` – Approve / Reject + comments; **Power Automate webhook** via `NotificationService` |
| **Role-based views** | `RoleService` – SharePoint groups; **My requests** / **All** / **Pending approval** pivots |
| **Search & advanced filter** | PnP OData filter + client search; department & date range |
| **Bulk actions** | Multi-select + bulk status update (approvers) |
| **Notifications** | Power Automate HTTP webhook on status/bulk change (`NotificationService`) |
| **Audit / history** | `AuditHistory` + `AuditService` – SharePoint **versioning** diffs |
| **Reporting** | `ReportingPanel` – summary stats, simple bar charts, **CSV export** |

## Request types

| Type | Line items | Notes |
|------|------------|--------|
| Stationery | Yes | Consumables |
| IT Equipment | Yes | Serial, warranty, specs |
| Travel | Yes | Budget vs actual |
| Leave | No | Balance field for Graph later |
| Facilities | Yes | Technician, location |
| Procurement | Yes | Threshold / dual approval |
| General | No | Catch-all |

## Architecture

```
src/
├── models/          # IAdminRequest, IChildItem, enums, IAttachment
├── config/          # requestTypeRegistry (schemas + thresholds)
├── services/
│   ├── SharePointRepository.ts
│   ├── RequestService.ts      # CRUD, bulk, KPIs
│   ├── RoleService.ts
│   ├── NotificationService.ts
│   ├── AuditService.ts
│   ├── ReportingService.ts
│   └── attachmentService.ts
├── components/
│   ├── AdminFormsApp.tsx
│   ├── dashboard/RequestsDashboard.tsx
│   ├── forms/DynamicRequestForm.tsx
│   ├── lists/ChildItemsList.tsx
│   ├── workflow/ApprovalPanel.tsx, AuditHistory.tsx
│   └── reporting/ReportingPanel.tsx
```

## Prerequisites

- Node.js 14+ / SPFx 1.16 / SharePoint Online

## Installation

```bash
git clone https://github.com/sachchinannam404/AdminForms.git
cd AdminForms
npm install
```

## SharePoint setup

### Admin Requests (parent)

Required columns: Title, Description, **RequestType** (Choice), RequesterName, RequesterEmail, Department, Status, Priority, TargetDeliveryDate, ApprovedBy, ApprovedDate, RejectionReason, TotalBudget, Comments, **DetailsJson** (multi-line text).

Enable **versioning** for audit history.

### Child lists (as needed)

Stationery Items, IT Equipment Items, Travel Items, Facilities Items, Procurement Items – same core columns (RequestId, ItemName, Category, Quantity, Unit, UnitPrice, TotalPrice, Description, Status, Vendor*, dates, Remarks, **ExtraJson**).

### Security groups (optional)

- `Admin Forms Approvers` → Approver view + bulk actions  
- `Admin Forms Admins` → Admin role  

### Power Automate (optional)

1. Create a flow with **When an HTTP request is received**.
2. Pass the URL into the web part:

```tsx
RequestService.initialize(this.context);
<AdminFormsApp
  currentUserName={this.context.pageContext.user.displayName}
  currentUserEmail={this.context.pageContext.user.email}
  powerAutomateWebhookUrl="https://prod-....logic.azure.com/workflows/..."
/>
```

Payload events: `AdminRequestStatusChanged`, `AdminRequestBulkStatusChanged`.

## Web part usage

```tsx
import { AdminFormsApp } from './components/AdminFormsApp';
import { RequestService } from './services/RequestService';

RequestService.initialize(this.context);

<AdminFormsApp
  currentUserName={this.context.pageContext.user.displayName}
  currentUserEmail={this.context.pageContext.user.email}
/>
```

## Adding a new request type

1. Add `RequestType` enum value.  
2. Register in `requestTypeRegistry.ts` (fields, optional child list, `approvalThreshold`).  
3. Create lists if needed.  
4. No new form components required.

## Development

```bash
npm run serve
npm run build
npm run package-solution
```

## License

MIT
