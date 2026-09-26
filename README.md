# Admin Forms SPFx Solution (v2 – Scalable)

A **scalable** SharePoint Framework (SPFx) solution for office administration: multi-type requests, config-driven forms, dashboard, approval workflow, and line items.

## Request types (out of the box)

| Type | Description | Line items |
|------|-------------|------------|
| **Stationery** | Office supplies | Yes |
| **IT Equipment** | Laptops, monitors, phones, etc. | Yes |
| **Travel** | Trips, flights, hotels, per-diem | Yes |
| **Leave** | Annual / sick / personal time off | No |
| **Facilities** | Maintenance, HVAC, furniture | Yes |
| **Procurement** | POs and vendor purchases | Yes |
| **General** | Catch-all admin requests | No |

Add a new type by registering it in `src/config/requestTypeRegistry.ts` (no new form components required for most cases).

## Architecture

```
src/
├── models/
│   ├── common/           # IAttachment, enums
│   ├── IAdminRequest.ts  # Base request + filters
│   ├── IChildItem.ts     # Generic line item
│   └── IStationeryItem.ts# Legacy stationery model
├── config/
│   └── requestTypeRegistry.ts   # Types, lists, form schemas
├── services/
│   ├── SharePointRepository.ts  # Generic list CRUD
│   ├── RequestService.ts        # High-level API
│   ├── attachmentService.ts
│   └── sharePointService.ts     # Backward-compatible wrapper
├── components/
│   ├── AdminFormsApp.tsx        # Shell: dashboard ↔ form ↔ detail
│   ├── dashboard/RequestsDashboard.tsx
│   ├── forms/DynamicRequestForm.tsx
│   ├── workflow/ApprovalPanel.tsx
│   ├── lists/ChildItemsList.tsx
│   └── (legacy) AdminRequestForm, StationeryItem*
```

## Features

- **Dashboard** – KPIs, search, filter by status / type / priority
- **Config-driven forms** – fields defined in the registry
- **Approval panel** – Approve / Reject with comments
- **Child line items** – per request type (stationery, IT, travel, etc.)
- **Attachments** – on parent requests (and extendable to children)
- **Generic repository** – one pattern for all SharePoint lists
- **Backward compatible** – existing Stationery forms still work via `SharePointService`

## Prerequisites

- Node.js 14.x or higher
- SharePoint Framework v1.16.0
- SharePoint Online tenant / Office 365

## Installation

```bash
git clone https://github.com/sachchinannam404/AdminForms.git
cd AdminForms
npm install
npm install -g @microsoft/sharepoint-cli   # optional
```

## SharePoint lists setup

### 1. Admin Requests (parent – shared by all types)

| Column | Type | Notes |
|--------|------|--------|
| Title | Single line text | |
| Description | Multiple lines | |
| **RequestType** | Choice | Stationery, ITEquipment, Travel, Leave, Facilities, Procurement, General |
| RequesterName | Single line text | |
| RequesterEmail | Single line text | |
| Department | Single line text | |
| Status | Choice | Draft, Pending, Approved, Rejected, InProgress, Completed, Cancelled |
| Priority | Choice | Low, Medium, High, Urgent |
| TargetDeliveryDate | Date and Time | |
| ApprovedBy | Single line text (or Person) | |
| ApprovedDate | Date and Time | |
| RejectionReason | Multiple lines | |
| TotalBudget | Currency | |
| Comments | Multiple lines | |
| **DetailsJson** | Multiple lines (plain) | JSON for type-specific fields |

### 2. Child lists (create as needed)

Use the same column set for: **Stationery Items**, **IT Equipment Items**, **Travel Items**, **Facilities Items**, **Procurement Items**:

| Column | Type |
|--------|------|
| RequestId | Single line text |
| ItemName | Single line text |
| Category | Single line text or Choice |
| Quantity | Number |
| Unit | Single line text |
| UnitPrice | Currency |
| TotalPrice | Currency |
| Description | Multiple lines |
| Status | Choice (Pending, Ordered, InStock, Delivered, Cancelled) |
| VendorName | Single line text |
| VendorEmail | Single line text |
| ExpectedDeliveryDate | Date and Time |
| ActualDeliveryDate | Date and Time |
| Remarks | Multiple lines |
| ExtraJson | Multiple lines (optional JSON) |

Leave and General types do not require a child list.

## Usage in a web part

```tsx
import { AdminFormsApp } from './components/AdminFormsApp';
import { RequestService } from './services/RequestService';

// In web part onInit / render:
RequestService.initialize(this.context);

// Render:
<AdminFormsApp
  currentUserName={this.context.pageContext.user.displayName}
  currentUserEmail={this.context.pageContext.user.email}
/>
```

## Adding a new request type

1. Add an enum value in `src/models/common/enums.ts` (`RequestType`).
2. Add a full entry in `src/config/requestTypeRegistry.ts` (display name, form fields, optional child list + child fields).
3. Create the SharePoint list(s) if you use children.
4. No new React form required – `DynamicRequestForm` and `ChildItemsList` pick up the config.

## Development

```bash
npm run serve
npm run build
npm run bundle
npm run package-solution
```

## Migration from v1

- Existing **Admin Requests** / **Stationery Items** lists still work.
- Add columns **RequestType** and **DetailsJson** on Admin Requests (default RequestType = Stationery for old items).
- Prefer `RequestService` and `AdminFormsApp` for new UI; legacy `AdminRequestForm` / `StationeryItemsList` remain available.

## License

MIT License

## Support

Open an issue on the GitHub repository for questions or enhancements.
