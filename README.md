# Admin Forms SPFx Solution (v2 – Full office admin suite)

Scalable SharePoint Framework solution for office administration: multi-type requests, dashboard, roles, bulk actions, approvals, audit, notifications, and reporting.

**Full site setup (web part, lists, groups):** see **[docs/SETUP.md](docs/SETUP.md)**.

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
├── webparts/adminForms/     # SPFx web part → AdminFormsApp + spfxContext
├── models/
├── config/requestTypeRegistry.ts
├── services/
│   ├── RequestService.ts
│   ├── SiteProvisioningService.ts   # lists, columns, groups
│   ├── RoleService.ts
│   └── ...
├── components/AdminFormsApp.tsx
scripts/Provision-AdminForms.ps1     # PnP PowerShell provisioning
docs/SETUP.md
```

## Quick start

```bash
git clone https://github.com/sachchinannam404/AdminForms.git
cd AdminForms
npm install
```

### 1. Web part (already wired)

`src/webparts/adminForms/AdminFormsWebPart.ts` initializes PnP with `this.context` and renders:

```tsx
<AdminFormsApp
  spfxContext={this.context}
  currentUserName={this.context.pageContext.user.displayName}
  currentUserEmail={this.context.pageContext.user.email}
  powerAutomateWebhookUrl={this.properties.powerAutomateWebhookUrl}
/>
```

Property pane: webhook URL + **Ensure lists & groups on load**.

### 2. Lists & columns

**Option A – in browser (site owner):** enable web part toggle *Ensure lists & groups on load*.

**Option B – PowerShell:**

```powershell
Connect-PnPOnline -Url "https://tenant.sharepoint.com/sites/YourSite" -Interactive
.\scripts\Provision-AdminForms.ps1
```

Creates **Admin Requests** (with **RequestType**, **DetailsJson**, versioning) and child lists (Stationery, IT Equipment, Travel, Facilities, Procurement Items).

### 3. Groups

Provisioning creates:

- **Admin Forms Approvers** – approve / bulk / “All requests” view  
- **Admin Forms Admins** – admin role  

Add members via Site settings → People and groups, or:

```powershell
Add-PnPGroupMember -Identity "Admin Forms Approvers" -LoginName "user@contoso.com"
```

### 4. Deploy

```bash
gulp bundle --ship
gulp package-solution --ship
```

Deploy `sharepoint/solution/admin-forms.sppkg` to the App Catalog, add **Admin Forms** to a page.

> If gulp/config scaffold is incomplete in a clone, use a standard SPFx 1.16 project and merge this `src/` (see [docs/SETUP.md](docs/SETUP.md)).

## Development

```bash
npm run serve
npm run build
npm run package-solution
```

## License

MIT
