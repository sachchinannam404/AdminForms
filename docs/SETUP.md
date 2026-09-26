# Admin Forms – site setup

## 1. Wire the SPFx web part

The web part lives at:

```
src/webparts/adminForms/AdminFormsWebPart.ts
```

It:

- Calls `RequestService.initialize(this.context)` in `onInit`
- Renders `<AdminFormsApp spfxContext={...} currentUserName={...} currentUserEmail={...} />`
- Optional property: **Power Automate webhook URL**
- Optional toggle: **Ensure lists & groups on load** (runs `SiteProvisioningService`)

### Build & deploy

```bash
npm install
gulp bundle --ship
gulp package-solution --ship
```

Upload the `.sppkg` from `sharepoint/solution/` to the tenant App Catalog, deploy, then add **Admin Forms** to a site page.

> If this repo is missing full SPFx scaffold files (`config/`, `gulpfile.js`, `tsconfig.json`), generate a 1.16 web part project and copy `src/` into it, or run `yo @microsoft/sharepoint` and replace the web part body with `AdminFormsWebPart.ts`.

### Property pane

| Property | Purpose |
|----------|---------|
| `powerAutomateWebhookUrl` | HTTP trigger for status notifications |
| `runProvisioningOnInit` | Create/update lists & groups when the page loads (needs elevated rights) |

---

## 2. SharePoint lists

### Option A – Web part toggle

1. Open the page as **site owner**.
2. Edit the web part → enable **Ensure lists & groups on load** → republish.
3. Reload once; check browser console for `[AdminForms] Provisioning:`.

### Option B – PnP PowerShell

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
Connect-PnPOnline -Url "https://YOURTENANT.sharepoint.com/sites/YOURSITE" -Interactive
.\scripts\Provision-AdminForms.ps1
```

### What gets created

| Artifact | Details |
|----------|---------|
| **Admin Requests** | Parent list; **versioning on**; columns including **RequestType**, **DetailsJson**, Status, Priority, budget, etc. |
| **Stationery Items** | Child line items |
| **IT Equipment Items** | Child |
| **Travel Items** | Child |
| **Facilities Items** | Child |
| **Procurement Items** | Child |

`RequestType` choices: Stationery, ITEquipment, Travel, Leave, Facilities, Procurement, General.  
`DetailsJson`: multi-line text for type-specific JSON.  
Child lists include `RequestId`, pricing, vendor, dates, `ExtraJson`.

---

## 3. Security groups

| Group | Purpose |
|-------|---------|
| **Admin Forms Approvers** | Approver pivot, bulk status, approve/reject |
| **Admin Forms Admins** | Admin role (same as approver + admin flag) |

Provisioning creates empty groups. Add people:

```powershell
Add-PnPGroupMember -Identity "Admin Forms Approvers" -LoginName "user@contoso.com"
Add-PnPGroupMember -Identity "Admin Forms Admins" -LoginName "admin@contoso.com"
```

Or: Site settings → People and groups → open each group → New.

`RoleService` matches group titles containing these names (and common variants like “Approvers”, “Site Owners”).

---

## 4. Verify

1. Add web part to a modern page.  
2. Create a Stationery request → add line items.  
3. Sign in as an Approver group member → see **All requests** / **Pending approval**.  
4. Approve/reject → optional webhook fires.  
5. **Load history** on a request (versioning must be on).
