# Admin Forms SPFx Solution

A comprehensive SharePoint Framework (SPFx) solution with parent and child forms for managing admin requests and stationery items with full attachment support.

## Features

- **Parent Form (Admin Request)**
  - Create and manage admin requests
  - Track request status and priority
  - Assign budgets and delivery dates
  - Add attachments (documents, images, etc.)
  - Approval workflow support

- **Child Form (Stationery Items)**
  - Manage stationery items linked to admin requests
  - Track item categories, quantities, and pricing
  - Monitor item status and delivery dates
  - Add vendor information
  - Support for multiple attachments per item

- **Common Features**
  - Full CRUD operations via SharePoint REST API
  - File attachment management
  - Responsive UI with Fluent UI components
  - Error handling and validation
  - Loading states and user feedback

## Prerequisites

- Node.js 14.x or higher
- SharePoint Framework v1.16.0
- SharePoint Online tenant
- Office 365 subscription
- Visual Studio Code or preferred TypeScript IDE

## Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/sachchinannam404/AdminForms.git
   cd AdminForms
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Install SharePoint Framework CLI**
   ```bash
   npm install -g @microsoft/sharepoint-cli
   ```

## Project Structure

```
src/
├── models/
│   ├── IAdminRequest.ts      # Admin Request model
│   └── IStationeryItem.ts    # Stationery Item model
├── services/
│   ├── sharePointService.ts  # SharePoint CRUD operations
│   └── attachmentService.ts  # File attachment handling
├── components/
│   ├── AdminRequestForm.tsx          # Parent form component
│   ├── AdminRequestForm.module.scss  # Parent form styles
│   ├── StationeryItemsList.tsx       # Child items list
│   ├── StationeryItemsList.module.scss
│   ├── StationeryItemForm.tsx        # Child form component
│   └── StationeryItemForm.module.scss
└── webparts/
    └── AdminFormWebPart/
        ├── AdminFormWebPart.tsx
        └── IAdminFormWebPartProps.ts
```

## SharePoint Lists Setup

You need to create two SharePoint lists:

### 1. Admin Requests List

Columns:
- **Title** (Single line text) - Request title
- **Description** (Multiple lines text) - Detailed description
- **RequesterName** (Single line text) - Name of requester
- **RequesterEmail** (Single line text) - Email of requester
- **Department** (Single line text) - Department name
- **Status** (Choice) - Pending, Approved, Rejected, InProgress, Completed
- **Priority** (Choice) - Low, Medium, High, Urgent
- **TargetDeliveryDate** (Date and Time) - Expected delivery date
- **ApprovedBy** (Person or Group) - Approver name
- **ApprovedDate** (Date and Time) - Approval date
- **RejectionReason** (Multiple lines text) - Reason for rejection
- **TotalBudget** (Currency) - Budget amount
- **Comments** (Multiple lines text) - Additional comments

### 2. Stationery Items List

Columns:
- **RequestId** (Single line text) - Link to Admin Request
- **ItemName** (Single line text) - Item name
- **Category** (Choice) - Paper, Pens, Notebooks, Folders, Tape, Ink, Other
- **Quantity** (Number) - Item quantity
- **Unit** (Single line text) - Unit type (pcs, ream, box, etc.)
- **UnitPrice** (Currency) - Price per unit
- **TotalPrice** (Currency) - Total cost
- **Description** (Multiple lines text) - Item description
- **Status** (Choice) - Pending, Ordered, InStock, Delivered
- **VendorName** (Single line text) - Supplier name
- **VendorEmail** (Single line text) - Supplier email
- **ExpectedDeliveryDate** (Date and Time) - Expected delivery
- **ActualDeliveryDate** (Date and Time) - Actual delivery
- **Remarks** (Multiple lines text) - Additional notes

## Development

1. **Start the local development server**
   ```bash
   npm run serve
   ```

2. **Build the solution**
   ```bash
   npm run build
   ```

3. **Bundle for production**
   ```bash
   npm run bundle
   ```

4. **Package the solution**
   ```bash
   npm run package-solution
   ```

## Usage

### Creating an Admin Request

1. Open the Admin Request Form
2. Fill in the required fields (Title, Description)
3. Add optional details (Budget, Delivery Date, etc.)
4. Attach supporting documents
5. Click "Save"

### Managing Stationery Items

1. Open the Admin Request
2. Navigate to the Stationery Items section
3. Click "Add Item" to create a new item
4. Fill in item details
5. Add attachments if needed
6. Save the item

## API Integration

The solution uses PnP/sp library for SharePoint REST API operations:

```typescript
// Example: Create Admin Request
const request = await SharePointService.createAdminRequest({
  title: 'Office Supplies Request',
  description: 'Need office supplies for Q2',
  requesterName: 'John Doe',
  requesterEmail: 'john@contoso.com'
});

// Example: Upload Attachment
const attachment = await AttachmentService.uploadRequestAttachment(
  requestId,
  file
);
```

## Troubleshooting

### SharePoint Lists Not Found
- Verify list names exactly match: "Admin Requests" and "Stationery Items"
- Check that lists are created in the same site where the web part is deployed

### Attachment Upload Fails
- Ensure user has Edit permissions on the SharePoint list
- Check file size limits (SharePoint default: 250MB)
- Verify file type is not blocked by organization policies

### Form Not Loading
- Check browser console for error messages
- Verify PnP/sp is properly initialized with context
- Ensure SharePoint list columns match the model definitions

## Contributing

Feel free to submit issues and enhancement requests!

## License

MIT License - see LICENSE file for details

## Support

For issues, questions, or suggestions, please create an issue in the GitHub repository.
