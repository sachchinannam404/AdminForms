import * as React from 'react';
import {
  TextField,
  Button,
  Stack,
  MessageBar,
  MessageBarType,
  Dropdown,
  IDropdownOption,
  DatePicker,
  Spinner,
  SpinnerSize,
  Label,
  Icon
} from '@fluentui/react';
import { IStationeryItem, ItemCategory, ItemStatus } from '../models/IStationeryItem';
import { SharePointService } from '../services/sharePointService';
import { AttachmentService } from '../services/attachmentService';
import styles from './StationeryItemForm.module.scss';

interface IStationeryItemFormProps {
  requestId: string;
  itemId?: string;
  onSave?: () => void;
  onCancel?: () => void;
}

interface IStationeryItemFormState {
  item: IStationeryItem;
  isLoading: boolean;
  isSaving: boolean;
  message: string;
  messageType: MessageBarType;
  selectedFiles: File[];
}

const categoryOptions: IDropdownOption[] = [
  { key: ItemCategory.Paper, text: 'Paper' },
  { key: ItemCategory.Pens, text: 'Pens' },
  { key: ItemCategory.Notebooks, text: 'Notebooks' },
  { key: ItemCategory.Folders, text: 'Folders' },
  { key: ItemCategory.Tape, text: 'Tape' },
  { key: ItemCategory.Ink, text: 'Ink' },
  { key: ItemCategory.Other, text: 'Other' }
];

const statusOptions: IDropdownOption[] = [
  { key: ItemStatus.Pending, text: 'Pending' },
  { key: ItemStatus.Ordered, text: 'Ordered' },
  { key: ItemStatus.InStock, text: 'In Stock' },
  { key: ItemStatus.Delivered, text: 'Delivered' }
];

export const StationeryItemForm: React.FC<IStationeryItemFormProps> = (props) => {
  const [state, setState] = React.useState<IStationeryItemFormState>({
    item: {
      requestId: props.requestId,
      itemName: '',
      category: ItemCategory.Other,
      quantity: 1,
      unit: 'pcs',
      unitPrice: 0,
      totalPrice: 0,
      description: '',
      status: ItemStatus.Pending,
      vendorName: '',
      vendorEmail: '',
      remarks: '',
      attachments: []
    },
    isLoading: false,
    isSaving: false,
    message: '',
    messageType: MessageBarType.info,
    selectedFiles: []
  });

  // Load item data if editing
  React.useEffect(() => {
    if (props.itemId) {
      loadItem(props.itemId);
    }
  }, [props.itemId]);

  const loadItem = async (itemId: string) => {
    setState((prev) => ({ ...prev, isLoading: true }));
    try {
      const item = await SharePointService.getStationeryItem(itemId);
      setState((prev) => ({
        ...prev,
        item,
        isLoading: false
      }));
    } catch (error) {
      setState((prev) => ({
        ...prev,
        isLoading: false,
        message: 'Error loading item',
        messageType: MessageBarType.error
      }));
    }
  };

  const handleSave = async () => {
    if (!state.item.itemName || !state.item.quantity) {
      setState((prev) => ({
        ...prev,
        message: 'Please fill in required fields',
        messageType: MessageBarType.warning
      }));
      return;
    }

    setState((prev) => ({ ...prev, isSaving: true }));
    try {
      // Calculate total price
      const totalPrice = state.item.quantity * (state.item.unitPrice || 0);
      state.item.totalPrice = totalPrice;

      if (state.item.id) {
        // Update existing item
        await SharePointService.updateStationeryItem(state.item.id, state.item);
      } else {
        // Create new item
        const result = await SharePointService.createStationeryItem(state.item);
        state.item.id = result.ID?.toString();
      }

      // Upload attachments
      if (state.selectedFiles.length > 0 && state.item.id) {
        for (const file of state.selectedFiles) {
          try {
            const attachment = await AttachmentService.uploadItemAttachment(
              state.item.id,
              file
            );
            state.item.attachments?.push(attachment);
          } catch (error) {
            console.error(`Error uploading file ${file.name}:`, error);
          }
        }
      }

      setState((prev) => ({
        ...prev,
        isSaving: false,
        message: 'Item saved successfully',
        messageType: MessageBarType.success,
        selectedFiles: []
      }));

      if (props.onSave) {
        props.onSave();
      }
    } catch (error) {
      setState((prev) => ({
        ...prev,
        isSaving: false,
        message: 'Error saving item',
        messageType: MessageBarType.error
      }));
    }
  };

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files) {
      setState((prev) => ({
        ...prev,
        selectedFiles: Array.from(event.target.files!)
      }));
    }
  };

  const handleFieldChange = (fieldName: string, value: any) => {
    setState((prev) => ({
      ...prev,
      item: {
        ...prev.item,
        [fieldName]: value
      }
    }));
  };

  if (state.isLoading) {
    return <Spinner size={SpinnerSize.large} label="Loading item..." />;
  }

  return (
    <Stack className={styles.formContainer}>
      <h2>Stationery Item Form</h2>

      {state.message && (
        <MessageBar messageBarType={state.messageType}>{state.message}</MessageBar>
      )}

      <Stack tokens={{ childrenGap: 15 }}>
        {/* Item Name */}
        <TextField
          label="Item Name"
          placeholder="Enter item name"
          required
          value={state.item.itemName}
          onChange={(e, value) => handleFieldChange('itemName', value)}
        />

        {/* Category and Quantity */}
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <Dropdown
            label="Category"
            options={categoryOptions}
            selectedKey={state.item.category}
            onChange={(e, option) => handleFieldChange('category', option?.key)}
            style={{ flex: 1 }}
          />
          <TextField
            label="Quantity"
            type="number"
            required
            style={{ flex: 1 }}
            value={state.item.quantity?.toString() || ''}
            onChange={(e, value) =>
              handleFieldChange('quantity', value ? parseInt(value) : 0)
            }
          />
        </Stack>

        {/* Unit and Pricing */}
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <TextField
            label="Unit"
            placeholder="e.g., pcs, ream, box"
            style={{ flex: 1 }}
            value={state.item.unit}
            onChange={(e, value) => handleFieldChange('unit', value)}
          />
          <TextField
            label="Unit Price"
            type="number"
            style={{ flex: 1 }}
            value={state.item.unitPrice?.toString() || ''}
            onChange={(e, value) =>
              handleFieldChange('unitPrice', value ? parseFloat(value) : 0)
            }
          />
          <TextField
            label="Total Price"
            type="number"
            disabled
            style={{ flex: 1 }}
            value={(
              (state.item.quantity || 0) * (state.item.unitPrice || 0)
            ).toFixed(2)}
          />
        </Stack>

        {/* Description */}
        <TextField
          label="Description"
          placeholder="Enter item description"
          multiline
          rows={3}
          value={state.item.description}
          onChange={(e, value) => handleFieldChange('description', value)}
        />

        {/* Status and Vendor */}
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <Dropdown
            label="Status"
            options={statusOptions}
            selectedKey={state.item.status}
            onChange={(e, option) => handleFieldChange('status', option?.key)}
            style={{ flex: 1 }}
          />
        </Stack>

        {/* Vendor Information */}
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <TextField
            label="Vendor Name"
            placeholder="Enter vendor name"
            style={{ flex: 1 }}
            value={state.item.vendorName}
            onChange={(e, value) => handleFieldChange('vendorName', value)}
          />
          <TextField
            label="Vendor Email"
            placeholder="Enter vendor email"
            type="email"
            style={{ flex: 1 }}
            value={state.item.vendorEmail}
            onChange={(e, value) => handleFieldChange('vendorEmail', value)}
          />
        </Stack>

        {/* Delivery Dates */}
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <DatePicker
            label="Expected Delivery Date"
            value={state.item.expectedDeliveryDate}
            onSelectDate={(date) => handleFieldChange('expectedDeliveryDate', date)}
            style={{ flex: 1 }}
          />
          <DatePicker
            label="Actual Delivery Date"
            value={state.item.actualDeliveryDate}
            onSelectDate={(date) => handleFieldChange('actualDeliveryDate', date)}
            style={{ flex: 1 }}
          />
        </Stack>

        {/* Remarks */}
        <TextField
          label="Remarks"
          placeholder="Enter any additional remarks"
          multiline
          rows={3}
          value={state.item.remarks}
          onChange={(e, value) => handleFieldChange('remarks', value)}
        />

        {/* Attachments */}
        <div>
          <Label>Attachments</Label>
          <input
            type="file"
            multiple
            onChange={handleFileChange}
            style={{ marginBottom: 10 }}
          />
          {state.selectedFiles.length > 0 && (
            <Stack tokens={{ childrenGap: 5 }}>
              {state.selectedFiles.map((file) => (
                <div key={file.name} className={styles.fileItem}>
                  <Icon iconName="Document" />
                  <span>{file.name}</span>
                </div>
              ))}
            </Stack>
          )}
          {state.item.attachments && state.item.attachments.length > 0 && (
            <div>
              <Label>Existing Attachments:</Label>
              <Stack tokens={{ childrenGap: 5 }}>
                {state.item.attachments.map((att) => (
                  <a
                    key={att.id}
                    href={att.fileUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className={styles.fileLink}
                  >
                    <Icon iconName="Document" />
                    {att.fileName}
                  </a>
                ))}
              </Stack>
            </div>
          )}
        </div>

        {/* Action Buttons */}
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <Button
            text="Save"
            onClick={handleSave}
            disabled={state.isSaving}
            primary
          />
          <Button
            text="Cancel"
            onClick={props.onCancel}
            disabled={state.isSaving}
          />
        </Stack>
      </Stack>
    </Stack>
  );
};
