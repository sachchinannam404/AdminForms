import * as React from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn,
  Button,
  Stack,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Dialog,
  DialogType,
  DefaultButton,
  PrimaryButton
} from '@fluentui/react';
import { IStationeryItem } from '../models/IStationeryItem';
import { SharePointService } from '../services/sharePointService';
import { StationeryItemForm } from './StationeryItemForm';
import styles from './StationeryItemsList.module.scss';

interface IStationeryItemsListProps {
  requestId: string;
}

interface IStationeryItemsListState {
  items: IStationeryItem[];
  isLoading: boolean;
  message: string;
  messageType: MessageBarType;
  showForm: boolean;
  selectedItem?: IStationeryItem;
  showDeleteDialog: boolean;
  deleteItemId?: string;
}

const columns: IColumn[] = [
  {
    key: 'column1',
    name: 'Item Name',
    fieldName: 'itemName',
    minWidth: 150,
    isResizable: true
  },
  {
    key: 'column2',
    name: 'Category',
    fieldName: 'category',
    minWidth: 120,
    isResizable: true
  },
  {
    key: 'column3',
    name: 'Quantity',
    fieldName: 'quantity',
    minWidth: 100,
    isResizable: true,
    onRender: (item: IStationeryItem) => `${item.quantity} ${item.unit || 'pcs'}`
  },
  {
    key: 'column4',
    name: 'Unit Price',
    fieldName: 'unitPrice',
    minWidth: 100,
    isResizable: true,
    onRender: (item: IStationeryItem) => `$${item.unitPrice?.toFixed(2) || '0.00'}`
  },
  {
    key: 'column5',
    name: 'Total Price',
    fieldName: 'totalPrice',
    minWidth: 100,
    isResizable: true,
    onRender: (item: IStationeryItem) => `$${item.totalPrice?.toFixed(2) || '0.00'}`
  },
  {
    key: 'column6',
    name: 'Status',
    fieldName: 'status',
    minWidth: 100,
    isResizable: true
  },
  {
    key: 'column7',
    name: 'Expected Delivery',
    fieldName: 'expectedDeliveryDate',
    minWidth: 140,
    isResizable: true,
    onRender: (item: IStationeryItem) =>
      item.expectedDeliveryDate
        ? new Date(item.expectedDeliveryDate).toLocaleDateString()
        : ''
  }
];

export const StationeryItemsList: React.FC<IStationeryItemsListProps> = (props) => {
  const [state, setState] = React.useState<IStationeryItemsListState>({
    items: [],
    isLoading: false,
    message: '',
    messageType: MessageBarType.info,
    showForm: false,
    showDeleteDialog: false
  });

  const selection = new Selection({
    onSelectionChanged: () => {
      const selectedItems = selection.getSelection() as IStationeryItem[];
      if (selectedItems.length > 0) {
        setState((prev) => ({
          ...prev,
          selectedItem: selectedItems[0]
        }));
      }
    }
  });

  // Load items on component mount
  React.useEffect(() => {
    loadItems();
  }, [props.requestId]);

  const loadItems = async () => {
    setState((prev) => ({ ...prev, isLoading: true }));
    try {
      const items = await SharePointService.getStationeryItemsByRequest(props.requestId);
      setState((prev) => ({
        ...prev,
        items,
        isLoading: false
      }));
    } catch (error) {
      setState((prev) => ({
        ...prev,
        isLoading: false,
        message: 'Error loading items',
        messageType: MessageBarType.error
      }));
    }
  };

  const handleAddItem = () => {
    setState((prev) => ({
      ...prev,
      showForm: true,
      selectedItem: undefined
    }));
  };

  const handleEditItem = () => {
    if (state.selectedItem) {
      setState((prev) => ({
        ...prev,
        showForm: true
      }));
    }
  };

  const handleDeleteItem = () => {
    if (state.selectedItem) {
      setState((prev) => ({
        ...prev,
        showDeleteDialog: true,
        deleteItemId: state.selectedItem?.id
      }));
    }
  };

  const confirmDelete = async () => {
    if (state.deleteItemId) {
      try {
        await SharePointService.deleteStationeryItem(state.deleteItemId);
        setState((prev) => ({
          ...prev,
          showDeleteDialog: false,
          deleteItemId: undefined,
          selectedItem: undefined,
          message: 'Item deleted successfully',
          messageType: MessageBarType.success
        }));
        loadItems();
      } catch (error) {
        setState((prev) => ({
          ...prev,
          showDeleteDialog: false,
          message: 'Error deleting item',
          messageType: MessageBarType.error
        }));
      }
    }
  };

  const handleFormSave = () => {
    setState((prev) => ({
      ...prev,
      showForm: false,
      message: 'Item saved successfully',
      messageType: MessageBarType.success
    }));
    loadItems();
  };

  const handleFormCancel = () => {
    setState((prev) => ({
      ...prev,
      showForm: false,
      selectedItem: undefined
    }));
  };

  if (state.isLoading) {
    return <Spinner size={SpinnerSize.large} label="Loading items..." />;
  }

  if (state.showForm) {
    return (
      <StationeryItemForm
        requestId={props.requestId}
        itemId={state.selectedItem?.id}
        onSave={handleFormSave}
        onCancel={handleFormCancel}
      />
    );
  }

  return (
    <Stack className={styles.container}>
      <h3>Stationery Items</h3>

      {state.message && (
        <MessageBar messageBarType={state.messageType}>{state.message}</MessageBar>
      )}

      <Stack horizontal tokens={{ childrenGap: 10 }} className={styles.buttonGroup}>
        <Button text="Add Item" onClick={handleAddItem} primary />
        <Button
          text="Edit"
          onClick={handleEditItem}
          disabled={!state.selectedItem}
        />
        <Button
          text="Delete"
          onClick={handleDeleteItem}
          disabled={!state.selectedItem}
        />
      </Stack>

      <DetailsList
        items={state.items}
        columns={columns}
        setKey="set"
        layoutMode={DetailsListLayoutMode.fixedColumns}
        selection={selection}
        selectionPreservedOnEmptyClick
        selectionMode={SelectionMode.single}
        isHeaderVisible
        onActiveItemChanged={() => {}}
      />

      <Dialog
        hidden={!state.showDeleteDialog}
        onDismiss={() => setState((prev) => ({ ...prev, showDeleteDialog: false }))}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Confirm Delete',
          closeButtonAriaLabel: 'Close',
          subText: 'Are you sure you want to delete this item?'
        }}
      >
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton text="Delete" onClick={confirmDelete} />
          <DefaultButton
            text="Cancel"
            onClick={() => setState((prev) => ({ ...prev, showDeleteDialog: false }))}
          />
        </Stack>
      </Dialog>
    </Stack>
  );
};
