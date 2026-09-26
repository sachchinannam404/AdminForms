import * as React from 'react';
import {
  Stack,
  PrimaryButton,
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Dialog,
  DialogType,
  TextField,
  Dropdown,
  DatePicker,
  Label
} from '@fluentui/react';
import { RequestType } from '../../models/common/enums';
import { IChildItem } from '../../models/IChildItem';
import { getRequestTypeConfig, IFormField } from '../../config/requestTypeRegistry';
import { RequestService } from '../../services/RequestService';

export interface IChildItemsListProps {
  requestId: string;
  requestType: RequestType;
}

export const ChildItemsList: React.FC<IChildItemsListProps> = (props) => {
  const config = getRequestTypeConfig(props.requestType);
  const [items, setItems] = React.useState<IChildItem[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [message, setMessage] = React.useState<{ text: string; type: MessageBarType } | null>(null);
  const [showForm, setShowForm] = React.useState(false);
  const [editing, setEditing] = React.useState<IChildItem | undefined>();
  const [formData, setFormData] = React.useState<Partial<IChildItem>>({});
  const [showDelete, setShowDelete] = React.useState(false);
  const [deleteId, setDeleteId] = React.useState<string | undefined>();

  const selectionRef = React.useRef(
    new Selection({
      onSelectionChanged: () => {
        const sel = selectionRef.current.getSelection() as IChildItem[];
        setEditing(sel.length ? sel[0] : undefined);
      }
    })
  );

  const load = React.useCallback(async () => {
    if (!config.supportsChildren) return;
    setIsLoading(true);
    try {
      const list = await RequestService.getChildItems(props.requestType, props.requestId);
      setItems(list);
    } catch {
      setMessage({ text: 'Failed to load items', type: MessageBarType.error });
    } finally {
      setIsLoading(false);
    }
  }, [props.requestId, props.requestType, config.supportsChildren]);

  React.useEffect(() => {
    load();
  }, [load]);

  if (!config.supportsChildren) {
    return null;
  }

  const columns: IColumn[] = [
    { key: 'name', name: 'Name', fieldName: 'itemName', minWidth: 140, isResizable: true },
    { key: 'category', name: 'Category', fieldName: 'category', minWidth: 100, isResizable: true },
    {
      key: 'qty',
      name: 'Qty',
      minWidth: 80,
      onRender: (i: IChildItem) => `${i.quantity ?? ''} ${i.unit || ''}`
    },
    {
      key: 'price',
      name: 'Total',
      minWidth: 90,
      onRender: (i: IChildItem) =>
        i.totalPrice != null ? `$${Number(i.totalPrice).toFixed(2)}` : '—'
    },
    { key: 'status', name: 'Status', fieldName: 'status', minWidth: 90, isResizable: true }
  ];

  const openAdd = () => {
    setEditing(undefined);
    setFormData({ requestId: props.requestId, quantity: 1, unit: 'pcs' });
    setShowForm(true);
  };

  const openEdit = () => {
    if (!editing) return;
    setFormData({ ...editing });
    setShowForm(true);
  };

  const saveItem = async () => {
    if (!formData.itemName) {
      setMessage({ text: 'Item name is required', type: MessageBarType.warning });
      return;
    }
    try {
      if (formData.id) {
        await RequestService.updateChildItem(props.requestType, formData.id, formData);
      } else {
        await RequestService.createChildItem(props.requestType, {
          ...formData,
          requestId: props.requestId
        });
      }
      setShowForm(false);
      setMessage({ text: 'Item saved', type: MessageBarType.success });
      load();
    } catch {
      setMessage({ text: 'Error saving item', type: MessageBarType.error });
    }
  };

  const confirmDelete = async () => {
    if (!deleteId) return;
    try {
      await RequestService.deleteChildItem(props.requestType, deleteId);
      setShowDelete(false);
      setDeleteId(undefined);
      setMessage({ text: 'Item deleted', type: MessageBarType.success });
      load();
    } catch {
      setMessage({ text: 'Error deleting item', type: MessageBarType.error });
    }
  };

  const renderChildField = (field: IFormField) => {
    const value = (formData as any)[field.key];
    if (field.type === 'dropdown') {
      return (
        <Dropdown
          key={field.key}
          label={field.label}
          options={field.options || []}
          selectedKey={value}
          onChange={(_, opt) => setFormData((p) => ({ ...p, [field.key]: opt?.key }))}
        />
      );
    }
    if (field.type === 'date') {
      return (
        <DatePicker
          key={field.key}
          label={field.label}
          value={value ? new Date(value) : undefined}
          onSelectDate={(d) => setFormData((p) => ({ ...p, [field.key]: d }))}
        />
      );
    }
    if (field.type === 'number' || field.type === 'currency') {
      return (
        <TextField
          key={field.key}
          label={field.label}
          type="number"
          required={field.required}
          value={value != null ? String(value) : ''}
          onChange={(_, v) =>
            setFormData((p) => ({ ...p, [field.key]: v ? parseFloat(v) : undefined }))
          }
        />
      );
    }
    if (field.type === 'multiline') {
      return (
        <TextField
          key={field.key}
          label={field.label}
          multiline
          rows={2}
          value={value || ''}
          onChange={(_, v) => setFormData((p) => ({ ...p, [field.key]: v }))}
        />
      );
    }
    return (
      <TextField
        key={field.key}
        label={field.label}
        required={field.required}
        placeholder={field.placeholder}
        value={value || ''}
        onChange={(_, v) => setFormData((p) => ({ ...p, [field.key]: v }))}
      />
    );
  };

  if (isLoading) {
    return <Spinner size={SpinnerSize.large} label="Loading items…" />;
  }

  if (showForm) {
    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <h3>{formData.id ? 'Edit item' : 'Add item'}</h3>
        {(config.childFormFields || []).map(renderChildField)}
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton text="Save" onClick={saveItem} />
          <DefaultButton text="Cancel" onClick={() => setShowForm(false)} />
        </Stack>
      </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <Label>Line items</Label>
      {message && (
        <MessageBar messageBarType={message.type} onDismiss={() => setMessage(null)}>
          {message.text}
        </MessageBar>
      )}
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton text="Add item" onClick={openAdd} />
        <DefaultButton text="Edit" onClick={openEdit} disabled={!editing} />
        <DefaultButton
          text="Delete"
          onClick={() => {
            if (editing?.id) {
              setDeleteId(editing.id);
              setShowDelete(true);
            }
          }}
          disabled={!editing}
        />
      </Stack>
      <DetailsList
        items={items}
        columns={columns}
        layoutMode={DetailsListLayoutMode.justified}
        selection={selectionRef.current}
        selectionMode={SelectionMode.single}
        selectionPreservedOnEmptyClick
      />
      <Dialog
        hidden={!showDelete}
        onDismiss={() => setShowDelete(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Confirm delete',
          subText: 'Delete this line item?'
        }}
      >
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton text="Delete" onClick={confirmDelete} />
          <DefaultButton text="Cancel" onClick={() => setShowDelete(false)} />
        </Stack>
      </Dialog>
    </Stack>
  );
};
