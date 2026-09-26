import * as React from 'react';
import {
  Stack,
  TextField,
  Dropdown,
  IDropdownOption,
  DatePicker,
  Checkbox,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Label,
  Icon
} from '@fluentui/react';
import {
  IAdminRequest,
  RequestStatus,
  PriorityLevel,
  RequestType
} from '../../models/IAdminRequest';
import { getRequestTypeConfig, IFormField } from '../../config/requestTypeRegistry';
import { RequestService } from '../../services/RequestService';
import { AttachmentService } from '../../services/attachmentService';
import styles from './DynamicRequestForm.module.scss';

export interface IDynamicRequestFormProps {
  requestType: RequestType;
  requestId?: string;
  onSave?: (request: IAdminRequest) => void;
  onCancel?: () => void;
}

export const DynamicRequestForm: React.FC<IDynamicRequestFormProps> = (props) => {
  const config = getRequestTypeConfig(props.requestType);
  const [request, setRequest] = React.useState<Partial<IAdminRequest>>({
    title: '',
    description: '',
    requestType: props.requestType,
    status: config.defaultStatus || RequestStatus.Pending,
    priority: config.defaultPriority || PriorityLevel.Medium,
    details: {},
    attachments: []
  });
  const [isLoading, setIsLoading] = React.useState(!!props.requestId);
  const [isSaving, setIsSaving] = React.useState(false);
  const [message, setMessage] = React.useState<{ text: string; type: MessageBarType } | null>(null);
  const [selectedFiles, setSelectedFiles] = React.useState<File[]>([]);

  React.useEffect(() => {
    if (!props.requestId) return;
    (async () => {
      setIsLoading(true);
      try {
        const data = await RequestService.getRequest(props.requestId!);
        const attachments = await AttachmentService.getRequestAttachments(props.requestId!);
        setRequest({ ...data, attachments });
      } catch {
        setMessage({ text: 'Error loading request', type: MessageBarType.error });
      } finally {
        setIsLoading(false);
      }
    })();
  }, [props.requestId]);

  const getValue = (field: IFormField): any => {
    if (field.inDetails) {
      return (request.details || {})[field.key];
    }
    return (request as any)[field.key];
  };

  const setValue = (field: IFormField, value: any) => {
    setRequest((prev) => {
      if (field.inDetails) {
        return {
          ...prev,
          details: { ...(prev.details || {}), [field.key]: value }
        };
      }
      return { ...prev, [field.key]: value };
    });
  };

  const validate = (): boolean => {
    for (const field of config.formFields) {
      if (!field.required) continue;
      const v = getValue(field);
      if (v === undefined || v === null || v === '') {
        setMessage({ text: `Please fill in required field: ${field.label}`, type: MessageBarType.warning });
        return false;
      }
    }
    return true;
  };

  const handleSave = async () => {
    if (!validate()) return;
    setIsSaving(true);
    try {
      let saved: IAdminRequest;
      if (request.id) {
        await RequestService.updateRequest(request.id, request);
        saved = request as IAdminRequest;
      } else {
        saved = await RequestService.createRequest(request);
      }

      if (selectedFiles.length && saved.id) {
        for (const file of selectedFiles) {
          try {
            await AttachmentService.uploadRequestAttachment(saved.id, file);
          } catch (e) {
            console.error(e);
          }
        }
      }

      setMessage({ text: 'Request saved successfully', type: MessageBarType.success });
      setSelectedFiles([]);
      if (props.onSave) props.onSave(saved);
    } catch {
      setMessage({ text: 'Error saving request', type: MessageBarType.error });
    } finally {
      setIsSaving(false);
    }
  };

  const renderField = (field: IFormField) => {
    const value = getValue(field);
    switch (field.type) {
      case 'multiline':
        return (
          <TextField
            key={field.key}
            label={field.label}
            multiline
            rows={3}
            required={field.required}
            placeholder={field.placeholder}
            value={value || ''}
            onChange={(_, v) => setValue(field, v)}
          />
        );
      case 'number':
      case 'currency':
        return (
          <TextField
            key={field.key}
            label={field.label}
            type="number"
            required={field.required}
            placeholder={field.placeholder}
            value={value != null ? String(value) : ''}
            onChange={(_, v) => setValue(field, v ? parseFloat(v) : undefined)}
          />
        );
      case 'date':
        return (
          <DatePicker
            key={field.key}
            label={field.label}
            value={value ? new Date(value) : undefined}
            onSelectDate={(d) => setValue(field, d)}
          />
        );
      case 'dropdown':
        return (
          <Dropdown
            key={field.key}
            label={field.label}
            required={field.required}
            options={(field.options || []) as IDropdownOption[]}
            selectedKey={value}
            onChange={(_, opt) => setValue(field, opt?.key)}
          />
        );
      case 'email':
        return (
          <TextField
            key={field.key}
            label={field.label}
            type="email"
            required={field.required}
            placeholder={field.placeholder}
            value={value || ''}
            onChange={(_, v) => setValue(field, v)}
          />
        );
      case 'checkbox':
        return (
          <Checkbox
            key={field.key}
            label={field.label}
            checked={!!value}
            onChange={(_, checked) => setValue(field, !!checked)}
          />
        );
      default:
        return (
          <TextField
            key={field.key}
            label={field.label}
            required={field.required}
            placeholder={field.placeholder}
            value={value || ''}
            onChange={(_, v) => setValue(field, v)}
          />
        );
    }
  };

  if (isLoading) {
    return <Spinner size={SpinnerSize.large} label="Loading request…" />;
  }

  return (
    <Stack className={styles.formContainer} tokens={{ childrenGap: 12 }}>
      <h2>{config.displayName}</h2>
      <p className={styles.subtitle}>{config.description}</p>

      {message && (
        <MessageBar messageBarType={message.type} onDismiss={() => setMessage(null)}>
          {message.text}
        </MessageBar>
      )}

      <Stack tokens={{ childrenGap: 12 }}>
        {config.formFields.map(renderField)}

        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <TextField
            label="Requester Name"
            value={request.requesterName || ''}
            onChange={(_, v) => setRequest((p) => ({ ...p, requesterName: v }))}
            styles={{ root: { flex: 1 } }}
          />
          <TextField
            label="Requester Email"
            type="email"
            value={request.requesterEmail || ''}
            onChange={(_, v) => setRequest((p) => ({ ...p, requesterEmail: v }))}
            styles={{ root: { flex: 1 } }}
          />
        </Stack>

        <div>
          <Label>Attachments</Label>
          <input
            type="file"
            multiple
            onChange={(e) => {
              if (e.target.files) setSelectedFiles(Array.from(e.target.files));
            }}
          />
          {selectedFiles.map((f) => (
            <div key={f.name} className={styles.fileItem}>
              <Icon iconName="Document" /> {f.name}
            </div>
          ))}
          {request.attachments && request.attachments.length > 0 && (
            <Stack tokens={{ childrenGap: 4 }}>
              <Label>Existing</Label>
              {request.attachments.map((a) => (
                <a key={a.id} href={a.fileUrl} target="_blank" rel="noopener noreferrer" className={styles.fileLink}>
                  <Icon iconName="Document" /> {a.fileName}
                </a>
              ))}
            </Stack>
          )}
        </div>

        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton text="Save" onClick={handleSave} disabled={isSaving} />
          <DefaultButton text="Cancel" onClick={props.onCancel} disabled={isSaving} />
        </Stack>
      </Stack>
    </Stack>
  );
};
