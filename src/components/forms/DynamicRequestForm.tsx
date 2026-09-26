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
import { validateRequiredFields, validateEmail } from '../../utils/validation';
import { useAttachments } from '../../hooks/useAttachments';
import { PeopleField } from '../common/PeopleField';
import { exceedsApprovalThreshold } from '../../config/requestTypeRegistry';
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
  const [fieldErrors, setFieldErrors] = React.useState<Record<string, string>>({});

  const {
    attachments,
    pendingFiles,
    setPendingFiles,
    uploadPending,
    isLoading: attachmentsLoading
  } = useAttachments('Admin Requests', props.requestId);

  React.useEffect(() => {
    if (!props.requestId) return;
    (async () => {
      setIsLoading(true);
      try {
        const data = await RequestService.getRequest(props.requestId!);
        setRequest({ ...data });
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
    setFieldErrors((prev) => {
      const next = { ...prev };
      delete next[field.key];
      return next;
    });
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
    const result = validateRequiredFields(config.formFields, getValue);
    const errors = { ...result.errors };
    if (request.requesterEmail && !validateEmail(request.requesterEmail)) {
      errors.requesterEmail = 'Requester email must be valid';
    }
    setFieldErrors(errors);
    if (Object.keys(errors).length) {
      setMessage({
        text: result.firstError || errors.requesterEmail || 'Please fix validation errors',
        type: MessageBarType.warning
      });
      return false;
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

      if (pendingFiles.length && saved.id) {
        await uploadPending(saved.id);
      }

      setMessage({ text: 'Request saved successfully', type: MessageBarType.success });
      if (props.onSave) props.onSave(saved);
    } catch {
      setMessage({ text: 'Error saving request', type: MessageBarType.error });
    } finally {
      setIsSaving(false);
    }
  };

  const renderField = (field: IFormField) => {
    const value = getValue(field);
    const err = fieldErrors[field.key];
    const requiredMark = field.required ? ' *' : '';

    switch (field.type) {
      case 'multiline':
        return (
          <TextField
            key={field.key}
            label={`${field.label}${requiredMark}`}
            multiline
            rows={3}
            required={field.required}
            placeholder={field.placeholder}
            value={value || ''}
            errorMessage={err}
            onChange={(_, v) => setValue(field, v)}
          />
        );
      case 'number':
      case 'currency':
        return (
          <TextField
            key={field.key}
            label={`${field.label}${requiredMark}`}
            type="number"
            required={field.required}
            placeholder={field.placeholder}
            value={value != null ? String(value) : ''}
            errorMessage={err}
            onChange={(_, v) => setValue(field, v ? parseFloat(v) : undefined)}
          />
        );
      case 'date':
        return (
          <DatePicker
            key={field.key}
            label={`${field.label}${requiredMark}`}
            value={value ? new Date(value) : undefined}
            onSelectDate={(d) => setValue(field, d)}
            isRequired={field.required}
          />
        );
      case 'dropdown':
        return (
          <Dropdown
            key={field.key}
            label={`${field.label}${requiredMark}`}
            required={field.required}
            options={(field.options || []) as IDropdownOption[]}
            selectedKey={value}
            errorMessage={err}
            onChange={(_, opt) => setValue(field, opt?.key)}
          />
        );
      case 'email':
        return (
          <TextField
            key={field.key}
            label={`${field.label}${requiredMark}`}
            type="email"
            required={field.required}
            placeholder={field.placeholder}
            value={value || ''}
            errorMessage={err}
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
            label={`${field.label}${requiredMark}`}
            required={field.required}
            placeholder={field.placeholder}
            value={value || ''}
            errorMessage={err}
            onChange={(_, v) => setValue(field, v)}
          />
        );
    }
  };

  if (isLoading || attachmentsLoading) {
    return <Spinner size={SpinnerSize.large} label="Loading request…" />;
  }

  const thresholdWarn = exceedsApprovalThreshold(props.requestType, request.totalBudget);

  return (
    <Stack className={styles.formContainer} tokens={{ childrenGap: 12 }}>
      <h2>{config.displayName}</h2>
      <p className={styles.subtitle}>{config.description}</p>

      {message && (
        <MessageBar messageBarType={message.type} onDismiss={() => setMessage(null)}>
          {message.text}
        </MessageBar>
      )}

      {thresholdWarn && (
        <MessageBar messageBarType={MessageBarType.warning}>
          Budget meets or exceeds the dual-approval threshold
          {config.approvalThreshold != null ? ` ($${config.approvalThreshold})` : ''}.
        </MessageBar>
      )}

      <Stack tokens={{ childrenGap: 12 }}>
        {config.formFields.map(renderField)}

        <PeopleField
          label="Requester"
          value={{
            displayName: request.requesterName || '',
            email: request.requesterEmail
          }}
          onChange={(p) =>
            setRequest((prev) => ({
              ...prev,
              requesterName: p.displayName,
              requesterEmail: p.email
            }))
          }
          errorMessage={fieldErrors.requesterEmail}
        />

        <div>
          <Label>Attachments</Label>
          <input
            type="file"
            multiple
            onChange={(e) => {
              if (e.target.files) setPendingFiles(Array.from(e.target.files));
            }}
          />
          {pendingFiles.map((f) => (
            <div key={f.name} className={styles.fileItem}>
              <Icon iconName="Document" /> {f.name} (pending)
            </div>
          ))}
          {attachments.length > 0 && (
            <Stack tokens={{ childrenGap: 4 }}>
              <Label>Existing attachments</Label>
              {attachments.map((a) => (
                <a
                  key={a.id}
                  href={a.fileUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  className={styles.fileLink}
                >
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
