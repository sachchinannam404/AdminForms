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
import { IAdminRequest, RequestStatus, PriorityLevel } from '../models/IAdminRequest';
import { SharePointService } from '../services/sharePointService';
import { AttachmentService } from '../services/attachmentService';
import styles from './AdminRequestForm.module.scss';

interface IAdminRequestFormProps {
  requestId?: string;
  onSave?: (request: IAdminRequest) => void;
  onCancel?: () => void;
}

interface IAdminRequestFormState {
  request: IAdminRequest;
  isLoading: boolean;
  isSaving: boolean;
  message: string;
  messageType: MessageBarType;
  selectedFiles: File[];
}

const statusOptions: IDropdownOption[] = [
  { key: RequestStatus.Pending, text: 'Pending' },
  { key: RequestStatus.Approved, text: 'Approved' },
  { key: RequestStatus.Rejected, text: 'Rejected' },
  { key: RequestStatus.InProgress, text: 'In Progress' },
  { key: RequestStatus.Completed, text: 'Completed' }
];

const priorityOptions: IDropdownOption[] = [
  { key: PriorityLevel.Low, text: 'Low' },
  { key: PriorityLevel.Medium, text: 'Medium' },
  { key: PriorityLevel.High, text: 'High' },
  { key: PriorityLevel.Urgent, text: 'Urgent' }
];

export const AdminRequestForm: React.FC<IAdminRequestFormProps> = (props) => {
  const [state, setState] = React.useState<IAdminRequestFormState>({
    request: {
      title: '',
      description: '',
      requesterName: '',
      requesterEmail: '',
      department: '',
      status: RequestStatus.Pending,
      priority: PriorityLevel.Medium,
      attachments: [],
      comments: ''
    },
    isLoading: false,
    isSaving: false,
    message: '',
    messageType: MessageBarType.info,
    selectedFiles: []
  });

  // Load request data if editing
  React.useEffect(() => {
    if (props.requestId) {
      loadRequest(props.requestId);
    }
  }, [props.requestId]);

  const loadRequest = async (requestId: string) => {
    setState((prev) => ({ ...prev, isLoading: true }));
    try {
      const request = await SharePointService.getAdminRequest(requestId);
      setState((prev) => ({
        ...prev,
        request,
        isLoading: false
      }));
    } catch (error) {
      setState((prev) => ({
        ...prev,
        isLoading: false,
        message: 'Error loading request',
        messageType: MessageBarType.error
      }));
    }
  };

  const handleSave = async () => {
    if (!state.request.title || !state.request.description) {
      setState((prev) => ({
        ...prev,
        message: 'Please fill in required fields',
        messageType: MessageBarType.warning
      }));
      return;
    }

    setState((prev) => ({ ...prev, isSaving: true }));
    try {
      if (state.request.id) {
        // Update existing request
        await SharePointService.updateAdminRequest(state.request.id, state.request);
      } else {
        // Create new request
        const result = await SharePointService.createAdminRequest(state.request);
        state.request.id = result.ID?.toString();
      }

      // Upload attachments
      if (state.selectedFiles.length > 0 && state.request.id) {
        for (const file of state.selectedFiles) {
          try {
            const attachment = await AttachmentService.uploadRequestAttachment(
              state.request.id,
              file
            );
            state.request.attachments?.push(attachment);
          } catch (error) {
            console.error(`Error uploading file ${file.name}:`, error);
          }
        }
      }

      setState((prev) => ({
        ...prev,
        isSaving: false,
        message: 'Request saved successfully',
        messageType: MessageBarType.success,
        selectedFiles: []
      }));

      if (props.onSave) {
        props.onSave(state.request);
      }
    } catch (error) {
      setState((prev) => ({
        ...prev,
        isSaving: false,
        message: 'Error saving request',
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
      request: {
        ...prev.request,
        [fieldName]: value
      }
    }));
  };

  if (state.isLoading) {
    return <Spinner size={SpinnerSize.large} label="Loading request..." />;
  }

  return (
    <Stack className={styles.formContainer}>
      <h2>Admin Request Form</h2>

      {state.message && (
        <MessageBar messageBarType={state.messageType}>{state.message}</MessageBar>
      )}

      <Stack tokens={{ childrenGap: 15 }}>
        {/* Title */}
        <TextField
          label="Request Title"
          placeholder="Enter request title"
          required
          value={state.request.title}
          onChange={(e, value) => handleFieldChange('title', value)}
        />

        {/* Description */}
        <TextField
          label="Description"
          placeholder="Enter detailed description"
          multiline
          rows={4}
          required
          value={state.request.description}
          onChange={(e, value) => handleFieldChange('description', value)}
        />

        {/* Requester Details */}
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <TextField
            label="Requester Name"
            placeholder="Enter your name"
            style={{ flex: 1 }}
            value={state.request.requesterName}
            onChange={(e, value) => handleFieldChange('requesterName', value)}
          />
          <TextField
            label="Email"
            placeholder="Enter your email"
            type="email"
            style={{ flex: 1 }}
            value={state.request.requesterEmail}
            onChange={(e, value) => handleFieldChange('requesterEmail', value)}
          />
        </Stack>

        {/* Department */}
        <TextField
          label="Department"
          placeholder="Enter department name"
          value={state.request.department}
          onChange={(e, value) => handleFieldChange('department', value)}
        />

        {/* Status and Priority */}
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <Dropdown
            label="Status"
            options={statusOptions}
            selectedKey={state.request.status}
            onChange={(e, option) => handleFieldChange('status', option?.key)}
            style={{ flex: 1 }}
          />
          <Dropdown
            label="Priority"
            options={priorityOptions}
            selectedKey={state.request.priority}
            onChange={(e, option) => handleFieldChange('priority', option?.key)}
            style={{ flex: 1 }}
          />
        </Stack>

        {/* Budget and Delivery Date */}
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <TextField
            label="Total Budget"
            placeholder="Enter budget amount"
            type="number"
            style={{ flex: 1 }}
            value={state.request.totalBudget?.toString() || ''}
            onChange={(e, value) =>
              handleFieldChange('totalBudget', value ? parseFloat(value) : 0)
            }
          />
          <DatePicker
            label="Target Delivery Date"
            value={state.request.targetDeliveryDate}
            onSelectDate={(date) => handleFieldChange('targetDeliveryDate', date)}
            style={{ flex: 1 }}
          />
        </Stack>

        {/* Comments */}
        <TextField
          label="Comments"
          placeholder="Enter any additional comments"
          multiline
          rows={3}
          value={state.request.comments}
          onChange={(e, value) => handleFieldChange('comments', value)}
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
          {state.request.attachments && state.request.attachments.length > 0 && (
            <div>
              <Label>Existing Attachments:</Label>
              <Stack tokens={{ childrenGap: 5 }}>
                {state.request.attachments.map((att) => (
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
