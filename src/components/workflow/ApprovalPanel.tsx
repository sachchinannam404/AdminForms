import * as React from 'react';
import {
  Stack,
  TextField,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Text
} from '@fluentui/react';
import { IAdminRequest, RequestStatus } from '../../models/IAdminRequest';
import { RequestService } from '../../services/RequestService';

export interface IApprovalPanelProps {
  request: IAdminRequest;
  currentUserName?: string;
  onComplete?: (updated: IAdminRequest) => void;
}

export const ApprovalPanel: React.FC<IApprovalPanelProps> = (props) => {
  const [comments, setComments] = React.useState('');
  const [rejectionReason, setRejectionReason] = React.useState('');
  const [busy, setBusy] = React.useState(false);
  const [message, setMessage] = React.useState<{ text: string; type: MessageBarType } | null>(null);

  const canAct =
    props.request.status === RequestStatus.Pending ||
    props.request.status === RequestStatus.Draft;

  const approve = async () => {
    if (!props.request.id) return;
    setBusy(true);
    try {
      await RequestService.approveRequest(
        props.request.id,
        props.currentUserName || 'Approver',
        comments || undefined
      );
      setMessage({ text: 'Request approved', type: MessageBarType.success });
      if (props.onComplete) {
        props.onComplete({
          ...props.request,
          status: RequestStatus.Approved,
          approvedBy: props.currentUserName || 'Approver',
          approvedDate: new Date(),
          comments: comments || props.request.comments
        });
      }
    } catch {
      setMessage({ text: 'Failed to approve', type: MessageBarType.error });
    } finally {
      setBusy(false);
    }
  };

  const reject = async () => {
    if (!props.request.id) return;
    if (!rejectionReason.trim()) {
      setMessage({ text: 'Rejection reason is required', type: MessageBarType.warning });
      return;
    }
    setBusy(true);
    try {
      await RequestService.rejectRequest(
        props.request.id,
        props.currentUserName || 'Approver',
        rejectionReason
      );
      setMessage({ text: 'Request rejected', type: MessageBarType.success });
      if (props.onComplete) {
        props.onComplete({
          ...props.request,
          status: RequestStatus.Rejected,
          approvedBy: props.currentUserName || 'Approver',
          approvedDate: new Date(),
          rejectionReason
        });
      }
    } catch {
      setMessage({ text: 'Failed to reject', type: MessageBarType.error });
    } finally {
      setBusy(false);
    }
  };

  if (!canAct) {
    return (
      <Stack tokens={{ childrenGap: 8 }}>
        <Text>
          Status: <strong>{props.request.status}</strong>
          {props.request.approvedBy && ` · By ${props.request.approvedBy}`}
        </Text>
        {props.request.rejectionReason && (
          <Text>Rejection reason: {props.request.rejectionReason}</Text>
        )}
      </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <Text variant="mediumPlus">Approval</Text>
      {message && (
        <MessageBar messageBarType={message.type} onDismiss={() => setMessage(null)}>
          {message.text}
        </MessageBar>
      )}
      <TextField
        label="Comments (optional)"
        multiline
        rows={2}
        value={comments}
        onChange={(_, v) => setComments(v || '')}
      />
      <TextField
        label="Rejection reason (required to reject)"
        multiline
        rows={2}
        value={rejectionReason}
        onChange={(_, v) => setRejectionReason(v || '')}
      />
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton text="Approve" onClick={approve} disabled={busy} />
        <DefaultButton text="Reject" onClick={reject} disabled={busy} />
      </Stack>
    </Stack>
  );
};
