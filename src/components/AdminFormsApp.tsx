/**
 * Top-level shell: Dashboard ↔ Create/Edit ↔ Detail (approval, children, audit) ↔ Reports
 */

import * as React from 'react';
import { Stack, DefaultButton, MessageBar, MessageBarType } from '@fluentui/react';
import { RequestsDashboard } from './dashboard/RequestsDashboard';
import { DynamicRequestForm } from './forms/DynamicRequestForm';
import { ApprovalPanel } from './workflow/ApprovalPanel';
import { AuditHistory } from './workflow/AuditHistory';
import { ChildItemsList } from './lists/ChildItemsList';
import { ReportingPanel } from './reporting/ReportingPanel';
import { IAdminRequest, RequestType } from '../models/IAdminRequest';
import { getRequestTypeConfig, exceedsApprovalThreshold } from '../config/requestTypeRegistry';
import { NotificationService } from '../services/NotificationService';

export interface IAdminFormsAppProps {
  currentUserName?: string;
  currentUserEmail?: string;
  /** Optional Power Automate HTTP webhook for status notifications */
  powerAutomateWebhookUrl?: string;
}

type ViewMode = 'dashboard' | 'form' | 'detail' | 'reporting';

export const AdminFormsApp: React.FC<IAdminFormsAppProps> = (props) => {
  const [view, setView] = React.useState<ViewMode>('dashboard');
  const [activeType, setActiveType] = React.useState<RequestType>(RequestType.Stationery);
  const [activeRequest, setActiveRequest] = React.useState<IAdminRequest | undefined>();

  React.useEffect(() => {
    if (props.powerAutomateWebhookUrl) {
      NotificationService.configure({
        powerAutomateWebhookUrl: props.powerAutomateWebhookUrl
      });
    }
  }, [props.powerAutomateWebhookUrl]);

  const openCreate = (type: RequestType) => {
    setActiveType(type);
    setActiveRequest(undefined);
    setView('form');
  };

  const openRequest = (request: IAdminRequest) => {
    setActiveRequest(request);
    setActiveType(request.requestType);
    setView('detail');
  };

  const backToDashboard = () => {
    setActiveRequest(undefined);
    setView('dashboard');
  };

  if (view === 'reporting') {
    return <ReportingPanel onBack={backToDashboard} />;
  }

  if (view === 'form') {
    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <DefaultButton text="← Back to dashboard" onClick={backToDashboard} />
        <DynamicRequestForm
          requestType={activeType}
          requestId={activeRequest?.id}
          onSave={(saved) => {
            setActiveRequest(saved);
            setView('detail');
          }}
          onCancel={backToDashboard}
        />
      </Stack>
    );
  }

  if (view === 'detail' && activeRequest) {
    const config = getRequestTypeConfig(activeRequest.requestType);
    const needsDual = exceedsApprovalThreshold(
      activeRequest.requestType,
      activeRequest.totalBudget
    );

    return (
      <Stack tokens={{ childrenGap: 20 }} styles={{ root: { padding: 16, maxWidth: 900 } }}>
        <DefaultButton text="← Back to dashboard" onClick={backToDashboard} />
        <h2>{activeRequest.title}</h2>
        <p>
          {config.displayName} · {activeRequest.status} · {activeRequest.priority}
          {activeRequest.totalBudget != null && ` · $${activeRequest.totalBudget}`}
        </p>
        <p>{activeRequest.description}</p>

        {needsDual && (
          <MessageBar messageBarType={MessageBarType.warning}>
            Amount meets or exceeds the approval threshold
            {config.approvalThreshold != null ? ` ($${config.approvalThreshold})` : ''}. Dual
            approval is recommended for this request type.
          </MessageBar>
        )}

        {activeRequest.requestType === 'Travel' && activeRequest.details?.actualCost != null && (
          <MessageBar messageBarType={MessageBarType.info}>
            Travel budget vs actual: estimated ${activeRequest.totalBudget ?? 0} / actual $
            {activeRequest.details.actualCost}
          </MessageBar>
        )}

        <ApprovalPanel
          request={activeRequest}
          currentUserName={props.currentUserName}
          onComplete={(updated) => setActiveRequest(updated)}
        />

        <DefaultButton text="Edit request" onClick={() => setView('form')} />

        {config.supportsChildren && activeRequest.id && (
          <ChildItemsList
            requestId={activeRequest.id}
            requestType={activeRequest.requestType}
          />
        )}

        {activeRequest.id && <AuditHistory requestId={activeRequest.id} />}
      </Stack>
    );
  }

  return (
    <RequestsDashboard
      currentUserName={props.currentUserName}
      currentUserEmail={props.currentUserEmail}
      onCreateRequest={openCreate}
      onOpenRequest={openRequest}
      onOpenReporting={() => setView('reporting')}
    />
  );
};
