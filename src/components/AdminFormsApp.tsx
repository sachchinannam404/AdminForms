/**
 * Top-level shell with ErrorBoundary, toast, and optional context provider.
 * Flow: Dashboard → Form → Detail (approval, children, audit) → Reports
 */

import * as React from 'react';
import { Stack, DefaultButton, MessageBar, MessageBarType } from '@fluentui/react';
import { RequestsDashboard } from './dashboard/RequestsDashboard';
import { DynamicRequestForm } from './forms/DynamicRequestForm';
import { ApprovalPanel } from './workflow/ApprovalPanel';
import { AuditHistory } from './workflow/AuditHistory';
import { ChildItemsList } from './lists/ChildItemsList';
import { ReportingPanel } from './reporting/ReportingPanel';
import { ErrorBoundary } from './common/ErrorBoundary';
import { MessageToastProvider } from './common/MessageToast';
import { AdminFormsProvider } from '../context/AdminFormsContext';
import { IAdminRequest, RequestType } from '../models/IAdminRequest';
import { getRequestTypeConfig, exceedsApprovalThreshold } from '../config/requestTypeRegistry';
import { NotificationService } from '../services/NotificationService';
import { RequestService } from '../services/RequestService';

export interface IAdminFormsAppProps {
  currentUserName?: string;
  currentUserEmail?: string;
  spfxContext?: any;
  powerAutomateWebhookUrl?: string;
}

type ViewMode = 'dashboard' | 'form' | 'detail' | 'reporting';

const AdminFormsAppInner: React.FC<IAdminFormsAppProps> = (props) => {
  const [view, setView] = React.useState<ViewMode>('dashboard');
  const [activeType, setActiveType] = React.useState<RequestType>(RequestType.Stationery);
  const [activeRequest, setActiveRequest] = React.useState<IAdminRequest | undefined>();

  React.useEffect(() => {
    if (props.spfxContext) {
      RequestService.initialize(props.spfxContext);
    }
    if (props.powerAutomateWebhookUrl) {
      NotificationService.configure({
        powerAutomateWebhookUrl: props.powerAutomateWebhookUrl
      });
    }
  }, [props.spfxContext, props.powerAutomateWebhookUrl]);

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
            approval is recommended.
          </MessageBar>
        )}

        {activeRequest.requestType === RequestType.Travel &&
          activeRequest.details?.actualCost != null && (
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

        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <DefaultButton text="Edit request" onClick={() => setView('form')} />
          {activeRequest.id && activeRequest.status !== 'Cancelled' && (
            <DefaultButton
              text="Archive (soft delete)"
              onClick={async () => {
                if (!activeRequest.id) return;
                await RequestService.archiveRequest(activeRequest.id);
                setActiveRequest({ ...activeRequest, status: 'Cancelled' as any });
              }}
            />
          )}
        </Stack>

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

export const AdminFormsApp: React.FC<IAdminFormsAppProps> = (props) => {
  return (
    <ErrorBoundary>
      <AdminFormsProvider
        spfxContext={props.spfxContext}
        displayName={props.currentUserName}
        email={props.currentUserEmail}
        notificationConfig={
          props.powerAutomateWebhookUrl
            ? { powerAutomateWebhookUrl: props.powerAutomateWebhookUrl }
            : undefined
        }
      >
        <MessageToastProvider>
          <AdminFormsAppInner {...props} />
        </MessageToastProvider>
      </AdminFormsProvider>
    </ErrorBoundary>
  );
};
