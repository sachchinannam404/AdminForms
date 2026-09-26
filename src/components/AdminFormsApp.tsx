/**
 * Top-level shell: Dashboard ↔ Create/Edit request ↔ Child items + Approval
 */

import * as React from 'react';
import { Stack, DefaultButton } from '@fluentui/react';
import { RequestsDashboard } from './dashboard/RequestsDashboard';
import { DynamicRequestForm } from './forms/DynamicRequestForm';
import { ApprovalPanel } from './workflow/ApprovalPanel';
import { ChildItemsList } from './lists/ChildItemsList';
import { IAdminRequest, RequestType } from '../models/IAdminRequest';
import { getRequestTypeConfig } from '../config/requestTypeRegistry';

export interface IAdminFormsAppProps {
  currentUserName?: string;
  currentUserEmail?: string;
}

type ViewMode = 'dashboard' | 'form' | 'detail';

export const AdminFormsApp: React.FC<IAdminFormsAppProps> = (props) => {
  const [view, setView] = React.useState<ViewMode>('dashboard');
  const [activeType, setActiveType] = React.useState<RequestType>(RequestType.Stationery);
  const [activeRequest, setActiveRequest] = React.useState<IAdminRequest | undefined>();

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
    return (
      <Stack tokens={{ childrenGap: 20 }} styles={{ root: { padding: 16, maxWidth: 900 } }}>
        <DefaultButton text="← Back to dashboard" onClick={backToDashboard} />
        <h2>{activeRequest.title}</h2>
        <p>{config.displayName} · {activeRequest.status} · {activeRequest.priority}</p>
        <p>{activeRequest.description}</p>

        <ApprovalPanel
          request={activeRequest}
          currentUserName={props.currentUserName}
          onComplete={(updated) => setActiveRequest(updated)}
        />

        <DefaultButton
          text="Edit request"
          onClick={() => setView('form')}
        />

        {config.supportsChildren && activeRequest.id && (
          <ChildItemsList
            requestId={activeRequest.id}
            requestType={activeRequest.requestType}
          />
        )}
      </Stack>
    );
  }

  return (
    <RequestsDashboard
      onCreateRequest={openCreate}
      onOpenRequest={openRequest}
    />
  );
};
