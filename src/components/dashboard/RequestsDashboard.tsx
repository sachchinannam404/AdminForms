import * as React from 'react';
import {
  Stack,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Text,
  Icon
} from '@fluentui/react';
import { IAdminRequest, RequestStatus, RequestType, PriorityLevel } from '../../models/IAdminRequest';
import { RequestService } from '../../services/RequestService';
import { getAllRequestTypes } from '../../config/requestTypeRegistry';
import styles from './RequestsDashboard.module.scss';

export interface IRequestsDashboardProps {
  onOpenRequest?: (request: IAdminRequest) => void;
  onCreateRequest?: (type: RequestType) => void;
}

interface IStats {
  total: number;
  pending: number;
  approved: number;
  rejected: number;
  inProgress: number;
  byType: Record<string, number>;
}

const statusOptions: IDropdownOption[] = [
  { key: '', text: 'All statuses' },
  ...Object.values(RequestStatus).map((s) => ({ key: s, text: s }))
];

const typeOptions: IDropdownOption[] = [
  { key: '', text: 'All types' },
  ...getAllRequestTypes().map((t) => ({ key: t.key, text: t.displayName }))
];

const priorityOptions: IDropdownOption[] = [
  { key: '', text: 'All priorities' },
  ...Object.values(PriorityLevel).map((p) => ({ key: p, text: p }))
];

export const RequestsDashboard: React.FC<IRequestsDashboardProps> = (props) => {
  const [items, setItems] = React.useState<IAdminRequest[]>([]);
  const [stats, setStats] = React.useState<IStats | null>(null);
  const [isLoading, setIsLoading] = React.useState(true);
  const [message, setMessage] = React.useState<{ text: string; type: MessageBarType } | null>(null);
  const [searchText, setSearchText] = React.useState('');
  const [statusFilter, setStatusFilter] = React.useState<string>('');
  const [typeFilter, setTypeFilter] = React.useState<string>('');
  const [priorityFilter, setPriorityFilter] = React.useState<string>('');
  const [createType, setCreateType] = React.useState<string>(RequestType.Stationery);

  const load = React.useCallback(async () => {
    setIsLoading(true);
    try {
      const [list, dashboardStats] = await Promise.all([
        RequestService.getRequests({
          searchText: searchText || undefined,
          status: statusFilter ? (statusFilter as RequestStatus) : undefined,
          requestType: typeFilter ? (typeFilter as RequestType) : undefined,
          priority: priorityFilter ? (priorityFilter as PriorityLevel) : undefined
        }),
        RequestService.getDashboardStats()
      ]);
      setItems(list);
      setStats(dashboardStats);
    } catch (e) {
      setMessage({ text: 'Failed to load requests', type: MessageBarType.error });
    } finally {
      setIsLoading(false);
    }
  }, [searchText, statusFilter, typeFilter, priorityFilter]);

  React.useEffect(() => {
    load();
  }, [load]);

  const columns: IColumn[] = [
    {
      key: 'title',
      name: 'Title',
      fieldName: 'title',
      minWidth: 160,
      isResizable: true,
      onRender: (item: IAdminRequest) => (
        <button
          type="button"
          className={styles.linkButton}
          onClick={() => props.onOpenRequest && props.onOpenRequest(item)}
        >
          {item.title}
        </button>
      )
    },
    {
      key: 'type',
      name: 'Type',
      fieldName: 'requestType',
      minWidth: 120,
      isResizable: true,
      onRender: (item: IAdminRequest) =>
        getAllRequestTypes().find((t) => t.key === item.requestType)?.displayName || item.requestType
    },
    { key: 'status', name: 'Status', fieldName: 'status', minWidth: 100, isResizable: true },
    { key: 'priority', name: 'Priority', fieldName: 'priority', minWidth: 90, isResizable: true },
    {
      key: 'requester',
      name: 'Requester',
      fieldName: 'requesterName',
      minWidth: 120,
      isResizable: true
    },
    {
      key: 'department',
      name: 'Department',
      fieldName: 'department',
      minWidth: 100,
      isResizable: true
    },
    {
      key: 'budget',
      name: 'Budget',
      fieldName: 'totalBudget',
      minWidth: 90,
      isResizable: true,
      onRender: (item: IAdminRequest) =>
        item.totalBudget != null ? `$${Number(item.totalBudget).toFixed(2)}` : '—'
    },
    {
      key: 'created',
      name: 'Created',
      fieldName: 'created',
      minWidth: 100,
      isResizable: true,
      onRender: (item: IAdminRequest) =>
        item.created ? item.created.toLocaleDateString() : ''
    }
  ];

  return (
    <Stack className={styles.container} tokens={{ childrenGap: 16 }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <h2 className={styles.title}>Admin Requests Dashboard</h2>
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <Dropdown
            selectedKey={createType}
            options={getAllRequestTypes().map((t) => ({ key: t.key, text: t.displayName }))}
            onChange={(_, opt) => opt && setCreateType(opt.key as string)}
            styles={{ root: { minWidth: 200 } }}
          />
          <PrimaryButton
            text="New request"
            iconProps={{ iconName: 'Add' }}
            onClick={() => props.onCreateRequest && props.onCreateRequest(createType as RequestType)}
          />
        </Stack>
      </Stack>

      {message && (
        <MessageBar messageBarType={message.type} onDismiss={() => setMessage(null)}>
          {message.text}
        </MessageBar>
      )}

      {stats && (
        <Stack horizontal wrap tokens={{ childrenGap: 12 }} className={styles.kpiRow}>
          <div className={styles.kpi}>
            <Icon iconName="AllApps" />
            <Text variant="large">{stats.total}</Text>
            <Text>Total</Text>
          </div>
          <div className={styles.kpi}>
            <Icon iconName="Clock" />
            <Text variant="large">{stats.pending}</Text>
            <Text>Pending</Text>
          </div>
          <div className={styles.kpi}>
            <Icon iconName="Accept" />
            <Text variant="large">{stats.approved}</Text>
            <Text>Approved</Text>
          </div>
          <div className={styles.kpi}>
            <Icon iconName="ProgressRingDots" />
            <Text variant="large">{stats.inProgress}</Text>
            <Text>In progress</Text>
          </div>
          <div className={styles.kpi}>
            <Icon iconName="Cancel" />
            <Text variant="large">{stats.rejected}</Text>
            <Text>Rejected</Text>
          </div>
        </Stack>
      )}

      <Stack horizontal wrap tokens={{ childrenGap: 12 }} verticalAlign="end">
        <TextField
          label="Search"
          placeholder="Title, description, requester…"
          value={searchText}
          onChange={(_, v) => setSearchText(v || '')}
          styles={{ root: { minWidth: 220 } }}
        />
        <Dropdown
          label="Status"
          options={statusOptions}
          selectedKey={statusFilter}
          onChange={(_, opt) => setStatusFilter((opt?.key as string) || '')}
          styles={{ root: { minWidth: 140 } }}
        />
        <Dropdown
          label="Type"
          options={typeOptions}
          selectedKey={typeFilter}
          onChange={(_, opt) => setTypeFilter((opt?.key as string) || '')}
          styles={{ root: { minWidth: 160 } }}
        />
        <Dropdown
          label="Priority"
          options={priorityOptions}
          selectedKey={priorityFilter}
          onChange={(_, opt) => setPriorityFilter((opt?.key as string) || '')}
          styles={{ root: { minWidth: 130 } }}
        />
        <DefaultButton text="Refresh" iconProps={{ iconName: 'Refresh' }} onClick={load} />
      </Stack>

      {isLoading ? (
        <Spinner size={SpinnerSize.large} label="Loading requests…" />
      ) : (
        <DetailsList
          items={items}
          columns={columns}
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
          isHeaderVisible
        />
      )}
    </Stack>
  );
};
