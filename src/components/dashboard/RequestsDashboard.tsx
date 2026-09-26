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
  Selection,
  SelectionMode,
  IColumn,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Text,
  Icon,
  DatePicker,
  Pivot,
  PivotItem
} from '@fluentui/react';
import {
  IAdminRequest,
  RequestStatus,
  RequestType,
  PriorityLevel
} from '../../models/IAdminRequest';
import { RequestService } from '../../services/RequestService';
import { getAllRequestTypes } from '../../config/requestTypeRegistry';
import { RoleService, ICurrentUserContext } from '../../services/RoleService';
import { ReportingService } from '../../services/ReportingService';
import styles from './RequestsDashboard.module.scss';

export interface IRequestsDashboardProps {
  currentUserName?: string;
  currentUserEmail?: string;
  onOpenRequest?: (request: IAdminRequest) => void;
  onCreateRequest?: (type: RequestType) => void;
  onOpenReporting?: () => void;
}

interface IStats {
  total: number;
  pending: number;
  approved: number;
  rejected: number;
  inProgress: number;
  overdue: number;
  totalBudget: number;
  byType: Record<string, number>;
}

type ViewMode = 'my' | 'all' | 'pendingApproval';

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

const bulkStatusOptions: IDropdownOption[] = Object.values(RequestStatus).map((s) => ({
  key: s,
  text: s
}));

export const RequestsDashboard: React.FC<IRequestsDashboardProps> = (props) => {
  const [items, setItems] = React.useState<IAdminRequest[]>([]);
  const [stats, setStats] = React.useState<IStats | null>(null);
  const [isLoading, setIsLoading] = React.useState(true);
  const [message, setMessage] = React.useState<{ text: string; type: MessageBarType } | null>(null);
  const [searchText, setSearchText] = React.useState('');
  const [statusFilter, setStatusFilter] = React.useState('');
  const [typeFilter, setTypeFilter] = React.useState('');
  const [priorityFilter, setPriorityFilter] = React.useState('');
  const [departmentFilter, setDepartmentFilter] = React.useState('');
  const [fromDate, setFromDate] = React.useState<Date | undefined>();
  const [toDate, setToDate] = React.useState<Date | undefined>();
  const [createType, setCreateType] = React.useState<string>(RequestType.Stationery);
  const [viewMode, setViewMode] = React.useState<ViewMode>('my');
  const [user, setUser] = React.useState<ICurrentUserContext | null>(null);
  const [bulkStatus, setBulkStatus] = React.useState<string>(RequestStatus.InProgress);
  const [selectedCount, setSelectedCount] = React.useState(0);

  const selectionRef = React.useRef(
    new Selection({
      onSelectionChanged: () => {
        setSelectedCount(selectionRef.current.getSelectedCount());
      }
    })
  );

  React.useEffect(() => {
    RoleService.getCurrentUser(props.currentUserName, props.currentUserEmail).then((u) => {
      setUser(u);
      if (u.isApprover) setViewMode('all');
    });
  }, [props.currentUserName, props.currentUserEmail]);

  const load = React.useCallback(async () => {
    if (!user) return;
    setIsLoading(true);
    try {
      const scope = RoleService.scopeFilter(user, viewMode);
      const list = await RequestService.getRequests({
        searchText: searchText || undefined,
        status: scope.statusPendingOnly
          ? RequestStatus.Pending
          : statusFilter
          ? (statusFilter as RequestStatus)
          : undefined,
        requestType: typeFilter ? (typeFilter as RequestType) : undefined,
        priority: priorityFilter ? (priorityFilter as PriorityLevel) : undefined,
        department: departmentFilter || undefined,
        requesterEmail: scope.requesterEmail,
        fromDate,
        toDate
      });
      setItems(list);
      const dashboardStats = await RequestService.getDashboardStats();
      setStats(dashboardStats);
    } catch {
      setMessage({ text: 'Failed to load requests', type: MessageBarType.error });
    } finally {
      setIsLoading(false);
    }
  }, [
    user,
    viewMode,
    searchText,
    statusFilter,
    typeFilter,
    priorityFilter,
    departmentFilter,
    fromDate,
    toDate
  ]);

  React.useEffect(() => {
    load();
  }, [load]);

  const handleBulkUpdate = async () => {
    const selected = selectionRef.current.getSelection() as IAdminRequest[];
    const ids = selected.map((s) => s.id!).filter(Boolean);
    if (!ids.length) {
      setMessage({ text: 'Select one or more requests', type: MessageBarType.warning });
      return;
    }
    const result = await RequestService.bulkUpdateStatus(
      ids,
      bulkStatus as RequestStatus,
      user?.displayName || 'User'
    );
    setMessage({
      text: `Updated ${result.success} request(s)${result.failed ? `, ${result.failed} failed` : ''}`,
      type: result.failed ? MessageBarType.warning : MessageBarType.success
    });
    selectionRef.current.setAllSelected(false);
    load();
  };

  const columns: IColumn[] = [
    {
      key: 'title',
      name: 'Title',
      fieldName: 'title',
      minWidth: 150,
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
      minWidth: 110,
      isResizable: true,
      onRender: (item: IAdminRequest) =>
        getAllRequestTypes().find((t) => t.key === item.requestType)?.displayName ||
        item.requestType
    },
    { key: 'status', name: 'Status', fieldName: 'status', minWidth: 90, isResizable: true },
    { key: 'priority', name: 'Priority', fieldName: 'priority', minWidth: 80, isResizable: true },
    {
      key: 'requester',
      name: 'Requester',
      fieldName: 'requesterName',
      minWidth: 110,
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
      minWidth: 80,
      onRender: (item: IAdminRequest) =>
        item.totalBudget != null ? `$${Number(item.totalBudget).toFixed(0)}` : '—'
    },
    {
      key: 'due',
      name: 'Due',
      minWidth: 90,
      onRender: (item: IAdminRequest) => {
        if (!item.targetDeliveryDate) return '';
        const overdue =
          item.targetDeliveryDate.getTime() < Date.now() &&
          item.status !== RequestStatus.Completed &&
          item.status !== RequestStatus.Rejected &&
          item.status !== RequestStatus.Cancelled;
        return (
          <span style={{ color: overdue ? '#a4262c' : undefined, fontWeight: overdue ? 600 : undefined }}>
            {item.targetDeliveryDate.toLocaleDateString()}
          </span>
        );
      }
    },
    {
      key: 'created',
      name: 'Created',
      minWidth: 90,
      onRender: (item: IAdminRequest) => (item.created ? item.created.toLocaleDateString() : '')
    }
  ];

  return (
    <Stack className={styles.container} tokens={{ childrenGap: 16 }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" wrap>
        <h2 className={styles.title}>Admin Requests Dashboard</h2>
        <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
          <Dropdown
            selectedKey={createType}
            options={getAllRequestTypes().map((t) => ({ key: t.key, text: t.displayName }))}
            onChange={(_, opt) => opt && setCreateType(opt.key as string)}
            styles={{ root: { minWidth: 180 } }}
          />
          <PrimaryButton
            text="New request"
            iconProps={{ iconName: 'Add' }}
            onClick={() =>
              props.onCreateRequest && props.onCreateRequest(createType as RequestType)
            }
          />
          <DefaultButton
            text="Export CSV"
            iconProps={{ iconName: 'ExcelDocument' }}
            onClick={() => ReportingService.exportToCsv(items)}
          />
          {props.onOpenReporting && (
            <DefaultButton text="Reports" iconProps={{ iconName: 'BarChart4' }} onClick={props.onOpenReporting} />
          )}
        </Stack>
      </Stack>

      {user && (
        <Pivot
          selectedKey={viewMode}
          onLinkClick={(item) => item && setViewMode(item.props.itemKey as ViewMode)}
        >
          <PivotItem headerText="My requests" itemKey="my" />
          {user.isApprover && <PivotItem headerText="All requests" itemKey="all" />}
          {user.isApprover && (
            <PivotItem headerText="Pending approval" itemKey="pendingApproval" />
          )}
        </Pivot>
      )}

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
            <Icon iconName="Warning" />
            <Text variant="large">{stats.overdue}</Text>
            <Text>Overdue</Text>
          </div>
          <div className={styles.kpi}>
            <Icon iconName="Accept" />
            <Text variant="large">{stats.approved}</Text>
            <Text>Approved</Text>
          </div>
          <div className={styles.kpi}>
            <Icon iconName="Money" />
            <Text variant="large">${stats.totalBudget.toFixed(0)}</Text>
            <Text>Budget</Text>
          </div>
        </Stack>
      )}

      <Stack horizontal wrap tokens={{ childrenGap: 12 }} verticalAlign="end">
        <TextField
          label="Search"
          placeholder="Title, description, requester, dept…"
          value={searchText}
          onChange={(_, v) => setSearchText(v || '')}
          styles={{ root: { minWidth: 200 } }}
        />
        <Dropdown
          label="Status"
          options={statusOptions}
          selectedKey={statusFilter}
          onChange={(_, opt) => setStatusFilter((opt?.key as string) || '')}
          styles={{ root: { minWidth: 130 } }}
        />
        <Dropdown
          label="Type"
          options={typeOptions}
          selectedKey={typeFilter}
          onChange={(_, opt) => setTypeFilter((opt?.key as string) || '')}
          styles={{ root: { minWidth: 140 } }}
        />
        <Dropdown
          label="Priority"
          options={priorityOptions}
          selectedKey={priorityFilter}
          onChange={(_, opt) => setPriorityFilter((opt?.key as string) || '')}
          styles={{ root: { minWidth: 120 } }}
        />
        <TextField
          label="Department"
          value={departmentFilter}
          onChange={(_, v) => setDepartmentFilter(v || '')}
          styles={{ root: { minWidth: 120 } }}
        />
        <DatePicker
          label="From"
          value={fromDate}
          onSelectDate={(d) => setFromDate(d || undefined)}
        />
        <DatePicker label="To" value={toDate} onSelectDate={(d) => setToDate(d || undefined)} />
        <DefaultButton text="Refresh" iconProps={{ iconName: 'Refresh' }} onClick={load} />
      </Stack>

      {user?.isApprover && (
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
          <Dropdown
            label="Bulk status"
            options={bulkStatusOptions}
            selectedKey={bulkStatus}
            onChange={(_, opt) => opt && setBulkStatus(opt.key as string)}
            styles={{ root: { minWidth: 140 } }}
          />
          <PrimaryButton
            text={`Update selected (${selectedCount})`}
            disabled={!selectedCount}
            onClick={handleBulkUpdate}
          />
        </Stack>
      )}

      {isLoading ? (
        <Spinner size={SpinnerSize.large} label="Loading requests…" />
      ) : (
        <DetailsList
          items={items}
          columns={columns}
          layoutMode={DetailsListLayoutMode.justified}
          selection={selectionRef.current}
          selectionMode={user?.isApprover ? SelectionMode.multiple : SelectionMode.none}
          selectionPreservedOnEmptyClick
          isHeaderVisible
        />
      )}
    </Stack>
  );
};
