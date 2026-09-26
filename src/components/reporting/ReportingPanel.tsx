import * as React from 'react';
import {
  Stack,
  Text,
  PrimaryButton,
  DefaultButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType
} from '@fluentui/react';
import { RequestService } from '../../services/RequestService';
import { ReportingService, IReportSummary } from '../../services/ReportingService';
import { IAdminRequest } from '../../models/IAdminRequest';

export interface IReportingPanelProps {
  onBack?: () => void;
}

export const ReportingPanel: React.FC<IReportingPanelProps> = (props) => {
  const [summary, setSummary] = React.useState<IReportSummary | null>(null);
  const [requests, setRequests] = React.useState<IAdminRequest[]>([]);
  const [loading, setLoading] = React.useState(true);

  React.useEffect(() => {
    (async () => {
      setLoading(true);
      try {
        const all = await RequestService.getRequests();
        setRequests(all);
        setSummary(ReportingService.buildSummary(all));
      } finally {
        setLoading(false);
      }
    })();
  }, []);

  if (loading) {
    return <Spinner size={SpinnerSize.large} label="Building report…" />;
  }

  if (!summary) {
    return <MessageBar messageBarType={MessageBarType.error}>Could not load report data.</MessageBar>;
  }

  const maxStatus = Math.max(...Object.values(summary.byStatus), 1);
  const maxType = Math.max(...Object.values(summary.byType), 1);

  return (
    <Stack tokens={{ childrenGap: 16 }} styles={{ root: { padding: 16, maxWidth: 800 } }}>
      <Stack horizontal horizontalAlign="space-between">
        <Text variant="xLarge">Reports</Text>
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text="Export CSV"
            iconProps={{ iconName: 'ExcelDocument' }}
            onClick={() => ReportingService.exportToCsv(requests)}
          />
          {props.onBack && <DefaultButton text="Back" onClick={props.onBack} />}
        </Stack>
      </Stack>

      <Stack horizontal wrap tokens={{ childrenGap: 16 }}>
        <Stat label="Total requests" value={String(summary.total)} />
        <Stat label="Pending" value={String(summary.pendingCount)} />
        <Stat label="Overdue" value={String(summary.overdueCount)} />
        <Stat label="Total budget" value={`$${summary.totalBudget.toFixed(0)}`} />
        <Stat label="Approved budget" value={`$${summary.approvedBudget.toFixed(0)}`} />
      </Stack>

      <Text variant="mediumPlus">By status</Text>
      <Stack tokens={{ childrenGap: 6 }}>
        {Object.keys(summary.byStatus).map((k) => (
          <Bar key={k} label={k} value={summary.byStatus[k]} max={maxStatus} />
        ))}
      </Stack>

      <Text variant="mediumPlus">By type</Text>
      <Stack tokens={{ childrenGap: 6 }}>
        {Object.keys(summary.byType).map((k) => (
          <Bar key={k} label={k} value={summary.byType[k]} max={maxType} />
        ))}
      </Stack>

      <Text variant="mediumPlus">By priority</Text>
      <Stack tokens={{ childrenGap: 6 }}>
        {Object.keys(summary.byPriority).map((k) => (
          <Bar
            key={k}
            label={k}
            value={summary.byPriority[k]}
            max={Math.max(...Object.values(summary.byPriority), 1)}
          />
        ))}
      </Stack>
    </Stack>
  );
};

const Stat: React.FC<{ label: string; value: string }> = ({ label, value }) => (
  <Stack
    styles={{
      root: {
        minWidth: 120,
        padding: 12,
        background: '#f3f2f1',
        borderRadius: 4
      }
    }}
  >
    <Text variant="large">{value}</Text>
    <Text variant="small">{label}</Text>
  </Stack>
);

const Bar: React.FC<{ label: string; value: number; max: number }> = ({ label, value, max }) => (
  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
    <Text styles={{ root: { minWidth: 100 } }}>{label}</Text>
    <div
      style={{
        flex: 1,
        height: 16,
        background: '#edebe9',
        borderRadius: 2,
        overflow: 'hidden'
      }}
    >
      <div
        style={{
          width: `${Math.round((value / max) * 100)}%`,
          height: '100%',
          background: '#0078d4'
        }}
      />
    </div>
    <Text styles={{ root: { minWidth: 32 } }}>{value}</Text>
  </Stack>
);
