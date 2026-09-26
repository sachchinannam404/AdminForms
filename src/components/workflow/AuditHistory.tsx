import * as React from 'react';
import {
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  DefaultButton
} from '@fluentui/react';
import { AuditService, IAuditEntry } from '../../services/AuditService';

export interface IAuditHistoryProps {
  requestId: string;
}

export const AuditHistory: React.FC<IAuditHistoryProps> = (props) => {
  const [entries, setEntries] = React.useState<IAuditEntry[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [loaded, setLoaded] = React.useState(false);

  const load = async () => {
    setLoading(true);
    try {
      const data = await AuditService.getRequestHistory(props.requestId);
      setEntries(data);
      setLoaded(true);
    } finally {
      setLoading(false);
    }
  };

  const columns: IColumn[] = [
    { key: 'ver', name: 'Version', fieldName: 'versionLabel', minWidth: 70 },
    {
      key: 'mod',
      name: 'Modified',
      minWidth: 120,
      onRender: (e: IAuditEntry) => e.modified.toLocaleString()
    },
    { key: 'editor', name: 'Editor', fieldName: 'editor', minWidth: 120 },
    { key: 'status', name: 'Status', fieldName: 'status', minWidth: 90 },
    {
      key: 'summary',
      name: 'Changes',
      fieldName: 'changesSummary',
      minWidth: 200,
      isResizable: true
    }
  ];

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Text variant="mediumPlus">Audit / history</Text>
        <DefaultButton
          text={loaded ? 'Refresh' : 'Load history'}
          onClick={load}
          disabled={loading}
        />
      </Stack>
      <Text variant="small">
        Requires versioning enabled on the Admin Requests list. Shows key field changes between
        versions.
      </Text>
      {loading && <Spinner size={SpinnerSize.small} label="Loading versions…" />}
      {loaded && !loading && entries.length === 0 && (
        <Text>No version history available.</Text>
      )}
      {entries.length > 0 && (
        <DetailsList
          items={entries}
          columns={columns}
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
        />
      )}
    </Stack>
  );
};
