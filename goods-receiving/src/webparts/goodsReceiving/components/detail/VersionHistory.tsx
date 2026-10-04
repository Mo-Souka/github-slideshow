import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { IVersionEntry } from '../../../../models/IReceivingRecord';
import { formatDateTime } from '../../../../logic/format';
import { useAppContext } from '../AppContext';
import { ErrorMessage } from '../common/ErrorMessage';
import styles from '../App.module.scss';

export interface IVersionHistoryProps {
  recordId: number;
  /** Changes whenever the record is saved, so the history reloads. */
  refreshKey: string;
}

/** Audit trail from SharePoint version history: who changed which field, when. */
export const VersionHistory: React.FC<IVersionHistoryProps> = ({ recordId, refreshKey }) => {
  const { services } = useAppContext();
  const [entries, setEntries] = React.useState<IVersionEntry[] | undefined>();
  const [error, setError] = React.useState<unknown>();
  const [attempt, setAttempt] = React.useState(0);

  React.useEffect(() => {
    let cancelled = false;
    setEntries(undefined);
    setError(undefined);
    services.records
      .getHistory(recordId)
      .then((result) => {
        if (!cancelled) setEntries(result);
      })
      .catch((e) => {
        if (!cancelled) setError(e);
      });
    return () => {
      cancelled = true;
    };
  }, [recordId, refreshKey, services, attempt]);

  if (error !== undefined) return <ErrorMessage error={error} onRetry={() => setAttempt(attempt + 1)} />;
  if (!entries) return <Spinner size={SpinnerSize.medium} label="Loading history..." />;
  if (entries.length === 0) return <div className={styles.empty}>No history available.</div>;

  return (
    <div>
      {entries.map((entry) => (
        <div key={entry.versionLabel} className={styles.historyEntry}>
          <div className={styles.historyMeta}>
            <strong>{entry.isCreation ? 'Created' : 'Changed'}</strong>
            {entry.modified ? ` ${formatDateTime(entry.modified)}` : ''}
            {entry.modifiedBy ? ` by ${entry.modifiedBy}` : ''} · version {entry.versionLabel}
          </div>
          {!entry.isCreation &&
            entry.changes.map((change) => (
              <div key={change.fieldKey} className={styles.historyChange}>
                <strong>{change.fieldName}:</strong>{' '}
                {change.oldValue ? <span className={styles.oldValue}>{change.oldValue}</span> : <em>(empty)</em>} →{' '}
                {change.newValue || <em>(empty)</em>}
              </div>
            ))}
        </div>
      ))}
    </div>
  );
};
