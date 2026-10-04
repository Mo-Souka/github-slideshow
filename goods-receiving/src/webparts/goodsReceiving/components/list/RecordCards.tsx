import * as React from 'react';
import { FieldKeys } from '../../../../config/fieldKeys';
import { IFieldConfig } from '../../../../models/IFieldConfig';
import { IReceivingRecord } from '../../../../models/IReceivingRecord';
import { formatFieldValue } from '../../../../logic/format';
import { hasBadge, StatusBadge } from '../common/StatusBadge';
import styles from '../App.module.scss';

export interface IRecordCardsProps {
  records: IReceivingRecord[];
  fields: IFieldConfig[];
  onOpen: (record: IReceivingRecord) => void;
}

/** Phone layout: one tappable card per record. Uses the same "showInTable" fields as the table. */
export const RecordCards: React.FC<IRecordCardsProps> = ({ records, fields, onOpen }) => {
  const recordIdField = fields.filter((f) => f.key === FieldKeys.recordId)[0];
  const approvalField = fields.filter((f) => f.key === FieldKeys.approvalStatus)[0];
  const badgeFields = fields.filter((f) => hasBadge(f) && f.key !== FieldKeys.approvalStatus);
  const rowFields = fields.filter((f) => f.key !== FieldKeys.recordId && !hasBadge(f));

  return (
    <div className={styles.cards}>
      {records.map((record) => {
        const approval = record.values[FieldKeys.approvalStatus];
        return (
          <button key={record.id} type="button" className={styles.card} onClick={() => onOpen(record)}>
            <div className={styles.cardHeader}>
              <span className={styles.cardTitle}>
                {recordIdField ? formatFieldValue(recordIdField, record.values[FieldKeys.recordId]) : `#${record.id}`}
              </span>
              {approvalField && <StatusBadge field={approvalField} value={typeof approval === 'string' ? approval : undefined} />}
            </div>
            {rowFields.map((field) => {
              const text = formatFieldValue(field, record.values[field.key]);
              if (!text) return null;
              return (
                <div key={field.key} className={styles.cardRow}>
                  <span className={styles.cardLabel}>{field.displayName}</span>
                  <span>{text}</span>
                </div>
              );
            })}
            {badgeFields.length > 0 && (
              <div className={styles.cardBadges}>
                {badgeFields.map((field) => {
                  const value = record.values[field.key];
                  return <StatusBadge key={field.key} field={field} value={typeof value === 'string' ? value : undefined} />;
                })}
              </div>
            )}
          </button>
        );
      })}
    </div>
  );
};
