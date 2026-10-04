import * as React from 'react';
import { ColumnActionsMode, ConstrainMode, DetailsList, DetailsListLayoutMode, IColumn, Link, SelectionMode } from '@fluentui/react';
import { FieldKeys } from '../../../../config/fieldKeys';
import { IFieldConfig } from '../../../../models/IFieldConfig';
import { IReceivingRecord } from '../../../../models/IReceivingRecord';
import { formatFieldValue } from '../../../../logic/format';
import { hasBadge, StatusBadge } from '../common/StatusBadge';
import styles from '../App.module.scss';

export interface IRecordTableProps {
  records: IReceivingRecord[];
  fields: IFieldConfig[];
  sortKey: string;
  sortAscending: boolean;
  onSort: (key: string) => void;
  onOpen: (record: IReceivingRecord) => void;
}

/** Desktop/tablet table. Columns come from the fields marked "showInTable". */
export const RecordTable: React.FC<IRecordTableProps> = ({ records, fields, sortKey, sortAscending, onSort, onOpen }) => {
  const columns: IColumn[] = fields.map((field) => {
    const width = field.tableWidth || 120;
    return {
      key: field.key,
      name: field.displayName,
      fieldName: field.key,
      minWidth: Math.round(width * 0.75),
      maxWidth: width,
      isResizable: true,
      isSorted: field.key === sortKey,
      isSortedDescending: field.key === sortKey && !sortAscending,
      sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A',
      columnActionsMode: field.sortable ? ColumnActionsMode.clickable : ColumnActionsMode.disabled,
      onColumnClick: field.sortable ? () => onSort(field.key) : undefined,
      onRender: (item: IReceivingRecord) => {
        const value = item.values[field.key];
        if (field.key === FieldKeys.recordId) {
          return (
            <Link className={styles.recordLink} onClick={() => onOpen(item)}>
              {formatFieldValue(field, value)}
            </Link>
          );
        }
        if (hasBadge(field)) {
          return <StatusBadge field={field} value={typeof value === 'string' ? value : undefined} />;
        }
        const text = formatFieldValue(field, value);
        return <span title={text}>{text}</span>;
      }
    };
  });

  return (
    <div style={{ overflowX: 'auto' }}>
      <DetailsList
        items={records}
        columns={columns}
        selectionMode={SelectionMode.none}
        layoutMode={DetailsListLayoutMode.justified}
        constrainMode={ConstrainMode.unconstrained}
        isHeaderVisible={true}
        getKey={(item: IReceivingRecord) => String(item.id)}
        onItemInvoked={(item: IReceivingRecord) => onOpen(item)}
        ariaLabelForGrid="Receiving records"
      />
    </div>
  );
};
