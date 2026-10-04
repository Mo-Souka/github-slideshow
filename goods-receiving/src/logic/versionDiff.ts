import { IFieldConfig } from '../models/IFieldConfig';
import { FieldValues, IFieldChange, IVersionEntry } from '../models/IReceivingRecord';
import { formatFieldValue } from './format';

/** One version of a list item, already mapped to field values. */
export interface IVersionSnapshot {
  versionId: number;
  versionLabel: string;
  modified?: Date;
  modifiedBy?: string;
  values: FieldValues;
}

/**
 * Turns SharePoint's version snapshots into a readable audit trail:
 * for each version, which fields changed and from what to what.
 * Built-in fields (Record ID, Created, Modified, ...) are not listed as changes,
 * and versions without any tracked change (such as the automatic Record ID
 * update right after creation) are left out. Returns the newest version first.
 */
export function buildVersionHistory(snapshots: IVersionSnapshot[], fields: IFieldConfig[]): IVersionEntry[] {
  const tracked = fields.filter((f) => !f.builtIn);
  const ordered = snapshots.slice().sort((a, b) => a.versionId - b.versionId);
  const entries: IVersionEntry[] = [];

  ordered.forEach((snapshot, index) => {
    const previous = index > 0 ? ordered[index - 1] : undefined;
    const changes: IFieldChange[] = [];
    tracked.forEach((field) => {
      const newValue = formatFieldValue(field, snapshot.values[field.key]);
      const oldValue = previous ? formatFieldValue(field, previous.values[field.key]) : '';
      if (newValue !== oldValue) {
        changes.push({ fieldKey: field.key, fieldName: field.displayName, oldValue, newValue });
      }
    });
    if (index > 0 && changes.length === 0) return;
    entries.push({
      versionLabel: snapshot.versionLabel,
      modified: snapshot.modified,
      modifiedBy: snapshot.modifiedBy,
      isCreation: index === 0,
      changes
    });
  });

  return entries.reverse();
}
