export interface IPersonValue {
  /** SharePoint site user ID (the ID in the site's User Information List). */
  id: number;
  title: string;
  email?: string;
  loginName?: string;
}

export interface ILookupValue {
  id: number;
  title: string;
}

/** `undefined` means "empty". */
export type FieldValue = string | number | Date | IPersonValue | ILookupValue | undefined;

/** Field values keyed by the `key` of each field in fields.json. */
export type FieldValues = Record<string, FieldValue>;

export interface IReceivingRecord {
  /** SharePoint list item ID. */
  id: number;
  /** ETag used to detect concurrent edits. Only present when loaded through the item endpoint. */
  etag?: string;
  values: FieldValues;
}

export interface IRecordPage {
  records: IReceivingRecord[];
  /** Opaque token for the next page, or undefined on the last page. */
  nextPageToken?: string;
}

export interface IFieldChange {
  fieldKey: string;
  fieldName: string;
  oldValue: string;
  newValue: string;
}

export interface IVersionEntry {
  versionLabel: string;
  modified?: Date;
  modifiedBy?: string;
  /** True for the first version (record created). */
  isCreation: boolean;
  changes: IFieldChange[];
}

/** True for person and lookup values (both have an id and a title). */
export function isReferenceValue(value: FieldValue): value is IPersonValue | ILookupValue {
  return typeof value === 'object' && !(value instanceof Date);
}
