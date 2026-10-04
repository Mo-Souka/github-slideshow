/**
 * Types for src/config/fields.json and src/config/solution.json.
 * The JSON schemas next to those files describe every property for editors.
 */

export type FieldType = 'Text' | 'Note' | 'Number' | 'Date' | 'DateTime' | 'Choice' | 'User' | 'Lookup';

export type BadgeTone = 'neutral' | 'info' | 'success' | 'warning' | 'danger';

export interface IFieldValidation {
  min?: number;
  max?: number;
  minExclusive?: number;
  maxExclusive?: number;
  integer?: boolean;
  maxLength?: number;
  pattern?: string;
  patternMessage?: string;
}

export interface IRequiredWhen {
  /** Key of the field whose value is checked. */
  field: string;
  /** The field becomes required when the other field has one of these values. */
  in: string[];
  message?: string;
}

export interface ILookupConfig {
  listTitle: string;
  showField: string;
}

export interface IFieldConfig {
  key: string;
  internalName: string;
  displayName: string;
  type: FieldType;
  description?: string;
  required: boolean;
  requiredForDraft?: boolean;
  requiredWhen?: IRequiredWhen;
  builtIn?: boolean;
  readOnly?: boolean;
  defaultValue?: string | number;
  choices?: string[];
  badgeTones?: Record<string, BadgeTone>;
  decimals?: number;
  validation?: IFieldValidation;
  lookup?: ILookupConfig;
  indexed?: boolean;
  order: number;
  section?: string;
  showInForm?: boolean;
  showInDetail?: boolean;
  showInTable: boolean;
  includeInExport: boolean;
  filterable?: boolean;
  searchable?: boolean;
  sortable?: boolean;
  tableWidth?: number;
}

export interface IFormSection {
  key: string;
  title: string;
}

export interface IFieldsConfigFile {
  sections: IFormSection[];
  fields: IFieldConfig[];
}

export interface IColumnName {
  internalName: string;
  displayName: string;
}

export interface ISolutionConfig {
  list: {
    title: string;
    url: string;
    description?: string;
    majorVersionLimit?: number;
  };
  library: {
    title: string;
    url: string;
    description?: string;
    majorVersionLimit?: number;
    recordLookupField: IColumnName;
    documentTypeField: IColumnName;
    documentTypes: string[];
    photoDocumentType: string;
    allowedExtensions: string[];
    maxFileSizeMB: number;
  };
  groups: {
    receivers: string;
    supervisors: string;
    supervisorPermissionLevel: string;
  };
  recordId: {
    prefix: string;
    digits: number;
  };
  listView: {
    dateRangeFieldKey: string;
    defaultDateWindowDays: number;
    pageSize: number;
    defaultSortFieldKey: string;
    defaultSortAscending: boolean;
  };
  export: {
    maxRows: number;
    fileNamePrefix: string;
    sheetName: string;
  };
}

/** Configuration after loading, validation and applying web part property overrides. */
export interface IAppConfig {
  sections: IFormSection[];
  /** All fields sorted by `order`. */
  fields: IFieldConfig[];
  solution: ISolutionConfig;
  /** Absolute URL of the site that holds the list and library. */
  siteUrl: string;
}
