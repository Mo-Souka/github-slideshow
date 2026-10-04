import { IFieldConfig } from '../models/IFieldConfig';
import { parseRecordId } from './recordId';

/**
 * Builds the CAML query that the record list sends to SharePoint.
 *
 * Large-list strategy (5,000 item list view threshold):
 *   - Conditions on indexed columns are placed first, starting with the
 *     Receiving Date range, so SharePoint can use the index to narrow the
 *     result before it evaluates anything else.
 *   - Free-text search ("contains") cannot use an index, so it only scales
 *     when combined with a selective indexed filter such as the date range.
 *   - A full record ID (GR-2026-000123) is turned into an ID lookup, which
 *     always works regardless of list size.
 */

export type FilterValue =
  | { kind: 'text'; value: string }
  | { kind: 'choice'; values: string[] }
  | { kind: 'person'; id: number; title: string }
  | { kind: 'lookup'; id: number; title: string }
  /** Dates as yyyy-mm-dd (local calendar dates). */
  | { kind: 'dateRange'; from?: string; to?: string }
  | { kind: 'numberRange'; min?: number; max?: number };

export interface IRecordQuery {
  search: string;
  /** Filter values keyed by field key. */
  filters: Record<string, FilterValue | undefined>;
  sortKey: string;
  sortAscending: boolean;
}

export interface ICamlOptions {
  fields: IFieldConfig[];
  recordIdPrefix: string;
  rowLimit: number;
  /** Fields returned for each row. */
  viewFieldKeys: string[];
  /** The primary date range field; its condition always comes first. */
  dateRangeFieldKey: string;
}

export function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function fieldRef(field: IFieldConfig, lookupId: boolean = false): string {
  return `<FieldRef Name="${escapeXml(field.internalName)}"${lookupId ? ' LookupId="TRUE"' : ''} />`;
}

function camlValueType(field: IFieldConfig): string {
  switch (field.type) {
    case 'Number':
      return 'Number';
    case 'Date':
    case 'DateTime':
      return 'DateTime';
    case 'Choice':
      return 'Choice';
    case 'User':
    case 'Lookup':
      return 'Integer';
    case 'Note':
      return 'Note';
    default:
      return 'Text';
  }
}

function isFilterEmpty(filter: FilterValue | undefined): boolean {
  if (!filter) return true;
  switch (filter.kind) {
    case 'text':
      return filter.value.trim() === '';
    case 'choice':
      return filter.values.length === 0;
    case 'person':
    case 'lookup':
      return !(filter.id > 0);
    case 'dateRange':
      return !filter.from && !filter.to;
    case 'numberRange':
      return filter.min === undefined && filter.max === undefined;
    default:
      return true;
  }
}

/** CAML conditions for one filter (a range produces two conditions). */
function filterConditions(field: IFieldConfig, filter: FilterValue): string[] {
  const type = camlValueType(field);
  switch (filter.kind) {
    case 'text':
      return [`<Contains>${fieldRef(field)}<Value Type="${type}">${escapeXml(filter.value.trim())}</Value></Contains>`];
    case 'choice':
      if (filter.values.length === 1) {
        return [`<Eq>${fieldRef(field)}<Value Type="${type}">${escapeXml(filter.values[0])}</Value></Eq>`];
      }
      return [
        `<In>${fieldRef(field)}<Values>${filter.values
          .map((v) => `<Value Type="${type}">${escapeXml(v)}</Value>`)
          .join('')}</Values></In>`
      ];
    case 'person':
    case 'lookup':
      return [`<Eq>${fieldRef(field, true)}<Value Type="Integer">${Math.floor(filter.id)}</Value></Eq>`];
    case 'dateRange': {
      const result: string[] = [];
      if (filter.from) {
        result.push(`<Geq>${fieldRef(field)}<Value Type="DateTime" IncludeTimeValue="FALSE">${escapeXml(filter.from)}</Value></Geq>`);
      }
      if (filter.to) {
        result.push(`<Leq>${fieldRef(field)}<Value Type="DateTime" IncludeTimeValue="FALSE">${escapeXml(filter.to)}</Value></Leq>`);
      }
      return result;
    }
    case 'numberRange': {
      const result: string[] = [];
      if (filter.min !== undefined) result.push(`<Geq>${fieldRef(field)}<Value Type="Number">${filter.min}</Value></Geq>`);
      if (filter.max !== undefined) result.push(`<Leq>${fieldRef(field)}<Value Type="Number">${filter.max}</Value></Leq>`);
      return result;
    }
    default:
      return [];
  }
}

/** Combines conditions with a binary operator: <And>a<And>b c</And></And>. */
function combine(operator: 'And' | 'Or', conditions: string[]): string {
  if (conditions.length === 0) return '';
  if (conditions.length === 1) return conditions[0];
  return `<${operator}>${conditions[0]}${combine(operator, conditions.slice(1))}</${operator}>`;
}

function buildViewFields(options: ICamlOptions): string {
  const names: string[] = ['ID'];
  options.viewFieldKeys.forEach((key) => {
    const field = options.fields.filter((f) => f.key === key)[0];
    if (field && names.indexOf(field.internalName) < 0) {
      names.push(field.internalName);
    }
  });
  return `<ViewFields>${names.map((n) => `<FieldRef Name="${escapeXml(n)}" />`).join('')}</ViewFields>`;
}

function buildOrderBy(query: IRecordQuery, fields: IFieldConfig[]): string {
  const field = fields.filter((f) => f.key === query.sortKey)[0];
  const name = field ? field.internalName : 'ID';
  return `<OrderBy><FieldRef Name="${escapeXml(name)}" Ascending="${query.sortAscending ? 'TRUE' : 'FALSE'}" /></OrderBy>`;
}

/** Builds the WHERE clause (without the <Where> element). */
export function buildWhere(query: IRecordQuery, options: ICamlOptions): string {
  const recordItemId = parseRecordId(options.recordIdPrefix, query.search);
  if (recordItemId !== undefined) {
    return `<Eq><FieldRef Name="ID" /><Value Type="Counter">${recordItemId}</Value></Eq>`;
  }

  const primary: string[] = [];
  const indexed: string[] = [];
  const other: string[] = [];

  options.fields.forEach((field) => {
    const filter = query.filters[field.key];
    if (!filter || isFilterEmpty(filter)) return;
    const conditions = filterConditions(field, filter);
    if (field.key === options.dateRangeFieldKey) {
      primary.push(...conditions);
    } else if (field.indexed) {
      indexed.push(...conditions);
    } else {
      other.push(...conditions);
    }
  });

  const search = (query.search || '').trim();
  if (search) {
    const searchConditions = options.fields
      .filter((f) => f.searchable)
      .map((f) => `<Contains>${fieldRef(f)}<Value Type="Text">${escapeXml(search)}</Value></Contains>`);
    if (searchConditions.length > 0) {
      other.push(combine('Or', searchConditions));
    }
  }

  return combine('And', primary.concat(indexed, other));
}

/** Full <View> XML for RenderListDataAsStream. */
export function buildViewXml(query: IRecordQuery, options: ICamlOptions): string {
  const where = buildWhere(query, options);
  const whereXml = where ? `<Where>${where}</Where>` : '';
  return (
    '<View>' +
    buildViewFields(options) +
    `<Query>${whereXml}${buildOrderBy(query, options.fields)}</Query>` +
    `<RowLimit Paged="TRUE">${Math.max(1, Math.floor(options.rowLimit))}</RowLimit>` +
    '</View>'
  );
}

/** Number of active filters (for the "Filters (3)" button label). */
export function countActiveFilters(query: IRecordQuery, excludeKeys: string[] = []): number {
  return Object.keys(query.filters).filter((key) => excludeKeys.indexOf(key) < 0 && !isFilterEmpty(query.filters[key]))
    .length;
}

export { isFilterEmpty };
