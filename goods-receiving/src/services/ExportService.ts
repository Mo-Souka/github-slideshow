import type { SheetData } from 'write-excel-file/browser';
import { IAppConfig, IFieldConfig } from '../models/IFieldConfig';
import { FieldValue, IReceivingRecord, isReferenceValue } from '../models/IReceivingRecord';
import { toIsoDateString } from '../logic/dates';

/** Minimal cell shape accepted by write-excel-file. */
export interface IExportCell {
  value?: string | number | Date;
  type?: StringConstructor | NumberConstructor | DateConstructor;
  format?: string;
  fontWeight?: 'bold';
  backgroundColor?: string;
  wrap?: boolean;
}

export interface IExportSheet {
  rows: IExportCell[][];
  columnWidths: number[];
}

const DATE_FORMAT = 'dd-mmm-yyyy';
const DATE_TIME_FORMAT = 'dd-mmm-yyyy hh:mm';

export function getExportFields(config: IAppConfig): IFieldConfig[] {
  return config.fields.filter((f) => f.includeInExport);
}

function toCell(field: IFieldConfig, value: FieldValue): IExportCell {
  if (value === undefined) return {};
  switch (field.type) {
    case 'Date':
      return value instanceof Date ? { value: dateOnlyForExcel(value), type: Date, format: DATE_FORMAT } : {};
    case 'DateTime':
      return value instanceof Date ? { value, type: Date, format: DATE_TIME_FORMAT } : {};
    case 'Number':
      return typeof value === 'number' && isFinite(value) ? { value, type: Number } : {};
    case 'User':
    case 'Lookup':
      return isReferenceValue(value) ? { value: value.title, type: String } : {};
    case 'Note':
      return { value: String(value), type: String, wrap: true };
    default:
      return { value: String(value), type: String };
  }
}

/**
 * Excel stores dates without a time zone. Convert the local calendar date to
 * midnight UTC so the cell shows the same date the user sees in the app.
 */
function dateOnlyForExcel(date: Date): Date {
  return new Date(`${toIsoDateString(date)}T00:00:00Z`);
}

function columnWidth(field: IFieldConfig): number {
  switch (field.type) {
    case 'Date':
      return 13;
    case 'DateTime':
      return 18;
    case 'Number':
      return 12;
    case 'Note':
      return 40;
    default:
      return Math.max(12, Math.min(40, Math.round((field.tableWidth || 120) / 7)));
  }
}

/** Builds header + data rows. Pure function so it can be unit tested. */
export function buildExportSheet(records: IReceivingRecord[], fields: IFieldConfig[]): IExportSheet {
  const header: IExportCell[] = fields.map((f) => ({
    value: f.displayName,
    type: String,
    fontWeight: 'bold',
    backgroundColor: '#E7E6E6'
  }));
  const rows = records.map((record) => fields.map((f) => toCell(f, record.values[f.key])));
  return { rows: [header].concat(rows), columnWidths: fields.map(columnWidth) };
}

export function buildExportFileName(prefix: string, now: Date = new Date()): string {
  const pad = (n: number): string => (n < 10 ? '0' + n : String(n));
  return `${prefix}-${toIsoDateString(now)}-${pad(now.getHours())}${pad(now.getMinutes())}.xlsx`;
}

/** Creates the .xlsx file in the browser and starts the download. */
export async function downloadExcel(sheet: IExportSheet, sheetName: string, fileName: string): Promise<void> {
  // Loaded on demand so the Excel library does not slow down the first page load.
  const module = await import(/* webpackChunkName: 'goods-receiving-excel' */ 'write-excel-file/browser');
  const writeXlsxFile = module.default;
  const data: SheetData = sheet.rows;
  await writeXlsxFile(data, {
    sheet: sheetName,
    columns: sheet.columnWidths.map((width) => ({ width })),
    stickyRowsCount: 1
  }).toFile(fileName);
}
