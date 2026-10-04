import { IFieldConfig } from '../models/IFieldConfig';
import { FieldValue, isReferenceValue } from '../models/IReceivingRecord';

/**
 * Display formatting. Dates use a fixed English format with the month name
 * (e.g. "04 Oct 2026") so they cannot be misread as day/month or month/day.
 */

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function pad2(value: number): string {
  return value < 10 ? '0' + value : String(value);
}

export function formatDate(date: Date): string {
  return `${pad2(date.getDate())} ${MONTHS[date.getMonth()]} ${date.getFullYear()}`;
}

export function formatDateTime(date: Date): string {
  return `${formatDate(date)} ${pad2(date.getHours())}:${pad2(date.getMinutes())}`;
}

export function formatNumber(value: number, decimals?: number): string {
  const maxDecimals = decimals === undefined ? 3 : decimals;
  const fixed = value.toFixed(maxDecimals);
  // Drop trailing zeros: 12.500 -> 12.5, 12.000 -> 12
  const trimmed = maxDecimals > 0 ? fixed.replace(/\.?0+$/, '') : fixed;
  const parts = trimmed.split('.');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return parts.join('.');
}

export function formatFileSize(bytes: number): string {
  if (!(bytes > 0)) return '0 KB';
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/** Text representation of a field value for tables, detail view and history. */
export function formatFieldValue(field: IFieldConfig, value: FieldValue): string {
  if (value === undefined) return '';
  switch (field.type) {
    case 'Date':
      return value instanceof Date ? formatDate(value) : String(value);
    case 'DateTime':
      return value instanceof Date ? formatDateTime(value) : String(value);
    case 'Number':
      return typeof value === 'number' && isFinite(value) ? formatNumber(value, field.decimals) : '';
    case 'User':
    case 'Lookup':
      return isReferenceValue(value) ? value.title : '';
    default:
      return typeof value === 'string' ? value : String(value);
  }
}
