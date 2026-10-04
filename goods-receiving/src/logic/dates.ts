/**
 * Date helpers.
 *
 * "Date only" fields (such as Receiving Date) are stored by the app at 12:00 UTC.
 * SharePoint shows the date part in the site's time zone, and noon UTC falls on
 * the same calendar day in every time zone between UTC-11 and UTC+11. Dates
 * entered through the standard SharePoint form are stored as local midnight,
 * which the browser converts back to the same calendar day as long as the
 * device and the site use the same time zone (see DEPLOYMENT.md).
 */

function pad2(value: number): string {
  return value < 10 ? '0' + value : String(value);
}

/** Calendar date in the browser's time zone as yyyy-mm-dd. */
export function toIsoDateString(date: Date): string {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}

/** Value to send to SharePoint for a date-only field. */
export function toDateOnlyPayload(date: Date): string {
  return `${toIsoDateString(date)}T12:00:00Z`;
}

/** Local date at midnight for a yyyy-mm-dd string. */
export function fromIsoDateString(text: string): Date | undefined {
  const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(text || '');
  if (!match) return undefined;
  const date = new Date(parseInt(match[1], 10), parseInt(match[2], 10) - 1, parseInt(match[3], 10));
  return isNaN(date.getTime()) ? undefined : date;
}

/** Today at local midnight. */
export function today(): Date {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

export function addDays(date: Date, days: number): Date {
  const result = new Date(date.getTime());
  result.setDate(result.getDate() + days);
  return result;
}

/** Strips the time part (local time). */
export function startOfDay(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

/**
 * Parses a date coming from SharePoint (ISO 8601 such as 2026-10-04T12:00:00Z).
 * Returns undefined for empty or unparseable values.
 */
export function parseSharePointDate(value: unknown): Date | undefined {
  if (value instanceof Date) return isNaN(value.getTime()) ? undefined : value;
  if (typeof value !== 'string' || value.trim() === '') return undefined;
  const text = value.trim();
  if (/^\d{4}-\d{2}-\d{2}T/.test(text)) {
    const date = new Date(text);
    return isNaN(date.getTime()) ? undefined : date;
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    return fromIsoDateString(text);
  }
  const fallback = new Date(text);
  return isNaN(fallback.getTime()) ? undefined : fallback;
}
