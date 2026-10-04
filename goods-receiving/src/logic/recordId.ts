/**
 * Record IDs look like GR-2026-000123: prefix, year the record was created,
 * and the SharePoint item ID padded with zeros. Because the number is the item
 * ID it is always unique; it does not restart every year.
 */

function pad(value: number, digits: number): string {
  let text = String(value);
  while (text.length < digits) {
    text = '0' + text;
  }
  return text;
}

export function formatRecordId(prefix: string, digits: number, itemId: number, created: Date): string {
  return `${prefix}-${created.getFullYear()}-${pad(itemId, digits)}`;
}

/**
 * Returns the SharePoint item ID if `text` is a full record ID such as
 * "GR-2026-000123" (case-insensitive, surrounding spaces ignored), otherwise undefined.
 */
export function parseRecordId(prefix: string, text: string): number | undefined {
  const match = /^\s*([A-Za-z]+)-(\d{4})-(\d+)\s*$/.exec(text || '');
  if (!match || match[1].toUpperCase() !== prefix.toUpperCase()) return undefined;
  const id = parseInt(match[3], 10);
  return id > 0 ? id : undefined;
}
