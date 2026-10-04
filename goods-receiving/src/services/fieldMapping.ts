import { IFieldConfig } from '../models/IFieldConfig';
import { FieldValue, FieldValues, isReferenceValue } from '../models/IReceivingRecord';
import { parseSharePointDate, toDateOnlyPayload } from '../logic/dates';

/**
 * Converts between SharePoint's JSON formats and the app's FieldValues.
 * SharePoint returns the same field in three different shapes depending on the
 * endpoint (item REST, RenderListDataAsStream, version history), so each has
 * its own reader. All of them are driven by fields.json.
 */

type Json = Record<string, unknown>;

// ---------------------------------------------------------------------------
// Item REST endpoint (items.getById)
// ---------------------------------------------------------------------------

export function getRestSelect(fields: IFieldConfig[]): { select: string[]; expand: string[] } {
  const select: string[] = ['Id'];
  const expand: string[] = [];
  fields.forEach((field) => {
    const name = field.internalName;
    if (field.type === 'User') {
      select.push(`${name}/Id`, `${name}/Title`, `${name}/EMail`, `${name}/Name`);
      expand.push(name);
    } else if (field.type === 'Lookup' && field.lookup) {
      select.push(`${name}/Id`, `${name}/${field.lookup.showField}`);
      expand.push(name);
    } else {
      select.push(name);
    }
  });
  return { select, expand };
}

function readText(value: unknown): string | undefined {
  if (value === undefined || value === null) return undefined;
  const text = String(value);
  return text === '' ? undefined : text;
}

function readNumber(value: unknown): number | undefined {
  if (typeof value === 'number') return isFinite(value) ? value : undefined;
  if (typeof value !== 'string' || value.trim() === '') return undefined;
  let text = value.trim().replace(/\s/g, '');
  const lastComma = text.lastIndexOf(',');
  const lastDot = text.lastIndexOf('.');
  if (lastComma > lastDot) {
    // 1.234,5 -> 1234.5
    text = text.replace(/\./g, '').replace(',', '.');
  } else {
    // 1,234.5 -> 1234.5
    text = text.replace(/,/g, '');
  }
  const number = parseFloat(text);
  return isFinite(number) ? number : undefined;
}

export function fromRestItem(item: Json, fields: IFieldConfig[]): FieldValues {
  const values: FieldValues = {};
  fields.forEach((field) => {
    const raw = item[field.internalName];
    values[field.key] = readRestValue(field, raw);
  });
  return values;
}

function readRestValue(field: IFieldConfig, raw: unknown): FieldValue {
  if (raw === undefined || raw === null) return undefined;
  switch (field.type) {
    case 'Number':
      return readNumber(raw);
    case 'Date':
    case 'DateTime':
      return parseSharePointDate(raw);
    case 'User': {
      const user = raw as Json;
      const id = Number(user.Id);
      if (!(id > 0)) return undefined;
      return { id, title: String(user.Title || ''), email: readText(user.EMail), loginName: readText(user.Name) };
    }
    case 'Lookup': {
      const lookup = raw as Json;
      const id = Number(lookup.Id);
      if (!(id > 0)) return undefined;
      const showField = field.lookup ? field.lookup.showField : 'Title';
      return { id, title: String(lookup[showField] || '') };
    }
    default:
      return readText(raw);
  }
}

export function getRestEtag(item: Json): string | undefined {
  const etag = item['odata.etag'] || item['@odata.etag'];
  if (etag) return String(etag);
  const metadata = item.__metadata as Json | undefined;
  return metadata && metadata.etag ? String(metadata.etag) : undefined;
}

// ---------------------------------------------------------------------------
// RenderListDataAsStream rows (record list, export)
// ---------------------------------------------------------------------------

const HTML_ENTITIES: Record<string, string> = { '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&#39;': "'", '&#160;': ' ', '&nbsp;': ' ' };

function decodeHtml(text: string): string {
  return text.replace(/&(amp|lt|gt|quot|nbsp|#39|#160);/g, (entity) => HTML_ENTITIES[entity] || entity);
}

function noteToPlainText(text: string): string {
  return decodeHtml(
    text
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/(p|div)>/gi, '\n')
      .replace(/<[^>]+>/g, '')
  ).replace(/\n{3,}/g, '\n\n').trim();
}

function firstOf(value: unknown): Json | undefined {
  if (Array.isArray(value)) return value.length > 0 ? (value[0] as Json) : undefined;
  if (value && typeof value === 'object') return value as Json;
  return undefined;
}

function readRenderValue(field: IFieldConfig, row: Json): FieldValue {
  const name = field.internalName;
  const raw = row[name];
  // For numbers and dates SharePoint adds a "Name." property with the
  // unformatted value; prefer it over the locale-formatted display text.
  const unformatted = row[`${name}.`];

  switch (field.type) {
    case 'Number':
      return readNumber(unformatted !== undefined && unformatted !== '' ? unformatted : raw);
    case 'Date':
    case 'DateTime':
      return parseSharePointDate(unformatted) || parseSharePointDate(raw);
    case 'User': {
      const user = firstOf(raw);
      if (!user) return undefined;
      const id = Number(user.id);
      if (!(id > 0)) return undefined;
      return { id, title: decodeHtml(String(user.title || '')), email: readText(user.email) };
    }
    case 'Lookup': {
      const lookup = firstOf(raw);
      if (!lookup) return undefined;
      const id = Number(lookup.lookupId);
      if (!(id > 0)) return undefined;
      return { id, title: decodeHtml(String(lookup.lookupValue || '')) };
    }
    case 'Note': {
      const text = readText(raw);
      return text === undefined ? undefined : noteToPlainText(text) || undefined;
    }
    default: {
      const text = readText(raw);
      return text === undefined ? undefined : decodeHtml(text);
    }
  }
}

export function fromRenderRow(row: Json, fields: IFieldConfig[]): FieldValues {
  const values: FieldValues = {};
  fields.forEach((field) => {
    values[field.key] = readRenderValue(field, row);
  });
  return values;
}

// ---------------------------------------------------------------------------
// Version history (items.getById(id).versions)
// ---------------------------------------------------------------------------

function readVersionValue(field: IFieldConfig, version: Json): FieldValue {
  const raw = version[field.internalName];
  if (raw === undefined || raw === null || raw === '') return undefined;
  switch (field.type) {
    case 'Number':
      return readNumber(raw);
    case 'Date':
    case 'DateTime':
      return parseSharePointDate(raw);
    case 'User':
    case 'Lookup': {
      const ref = firstOf(raw);
      if (!ref) return undefined;
      const id = Number(ref.LookupId);
      if (!(id > 0)) return undefined;
      return { id, title: String(ref.LookupValue || '') };
    }
    default:
      return readText(raw);
  }
}

export function fromVersion(version: Json, fields: IFieldConfig[]): FieldValues {
  const values: FieldValues = {};
  fields.forEach((field) => {
    values[field.key] = readVersionValue(field, version);
  });
  return values;
}

// ---------------------------------------------------------------------------
// Writing
// ---------------------------------------------------------------------------

/**
 * Builds the REST payload for add/update. Empty values are sent as null so
 * that clearing a field in the form also clears it in SharePoint.
 */
export function toRestPayload(values: FieldValues, fields: IFieldConfig[]): Record<string, unknown> {
  const payload: Record<string, unknown> = {};
  fields.forEach((field) => {
    if (!(field.key in values)) return;
    const value = values[field.key];
    const name = field.internalName;
    const empty = null;

    switch (field.type) {
      case 'User':
      case 'Lookup':
        payload[`${name}Id`] = isReferenceValue(value) && value.id > 0 ? value.id : empty;
        break;
      case 'Number':
        payload[name] = typeof value === 'number' && isFinite(value) ? value : empty;
        break;
      case 'Date':
        payload[name] = value instanceof Date ? toDateOnlyPayload(value) : empty;
        break;
      case 'DateTime':
        payload[name] = value instanceof Date ? value.toISOString() : empty;
        break;
      default:
        payload[name] = typeof value === 'string' && value.trim() !== '' ? value.trim() : empty;
        break;
    }
  });
  return payload;
}
