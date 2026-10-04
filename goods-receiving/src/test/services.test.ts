import { toAppError, extractSharePointMessage } from '../services/errors';
import { buildExportFileName, buildExportSheet, getExportFields } from '../services/ExportService';
import { buildVersionHistory } from '../logic/versionDiff';
import { makeUniqueFileName, sanitizeFileName, validateFile } from '../logic/files';
import { AppError } from '../models/AppError';
import { testConfig } from './testConfig';

const config = testConfig();

function httpError(status: number, spMessage: string): unknown {
  return {
    isHttpRequestError: true,
    status,
    message: `Error making HttpClient request in queryable [${status}] ::> ${JSON.stringify({
      'odata.error': { code: '-1', message: { lang: 'en-US', value: spMessage } }
    })}`
  };
}

describe('error mapping', () => {
  it('extracts the SharePoint message', () => {
    expect(extractSharePointMessage('x ::> {"odata.error":{"message":{"value":"Boom"}}}')).toBe('Boom');
    expect(extractSharePointMessage('plain text')).toBe('plain text');
  });

  it('maps HTTP statuses to friendly errors', () => {
    expect(toAppError(httpError(403, 'Access denied.'), { action: 'save' }).kind).toBe('PermissionDenied');
    const notFound = toAppError(httpError(404, "List 'Material Receiving Records' does not exist at site with URL 'x'."), {
      listTitle: 'Material Receiving Records'
    });
    expect(notFound.kind).toBe('ListNotFound');
    expect(notFound.userMessage).toContain('provisioning script');
    expect(toAppError(httpError(404, 'Item does not exist.')).kind).toBe('NotFound');
    expect(toAppError(httpError(412, 'The version of the item has changed.')).kind).toBe('Conflict');
    expect(toAppError(httpError(429, 'Too many requests')).kind).toBe('Throttled');
    expect(
      toAppError(httpError(500, 'Microsoft.SharePoint.SPQueryThrottledException: The attempted operation is prohibited because it exceeds the list view threshold.'))
        .kind
    ).toBe('ThresholdExceeded');
    expect(toAppError(httpError(400, "Column 'Foo' does not exist.")).kind).toBe('ConfigurationMismatch');
  });

  it('recognises network failures', () => {
    expect(toAppError(new TypeError('Failed to fetch')).kind).toBe('Network');
  });

  it('passes AppErrors through unchanged', () => {
    const original = new AppError('Validation', 'Bad');
    expect(toAppError(original)).toBe(original);
    expect(original instanceof AppError).toBe(true);
  });
});

describe('version history', () => {
  it('lists changed fields per version and skips versions without tracked changes', () => {
    const history = buildVersionHistory(
      [
        { versionId: 512, versionLabel: '1.0', modifiedBy: 'Jane', values: { supplierName: 'ACME', quantityReceived: 5 } },
        { versionId: 1024, versionLabel: '2.0', modifiedBy: 'Jane', values: { supplierName: 'ACME', quantityReceived: 5, recordId: 'GR-2026-000001' } },
        { versionId: 1536, versionLabel: '3.0', modifiedBy: 'Bob', values: { supplierName: 'ACME Ltd', quantityReceived: 6, recordId: 'GR-2026-000001' } }
      ],
      config.fields
    );
    expect(history.map((h) => h.versionLabel)).toEqual(['3.0', '1.0']);
    expect(history[0].changes).toEqual([
      { fieldKey: 'supplierName', fieldName: 'Supplier Name', oldValue: 'ACME', newValue: 'ACME Ltd' },
      { fieldKey: 'quantityReceived', fieldName: 'Quantity Received', oldValue: '5', newValue: '6' }
    ]);
    expect(history[1].isCreation).toBe(true);
  });
});

describe('files', () => {
  it('sanitizes names and avoids duplicates', () => {
    expect(sanitizeFileName('a:b*c?.pdf')).toBe('a_b_c_.pdf');
    expect(sanitizeFileName('~$report.docx ')).toBe('report.docx');
    expect(makeUniqueFileName('image.jpg', ['IMAGE.jpg', 'image (2).jpg'])).toBe('image (3).jpg');
    expect(makeUniqueFileName('new.pdf', ['image.jpg'])).toBe('new.pdf');
  });

  it('validates type and size', () => {
    const allowed = config.solution.library.allowedExtensions;
    expect(validateFile({ name: 'photo.HEIC', size: 10 }, allowed, 250)).toBeUndefined();
    expect(validateFile({ name: 'virus.exe', size: 10 }, allowed, 250)).toContain('cannot be uploaded');
    expect(validateFile({ name: 'big.pdf', size: 300 * 1024 * 1024 }, allowed, 250)).toContain('larger than 250 MB');
  });
});

describe('Excel export', () => {
  it('uses display names as headers and real dates and numbers', () => {
    const fields = getExportFields(config);
    const sheet = buildExportSheet(
      [{ id: 1, values: { recordId: 'GR-2026-000001', receivingDate: new Date(2026, 9, 4), quantityReceived: 2.5, receiverName: { id: 1, title: 'Jane' } } }],
      fields
    );
    const headers = sheet.rows[0].map((c) => c.value);
    expect(headers[0]).toBe('Record ID');
    expect(headers).toContain('Purchase Order Number');
    expect(headers).toContain('Created Date');
    const dataRow = sheet.rows[1];
    const dateCell = dataRow[fields.map((f) => f.key).indexOf('receivingDate')];
    expect(dateCell.type).toBe(Date);
    expect((dateCell.value as Date).toISOString()).toBe('2026-10-04T00:00:00.000Z');
    expect(dataRow[fields.map((f) => f.key).indexOf('quantityReceived')]).toEqual({ value: 2.5, type: Number });
    expect(dataRow[fields.map((f) => f.key).indexOf('receiverName')].value).toBe('Jane');
    expect(sheet.columnWidths.length).toBe(fields.length);
  });

  it('names the file with date and time', () => {
    expect(buildExportFileName('Goods-Receiving', new Date(2026, 9, 4, 9, 5))).toBe('Goods-Receiving-2026-10-04-0905.xlsx');
  });
});
