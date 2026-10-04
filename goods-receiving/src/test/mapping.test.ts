import { fromRenderRow, fromRestItem, fromVersion, getRestEtag, getRestSelect, toRestPayload } from '../services/fieldMapping';
import { formatRecordId, parseRecordId } from '../logic/recordId';
import { formatDate, formatFieldValue, formatNumber } from '../logic/format';
import { field, testConfig } from './testConfig';

const config = testConfig();

describe('record ID', () => {
  it('formats GR-YYYY-000001 from the item ID and creation year', () => {
    expect(formatRecordId('GR', 6, 1, new Date(2026, 0, 15))).toBe('GR-2026-000001');
    expect(formatRecordId('GR', 6, 1234567, new Date(2027, 5, 1))).toBe('GR-2027-1234567');
  });

  it('parses full record IDs only', () => {
    expect(parseRecordId('GR', 'GR-2026-000123')).toBe(123);
    expect(parseRecordId('GR', 'gr-2026-42')).toBe(42);
    expect(parseRecordId('GR', '000123')).toBeUndefined();
    expect(parseRecordId('GR', 'GR-2026-000000')).toBeUndefined();
  });
});

describe('RenderListDataAsStream mapping', () => {
  const row = {
    ID: '42',
    Title: 'GR-2026-000042',
    ReceivingDate: '10/4/2026',
    'ReceivingDate.': '2026-10-04T12:00:00Z',
    SupplierName: 'Smith &amp; Sons',
    QuantityReceived: '1,234.5',
    'QuantityReceived.': '1234.5',
    UnitOfMeasure: 'KG',
    ReceiverName: [{ id: '15', title: 'Jane Receiver', email: 'jane@contoso.com' }],
    ReceivingComments: 'Line one<br>Line two',
    Author: [{ id: '15', title: 'Jane Receiver' }],
    Created: '2026-10-04T08:15:00Z',
    InspectionStatus: ''
  };

  it('reads typed values', () => {
    const values = fromRenderRow(row, config.fields);
    expect(values.recordId).toBe('GR-2026-000042');
    expect((values.receivingDate as Date).toISOString()).toBe('2026-10-04T12:00:00.000Z');
    expect(values.supplierName).toBe('Smith & Sons');
    expect(values.quantityReceived).toBe(1234.5);
    expect(values.receiverName).toEqual({ id: 15, title: 'Jane Receiver', email: 'jane@contoso.com' });
    expect(values.receivingComments).toBe('Line one\nLine two');
    expect(values.createdBy).toEqual({ id: 15, title: 'Jane Receiver', email: undefined });
    expect(values.inspectionStatus).toBeUndefined();
  });

  it('falls back to the formatted number when the raw value is missing', () => {
    const values = fromRenderRow({ QuantityReceived: '1.234,5' }, config.fields);
    expect(values.quantityReceived).toBe(1234.5);
  });
});

describe('REST item mapping', () => {
  it('builds select and expand for people fields', () => {
    const { select, expand } = getRestSelect(config.fields);
    expect(select).toContain('ReceiverName/Title');
    expect(select).toContain('Author/Id');
    expect(expand).toEqual(expect.arrayContaining(['ReceiverName', 'Author', 'Editor', 'ReviewedBy']));
  });

  it('reads an item', () => {
    const item = {
      'odata.etag': '"3"',
      Title: 'GR-2026-000007',
      ReceivingDate: '2026-10-04T12:00:00Z',
      QuantityReceived: 5,
      ReceiverName: { Id: 15, Title: 'Jane', EMail: 'jane@contoso.com', Name: 'i:0#.f|membership|jane@contoso.com' },
      ReviewedBy: null
    };
    const values = fromRestItem(item, config.fields);
    expect(getRestEtag(item)).toBe('"3"');
    expect(values.quantityReceived).toBe(5);
    expect(values.receiverName).toEqual({
      id: 15,
      title: 'Jane',
      email: 'jane@contoso.com',
      loginName: 'i:0#.f|membership|jane@contoso.com'
    });
    expect(values.reviewedBy).toBeUndefined();
  });

  it('builds a payload with date-only at noon UTC, person IDs and nulls for cleared fields', () => {
    const editable = config.fields.filter((f) => !f.readOnly && !f.builtIn);
    const payload = toRestPayload(
      {
        receivingDate: new Date(2026, 9, 4),
        supplierName: '  ACME  ',
        quantityReceived: 3,
        receiverName: { id: 15, title: 'Jane' },
        batchNumber: '',
        approvalStatus: 'Approved'
      },
      editable
    );
    expect(payload).toEqual({
      ReceivingDate: '2026-10-04T12:00:00Z',
      SupplierName: 'ACME',
      QuantityReceived: 3,
      ReceiverNameId: 15,
      BatchNumber: null
    });
  });
});

describe('version mapping', () => {
  it('reads version values', () => {
    const values = fromVersion(
      { SupplierName: 'ACME', QuantityReceived: 2, ReceiverName: { LookupId: 4, LookupValue: 'Bob', Email: 'b@x' } },
      config.fields
    );
    expect(values.supplierName).toBe('ACME');
    expect(values.quantityReceived).toBe(2);
    expect(values.receiverName).toEqual({ id: 4, title: 'Bob' });
  });
});

describe('formatting', () => {
  it('formats dates unambiguously and numbers with separators', () => {
    expect(formatDate(new Date(2026, 9, 4))).toBe('04 Oct 2026');
    expect(formatNumber(1234.5, 3)).toBe('1,234.5');
    expect(formatNumber(12, 3)).toBe('12');
    expect(formatFieldValue(field(config, 'receiverName'), { id: 1, title: 'Jane' })).toBe('Jane');
    expect(formatFieldValue(field(config, 'quantityReceived'), undefined)).toBe('');
  });
});
