import { buildViewXml, buildWhere, countActiveFilters, ICamlOptions, IRecordQuery } from '../logic/caml';
import { testConfig } from './testConfig';

const config = testConfig();
const options: ICamlOptions = {
  fields: config.fields,
  recordIdPrefix: 'GR',
  rowLimit: 25,
  viewFieldKeys: ['recordId', 'receivingDate', 'supplierName'],
  dateRangeFieldKey: 'receivingDate'
};

function query(overrides: Partial<IRecordQuery> = {}): IRecordQuery {
  return { search: '', filters: {}, sortKey: 'receivingDate', sortAscending: false, ...overrides };
}

describe('CAML query builder', () => {
  it('puts the date range first so SharePoint can use the index', () => {
    const where = buildWhere(
      query({
        search: 'acme',
        filters: {
          inspectionStatus: { kind: 'choice', values: ['Rejected'] },
          receivingDate: { kind: 'dateRange', from: '2026-09-04', to: '2026-10-04' }
        }
      }),
      options
    );
    expect(where.indexOf('<And><Geq><FieldRef Name="ReceivingDate" />')).toBe(0);
    expect(where.indexOf('ReceivingDate')).toBeLessThan(where.indexOf('InspectionStatus'));
    expect(where.indexOf('InspectionStatus')).toBeLessThan(where.indexOf('<Or>'));
    expect(where).toContain('<Value Type="DateTime" IncludeTimeValue="FALSE">2026-10-04</Value></Leq>');
  });

  it('searches all searchable text fields with OR', () => {
    const where = buildWhere(query({ search: '4500' }), options);
    ['Title', 'SupplierName', 'PurchaseOrderNumber', 'DeliveryNoteNumber', 'MaterialNumber', 'BatchNumber'].forEach((name) => {
      expect(where).toContain(`<Contains><FieldRef Name="${name}" /><Value Type="Text">4500</Value></Contains>`);
    });
    expect(where).not.toContain('InvoiceNumber');
  });

  it('turns a full record ID into an ID lookup and ignores other filters', () => {
    const where = buildWhere(
      query({ search: ' gr-2026-000123 ', filters: { receivingDate: { kind: 'dateRange', from: '2026-10-01' } } }),
      options
    );
    expect(where).toBe('<Eq><FieldRef Name="ID" /><Value Type="Counter">123</Value></Eq>');
  });

  it('escapes user input', () => {
    const where = buildWhere(query({ search: 'A&B <x>' }), options);
    expect(where).toContain('A&amp;B &lt;x&gt;');
    expect(where).not.toContain('<x>');
  });

  it('uses Eq for one choice, In for several, and LookupId for people', () => {
    const one = buildWhere(query({ filters: { approvalStatus: { kind: 'choice', values: ['Pending Approval'] } } }), options);
    expect(one).toBe('<Eq><FieldRef Name="ApprovalStatus" /><Value Type="Choice">Pending Approval</Value></Eq>');
    const many = buildWhere(query({ filters: { approvalStatus: { kind: 'choice', values: ['Draft', 'Rejected'] } } }), options);
    expect(many).toContain('<In><FieldRef Name="ApprovalStatus" /><Values><Value Type="Choice">Draft</Value>');
    const person = buildWhere(query({ filters: { receiverName: { kind: 'person', id: 7, title: 'Jane' } } }), options);
    expect(person).toBe('<Eq><FieldRef Name="ReceiverName" LookupId="TRUE" /><Value Type="Integer">7</Value></Eq>');
  });

  it('builds a paged view with fields and sort order', () => {
    const xml = buildViewXml(query({ sortKey: 'supplierName', sortAscending: true }), options);
    expect(xml).toContain('<ViewFields><FieldRef Name="ID" /><FieldRef Name="Title" /><FieldRef Name="ReceivingDate" />');
    expect(xml).toContain('<OrderBy><FieldRef Name="SupplierName" Ascending="TRUE" /></OrderBy>');
    expect(xml).toContain('<RowLimit Paged="TRUE">25</RowLimit>');
    expect(xml).not.toContain('<Where>');
  });

  it('counts active filters', () => {
    const q = query({
      filters: {
        receivingDate: { kind: 'dateRange', from: '2026-01-01' },
        supplierName: { kind: 'text', value: ' ' },
        approvalStatus: { kind: 'choice', values: ['Draft'] }
      }
    });
    expect(countActiveFilters(q)).toBe(2);
    expect(countActiveFilters(q, ['receivingDate'])).toBe(1);
  });
});
