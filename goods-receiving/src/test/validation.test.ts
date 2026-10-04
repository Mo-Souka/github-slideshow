import { getFormFields } from '../config/appConfig';
import { FieldValues } from '../models/IReceivingRecord';
import { hasErrors, isEmptyValue, validateField, validateValues } from '../logic/validation';
import { field, testConfig } from './testConfig';

const config = testConfig();
const formFields = getFormFields(config);

function completeValues(): FieldValues {
  return {
    receivingDate: new Date(2026, 9, 4),
    supplierName: 'ACME Steel',
    purchaseOrderNumber: '4500012345',
    deliveryNoteNumber: 'DN-1',
    invoiceNumber: 'INV-1',
    truckNumber: 'B-AB 123',
    materialNumber: 'MAT-100',
    materialDescription: 'Steel sheet 1mm',
    quantityReceived: 12.5,
    unitOfMeasure: 'KG',
    batchNumber: 'B-77',
    warehouseLocation: 'A-01-02',
    receiverName: { id: 10, title: 'Jane Receiver' },
    inspectionStatus: 'Accepted'
  };
}

describe('validation', () => {
  it('accepts a complete record for submission', () => {
    expect(validateValues(formFields, completeValues(), 'submit')).toEqual({});
  });

  it('only requires the draft fields when saving a draft', () => {
    const errors = validateValues(formFields, { receivingDate: new Date(), supplierName: 'X', purchaseOrderNumber: '1' }, 'draft');
    expect(errors).toEqual({});
    const missing = validateValues(formFields, {}, 'draft');
    expect(Object.keys(missing).sort()).toEqual(['purchaseOrderNumber', 'receivingDate', 'supplierName']);
  });

  it('requires all required fields when submitting', () => {
    const errors = validateValues(formFields, { supplierName: 'X' }, 'submit');
    expect(errors.materialNumber).toBe('Material Number is required.');
    expect(errors.quantityReceived).toBe('Quantity Received is required.');
    expect(errors.receiverName).toBe('Receiver Name is required.');
    expect(errors.supplierName).toBeUndefined();
  });

  it('requires quantity greater than zero, in drafts too', () => {
    const values = { ...completeValues(), quantityReceived: 0 };
    expect(validateValues(formFields, values, 'submit').quantityReceived).toBe('Quantity Received must be greater than 0.');
    expect(validateValues(formFields, values, 'draft').quantityReceived).toBe('Quantity Received must be greater than 0.');
    expect(validateField(field(config, 'quantityReceived'), { quantityReceived: -1 }, 'draft')).toBeDefined();
    expect(validateField(field(config, 'quantityReceived'), { quantityReceived: NaN }, 'draft')).toBe(
      'Quantity Received must be a number.'
    );
    expect(validateField(field(config, 'quantityReceived'), { quantityReceived: 1.2345 }, 'draft')).toBe(
      'Quantity Received can have at most 3 decimal places.'
    );
  });

  it('requires comments when inspection is rejected or partially accepted', () => {
    const rejected = { ...completeValues(), inspectionStatus: 'Rejected' };
    const partial = { ...completeValues(), inspectionStatus: 'Partially Accepted' };
    const message = 'Comments are required when the inspection status is Rejected or Partially Accepted.';
    expect(validateValues(formFields, rejected, 'submit').receivingComments).toBe(message);
    expect(validateValues(formFields, partial, 'submit').receivingComments).toBe(message);
    expect(validateValues(formFields, { ...rejected, receivingComments: '  ' }, 'submit').receivingComments).toBe(message);
    expect(validateValues(formFields, { ...rejected, receivingComments: 'Two pallets damaged' }, 'submit')).toEqual({});
    expect(validateValues(formFields, completeValues(), 'submit').receivingComments).toBeUndefined();
  });

  it('requires supervisor comments when rejecting', () => {
    const supervisorComments = field(config, 'supervisorComments');
    expect(validateField(supervisorComments, { approvalStatus: 'Rejected' }, 'submit')).toBe(
      'Please explain why the record is rejected.'
    );
    expect(validateField(supervisorComments, { approvalStatus: 'Approved' }, 'submit')).toBeUndefined();
  });

  it('checks maximum length and choices', () => {
    const values = { ...completeValues(), purchaseOrderNumber: 'x'.repeat(51), unitOfMeasure: 'TONS' };
    const errors = validateValues(formFields, values, 'submit');
    expect(errors.purchaseOrderNumber).toContain('at most 50 characters');
    expect(errors.unitOfMeasure).toContain('must be one of');
    expect(hasErrors(errors)).toBe(true);
  });

  it('treats blank strings and missing people as empty', () => {
    expect(isEmptyValue('   ')).toBe(true);
    expect(isEmptyValue(undefined)).toBe(true);
    expect(isEmptyValue({ id: 0, title: '' })).toBe(true);
    expect(isEmptyValue(0)).toBe(false);
  });
});
