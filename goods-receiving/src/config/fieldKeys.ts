import { FieldType } from '../models/IFieldConfig';

/**
 * Field keys the application logic depends on (workflow, record ID, audit info).
 * They must exist in fields.json with the listed type. Every other field is
 * fully driven by the configuration and can be added or removed freely.
 */
export const FieldKeys = {
  recordId: 'recordId',
  receivingDate: 'receivingDate',
  receiverName: 'receiverName',
  approvalStatus: 'approvalStatus',
  supervisorComments: 'supervisorComments',
  reviewedBy: 'reviewedBy',
  reviewedOn: 'reviewedOn',
  createdBy: 'createdBy',
  created: 'created',
  modifiedBy: 'modifiedBy',
  modified: 'modified'
} as const;

export const RequiredFieldTypes: Record<string, FieldType> = {
  [FieldKeys.recordId]: 'Text',
  [FieldKeys.receivingDate]: 'Date',
  [FieldKeys.receiverName]: 'User',
  [FieldKeys.approvalStatus]: 'Choice',
  [FieldKeys.supervisorComments]: 'Note',
  [FieldKeys.reviewedBy]: 'User',
  [FieldKeys.reviewedOn]: 'DateTime',
  [FieldKeys.createdBy]: 'User',
  [FieldKeys.created]: 'DateTime',
  [FieldKeys.modifiedBy]: 'User',
  [FieldKeys.modified]: 'DateTime'
};

/** Approval status values. They must match the choices of the approvalStatus field. */
export const ApprovalStatus = {
  Draft: 'Draft',
  PendingApproval: 'Pending Approval',
  Approved: 'Approved',
  Rejected: 'Rejected'
} as const;

export type ApprovalStatusValue = (typeof ApprovalStatus)[keyof typeof ApprovalStatus];
