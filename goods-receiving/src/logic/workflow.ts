import { ApprovalStatus, ApprovalStatusValue, FieldKeys } from '../config/fieldKeys';
import { IReceivingRecord, isReferenceValue } from '../models/IReceivingRecord';
import { IUserContext } from '../models/IUserContext';

/**
 * The approval workflow in one place. Every screen asks these functions what
 * the current user may do, so the rules cannot drift apart between screens.
 *
 *   Draft ──submit──> Pending Approval ──approve──> Approved (read-only)
 *     ^                     │
 *     │                  reject
 *     │                     v
 *     └──(edit, resubmit)── Rejected
 *
 * IMPORTANT: this is enforced in the user interface only. See README.md,
 * "Security model", for list settings that enforce it in SharePoint as well.
 */

export type WorkflowAction = 'saveDraft' | 'submit' | 'approve' | 'reject';

export interface IRecordPermissions {
  canEdit: boolean;
  canSubmit: boolean;
  canApprove: boolean;
  canReject: boolean;
  canManageAttachments: boolean;
  /** Explains why the record is read-only or why approval is not possible, if relevant. */
  readOnlyReason?: string;
  approvalBlockedReason?: string;
}

export function getStatus(record: IReceivingRecord | undefined): ApprovalStatusValue {
  if (!record) return ApprovalStatus.Draft;
  const value = record.values[FieldKeys.approvalStatus];
  return (typeof value === 'string' && value ? value : ApprovalStatus.Draft) as ApprovalStatusValue;
}

function personId(record: IReceivingRecord, key: string): number | undefined {
  const value = record.values[key];
  return isReferenceValue(value) ? value.id : undefined;
}

/** The user created the record or is named as its receiver. */
export function isOwnRecord(user: IUserContext, record: IReceivingRecord): boolean {
  return personId(record, FieldKeys.createdBy) === user.id || personId(record, FieldKeys.receiverName) === user.id;
}

export function canCreateRecords(user: IUserContext): boolean {
  return user.isReceiver || user.isSupervisor;
}

const EDITABLE_FOR_RECEIVERS: string[] = [ApprovalStatus.Draft, ApprovalStatus.Rejected];
const EDITABLE_FOR_SUPERVISORS: string[] = [ApprovalStatus.Draft, ApprovalStatus.Rejected, ApprovalStatus.PendingApproval];
const SUBMITTABLE: string[] = [ApprovalStatus.Draft, ApprovalStatus.Rejected];

/** What the user may do with an existing record. Pass undefined for a new record. */
export function getRecordPermissions(user: IUserContext, record: IReceivingRecord | undefined): IRecordPermissions {
  if (!record) {
    const canCreate = canCreateRecords(user);
    return {
      canEdit: canCreate,
      canSubmit: canCreate,
      canApprove: false,
      canReject: false,
      canManageAttachments: canCreate,
      readOnlyReason: canCreate ? undefined : 'You are not a member of the receiving team, so you cannot create records.'
    };
  }

  const status = getStatus(record);
  const own = isOwnRecord(user, record);
  let canEdit = false;
  let readOnlyReason: string | undefined;

  if (status === ApprovalStatus.Approved) {
    readOnlyReason = 'This record is approved and can no longer be changed.';
  } else if (user.isSupervisor) {
    canEdit = EDITABLE_FOR_SUPERVISORS.indexOf(status) >= 0;
  } else if (user.isReceiver) {
    if (!own) {
      readOnlyReason = 'You can only edit records you created or received yourself.';
    } else if (EDITABLE_FOR_RECEIVERS.indexOf(status) < 0) {
      readOnlyReason = 'This record is waiting for approval and cannot be changed until a supervisor approves or rejects it.';
    } else {
      canEdit = true;
    }
  } else {
    readOnlyReason = 'You have read-only access to receiving records.';
  }

  const isPending = status === ApprovalStatus.PendingApproval;
  let approvalBlockedReason: string | undefined;
  let canDecide = false;
  if (user.isSupervisor && isPending) {
    if (own) {
      approvalBlockedReason =
        'You created or received this record yourself, so another supervisor must approve or reject it.';
    } else {
      canDecide = true;
    }
  }

  return {
    canEdit,
    canSubmit: canEdit && SUBMITTABLE.indexOf(status) >= 0,
    canApprove: canDecide,
    canReject: canDecide,
    canManageAttachments: canEdit,
    readOnlyReason,
    approvalBlockedReason
  };
}

/** Status after an action. "Save" keeps the current status (new records start as Draft). */
export function getNextStatus(action: WorkflowAction, current: ApprovalStatusValue | undefined): ApprovalStatusValue {
  switch (action) {
    case 'submit':
      return ApprovalStatus.PendingApproval;
    case 'approve':
      return ApprovalStatus.Approved;
    case 'reject':
      return ApprovalStatus.Rejected;
    case 'saveDraft':
    default:
      return current || ApprovalStatus.Draft;
  }
}
