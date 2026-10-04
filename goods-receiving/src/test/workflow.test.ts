import { ApprovalStatus } from '../config/fieldKeys';
import { getNextStatus, getRecordPermissions } from '../logic/workflow';
import { record, user } from './testConfig';

const receiver = user({ id: 10 });
const otherReceiver = user({ id: 11, displayName: 'Other' });
const supervisor = user({ id: 20, isSupervisor: true });
const viewer = user({ id: 30, isReceiver: false, isSupervisor: false });

function recordBy(authorId: number, status: string, receiverId: number = authorId): ReturnType<typeof record> {
  return record({
    approvalStatus: status,
    createdBy: { id: authorId, title: 'Author' },
    receiverName: { id: receiverId, title: 'Receiver' }
  });
}

describe('workflow permissions', () => {
  it('lets receivers create records but not viewers', () => {
    expect(getRecordPermissions(receiver, undefined).canEdit).toBe(true);
    expect(getRecordPermissions(viewer, undefined).canEdit).toBe(false);
  });

  it('lets receivers edit and submit their own Draft and Rejected records', () => {
    [ApprovalStatus.Draft, ApprovalStatus.Rejected].forEach((status) => {
      const p = getRecordPermissions(receiver, recordBy(10, status));
      expect(p.canEdit).toBe(true);
      expect(p.canSubmit).toBe(true);
      expect(p.canApprove).toBe(false);
    });
  });

  it('counts a record as own when the user is the named receiver', () => {
    expect(getRecordPermissions(receiver, recordBy(99, ApprovalStatus.Draft, 10)).canEdit).toBe(true);
  });

  it("does not let receivers edit other people's records", () => {
    const p = getRecordPermissions(otherReceiver, recordBy(10, ApprovalStatus.Draft));
    expect(p.canEdit).toBe(false);
    expect(p.readOnlyReason).toContain('only edit records you created');
  });

  it('locks pending records for receivers', () => {
    const p = getRecordPermissions(receiver, recordBy(10, ApprovalStatus.PendingApproval));
    expect(p.canEdit).toBe(false);
    expect(p.canManageAttachments).toBe(false);
  });

  it('makes approved records read-only for everyone', () => {
    [receiver, supervisor].forEach((u) => {
      const p = getRecordPermissions(u, recordBy(10, ApprovalStatus.Approved));
      expect(p.canEdit).toBe(false);
      expect(p.canApprove).toBe(false);
      expect(p.readOnlyReason).toContain('approved');
    });
  });

  it("lets supervisors edit and decide on other people's pending records", () => {
    const p = getRecordPermissions(supervisor, recordBy(10, ApprovalStatus.PendingApproval));
    expect(p.canEdit).toBe(true);
    expect(p.canSubmit).toBe(false);
    expect(p.canApprove).toBe(true);
    expect(p.canReject).toBe(true);
  });

  it('blocks supervisors from approving their own records', () => {
    const p = getRecordPermissions(supervisor, recordBy(20, ApprovalStatus.PendingApproval));
    expect(p.canApprove).toBe(false);
    expect(p.approvalBlockedReason).toContain('another supervisor');
    const asReceiver = getRecordPermissions(supervisor, recordBy(10, ApprovalStatus.PendingApproval, 20));
    expect(asReceiver.canApprove).toBe(false);
  });

  it('only offers approval for pending records', () => {
    expect(getRecordPermissions(supervisor, recordBy(10, ApprovalStatus.Draft)).canApprove).toBe(false);
  });

  it('gives viewers read-only access', () => {
    const p = getRecordPermissions(viewer, recordBy(10, ApprovalStatus.Draft));
    expect(p.canEdit).toBe(false);
    expect(p.readOnlyReason).toContain('read-only');
  });
});

describe('status transitions', () => {
  it('follows Draft -> Pending -> Approved/Rejected', () => {
    expect(getNextStatus('saveDraft', undefined)).toBe(ApprovalStatus.Draft);
    expect(getNextStatus('saveDraft', ApprovalStatus.Rejected)).toBe(ApprovalStatus.Rejected);
    expect(getNextStatus('submit', ApprovalStatus.Draft)).toBe(ApprovalStatus.PendingApproval);
    expect(getNextStatus('submit', ApprovalStatus.Rejected)).toBe(ApprovalStatus.PendingApproval);
    expect(getNextStatus('approve', ApprovalStatus.PendingApproval)).toBe(ApprovalStatus.Approved);
    expect(getNextStatus('reject', ApprovalStatus.PendingApproval)).toBe(ApprovalStatus.Rejected);
  });
});
