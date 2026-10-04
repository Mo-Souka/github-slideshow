import * as React from 'react';
import {
  DefaultButton,
  MessageBar,
  MessageBarType,
  Pivot,
  PivotItem,
  PrimaryButton,
  Spinner,
  SpinnerSize
} from '@fluentui/react';
import { getField, getFormFields } from '../../../../config/appConfig';
import { ApprovalStatus, FieldKeys } from '../../../../config/fieldKeys';
import { IFieldConfig } from '../../../../models/IFieldConfig';
import { IReceivingRecord } from '../../../../models/IReceivingRecord';
import { IPendingFile } from '../../../../models/IAttachment';
import { formatFieldValue } from '../../../../logic/format';
import { hasErrors, validateValues } from '../../../../logic/validation';
import { getRecordPermissions, getStatus } from '../../../../logic/workflow';
import { useAppContext } from '../AppContext';
import { ErrorMessage } from '../common/ErrorMessage';
import { hasBadge, StatusBadge } from '../common/StatusBadge';
import { fireAndForget } from '../common/async';
import { AttachmentManager } from '../attachments/AttachmentManager';
import { groupBySection } from '../form/RecordForm';
import { isWideField } from '../form/FieldEditor';
import { ApprovalDialog } from './ApprovalDialog';
import { VersionHistory } from './VersionHistory';
import styles from '../App.module.scss';

export interface IRecordDetailProps {
  recordId: number;
}

const FieldDisplay: React.FC<{ field: IFieldConfig; record: IReceivingRecord }> = ({ field, record }) => {
  const value = record.values[field.key];
  const text = formatFieldValue(field, value);
  return (
    <div className={`${styles.displayField} ${isWideField(field) ? styles.fullWidth : ''}`}>
      <div className={styles.displayLabel}>{field.displayName}</div>
      <div className={styles.displayValue}>
        {hasBadge(field) && text ? (
          <StatusBadge field={field} value={text} />
        ) : text ? (
          text
        ) : (
          <span className={styles.empty}>-</span>
        )}
      </div>
    </div>
  );
};

export const RecordDetail: React.FC<IRecordDetailProps> = ({ recordId }) => {
  const { config, user, services, navigate, notify } = useAppContext();
  const [record, setRecord] = React.useState<IReceivingRecord | undefined>();
  const [loadError, setLoadError] = React.useState<unknown>();
  const [actionError, setActionError] = React.useState<unknown>();
  const [incomplete, setIncomplete] = React.useState<string[]>([]);
  const [decision, setDecision] = React.useState<'approve' | 'reject' | undefined>();
  const [submitting, setSubmitting] = React.useState(false);
  const [reloadCounter, setReloadCounter] = React.useState(0);
  const [pendingFiles, setPendingFiles] = React.useState<IPendingFile[]>([]);

  const detailFields = React.useMemo(() => config.fields.filter((f) => f.showInDetail !== false), [config]);
  const sections = React.useMemo(() => groupBySection(config, detailFields), [config, detailFields]);
  const formFields = React.useMemo(() => getFormFields(config), [config]);

  React.useEffect(() => {
    let cancelled = false;
    setLoadError(undefined);
    services.records
      .getById(recordId)
      .then((loaded) => {
        if (!cancelled) setRecord(loaded);
      })
      .catch((error) => {
        if (!cancelled) setLoadError(error);
      });
    return () => {
      cancelled = true;
    };
  }, [recordId, services, reloadCounter]);

  if (loadError !== undefined) {
    return (
      <>
        <ErrorMessage error={loadError} onRetry={() => setReloadCounter(reloadCounter + 1)} />
        <div className={styles.toolbar} style={{ marginTop: 12 }}>
          <DefaultButton text="Back to list" onClick={() => navigate({ name: 'list' })} />
        </div>
      </>
    );
  }
  if (!record) {
    return <Spinner className={styles.loading} size={SpinnerSize.large} label="Loading record..." />;
  }

  const permissions = getRecordPermissions(user, record);
  const status = getStatus(record);
  const approvalField = getField(config, FieldKeys.approvalStatus);
  const recordIdText = String(record.values[FieldKeys.recordId] || record.id);
  const otherBadgeFields = config.fields.filter((f) => hasBadge(f) && f.key !== FieldKeys.approvalStatus);

  /** Fields that would block submission or approval. */
  const findIncomplete = (): string[] => {
    const errors = validateValues(formFields, record.values, 'submit');
    return hasErrors(errors) ? formFields.filter((f) => errors[f.key]).map((f) => errors[f.key]) : [];
  };

  const submit = async (): Promise<void> => {
    const problems = findIncomplete();
    setIncomplete(problems);
    if (problems.length > 0) return;
    setSubmitting(true);
    setActionError(undefined);
    try {
      const updated = await services.records.update(record, {}, ApprovalStatus.PendingApproval);
      setRecord(updated);
      notify({ type: 'success', text: `Record ${recordIdText} was submitted for approval.` });
    } catch (error) {
      setActionError(error);
    } finally {
      setSubmitting(false);
    }
  };

  const startDecision = (next: 'approve' | 'reject'): void => {
    if (next === 'approve') {
      const problems = findIncomplete();
      setIncomplete(problems);
      if (problems.length > 0) return;
    }
    setDecision(next);
  };

  return (
    <div>
      <div className={styles.header}>
        <div className={styles.titleRow}>
          <h2 className={styles.title}>{recordIdText}</h2>
          <StatusBadge field={approvalField} value={status} />
          {otherBadgeFields.map((f) => {
            const value = record.values[f.key];
            return <StatusBadge key={f.key} field={f} value={typeof value === 'string' ? value : undefined} />;
          })}
        </div>
        <div className={styles.toolbar}>
          {permissions.canApprove && (
            <PrimaryButton text="Approve" iconProps={{ iconName: 'CheckMark' }} onClick={() => startDecision('approve')} />
          )}
          {permissions.canReject && (
            <DefaultButton text="Reject" iconProps={{ iconName: 'Cancel' }} onClick={() => startDecision('reject')} />
          )}
          {permissions.canSubmit && (
            <PrimaryButton
              text={submitting ? 'Submitting...' : status === ApprovalStatus.Rejected ? 'Resubmit for Approval' : 'Submit for Approval'}
              iconProps={{ iconName: 'Send' }}
              disabled={submitting}
              onClick={() => fireAndForget(submit())}
            />
          )}
          {permissions.canEdit && (
            <DefaultButton text="Edit" iconProps={{ iconName: 'Edit' }} onClick={() => navigate({ name: 'edit', id: record.id })} />
          )}
          <DefaultButton text="Back" iconProps={{ iconName: 'Back' }} onClick={() => navigate({ name: user.isSupervisor && status === ApprovalStatus.PendingApproval ? 'pending' : 'list' })} />
        </div>
      </div>

      {permissions.approvalBlockedReason && (
        <MessageBar messageBarType={MessageBarType.info} className={styles.notification}>
          {permissions.approvalBlockedReason}
        </MessageBar>
      )}
      {!permissions.canEdit && permissions.readOnlyReason && (
        <MessageBar messageBarType={MessageBarType.info} className={styles.notification}>
          {permissions.readOnlyReason}
        </MessageBar>
      )}
      {incomplete.length > 0 && (
        <MessageBar
          messageBarType={MessageBarType.warning}
          isMultiline={true}
          className={styles.notification}
          onDismiss={() => setIncomplete([])}
          actions={
            permissions.canEdit ? (
              <DefaultButton text="Edit record" onClick={() => navigate({ name: 'edit', id: record.id })} />
            ) : undefined
          }
        >
          This record is incomplete:
          <ul style={{ margin: '4px 0 0', paddingLeft: 20 }}>
            {incomplete.map((m) => (
              <li key={m}>{m}</li>
            ))}
          </ul>
        </MessageBar>
      )}
      {actionError !== undefined && (
        <div className={styles.notification}>
          <ErrorMessage
            error={actionError}
            onDismiss={() => setActionError(undefined)}
            onRetry={() => {
              setActionError(undefined);
              setReloadCounter(reloadCounter + 1);
            }}
          />
        </div>
      )}

      <Pivot aria-label="Record sections" overflowBehavior="menu">
        <PivotItem headerText="Details" itemIcon="Info">
          <div style={{ paddingTop: 12 }}>
            {sections.map((section) => (
              <section key={section.key} className={styles.section} aria-label={section.title}>
                <h3 className={styles.sectionTitle}>{section.title}</h3>
                <div className={styles.grid}>
                  {section.fields.map((field) => (
                    <FieldDisplay key={field.key} field={field} record={record} />
                  ))}
                </div>
              </section>
            ))}
          </div>
        </PivotItem>
        <PivotItem headerText="Attachments" itemIcon="Attach">
          <div style={{ paddingTop: 12 }}>
            <AttachmentManager
              record={record}
              editable={permissions.canManageAttachments}
              pendingFiles={pendingFiles}
              onPendingFilesChange={setPendingFiles}
            />
          </div>
        </PivotItem>
        <PivotItem headerText="History" itemIcon="History">
          <div style={{ paddingTop: 12 }}>
            <VersionHistory recordId={record.id} refreshKey={record.etag || String(reloadCounter)} />
          </div>
        </PivotItem>
      </Pivot>

      <ApprovalDialog
        decision={decision}
        record={record}
        onDismiss={() => setDecision(undefined)}
        onDone={(updated) => {
          setDecision(undefined);
          setRecord(updated);
          const verb = getStatus(updated) === ApprovalStatus.Approved ? 'approved' : 'rejected';
          notify({ type: 'success', text: `Record ${recordIdText} was ${verb}.` });
        }}
      />
    </div>
  );
};
