import * as React from 'react';
import { DefaultButton, MessageBar, MessageBarType, PrimaryButton, Spinner, SpinnerSize } from '@fluentui/react';
import { getField, getFormFields } from '../../../../config/appConfig';
import { ApprovalStatus, FieldKeys } from '../../../../config/fieldKeys';
import { IAppConfig, IFieldConfig } from '../../../../models/IFieldConfig';
import { FieldValue, FieldValues, IReceivingRecord, isReferenceValue } from '../../../../models/IReceivingRecord';
import { IUserContext } from '../../../../models/IUserContext';
import { IPendingFile } from '../../../../models/IAttachment';
import { AppError } from '../../../../models/AppError';
import { today } from '../../../../logic/dates';
import { formatDateTime } from '../../../../logic/format';
import { hasErrors, isFieldRequired, validateField, validateValues, ValidationErrors, ValidationMode } from '../../../../logic/validation';
import { getNextStatus, getRecordPermissions, getStatus, WorkflowAction } from '../../../../logic/workflow';
import { useAppContext } from '../AppContext';
import { ErrorMessage } from '../common/ErrorMessage';
import { fireAndForget } from '../common/async';
import { StatusBadge } from '../common/StatusBadge';
import { AttachmentManager, uploadPendingFiles } from '../attachments/AttachmentManager';
import { FieldEditor, isWideField } from './FieldEditor';
import styles from '../App.module.scss';

export interface IRecordFormProps {
  /** Undefined for a new record. */
  recordId?: number;
}

/** Initial values for a new record, from "defaultValue" in fields.json. */
export function getDefaultValues(fields: IFieldConfig[], user: IUserContext): FieldValues {
  const values: FieldValues = {};
  fields.forEach((field) => {
    const defaultValue = field.defaultValue;
    if (defaultValue === undefined) return;
    if (defaultValue === '[today]') {
      values[field.key] = today();
    } else if (defaultValue === '[me]') {
      values[field.key] = { id: user.id, title: user.displayName, email: user.email, loginName: user.loginName };
    } else {
      values[field.key] = defaultValue;
    }
  });
  return values;
}

function pickValues(record: IReceivingRecord, fields: IFieldConfig[]): FieldValues {
  const values: FieldValues = {};
  fields.forEach((f) => {
    values[f.key] = record.values[f.key];
  });
  return values;
}

/** Groups form fields by the sections defined in fields.json. */
export function groupBySection(config: IAppConfig, fields: IFieldConfig[]): { key: string; title: string; fields: IFieldConfig[] }[] {
  const groups = config.sections
    .map((s) => ({ key: s.key, title: s.title, fields: fields.filter((f) => f.section === s.key) }))
    .filter((g) => g.fields.length > 0);
  const known = config.sections.map((s) => s.key);
  const other = fields.filter((f) => !f.section || known.indexOf(f.section) < 0);
  if (other.length > 0) groups.push({ key: 'other', title: 'Other', fields: other });
  return groups;
}

export const RecordForm: React.FC<IRecordFormProps> = ({ recordId }) => {
  const { config, user, services, navigate, notify } = useAppContext();
  const formFields = React.useMemo(() => getFormFields(config), [config]);
  const sections = React.useMemo(() => groupBySection(config, formFields), [config, formFields]);

  const [record, setRecord] = React.useState<IReceivingRecord | undefined>();
  const [values, setValues] = React.useState<FieldValues>(() => (recordId ? {} : getDefaultValues(formFields, user)));
  const [errors, setErrors] = React.useState<ValidationErrors>({});
  const [validationMode, setValidationMode] = React.useState<ValidationMode | undefined>();
  const [loading, setLoading] = React.useState<boolean>(!!recordId);
  const [loadError, setLoadError] = React.useState<unknown>();
  const [saveError, setSaveError] = React.useState<unknown>();
  const [saving, setSaving] = React.useState<WorkflowAction | undefined>();
  const [dirty, setDirty] = React.useState(false);
  const [pendingFiles, setPendingFiles] = React.useState<IPendingFile[]>([]);
  const [reloadCounter, setReloadCounter] = React.useState(0);

  React.useEffect(() => {
    if (!recordId) return undefined;
    let cancelled = false;
    setLoading(true);
    setLoadError(undefined);
    services.records
      .getById(recordId)
      .then((loaded) => {
        if (cancelled) return;
        setRecord(loaded);
        setValues(pickValues(loaded, formFields));
        setLoading(false);
      })
      .catch((error) => {
        if (cancelled) return;
        setLoadError(error);
        setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [recordId, services, formFields, reloadCounter]);

  // Warn before leaving the page with unsaved changes.
  React.useEffect(() => {
    if (!dirty) return undefined;
    const handler = (e: BeforeUnloadEvent): void => {
      e.preventDefault();
      e.returnValue = '';
    };
    window.addEventListener('beforeunload', handler);
    return () => window.removeEventListener('beforeunload', handler);
  }, [dirty]);

  // Latest values, so a change handler never works on stale state.
  const valuesRef = React.useRef<FieldValues>(values);
  valuesRef.current = values;

  const onFieldChange = React.useCallback(
    (key: string, value: FieldValue) => {
      const next = { ...valuesRef.current, [key]: value };
      valuesRef.current = next;
      setValues(next);
      setDirty(true);
      if (validationMode) {
        // After a failed save, re-check the changed field and the fields that depend on it.
        setErrors((previousErrors) => {
          const updated = { ...previousErrors };
          formFields
            .filter((f) => f.key === key || (f.requiredWhen && f.requiredWhen.field === key))
            .forEach((f) => {
              const message = validateField(f, next, validationMode);
              if (message) updated[f.key] = message;
              else delete updated[f.key];
            });
          return updated;
        });
      }
    },
    [formFields, validationMode]
  );

  const permissions = getRecordPermissions(user, record);
  const status = getStatus(record);
  const isNew = !recordId;

  const goBack = (): void => {
    if (dirty && !window.confirm('Discard your unsaved changes?')) return;
    setDirty(false);
    navigate(recordId ? { name: 'view', id: recordId } : { name: 'list' });
  };

  const save = async (action: WorkflowAction): Promise<void> => {
    const mode: ValidationMode = action === 'submit' ? 'submit' : 'draft';
    const validationErrors = validateValues(formFields, values, mode);
    setErrors(validationErrors);
    setValidationMode(mode);
    if (hasErrors(validationErrors)) {
      const count = Object.keys(validationErrors).length;
      setSaveError(new AppError('Validation', `Please correct the ${count === 1 ? 'field' : `${count} fields`} marked in red.`));
      const first = formFields.filter((f) => validationErrors[f.key])[0];
      const element = first ? document.getElementById(`gr-field-${first.key}`) : undefined;
      if (element) element.scrollIntoView({ behavior: 'smooth', block: 'center' });
      return;
    }

    setSaving(action);
    setSaveError(undefined);
    try {
      const nextStatus = getNextStatus(action, record ? status : undefined);
      const saved = record
        ? await services.records.update(record, values, action === 'submit' ? nextStatus : undefined)
        : await services.records.create(values, nextStatus);
      setRecord(saved);
      setDirty(false);

      const recordIdText = String(saved.values[FieldKeys.recordId] || '');
      const uploadErrors = await uploadPendingFiles(services.attachments, saved, pendingFiles, setPendingFiles);
      const verb = action === 'submit' ? 'submitted for approval' : 'saved';
      if (uploadErrors.length > 0) {
        notify(
          {
            type: 'warning',
            text: `Record ${recordIdText} was ${verb}, but ${uploadErrors.length} file(s) could not be uploaded: ${uploadErrors.join(' ')} Open the record and add them again.`
          },
          { name: 'view', id: saved.id }
        );
      } else {
        notify({ type: 'success', text: `Record ${recordIdText} was ${verb}.` }, { name: 'view', id: saved.id });
      }
    } catch (error) {
      setSaveError(error);
    } finally {
      setSaving(undefined);
    }
  };

  if (loading) {
    return <Spinner className={styles.loading} size={SpinnerSize.large} label="Loading record..." />;
  }
  if (loadError) {
    return (
      <>
        <ErrorMessage error={loadError} onRetry={() => setReloadCounter(reloadCounter + 1)} />
        <div className={styles.toolbar} style={{ marginTop: 12 }}>
          <DefaultButton text="Back to list" onClick={() => navigate({ name: 'list' })} />
        </div>
      </>
    );
  }
  if (!permissions.canEdit) {
    return (
      <>
        <MessageBar messageBarType={MessageBarType.info}>{permissions.readOnlyReason || 'This record cannot be edited.'}</MessageBar>
        <div className={styles.toolbar} style={{ marginTop: 12 }}>
          <DefaultButton
            text={record ? 'Back to record' : 'Back to list'}
            onClick={() => navigate(record ? { name: 'view', id: record.id } : { name: 'list' })}
          />
        </div>
      </>
    );
  }

  const approvalField = getField(config, FieldKeys.approvalStatus);
  const recordIdText = record ? String(record.values[FieldKeys.recordId] || '') : '';
  const supervisorComments = record ? record.values[FieldKeys.supervisorComments] : undefined;
  const reviewedBy = record ? record.values[FieldKeys.reviewedBy] : undefined;
  const reviewedOn = record ? record.values[FieldKeys.reviewedOn] : undefined;
  const busy = !!saving;
  const isDraftLike = isNew || status === ApprovalStatus.Draft;

  return (
    <div>
      <div className={styles.header}>
        <div className={styles.titleRow}>
          <h2 className={styles.title}>{isNew ? 'New receiving record' : `Edit ${recordIdText}`}</h2>
          {!isNew && <StatusBadge field={approvalField} value={status} />}
        </div>
      </div>

      {status === ApprovalStatus.Rejected && (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline={true} className={styles.notification}>
          <strong>Rejected</strong>
          {isReferenceValue(reviewedBy) ? ` by ${reviewedBy.title}` : ''}
          {reviewedOn instanceof Date ? ` on ${formatDateTime(reviewedOn)}` : ''}
          {typeof supervisorComments === 'string' && supervisorComments ? `: ${supervisorComments}` : '.'} Correct the record and
          resubmit it for approval.
        </MessageBar>
      )}

      {sections.map((section) => (
        <section key={section.key} className={styles.section} aria-labelledby={`gr-section-${section.key}`}>
          <h3 id={`gr-section-${section.key}`} className={styles.sectionTitle}>
            {section.title}
          </h3>
          <div className={styles.grid}>
            {section.fields.map((field) => (
              <div key={field.key} id={`gr-field-${field.key}`} className={isWideField(field) ? styles.fullWidth : undefined}>
                <FieldEditor
                  field={field}
                  value={values[field.key]}
                  onChange={onFieldChange}
                  required={isFieldRequired(field, values, 'submit')}
                  errorMessage={errors[field.key]}
                  disabled={busy}
                />
              </div>
            ))}
          </div>
        </section>
      ))}

      <section className={styles.section} aria-labelledby="gr-section-attachments">
        <h3 id="gr-section-attachments" className={styles.sectionTitle}>
          Attachments
        </h3>
        <AttachmentManager
          record={record}
          editable={!busy}
          pendingFiles={pendingFiles}
          onPendingFilesChange={setPendingFiles}
        />
      </section>

      <div className={styles.actions}>
        {saveError !== undefined && (
          <div className={styles.fullRow}>
            <ErrorMessage error={saveError} onDismiss={() => setSaveError(undefined)} />
          </div>
        )}
        {permissions.canSubmit && (
          <PrimaryButton
            text={status === ApprovalStatus.Rejected ? 'Resubmit for Approval' : 'Submit for Approval'}
            iconProps={{ iconName: 'Send' }}
            disabled={busy}
            onClick={() => fireAndForget(save('submit'))}
          />
        )}
        <DefaultButton
          text={saving === 'saveDraft' ? 'Saving...' : isDraftLike ? 'Save as Draft' : 'Save changes'}
          iconProps={{ iconName: 'Save' }}
          disabled={busy}
          onClick={() => fireAndForget(save('saveDraft'))}
        />
        <DefaultButton text="Cancel" disabled={busy} onClick={goBack} />
        {saving === 'submit' && <Spinner size={SpinnerSize.small} label="Submitting..." labelPosition="right" />}
      </div>
    </div>
  );
};
