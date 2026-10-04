import * as React from 'react';
import { DefaultButton, Dialog, DialogFooter, DialogType, PrimaryButton, TextField } from '@fluentui/react';
import { getField } from '../../../../config/appConfig';
import { ApprovalStatus, FieldKeys } from '../../../../config/fieldKeys';
import { IReceivingRecord } from '../../../../models/IReceivingRecord';
import { validateField } from '../../../../logic/validation';
import { useAppContext } from '../AppContext';
import { ErrorMessage } from '../common/ErrorMessage';
import { fireAndForget } from '../common/async';

export interface IApprovalDialogProps {
  decision: 'approve' | 'reject' | undefined;
  record: IReceivingRecord;
  onDismiss: () => void;
  onDone: (updated: IReceivingRecord) => void;
}

/** Approve or reject a pending record. A comment is required when rejecting (see fields.json). */
export const ApprovalDialog: React.FC<IApprovalDialogProps> = ({ decision, record, onDismiss, onDone }) => {
  const { config, user, services } = useAppContext();
  const commentsField = getField(config, FieldKeys.supervisorComments);
  const [comments, setComments] = React.useState('');
  const [error, setError] = React.useState<string | undefined>();
  const [saveError, setSaveError] = React.useState<unknown>();
  const [saving, setSaving] = React.useState(false);

  React.useEffect(() => {
    if (decision) {
      setComments('');
      setError(undefined);
      setSaveError(undefined);
    }
  }, [decision]);

  const status = decision === 'reject' ? ApprovalStatus.Rejected : ApprovalStatus.Approved;

  const confirm = async (): Promise<void> => {
    const message = validateField(
      commentsField,
      { ...record.values, [FieldKeys.approvalStatus]: status, [FieldKeys.supervisorComments]: comments },
      'submit'
    );
    setError(message);
    if (message) return;
    setSaving(true);
    setSaveError(undefined);
    try {
      const updated = await services.records.decide(record, status, comments, user.id);
      onDone(updated);
    } catch (e) {
      setSaveError(e);
    } finally {
      setSaving(false);
    }
  };

  const recordId = String(record.values[FieldKeys.recordId] || record.id);
  return (
    <Dialog
      hidden={!decision}
      onDismiss={saving ? undefined : onDismiss}
      minWidth={320}
      maxWidth={560}
      dialogContentProps={{
        type: DialogType.normal,
        title: decision === 'reject' ? `Reject ${recordId}?` : `Approve ${recordId}?`,
        subText:
          decision === 'reject'
            ? 'The record goes back to the receiver, who can correct it and submit it again.'
            : 'Approved records can no longer be changed.'
      }}
      modalProps={{ isBlocking: true }}
    >
      <TextField
        label={commentsField.displayName}
        multiline={true}
        rows={4}
        required={decision === 'reject'}
        value={comments}
        onChange={(e, text) => {
          setComments(text || '');
          if (error) setError(undefined);
        }}
        errorMessage={error}
        disabled={saving}
        placeholder={decision === 'reject' ? 'What needs to be corrected?' : 'Optional'}
      />
      {saveError !== undefined && (
        <div style={{ marginTop: 12 }}>
          <ErrorMessage error={saveError} />
        </div>
      )}
      <DialogFooter>
        <PrimaryButton
          text={saving ? 'Saving...' : decision === 'reject' ? 'Reject' : 'Approve'}
          disabled={saving}
          onClick={() => fireAndForget(confirm())}
        />
        <DefaultButton text="Cancel" disabled={saving} onClick={onDismiss} />
      </DialogFooter>
    </Dialog>
  );
};
