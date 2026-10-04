import * as React from 'react';
import {
  ActionButton,
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  Icon,
  IconButton,
  Link,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  ProgressIndicator,
  Spinner,
  SpinnerSize
} from '@fluentui/react';
import { FieldKeys } from '../../../../config/fieldKeys';
import { IAttachment, IPendingFile } from '../../../../models/IAttachment';
import { IReceivingRecord } from '../../../../models/IReceivingRecord';
import { AttachmentService } from '../../../../services/AttachmentService';
import { toAppError } from '../../../../services/errors';
import { formatDateTime, formatFileSize } from '../../../../logic/format';
import { getFileIconName, isBrowserImage } from '../../../../logic/files';
import { useAppContext } from '../AppContext';
import { ErrorMessage } from '../common/ErrorMessage';
import { fireAndForget } from '../common/async';
import styles from '../App.module.scss';

export interface IAttachmentManagerProps {
  /** Undefined while creating a new record: files are queued and uploaded after the first save. */
  record: IReceivingRecord | undefined;
  /** Whether files can be added or deleted (record is editable by the current user). */
  editable: boolean;
  pendingFiles: IPendingFile[];
  onPendingFilesChange: (files: IPendingFile[]) => void;
}

interface IUploadState {
  localId: string;
  name: string;
  documentType: string;
  progress: number;
}

let localIdCounter = 0;
function newLocalId(): string {
  localIdCounter += 1;
  return `pending-${new Date().getTime()}-${localIdCounter}`;
}

function recordCode(record: IReceivingRecord): string {
  return String(record.values[FieldKeys.recordId] || record.id);
}

/**
 * Uploads files that were added before the record existed. Returns error
 * messages for files that failed; successful files are removed from the queue.
 */
export async function uploadPendingFiles(
  service: AttachmentService,
  record: IReceivingRecord,
  files: IPendingFile[],
  onChange: (files: IPendingFile[]) => void
): Promise<string[]> {
  const errors: string[] = [];
  let remaining = files.slice();
  for (const pending of files) {
    try {
      await service.upload(record.id, recordCode(record), pending.file, pending.documentType);
      remaining = remaining.filter((f) => f.localId !== pending.localId);
      onChange(remaining);
    } catch (error) {
      errors.push(toAppError(error).userMessage);
    }
  }
  return errors;
}

/** Image preview with a file-type icon as fallback (e.g. when the preview cannot be generated). */
const Thumbnail: React.FC<{ url?: string; iconName: string }> = ({ url, iconName }) => {
  const [failed, setFailed] = React.useState(false);
  if (url && !failed) {
    return <img className={styles.fileThumb} src={url} alt="" loading="lazy" onError={() => setFailed(true)} />;
  }
  return (
    <span className={styles.fileThumb}>
      <Icon iconName={iconName} />
    </span>
  );
};

export const AttachmentManager: React.FC<IAttachmentManagerProps> = ({ record, editable, pendingFiles, onPendingFilesChange }) => {
  const { config, services } = useAppContext();
  const library = config.solution.library;
  const [attachments, setAttachments] = React.useState<IAttachment[]>([]);
  const [loading, setLoading] = React.useState<boolean>(!!record);
  const [loadError, setLoadError] = React.useState<unknown>();
  const [uploads, setUploads] = React.useState<IUploadState[]>([]);
  const [messages, setMessages] = React.useState<string[]>([]);
  const [actionError, setActionError] = React.useState<unknown>();
  const [toDelete, setToDelete] = React.useState<IAttachment | undefined>();
  const [deleting, setDeleting] = React.useState(false);
  const [dragOver, setDragOver] = React.useState<string | undefined>();
  const [reloadCounter, setReloadCounter] = React.useState(0);
  const fileInputs = React.useRef<Record<string, HTMLInputElement | null>>({});
  const cameraInput = React.useRef<HTMLInputElement | null>(null);
  const recordItemId = record ? record.id : undefined;

  React.useEffect(() => {
    if (!recordItemId) {
      setLoading(false);
      return undefined;
    }
    let cancelled = false;
    setLoading(true);
    setLoadError(undefined);
    services.attachments
      .list(recordItemId)
      .then((result) => {
        if (cancelled) return;
        setAttachments(result);
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
  }, [recordItemId, services, reloadCounter]);

  const uploadNow = async (files: File[], documentType: string): Promise<void> => {
    if (!record) return;
    for (const file of files) {
      const localId = newLocalId();
      setUploads((current) => current.concat([{ localId, name: file.name, documentType, progress: 0 }]));
      try {
        const saved = await services.attachments.upload(record.id, recordCode(record), file, documentType, (fraction) =>
          setUploads((current) => current.map((u) => (u.localId === localId ? { ...u, progress: fraction } : u)))
        );
        setAttachments((current) => current.concat([saved]));
      } catch (error) {
        setMessages((current) => current.concat([toAppError(error).userMessage]));
      } finally {
        setUploads((current) => current.filter((u) => u.localId !== localId));
      }
    }
  };

  const addFiles = (fileList: FileList | File[] | null, documentType: string): void => {
    if (!fileList || !editable) return;
    const files = Array.prototype.slice.call(fileList) as File[];
    const problems: string[] = [];
    const accepted: File[] = [];
    files.forEach((file) => {
      const problem = services.attachments.validate(file);
      if (problem) problems.push(problem);
      else accepted.push(file);
    });
    setMessages(problems);
    if (accepted.length === 0) return;

    if (record) {
      fireAndForget(uploadNow(accepted, documentType));
    } else {
      onPendingFilesChange(
        pendingFiles.concat(accepted.map((file) => ({ localId: newLocalId(), file, documentType })))
      );
    }
  };

  const confirmDelete = async (): Promise<void> => {
    if (!toDelete) return;
    setDeleting(true);
    setActionError(undefined);
    try {
      await services.attachments.remove(toDelete);
      setAttachments((current) => current.filter((a) => a.serverRelativeUrl !== toDelete.serverRelativeUrl));
      setToDelete(undefined);
    } catch (error) {
      setActionError(error);
      setToDelete(undefined);
    } finally {
      setDeleting(false);
    }
  };

  const onDrop = (e: React.DragEvent<HTMLDivElement>, documentType: string): void => {
    e.preventDefault();
    e.stopPropagation();
    setDragOver(undefined);
    addFiles(e.dataTransfer.files, documentType);
  };

  const accept = library.allowedExtensions.join(',');

  if (loading) {
    return <Spinner size={SpinnerSize.medium} label="Loading attachments..." />;
  }

  return (
    <div>
      {loadError !== undefined && <ErrorMessage error={loadError} onRetry={() => setReloadCounter(reloadCounter + 1)} />}
      {actionError !== undefined && <ErrorMessage error={actionError} onDismiss={() => setActionError(undefined)} />}
      {messages.length > 0 && (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline={true} onDismiss={() => setMessages([])} className={styles.notification}>
          {messages.map((m, i) => (
            <div key={i}>{m}</div>
          ))}
        </MessageBar>
      )}
      {editable && (
        <div className={styles.help} style={{ marginBottom: 8 }}>
          {record
            ? 'Files are uploaded as soon as you add them.'
            : 'Files are uploaded when you save the record.'}{' '}
          Allowed: {library.allowedExtensions.join(', ')} (max. {library.maxFileSizeMB} MB each).
        </div>
      )}

      <div className={styles.attachmentGroups}>
        {library.documentTypes.map((documentType) => {
          const saved = attachments.filter((a) => a.documentType === documentType);
          const queued = pendingFiles.filter((p) => p.documentType === documentType);
          const uploading = uploads.filter((u) => u.documentType === documentType);
          const count = saved.length + queued.length;
          const isPhoto = documentType === library.photoDocumentType;

          return (
            <div
              key={documentType}
              className={styles.attachmentGroup}
              onDragOver={(e) => {
                if (!editable) return;
                e.preventDefault();
                setDragOver(documentType);
              }}
              onDragLeave={() => setDragOver(undefined)}
              onDrop={(e) => onDrop(e, documentType)}
            >
              <div className={styles.attachmentGroupHeader}>
                <span className={styles.attachmentGroupTitle}>
                  {documentType} {count > 0 ? `(${count})` : ''}
                </span>
                {editable && (
                  <span>
                    {isPhoto && (
                      <ActionButton
                        iconProps={{ iconName: 'Camera' }}
                        text="Take photo"
                        onClick={() => cameraInput.current && cameraInput.current.click()}
                      />
                    )}
                    <ActionButton
                      iconProps={{ iconName: 'Attach' }}
                      text="Add files"
                      onClick={() => {
                        const input = fileInputs.current[documentType];
                        if (input) input.click();
                      }}
                    />
                  </span>
                )}
              </div>

              {saved.map((attachment) => (
                <div key={attachment.serverRelativeUrl} className={styles.fileRow}>
                  <Thumbnail
                    url={isBrowserImage(attachment.name) ? services.attachments.getThumbnailUrl(attachment) : undefined}
                    iconName={getFileIconName(attachment.name)}
                  />
                  <div className={styles.fileInfo}>
                    <Link
                      className={styles.fileName}
                      href={services.attachments.getPreviewUrl(attachment)}
                      target="_blank"
                      rel="noopener noreferrer"
                      title={`Open ${attachment.name}`}
                    >
                      {attachment.name}
                    </Link>
                    <div className={styles.fileMeta}>
                      {formatFileSize(attachment.sizeBytes)}
                      {attachment.created ? ` · ${formatDateTime(attachment.created)}` : ''}
                      {attachment.createdBy ? ` · ${attachment.createdBy}` : ''}
                    </div>
                  </div>
                  <IconButton
                    iconProps={{ iconName: 'Download' }}
                    title="Download"
                    ariaLabel={`Download ${attachment.name}`}
                    href={services.attachments.getDownloadUrl(attachment)}
                  />
                  {editable && (
                    <IconButton
                      iconProps={{ iconName: 'Delete' }}
                      title="Delete"
                      ariaLabel={`Delete ${attachment.name}`}
                      onClick={() => setToDelete(attachment)}
                    />
                  )}
                </div>
              ))}

              {queued.map((pending) => (
                <div key={pending.localId} className={styles.fileRow}>
                  <span className={styles.fileThumb}>
                    <Icon iconName={getFileIconName(pending.file.name)} />
                  </span>
                  <div className={styles.fileInfo}>
                    <span className={styles.fileName}>{pending.file.name}</span>
                    <div className={styles.fileMeta}>{formatFileSize(pending.file.size)} · not uploaded yet</div>
                  </div>
                  {editable && (
                    <IconButton
                      iconProps={{ iconName: 'Cancel' }}
                      title="Remove"
                      ariaLabel={`Remove ${pending.file.name}`}
                      onClick={() => onPendingFilesChange(pendingFiles.filter((p) => p.localId !== pending.localId))}
                    />
                  )}
                </div>
              ))}

              {uploading.map((upload) => (
                <ProgressIndicator
                  key={upload.localId}
                  label={`Uploading ${upload.name}`}
                  percentComplete={upload.progress}
                  description={`${Math.round(upload.progress * 100)}%`}
                />
              ))}

              {count === 0 && uploading.length === 0 && !editable && <div className={styles.empty}>No files</div>}

              {editable && (
                <div className={`${styles.dropZone} ${dragOver === documentType ? styles.dropZoneActive : ''}`}>
                  Drag files here or use &quot;Add files&quot;
                </div>
              )}

              <input
                ref={(element) => {
                  fileInputs.current[documentType] = element;
                }}
                className={styles.hiddenInput}
                type="file"
                multiple={true}
                accept={accept}
                onChange={(e) => {
                  addFiles(e.target.files, documentType);
                  e.target.value = '';
                }}
              />
            </div>
          );
        })}
      </div>

      {/* Opens the camera directly on phones and tablets. */}
      <input
        ref={cameraInput}
        className={styles.hiddenInput}
        type="file"
        accept="image/*"
        capture="environment"
        onChange={(e) => {
          addFiles(e.target.files, library.photoDocumentType);
          e.target.value = '';
        }}
      />

      <Dialog
        hidden={!toDelete}
        onDismiss={() => setToDelete(undefined)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Delete file?',
          subText: toDelete ? `"${toDelete.name}" will be moved to the site recycle bin.` : ''
        }}
      >
        <DialogFooter>
          <PrimaryButton text={deleting ? 'Deleting...' : 'Delete'} disabled={deleting} onClick={() => fireAndForget(confirmDelete())} />
          <DefaultButton text="Cancel" disabled={deleting} onClick={() => setToDelete(undefined)} />
        </DialogFooter>
      </Dialog>
    </div>
  );
};
