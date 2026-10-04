import * as React from 'react';
import { IAttachment, IPendingFile } from '../../../../models/IAttachment';
import { IReceivingRecord } from '../../../../models/IReceivingRecord';
import { AttachmentService } from '../../../../services/AttachmentService';

export interface IAttachmentManagerProps {
  record: IReceivingRecord | undefined;
  editable: boolean;
  pendingFiles: IPendingFile[];
  onPendingFilesChange: (files: IPendingFile[]) => void;
  onAttachmentsChange?: (attachments: IAttachment[]) => void;
}

/** Placeholder until stage 5 (attachments). */
export const AttachmentManager: React.FC<IAttachmentManagerProps> = () => <div>Attachments will be available in stage 5.</div>;

export async function uploadPendingFiles(
  service: AttachmentService,
  record: IReceivingRecord,
  files: IPendingFile[],
  onChange: (files: IPendingFile[]) => void
): Promise<string[]> {
  return [];
}
