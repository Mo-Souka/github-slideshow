export interface IAttachment {
  /** List item ID of the file in the document library. */
  id: number;
  name: string;
  /** Server-relative URL of the file. */
  serverRelativeUrl: string;
  documentType: string;
  sizeBytes: number;
  created?: Date;
  createdBy?: string;
}

/** A file chosen in the form that is uploaded after the record is saved. */
export interface IPendingFile {
  /** Client-side identifier, only used for React keys and removal. */
  localId: string;
  file: File;
  documentType: string;
}
