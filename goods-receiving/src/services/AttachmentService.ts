import { SPFI } from '@pnp/sp';
import { IList } from '@pnp/sp/lists';
import { IAppConfig } from '../models/IFieldConfig';
import { IAttachment } from '../models/IAttachment';
import { AppError } from '../models/AppError';
import { makeUniqueFileName, sanitizeFileName, validateFile } from '../logic/files';
import { parseSharePointDate } from '../logic/dates';
import { toAppError } from './errors';

type Json = Record<string, unknown>;

/** Upload chunk size. Files larger than this are uploaded in several requests. */
const CHUNK_SIZE = 4 * 1024 * 1024;

/**
 * Files in the "Receiving Documents" library. Each record gets its own folder
 * (named after the Record ID), and every file is linked to the record through
 * the lookup column and tagged with a document type.
 */
export class AttachmentService {
  private readonly _sp: SPFI;
  private readonly _config: IAppConfig;
  private _libraryRootUrl: string | undefined;

  public constructor(sp: SPFI, config: IAppConfig) {
    this._sp = sp;
    this._config = config;
  }

  private get _library(): IList {
    return this._sp.web.lists.getByTitle(this._config.solution.library.title);
  }

  private get _lookupField(): string {
    return this._config.solution.library.recordLookupField.internalName;
  }

  private get _typeField(): string {
    return this._config.solution.library.documentTypeField.internalName;
  }

  /** Throws a friendly error if the library is missing or not accessible. */
  public async ensureLibraryExists(): Promise<void> {
    await this._getLibraryRootUrl();
  }

  public validate(file: File): string | undefined {
    const { allowedExtensions, maxFileSizeMB } = this._config.solution.library;
    return validateFile(file, allowedExtensions, maxFileSizeMB);
  }

  public async list(recordItemId: number): Promise<IAttachment[]> {
    try {
      const items: Json[] = await this._library.items
        .select(...this._selectFields())
        .expand('Author', 'File')
        .filter(`${this._lookupField}Id eq ${Math.floor(recordItemId)}`)
        .orderBy('Created', true)
        .top(500)();
      return items.map((item) => this._toAttachment(item));
    } catch (error) {
      throw toAppError(error, { action: 'view attachments', listTitle: this._config.solution.library.title });
    }
  }

  /**
   * Uploads one file into the record's folder and tags it. Large files are sent
   * in chunks; `onProgress` receives a value between 0 and 1.
   */
  public async upload(
    recordItemId: number,
    recordId: string,
    file: File,
    documentType: string,
    onProgress?: (fraction: number) => void
  ): Promise<IAttachment> {
    const problem = this.validate(file);
    if (problem) throw new AppError('Validation', problem);

    try {
      const folderUrl = await this._ensureFolder(recordId);
      const folder = this._sp.web.getFolderByServerRelativePath(folderUrl);
      const existing: { Name: string }[] = await folder.files.select('Name')();
      const name = makeUniqueFileName(
        sanitizeFileName(file.name),
        existing.map((f) => f.Name)
      );

      if (onProgress) onProgress(0);
      const info = await folder.files.addChunked(name, file, {
        Overwrite: false,
        chunkSize: CHUNK_SIZE,
        progress: (data) => {
          if (onProgress && file.size > 0) {
            onProgress(data.stage === 'finishing' ? 1 : Math.min(0.99, data.offset / file.size));
          }
        }
      });

      const item = await this._sp.web.getFileByServerRelativePath(info.ServerRelativeUrl).getItem();
      await item.update({
        [`${this._lookupField}Id`]: recordItemId,
        [this._typeField]: documentType
      });
      if (onProgress) onProgress(1);

      const itemId = Number((item as unknown as { Id?: number }).Id || 0);
      const fallback: Json = {
        Id: itemId,
        FileLeafRef: info.Name || name,
        FileRef: info.ServerRelativeUrl,
        File: { Length: file.size },
        [this._typeField]: documentType,
        Created: new Date().toISOString()
      };
      if (!(itemId > 0)) return this._toAttachment(fallback);
      const saved: Json = await this._library.items
        .getById(itemId)
        .select(...this._selectFields())
        .expand('Author', 'File')()
        .catch(() => fallback);
      return this._toAttachment(saved);
    } catch (error) {
      throw toAppError(error, { action: `upload "${file.name}"`, listTitle: this._config.solution.library.title });
    }
  }

  /** Moves the file to the site recycle bin (recoverable for 93 days). */
  public async remove(attachment: IAttachment): Promise<void> {
    try {
      await this._sp.web.getFileByServerRelativePath(attachment.serverRelativeUrl).recycle();
    } catch (error) {
      throw toAppError(error, { action: `delete "${attachment.name}"`, listTitle: this._config.solution.library.title });
    }
  }

  /** Opens the file in SharePoint's viewer (works for PDF, images incl. HEIC, Office files). */
  public getPreviewUrl(attachment: IAttachment): string {
    const origin = this._getOrigin();
    const folder = attachment.serverRelativeUrl.substring(0, attachment.serverRelativeUrl.lastIndexOf('/'));
    const root = this._libraryRootUrl || folder.substring(0, folder.lastIndexOf('/'));
    return `${origin}${encodeURI(root)}/Forms/AllItems.aspx?id=${encodeURIComponent(attachment.serverRelativeUrl)}&parent=${encodeURIComponent(folder)}`;
  }

  public getDownloadUrl(attachment: IAttachment): string {
    return `${this._getOrigin()}${encodeURI(attachment.serverRelativeUrl)}?download=1`;
  }

  /** Direct file URL, used for image thumbnails. */
  public getFileUrl(attachment: IAttachment): string {
    return `${this._getOrigin()}${encodeURI(attachment.serverRelativeUrl)}`;
  }

  private _getOrigin(): string {
    const match = /^https?:\/\/[^/]+/i.exec(this._config.siteUrl);
    return match ? match[0] : '';
  }

  private async _getLibraryRootUrl(): Promise<string> {
    if (this._libraryRootUrl) return this._libraryRootUrl;
    try {
      const root = await this._library.rootFolder.select('ServerRelativeUrl')();
      this._libraryRootUrl = root.ServerRelativeUrl;
      return this._libraryRootUrl;
    } catch (error) {
      throw toAppError(error, { action: 'open the documents library', listTitle: this._config.solution.library.title });
    }
  }

  private async _ensureFolder(recordId: string): Promise<string> {
    const root = await this._getLibraryRootUrl();
    const folderName = sanitizeFileName(recordId);
    const folderUrl = `${root}/${folderName}`;
    try {
      const folder = await this._sp.web.getFolderByServerRelativePath(folderUrl).select('Exists')();
      if (folder.Exists) return folderUrl;
    } catch {
      // Not found: create it below.
    }
    await this._library.rootFolder.folders.addUsingPath(folderName);
    return folderUrl;
  }

  private _selectFields(): string[] {
    return ['Id', 'FileLeafRef', 'FileRef', 'File/Length', this._typeField, 'Created', 'Author/Title'];
  }

  private _toAttachment(item: Json): IAttachment {
    const author = item.Author as Json | undefined;
    const file = item.File as Json | undefined;
    return {
      id: Number(item.Id || 0),
      name: String(item.FileLeafRef || ''),
      serverRelativeUrl: String(item.FileRef || ''),
      documentType: String(item[this._typeField] || ''),
      sizeBytes: Number((file && file.Length) || 0),
      created: parseSharePointDate(item.Created),
      createdBy: author ? String(author.Title || '') : undefined
    };
  }
}
