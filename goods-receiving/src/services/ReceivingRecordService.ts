import { SPFI } from '@pnp/sp';
import { IList } from '@pnp/sp/lists';
import { RenderListDataOptions } from '@pnp/sp/lists';
import { ApprovalStatusValue, FieldKeys } from '../config/fieldKeys';
import { IAppConfig, IFieldConfig } from '../models/IFieldConfig';
import { FieldValues, IReceivingRecord, IRecordPage, IVersionEntry } from '../models/IReceivingRecord';
import { buildViewXml, IRecordQuery } from '../logic/caml';
import { formatRecordId } from '../logic/recordId';
import { buildVersionHistory, IVersionSnapshot } from '../logic/versionDiff';
import { parseSharePointDate } from '../logic/dates';
import { fromRenderRow, fromRestItem, fromVersion, getRestEtag, getRestSelect, toRestPayload } from './fieldMapping';
import { toAppError } from './errors';

type Json = Record<string, unknown>;

export interface IExportProgress {
  loaded: number;
}

/**
 * All reads and writes of receiving records. UI components never talk to
 * SharePoint directly; they call this service.
 */
export class ReceivingRecordService {
  private readonly _sp: SPFI;
  private readonly _config: IAppConfig;

  public constructor(sp: SPFI, config: IAppConfig) {
    this._sp = sp;
    this._config = config;
  }

  private get _list(): IList {
    return this._sp.web.lists.getByTitle(this._config.solution.list.title);
  }

  private get _listTitle(): string {
    return this._config.solution.list.title;
  }

  /** Throws a friendly ListNotFound/PermissionDenied error if the list cannot be read. */
  public async ensureListExists(): Promise<void> {
    try {
      await this._list.select('Id')();
    } catch (error) {
      throw toAppError(error, { action: 'open the receiving records', listTitle: this._listTitle });
    }
  }

  public async getById(id: number): Promise<IReceivingRecord> {
    const { select, expand } = getRestSelect(this._config.fields);
    try {
      const item: Json = await this._list.items
        .getById(id)
        .select(...select)
        .expand(...expand)();
      const record: IReceivingRecord = {
        id,
        etag: getRestEtag(item),
        values: fromRestItem(item, this._config.fields)
      };
      this._fillMissingRecordId(record);
      return record;
    } catch (error) {
      throw toAppError(error, { action: 'open this record', listTitle: this._listTitle });
    }
  }

  /**
   * One page of records for the list view. `pageToken` comes from the
   * previous page's `nextPageToken` (undefined for the first page).
   */
  public async query(query: IRecordQuery, pageToken?: string, pageSize?: number): Promise<IRecordPage> {
    const viewFieldKeys = this._config.fields.filter((f) => f.showInTable || f.includeInExport).map((f) => f.key);
    viewFieldKeys.push(FieldKeys.approvalStatus, FieldKeys.createdBy, FieldKeys.receiverName, FieldKeys.created);
    const viewXml = buildViewXml(query, {
      fields: this._config.fields,
      recordIdPrefix: this._config.solution.recordId.prefix,
      rowLimit: pageSize || this._config.solution.listView.pageSize,
      viewFieldKeys,
      dateRangeFieldKey: this._config.solution.listView.dateRangeFieldKey
    });

    const pagingQuery = new Map<string, string>();
    if (pageToken) {
      pageToken.split('&').forEach((pair) => {
        const index = pair.indexOf('=');
        if (index <= 0) return;
        const key = decodeURIComponent(pair.substring(0, index));
        // "View" would make SharePoint use a saved view instead of our query.
        if (key.toLowerCase() === 'view') return;
        pagingQuery.set(key, decodeURIComponent(pair.substring(index + 1)));
      });
    }

    try {
      const result = await this._list.renderListDataAsStream(
        {
          ViewXml: viewXml,
          RenderOptions: RenderListDataOptions.ListData,
          DatesInUtc: true
        },
        undefined,
        pagingQuery
      );
      const rows = (result.Row || []) as Json[];
      const records = rows.map((row) => {
        const record: IReceivingRecord = { id: Number(row.ID), values: fromRenderRow(row, this._config.fields) };
        this._fillMissingRecordId(record);
        return record;
      });
      const next = result.NextHref ? result.NextHref.replace(/^\?/, '') : undefined;
      return { records, nextPageToken: next };
    } catch (error) {
      throw toAppError(error, { action: 'load the receiving records', listTitle: this._listTitle });
    }
  }

  /** All records matching the query (for Excel export), up to `maxRows`. */
  public async queryAll(
    query: IRecordQuery,
    maxRows: number,
    onProgress?: (progress: IExportProgress) => void
  ): Promise<{ records: IReceivingRecord[]; truncated: boolean }> {
    const all: IReceivingRecord[] = [];
    let token: string | undefined;
    do {
      const page = await this.query(query, token, 500);
      all.push(...page.records);
      token = page.nextPageToken;
      if (onProgress) onProgress({ loaded: all.length });
    } while (token && all.length < maxRows);
    return { records: all.slice(0, maxRows), truncated: !!token || all.length > maxRows };
  }

  /** Creates a record, then writes its Record ID (derived from the new item ID). */
  public async create(values: FieldValues, status: ApprovalStatusValue): Promise<IReceivingRecord> {
    const payload = toRestPayload(values, this._editableFields());
    payload[this._internalName(FieldKeys.approvalStatus)] = status;
    let newId: number;
    let created: Date;
    try {
      const result: Json = await this._list.items.add(payload);
      newId = Number(result.Id || result.ID);
      created = parseSharePointDate(result.Created) || new Date();
    } catch (error) {
      throw toAppError(error, { action: 'create a receiving record', listTitle: this._listTitle });
    }

    try {
      await this._list.items.getById(newId).update({ Title: this.formatRecordId(newId, created) });
    } catch {
      // Not fatal: the ID is shown from the item ID and repaired on the next save.
    }
    return this.getById(newId);
  }

  /**
   * Saves changes. Uses the record's ETag so a concurrent change by someone
   * else results in a "Conflict" error instead of silently overwriting it.
   */
  public async update(record: IReceivingRecord, values: FieldValues, status?: ApprovalStatusValue): Promise<IReceivingRecord> {
    const payload = toRestPayload(values, this._editableFields());
    if (status) {
      payload[this._internalName(FieldKeys.approvalStatus)] = status;
    }
    const recordIdValue = record.values[FieldKeys.recordId];
    if (typeof recordIdValue === 'string' && recordIdValue) {
      payload.Title = recordIdValue;
    }
    try {
      await this._list.items.getById(record.id).update(payload, record.etag || '*');
    } catch (error) {
      throw toAppError(error, { action: 'save this record', listTitle: this._listTitle });
    }
    return this.getById(record.id);
  }

  /** Approve or reject. Records who decided and when. */
  public async decide(
    record: IReceivingRecord,
    status: ApprovalStatusValue,
    supervisorComments: string | undefined,
    reviewerUserId: number
  ): Promise<IReceivingRecord> {
    const payload: Record<string, unknown> = {
      [this._internalName(FieldKeys.approvalStatus)]: status,
      [this._internalName(FieldKeys.supervisorComments)]: supervisorComments && supervisorComments.trim() ? supervisorComments.trim() : '',
      [`${this._internalName(FieldKeys.reviewedBy)}Id`]: reviewerUserId,
      [this._internalName(FieldKeys.reviewedOn)]: new Date().toISOString()
    };
    try {
      await this._list.items.getById(record.id).update(payload, record.etag || '*');
    } catch (error) {
      throw toAppError(error, { action: 'approve or reject records', listTitle: this._listTitle });
    }
    return this.getById(record.id);
  }

  /** Change history built from SharePoint version history. Newest first. */
  public async getHistory(id: number): Promise<IVersionEntry[]> {
    try {
      const versions: Json[] = await this._list.items.getById(id).versions();
      const snapshots: IVersionSnapshot[] = versions.map((version) => {
        const editor = version.Editor as Json | undefined;
        return {
          versionId: Number(version.VersionId),
          versionLabel: String(version.VersionLabel || ''),
          modified: parseSharePointDate(version.Modified) || parseSharePointDate(version.Created),
          modifiedBy: editor ? String(editor.LookupValue || '') : undefined,
          values: fromVersion(version, this._config.fields)
        };
      });
      return buildVersionHistory(snapshots, this._config.fields);
    } catch (error) {
      throw toAppError(error, { action: 'view the change history', listTitle: this._listTitle });
    }
  }

  public formatRecordId(itemId: number, created: Date): string {
    const { prefix, digits } = this._config.solution.recordId;
    return formatRecordId(prefix, digits, itemId, created);
  }

  /** Fields the app writes from the form (excludes read-only and built-in fields). */
  private _editableFields(): IFieldConfig[] {
    return this._config.fields.filter((f) => !f.readOnly && !f.builtIn);
  }

  private _internalName(key: string): string {
    const field = this._config.fields.filter((f) => f.key === key)[0];
    return field ? field.internalName : key;
  }

  /** Shows a computed Record ID if the Title update after creation did not happen. */
  private _fillMissingRecordId(record: IReceivingRecord): void {
    if (record.values[FieldKeys.recordId]) return;
    const created = record.values[FieldKeys.created];
    record.values[FieldKeys.recordId] = this.formatRecordId(record.id, created instanceof Date ? created : new Date());
  }
}
