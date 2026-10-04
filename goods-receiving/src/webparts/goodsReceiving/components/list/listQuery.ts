import { ApprovalStatus, FieldKeys } from '../../../../config/fieldKeys';
import { IAppConfig } from '../../../../models/IFieldConfig';
import { IRecordQuery } from '../../../../logic/caml';
import { addDays, today, toIsoDateString } from '../../../../logic/dates';

/**
 * Default query for the record list:
 *  - normal view: Receiving Date within the last N days (web part setting),
 *    newest first. The date range keeps queries fast on large lists.
 *  - pending view (supervisors): all Pending Approval records, oldest first.
 */
export function getDefaultQuery(config: IAppConfig, pendingOnly: boolean): IRecordQuery {
  const { listView } = config.solution;
  if (pendingOnly) {
    return {
      search: '',
      filters: { [FieldKeys.approvalStatus]: { kind: 'choice', values: [ApprovalStatus.PendingApproval] } },
      sortKey: listView.dateRangeFieldKey,
      sortAscending: true
    };
  }
  const end = today();
  const start = addDays(end, -listView.defaultDateWindowDays);
  return {
    search: '',
    filters: {
      [listView.dateRangeFieldKey]: { kind: 'dateRange', from: toIsoDateString(start), to: toIsoDateString(end) }
    },
    sortKey: listView.defaultSortFieldKey,
    sortAscending: listView.defaultSortAscending
  };
}

const STORAGE_PREFIX = 'goods-receiving-query-';

/** Remembers filters while the user opens records and comes back (this browser tab only). */
export function loadStoredQuery(pendingOnly: boolean, listTitle: string): IRecordQuery | undefined {
  try {
    const raw = window.sessionStorage.getItem(`${STORAGE_PREFIX}${pendingOnly ? 'pending' : 'all'}-${listTitle}`);
    if (!raw) return undefined;
    const parsed = JSON.parse(raw) as IRecordQuery;
    return parsed && typeof parsed.search === 'string' && parsed.filters ? parsed : undefined;
  } catch {
    return undefined;
  }
}

export function storeQuery(pendingOnly: boolean, listTitle: string, query: IRecordQuery): void {
  try {
    window.sessionStorage.setItem(`${STORAGE_PREFIX}${pendingOnly ? 'pending' : 'all'}-${listTitle}`, JSON.stringify(query));
  } catch {
    // Storage may be disabled; remembering filters is only a convenience.
  }
}
