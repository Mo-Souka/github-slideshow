import * as React from 'react';
import {
  ActionButton,
  DefaultButton,
  Dropdown,
  IconButton,
  PrimaryButton,
  SearchBox,
  Spinner,
  SpinnerSize
} from '@fluentui/react';
import { FieldKeys } from '../../../../config/fieldKeys';
import { IFieldConfig } from '../../../../models/IFieldConfig';
import { IReceivingRecord, IRecordPage } from '../../../../models/IReceivingRecord';
import { countActiveFilters, FilterValue, IRecordQuery, isFilterEmpty } from '../../../../logic/caml';
import { formatDate } from '../../../../logic/format';
import { fromIsoDateString } from '../../../../logic/dates';
import { canCreateRecords } from '../../../../logic/workflow';
import { NARROW_WIDTH, useAppContext } from '../AppContext';
import { ErrorMessage } from '../common/ErrorMessage';
import { ExportButton } from './ExportButton';
import { DateRangeFilter } from './DateRangeFilter';
import { FilterPanel } from './FilterPanel';
import { getDefaultQuery, loadStoredQuery, storeQuery } from './listQuery';
import { RecordCards } from './RecordCards';
import { RecordTable } from './RecordTable';
import styles from '../App.module.scss';

export interface IRecordListProps {
  /** Supervisor quick view: only records waiting for approval. */
  pendingOnly: boolean;
}

function describeFilter(field: IFieldConfig, filter: FilterValue): string {
  switch (filter.kind) {
    case 'choice':
      return `${field.displayName}: ${filter.values.join(', ')}`;
    case 'person':
    case 'lookup':
      return `${field.displayName}: ${filter.title}`;
    case 'text':
      return `${field.displayName} contains "${filter.value}"`;
    case 'dateRange': {
      const from = filter.from ? fromIsoDateString(filter.from) : undefined;
      const to = filter.to ? fromIsoDateString(filter.to) : undefined;
      return `${field.displayName}: ${from ? formatDate(from) : '...'} - ${to ? formatDate(to) : '...'}`;
    }
    case 'numberRange':
      return `${field.displayName}: ${filter.min === undefined ? '...' : filter.min} - ${filter.max === undefined ? '...' : filter.max}`;
    default:
      return field.displayName;
  }
}

export const RecordList: React.FC<IRecordListProps> = ({ pendingOnly }) => {
  const { config, user, services, navigate, width } = useAppContext();
  const listTitle = config.solution.list.title;
  const dateKey = config.solution.listView.dateRangeFieldKey;
  const dateField = config.fields.filter((f) => f.key === dateKey)[0];
  const tableFields = React.useMemo(() => config.fields.filter((f) => f.showInTable), [config]);
  const searchableNames = React.useMemo(
    () => config.fields.filter((f) => f.searchable).map((f) => f.displayName.replace(/ Number$/, '').toLowerCase()),
    [config]
  );
  // Filters shown in the panel. The date range has its own controls; the
  // pending view always filters on approval status.
  const panelFields = React.useMemo(
    () =>
      config.fields.filter(
        (f) => f.filterable && f.key !== dateKey && !(pendingOnly && f.key === FieldKeys.approvalStatus)
      ),
    [config, dateKey, pendingOnly]
  );

  const [query, setQuery] = React.useState<IRecordQuery>(
    () => loadStoredQuery(pendingOnly, listTitle) || getDefaultQuery(config, pendingOnly)
  );
  const [searchText, setSearchText] = React.useState<string>(query.search);
  const [pageTokens, setPageTokens] = React.useState<(string | undefined)[]>([undefined]);
  const [pageIndex, setPageIndex] = React.useState(0);
  const [page, setPage] = React.useState<IRecordPage | undefined>();
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<unknown>();
  const [filtersOpen, setFiltersOpen] = React.useState(false);
  const [reloadCounter, setReloadCounter] = React.useState(0);

  const pageSize = config.solution.listView.pageSize;
  const isNarrow = width < NARROW_WIDTH;

  const updateQuery = React.useCallback(
    (changes: Partial<IRecordQuery>) => {
      setQuery((previous) => {
        const next = { ...previous, ...changes };
        storeQuery(pendingOnly, listTitle, next);
        return next;
      });
      setPageTokens([undefined]);
      setPageIndex(0);
    },
    [pendingOnly, listTitle]
  );

  // Search as you type (after a short pause) or immediately on Enter.
  React.useEffect(() => {
    if (searchText === query.search) return undefined;
    const timer = window.setTimeout(() => updateQuery({ search: searchText }), 500);
    return () => window.clearTimeout(timer);
  }, [searchText, query.search, updateQuery]);

  React.useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(undefined);
    services.records
      .query(query, pageTokens[pageIndex], pageSize)
      .then((result) => {
        if (cancelled) return;
        setPage(result);
        setLoading(false);
        setPageTokens((tokens) => {
          const next = tokens.slice(0, pageIndex + 1);
          next[pageIndex + 1] = result.nextPageToken;
          return next;
        });
      })
      .catch((e) => {
        if (cancelled) return;
        setError(e);
        setLoading(false);
      });
    return () => {
      cancelled = true;
    };
    // pageTokens is intentionally not a dependency: it is updated by this effect.
  }, [services, query, pageIndex, pageSize, reloadCounter]);

  const openRecord = (record: IReceivingRecord): void => navigate({ name: 'view', id: record.id });

  const onSort = (key: string): void => {
    updateQuery({ sortKey: key, sortAscending: query.sortKey === key ? !query.sortAscending : true });
  };

  const setFilter = (key: string, value: FilterValue | undefined): void => {
    updateQuery({ filters: { ...query.filters, [key]: value } });
  };

  const resetFilters = (): void => {
    const defaults = getDefaultQuery(config, pendingOnly);
    setSearchText('');
    updateQuery({ ...defaults, sortKey: query.sortKey, sortAscending: query.sortAscending });
  };

  const activeFilterCount = countActiveFilters(query, pendingOnly ? [dateKey, FieldKeys.approvalStatus] : [dateKey]);
  const chips = panelFields
    .map((f) => ({ field: f, filter: query.filters[f.key] }))
    .filter((c) => c.filter && !isFilterEmpty(c.filter));

  const sortableFields = tableFields.filter((f) => f.sortable);
  const records = page ? page.records : [];
  const firstRow = pageIndex * pageSize + 1;
  const hasNext = !!pageTokens[pageIndex + 1];
  const isSupervisor = user.isSupervisor;

  return (
    <div>
      <div className={styles.header}>
        <div className={styles.titleRow}>
          <h2 className={styles.title}>{pendingOnly ? 'Pending approval' : 'Goods Receiving'}</h2>
        </div>
        <div className={styles.toolbar}>
          {canCreateRecords(user) && (
            <PrimaryButton text="New record" iconProps={{ iconName: 'Add' }} onClick={() => navigate({ name: 'new' })} />
          )}
          {isSupervisor && (
            <DefaultButton
              text={pendingOnly ? 'All records' : 'Pending approval'}
              iconProps={{ iconName: pendingOnly ? 'BulletedList' : 'ReviewRequestSolid' }}
              onClick={() => navigate({ name: pendingOnly ? 'list' : 'pending' })}
            />
          )}
          <ExportButton query={query} disabled={loading && !page} />
        </div>
      </div>

      <div className={styles.filterBar}>
        <div className={styles.search}>
          <SearchBox
            placeholder={`Search: record ID, ${searchableNames.filter((n) => n !== 'record id').join(', ')}`}
            value={searchText}
            onChange={(e, text) => setSearchText(text || '')}
            onSearch={(text) => updateQuery({ search: text || '' })}
            onClear={() => {
              setSearchText('');
              updateQuery({ search: '' });
            }}
            ariaLabel="Search records"
          />
        </div>
        {dateField && (
          <DateRangeFilter
            inline={true}
            label={dateField.displayName}
            value={(() => {
              const f = query.filters[dateKey];
              return f && f.kind === 'dateRange' ? f : undefined;
            })()}
            onChange={(value) => setFilter(dateKey, value)}
          />
        )}
        <DefaultButton
          text={activeFilterCount > 0 ? `Filters (${activeFilterCount})` : 'Filters'}
          iconProps={{ iconName: 'Filter' }}
          onClick={() => setFiltersOpen(true)}
        />
        <IconButton iconProps={{ iconName: 'Refresh' }} title="Refresh" ariaLabel="Refresh" onClick={() => setReloadCounter(reloadCounter + 1)} />
      </div>

      {(chips.length > 0 || query.search) && (
        <div className={styles.chips}>
          {chips.map(({ field, filter }) => (
            <span key={field.key} className={styles.chip}>
              {describeFilter(field, filter!)}
              <IconButton
                iconProps={{ iconName: 'Cancel' }}
                title="Remove filter"
                ariaLabel={`Remove filter ${field.displayName}`}
                styles={{ root: { width: 24, height: 24 }, icon: { fontSize: 10 } }}
                onClick={() => setFilter(field.key, undefined)}
              />
            </span>
          ))}
          <ActionButton text="Reset all" iconProps={{ iconName: 'ClearFilter' }} onClick={resetFilters} styles={{ root: { height: 28 } }} />
        </div>
      )}

      {isNarrow && sortableFields.length > 0 && (
        <div style={{ display: 'flex', gap: 8, alignItems: 'flex-end', marginBottom: 8 }}>
          <Dropdown
            label="Sort by"
            styles={{ root: { flex: 1 } }}
            options={sortableFields.map((f) => ({ key: f.key, text: f.displayName }))}
            selectedKey={query.sortKey}
            onChange={(e, option) => option && updateQuery({ sortKey: String(option.key), sortAscending: true })}
          />
          <IconButton
            iconProps={{ iconName: query.sortAscending ? 'SortUp' : 'SortDown' }}
            title={query.sortAscending ? 'Ascending' : 'Descending'}
            ariaLabel="Change sort direction"
            onClick={() => updateQuery({ sortAscending: !query.sortAscending })}
          />
        </div>
      )}

      {error !== undefined && <ErrorMessage error={error} onRetry={() => setReloadCounter(reloadCounter + 1)} />}

      {loading && <Spinner className={styles.loading} size={SpinnerSize.large} label="Loading records..." />}

      {!loading && !error && records.length === 0 && (
        <div className={styles.emptyState}>
          {pendingOnly
            ? 'There are no records waiting for approval.'
            : query.search || activeFilterCount > 0
              ? 'No records match your search and filters.'
              : 'No records in this date range.'}
        </div>
      )}

      {!loading && records.length > 0 && (
        <>
          {isNarrow ? (
            <RecordCards records={records} fields={tableFields} onOpen={openRecord} />
          ) : (
            <RecordTable
              records={records}
              fields={tableFields}
              sortKey={query.sortKey}
              sortAscending={query.sortAscending}
              onSort={onSort}
              onOpen={openRecord}
            />
          )}
        </>
      )}

      {!loading && (pageIndex > 0 || hasNext) && (
        <div className={styles.pager}>
          <DefaultButton
            text="Previous"
            iconProps={{ iconName: 'ChevronLeft' }}
            disabled={pageIndex === 0}
            onClick={() => setPageIndex(pageIndex - 1)}
          />
          <span className={styles.subtitle}>
            Page {pageIndex + 1} · records {firstRow}-{firstRow + records.length - 1}
          </span>
          <DefaultButton
            text="Next"
            menuIconProps={{ iconName: 'ChevronRight' }}
            disabled={!hasNext}
            onClick={() => setPageIndex(pageIndex + 1)}
          />
        </div>
      )}

      {pendingOnly && !loading && records.length > 0 && (
        <p className={styles.subtitle}>
          Oldest records first. Records you created or received yourself must be approved by another supervisor.
        </p>
      )}

      <FilterPanel
        isOpen={filtersOpen}
        fields={panelFields}
        filters={query.filters}
        onDismiss={() => setFiltersOpen(false)}
        onApply={(filters) => {
          setFiltersOpen(false);
          updateQuery({ filters });
        }}
      />
    </div>
  );
};
