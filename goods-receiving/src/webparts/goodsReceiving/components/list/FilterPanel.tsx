import * as React from 'react';
import { DefaultButton, Dropdown, IDropdownOption, Panel, PanelType, PrimaryButton, TextField } from '@fluentui/react';
import { IFieldConfig } from '../../../../models/IFieldConfig';
import { FilterValue } from '../../../../logic/caml';
import { PeoplePicker } from '../common/PeoplePicker';
import { DateRangeFilter } from './DateRangeFilter';

export interface IFilterPanelProps {
  isOpen: boolean;
  fields: IFieldConfig[];
  filters: Record<string, FilterValue | undefined>;
  onApply: (filters: Record<string, FilterValue | undefined>) => void;
  onDismiss: () => void;
}

/** Filters generated from the fields marked "filterable" in fields.json. */
export const FilterPanel: React.FC<IFilterPanelProps> = ({ isOpen, fields, filters, onApply, onDismiss }) => {
  const [draft, setDraft] = React.useState<Record<string, FilterValue | undefined>>(filters);

  React.useEffect(() => {
    if (isOpen) setDraft(filters);
  }, [isOpen, filters]);

  const set = (key: string, value: FilterValue | undefined): void => setDraft({ ...draft, [key]: value });

  const renderFilter = (field: IFieldConfig): React.ReactNode => {
    const current = draft[field.key];
    switch (field.type) {
      case 'Choice': {
        const selected = current && current.kind === 'choice' ? current.values : [];
        const options: IDropdownOption[] = (field.choices || []).map((c) => ({ key: c, text: c }));
        return (
          <Dropdown
            label={field.displayName}
            multiSelect={true}
            options={options}
            selectedKeys={selected}
            placeholder="All"
            onChange={(e, option) => {
              if (!option) return;
              const key = String(option.key);
              const next = option.selected ? selected.concat([key]) : selected.filter((v) => v !== key);
              set(field.key, next.length > 0 ? { kind: 'choice', values: next } : undefined);
            }}
          />
        );
      }
      case 'User':
        return (
          <PeoplePicker
            label={field.displayName}
            value={current && current.kind === 'person' ? { id: current.id, title: current.title } : undefined}
            onChange={(person) => set(field.key, person ? { kind: 'person', id: person.id, title: person.title } : undefined)}
            placeholder="Anyone"
          />
        );
      case 'Date':
      case 'DateTime':
        return (
          <DateRangeFilter
            label={field.displayName}
            value={current && current.kind === 'dateRange' ? current : undefined}
            onChange={(value) => set(field.key, value)}
          />
        );
      case 'Number': {
        const range = current && current.kind === 'numberRange' ? current : { kind: 'numberRange' as const };
        const parse = (text: string | undefined): number | undefined => {
          const n = parseFloat((text || '').replace(',', '.'));
          return isFinite(n) ? n : undefined;
        };
        const update = (min: number | undefined, max: number | undefined): void =>
          set(field.key, min === undefined && max === undefined ? undefined : { kind: 'numberRange', min, max });
        return (
          <div style={{ display: 'flex', gap: 8 }}>
            <TextField
              label={`${field.displayName} from`}
              inputMode="decimal"
              value={range.min === undefined ? '' : String(range.min)}
              onChange={(e, text) => update(parse(text), range.max)}
            />
            <TextField
              label="to"
              inputMode="decimal"
              value={range.max === undefined ? '' : String(range.max)}
              onChange={(e, text) => update(range.min, parse(text))}
            />
          </div>
        );
      }
      default:
        return (
          <TextField
            label={`${field.displayName} contains`}
            value={current && current.kind === 'text' ? current.value : ''}
            onChange={(e, text) => set(field.key, text ? { kind: 'text', value: text } : undefined)}
          />
        );
    }
  };

  const clearAll = (): void => {
    const cleared: Record<string, FilterValue | undefined> = { ...draft };
    fields.forEach((f) => {
      cleared[f.key] = undefined;
    });
    setDraft(cleared);
  };

  return (
    <Panel
      isOpen={isOpen}
      type={PanelType.smallFixedFar}
      headerText="Filters"
      closeButtonAriaLabel="Close"
      onDismiss={onDismiss}
      isFooterAtBottom={true}
      onRenderFooterContent={() => (
        <div style={{ display: 'flex', gap: 8 }}>
          <PrimaryButton text="Apply" onClick={() => onApply(draft)} />
          <DefaultButton text="Clear filters" onClick={clearAll} />
        </div>
      )}
    >
      <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        {fields.map((field) => (
          <div key={field.key}>{renderFilter(field)}</div>
        ))}
      </div>
    </Panel>
  );
};
