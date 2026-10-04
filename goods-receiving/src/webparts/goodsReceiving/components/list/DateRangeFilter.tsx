import * as React from 'react';
import { DatePicker, DayOfWeek } from '@fluentui/react';
import { formatDate } from '../../../../logic/format';
import { fromIsoDateString, toIsoDateString } from '../../../../logic/dates';
import { parseDateInput } from '../form/FieldEditor';
import styles from '../App.module.scss';

export interface IDateRange {
  kind: 'dateRange';
  from?: string;
  to?: string;
}

export interface IDateRangeFilterProps {
  label: string;
  value: IDateRange | undefined;
  onChange: (value: IDateRange | undefined) => void;
  /** Render the two pickers as separate items of a flex container (filter bar). */
  inline?: boolean;
}

/** "From" and "To" date pickers producing yyyy-mm-dd strings. */
export const DateRangeFilter: React.FC<IDateRangeFilterProps> = ({ label, value, onChange, inline }) => {
  const from = value && value.from ? fromIsoDateString(value.from) : undefined;
  const to = value && value.to ? fromIsoDateString(value.to) : undefined;

  const update = (nextFrom: Date | undefined, nextTo: Date | undefined): void => {
    if (!nextFrom && !nextTo) {
      onChange(undefined);
      return;
    }
    onChange({
      kind: 'dateRange',
      from: nextFrom ? toIsoDateString(nextFrom) : undefined,
      to: nextTo ? toIsoDateString(nextTo) : undefined
    });
  };

  const common = {
    formatDate: (date?: Date) => (date ? formatDate(date) : ''),
    parseDateFromString: (text: string) => parseDateInput(text) || null,
    allowTextInput: true,
    firstDayOfWeek: DayOfWeek.Monday
  };

  const pickers = (
    <>
      <div className={inline ? styles.dateFilter : undefined}>
        <DatePicker
          {...common}
          label={`${label} from`}
          placeholder="Any date"
          value={from}
          maxDate={to}
          onSelectDate={(date) => update(date || undefined, to)}
        />
      </div>
      <div className={inline ? styles.dateFilter : undefined}>
        <DatePicker
          {...common}
          label="to"
          placeholder="Any date"
          value={to}
          minDate={from}
          onSelectDate={(date) => update(from, date || undefined)}
        />
      </div>
    </>
  );

  return inline ? pickers : <div style={{ display: 'flex', gap: 8 }}>{pickers}</div>;
};
