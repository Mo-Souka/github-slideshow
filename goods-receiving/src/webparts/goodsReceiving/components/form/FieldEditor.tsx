import * as React from 'react';
import {
  ComboBox,
  DatePicker,
  DayOfWeek,
  defaultDatePickerStrings,
  Dropdown,
  IComboBoxOption,
  IDropdownOption,
  TextField
} from '@fluentui/react';
import { IFieldConfig } from '../../../../models/IFieldConfig';
import { FieldValue, ILookupValue, IPersonValue, isReferenceValue } from '../../../../models/IReceivingRecord';
import { formatDate, formatNumber } from '../../../../logic/format';
import { toAppError } from '../../../../services/errors';
import { PeoplePicker } from '../common/PeoplePicker';
import { useAppContext } from '../AppContext';
import styles from '../App.module.scss';

export interface IFieldEditorProps {
  field: IFieldConfig;
  value: FieldValue;
  onChange: (key: string, value: FieldValue) => void;
  required: boolean;
  errorMessage?: string;
  disabled?: boolean;
}

/** Parses "12,5" or "1,234.5" typed on any keyboard. Returns NaN for invalid input. */
export function parseNumberInput(text: string): number | undefined {
  const trimmed = text.trim();
  if (trimmed === '') return undefined;
  let normalized = trimmed.replace(/\s/g, '');
  if (normalized.indexOf(',') >= 0 && normalized.indexOf('.') < 0) {
    normalized = normalized.replace(',', '.');
  } else {
    normalized = normalized.replace(/,/g, '');
  }
  return /^-?\d*\.?\d+$|^-?\d+\.$/.test(normalized) ? parseFloat(normalized) : NaN;
}

/** Accepts dd.mm.yyyy, dd/mm/yyyy, yyyy-mm-dd and "04 Oct 2026" when typed by hand. */
export function parseDateInput(text: string): Date | undefined {
  const value = (text || '').trim();
  let match = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(value);
  if (match) return new Date(+match[1], +match[2] - 1, +match[3]);
  match = /^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/.exec(value);
  if (match) return new Date(+match[3], +match[2] - 1, +match[1]);
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? undefined : new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

const NumberEditor: React.FC<IFieldEditorProps> = ({ field, value, onChange, required, errorMessage, disabled }) => {
  const [text, setText] = React.useState<string>(typeof value === 'number' && isFinite(value) ? String(value) : '');

  // Keep the text in sync when the value is replaced from outside (e.g. form reset).
  React.useEffect(() => {
    const parsed = parseNumberInput(text);
    const same = parsed === value || (parsed !== undefined && value !== undefined && isNaN(parsed) && isNaN(value as number));
    if (!same) setText(typeof value === 'number' && isFinite(value) ? String(value) : '');
  }, [value]);

  return (
    <TextField
      label={field.displayName}
      required={required}
      value={text}
      inputMode="decimal"
      autoComplete="off"
      disabled={disabled}
      errorMessage={errorMessage}
      description={field.description}
      onChange={(e, newText) => {
        setText(newText || '');
        onChange(field.key, parseNumberInput(newText || ''));
      }}
      onBlur={() => {
        if (typeof value === 'number' && isFinite(value)) setText(formatNumber(value, field.decimals).replace(/,/g, ''));
      }}
    />
  );
};

const LookupEditor: React.FC<IFieldEditorProps> = ({ field, value, onChange, required, errorMessage, disabled }) => {
  const { services } = useAppContext();
  const [options, setOptions] = React.useState<ILookupValue[]>([]);
  const [loadError, setLoadError] = React.useState<string | undefined>();

  React.useEffect(() => {
    let cancelled = false;
    services.records
      .getLookupOptions(field)
      .then((result) => {
        if (!cancelled) setOptions(result);
      })
      .catch((error) => {
        if (!cancelled) setLoadError(toAppError(error).userMessage);
      });
    return () => {
      cancelled = true;
    };
  }, [field, services]);

  const comboOptions: IComboBoxOption[] = options.map((o) => ({ key: o.id, text: o.title }));
  const selected = isReferenceValue(value) ? value.id : undefined;
  return (
    <ComboBox
      label={field.displayName}
      required={required}
      options={comboOptions}
      selectedKey={selected === undefined ? null : selected}
      allowFreeform={false}
      autoComplete="on"
      disabled={disabled}
      errorMessage={errorMessage || loadError}
      placeholder="Select..."
      onChange={(e, option) => {
        onChange(field.key, option ? { id: Number(option.key), title: option.text } : undefined);
      }}
    />
  );
};

/** Renders the right input for a field based on its type in fields.json. */
export const FieldEditor: React.FC<IFieldEditorProps> = (props) => {
  const { field, value, onChange, required, errorMessage, disabled } = props;

  switch (field.type) {
    case 'Number':
      return <NumberEditor {...props} />;

    case 'Date':
    case 'DateTime':
      return (
        <div>
          <DatePicker
            label={field.displayName}
            isRequired={required}
            value={value instanceof Date ? value : undefined}
            onSelectDate={(date) => onChange(field.key, date || undefined)}
            formatDate={(date) => (date ? formatDate(date) : '')}
            parseDateFromString={(text) => parseDateInput(text) || null}
            allowTextInput={true}
            firstDayOfWeek={DayOfWeek.Monday}
            disabled={disabled}
            placeholder="Select a date"
            strings={{
              ...defaultDatePickerStrings,
              isRequiredErrorMessage: `${field.displayName} is required.`,
              invalidInputErrorMessage: 'Please enter a valid date, e.g. 04.10.2026.'
            }}
            textField={errorMessage ? { errorMessage } : undefined}
          />
          {field.description && <div className={styles.help}>{field.description}</div>}
        </div>
      );

    case 'Choice': {
      const options: IDropdownOption[] = (field.choices || []).map((c) => ({ key: c, text: c }));
      return (
        <Dropdown
          label={field.displayName}
          required={required}
          options={options}
          selectedKey={typeof value === 'string' && value ? value : null}
          placeholder="Select..."
          disabled={disabled}
          errorMessage={errorMessage}
          onChange={(e, option) => onChange(field.key, option ? String(option.key) : undefined)}
        />
      );
    }

    case 'User':
      return (
        <div>
          <PeoplePicker
            label={field.displayName}
            value={isReferenceValue(value) ? (value as IPersonValue) : undefined}
            onChange={(person) => onChange(field.key, person)}
            required={required}
            disabled={disabled}
            errorMessage={errorMessage}
          />
          {field.description && <div className={styles.help}>{field.description}</div>}
        </div>
      );

    case 'Lookup':
      return <LookupEditor {...props} />;

    case 'Note':
      return (
        <TextField
          label={field.displayName}
          required={required}
          multiline={true}
          rows={4}
          autoAdjustHeight={true}
          value={typeof value === 'string' ? value : ''}
          disabled={disabled}
          errorMessage={errorMessage}
          description={field.description}
          onChange={(e, text) => onChange(field.key, text || undefined)}
        />
      );

    default:
      return (
        <TextField
          label={field.displayName}
          required={required}
          value={typeof value === 'string' ? value : ''}
          autoComplete="off"
          disabled={disabled}
          errorMessage={errorMessage}
          description={field.description}
          onChange={(e, text) => onChange(field.key, text || undefined)}
        />
      );
  }
};

/** Notes and other long fields span the full width of the form grid. */
export function isWideField(field: IFieldConfig): boolean {
  return field.type === 'Note';
}
