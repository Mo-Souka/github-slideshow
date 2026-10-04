import * as React from 'react';
import { IPersonaProps, Label, NormalPeoplePicker } from '@fluentui/react';
import { IPersonValue } from '../../../../models/IReceivingRecord';
import { IPersonSuggestion } from '../../../../services/UserService';
import { toAppError } from '../../../../services/errors';
import { useAppContext } from '../AppContext';
import styles from '../App.module.scss';

export interface IPeoplePickerProps {
  label: string;
  value: IPersonValue | undefined;
  onChange: (value: IPersonValue | undefined) => void;
  required?: boolean;
  disabled?: boolean;
  errorMessage?: string;
  placeholder?: string;
}

interface IPersonaWithSuggestion extends IPersonaProps {
  suggestion?: IPersonSuggestion;
}

/** Single-person picker that searches the organisation directory. */
export const PeoplePicker: React.FC<IPeoplePickerProps> = (props) => {
  const { services } = useAppContext();
  const [searchError, setSearchError] = React.useState<string | undefined>();
  const [resolving, setResolving] = React.useState(false);
  const id = React.useMemo(() => `gr-people-${Math.random().toString(36).substring(2)}`, []);

  const selected: IPersonaProps[] = props.value
    ? [{ key: String(props.value.id), text: props.value.title, secondaryText: props.value.email }]
    : [];

  const onResolveSuggestions = async (filter: string): Promise<IPersonaProps[]> => {
    try {
      setSearchError(undefined);
      const results = await services.users.searchPeople(filter);
      return results.map<IPersonaWithSuggestion>((s) => ({
        key: s.key,
        text: s.displayName,
        secondaryText: s.email || s.jobTitle,
        suggestion: s
      }));
    } catch (error) {
      setSearchError(toAppError(error).userMessage);
      return [];
    }
  };

  const onPickerChange = (items?: IPersonaProps[]): void => {
    const item = items && items.length > 0 ? (items[items.length - 1] as IPersonaWithSuggestion) : undefined;
    if (!item) {
      props.onChange(undefined);
      return;
    }
    if (!item.suggestion) return;
    setResolving(true);
    services.users
      .ensureUser(item.suggestion)
      .then((person) => {
        setResolving(false);
        props.onChange(person);
      })
      .catch((error) => {
        setResolving(false);
        setSearchError(toAppError(error).userMessage);
      });
  };

  const error = props.errorMessage || searchError;
  return (
    <div>
      <Label htmlFor={id} required={props.required} disabled={props.disabled}>
        {props.label}
      </Label>
      <NormalPeoplePicker
        selectedItems={selected}
        onResolveSuggestions={onResolveSuggestions}
        onChange={onPickerChange}
        itemLimit={1}
        resolveDelay={300}
        disabled={props.disabled || resolving}
        inputProps={{
          id,
          placeholder: selected.length === 0 ? props.placeholder || 'Type a name or email' : undefined,
          'aria-invalid': !!error
        }}
        pickerSuggestionsProps={{
          suggestionsHeaderText: 'People',
          noResultsFoundText: 'No people found',
          loadingText: 'Searching...'
        }}
      />
      {error && (
        <div className={styles.fieldError} role="alert">
          {error}
        </div>
      )}
    </div>
  );
};
