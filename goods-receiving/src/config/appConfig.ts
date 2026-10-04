import fieldsJson from './fields.json';
import solutionJson from './solution.json';
import { ApprovalStatus, FieldKeys, RequiredFieldTypes } from './fieldKeys';
import { IAppConfig, IFieldConfig, IFieldsConfigFile, ISolutionConfig } from '../models/IFieldConfig';
import { AppError } from '../models/AppError';

const FIELD_TYPES = ['Text', 'Note', 'Number', 'Date', 'DateTime', 'Choice', 'User', 'Lookup'];

export const defaultFieldsConfig: IFieldsConfigFile = fieldsJson as unknown as IFieldsConfigFile;
export const defaultSolutionConfig: ISolutionConfig = solutionJson as unknown as ISolutionConfig;

/**
 * Checks fields.json / solution.json for mistakes that would break the app.
 * Returns a list of human-readable problems (empty when the configuration is valid).
 */
export function validateConfig(fieldsFile: IFieldsConfigFile, solution: ISolutionConfig): string[] {
  const problems: string[] = [];
  const fields = fieldsFile.fields || [];
  const byKey: Record<string, IFieldConfig> = {};
  const internalNames: Record<string, boolean> = {};
  const sectionKeys = (fieldsFile.sections || []).map((s) => s.key);

  fields.forEach((field, index) => {
    const label = field.key ? `Field "${field.key}"` : `Field #${index + 1}`;
    if (!field.key) problems.push(`${label} has no key.`);
    if (!field.internalName) problems.push(`${label} has no internalName.`);
    if (!field.displayName) problems.push(`${label} has no displayName.`);
    if (byKey[field.key]) problems.push(`${label}: the key is used more than once.`);
    if (internalNames[field.internalName]) problems.push(`${label}: internalName "${field.internalName}" is used more than once.`);
    byKey[field.key] = field;
    internalNames[field.internalName] = true;

    if (FIELD_TYPES.indexOf(field.type) < 0) {
      problems.push(`${label}: unknown type "${field.type}".`);
    }
    if (field.type === 'Choice' && (!field.choices || field.choices.length === 0)) {
      problems.push(`${label}: Choice fields need a "choices" list.`);
    }
    if (field.type === 'Lookup' && !field.lookup) {
      problems.push(`${label}: Lookup fields need a "lookup" setting.`);
    }
    if (field.searchable && field.type !== 'Text') {
      problems.push(`${label}: only Text fields can be searchable.`);
    }
    if (field.section && sectionKeys.indexOf(field.section) < 0) {
      problems.push(`${label}: section "${field.section}" is not defined in "sections".`);
    }
    if (field.badgeTones && field.choices) {
      Object.keys(field.badgeTones).forEach((choice) => {
        if (field.choices!.indexOf(choice) < 0) {
          problems.push(`${label}: badgeTones has "${choice}", which is not one of its choices.`);
        }
      });
    }
    if (field.validation && field.validation.pattern) {
      try {
        // eslint-disable-next-line no-new, @rushstack/security/no-unsafe-regexp
        new RegExp(field.validation.pattern);
      } catch {
        problems.push(`${label}: validation.pattern is not a valid regular expression.`);
      }
    }
  });

  fields.forEach((field) => {
    if (!field.requiredWhen) return;
    const other = byKey[field.requiredWhen.field];
    if (!other) {
      problems.push(`Field "${field.key}": requiredWhen refers to unknown field "${field.requiredWhen.field}".`);
    } else if (other.choices) {
      field.requiredWhen.in.forEach((value) => {
        if (other.choices!.indexOf(value) < 0) {
          problems.push(`Field "${field.key}": requiredWhen value "${value}" is not a choice of "${other.key}".`);
        }
      });
    }
  });

  Object.keys(RequiredFieldTypes).forEach((key) => {
    const field = byKey[key];
    if (!field) {
      problems.push(`Required field "${key}" is missing. The app logic depends on it.`);
    } else if (field.type !== RequiredFieldTypes[key]) {
      problems.push(`Field "${key}" must have type ${RequiredFieldTypes[key]}.`);
    }
  });

  const approval = byKey[FieldKeys.approvalStatus];
  if (approval && approval.choices) {
    Object.keys(ApprovalStatus).forEach((name) => {
      const value = (ApprovalStatus as Record<string, string>)[name];
      if (approval.choices!.indexOf(value) < 0) {
        problems.push(`Field "approvalStatus" must include the choice "${value}".`);
      }
    });
  }

  const dateField = byKey[solution.listView.dateRangeFieldKey];
  if (!dateField || (dateField.type !== 'Date' && dateField.type !== 'DateTime')) {
    problems.push(`listView.dateRangeFieldKey "${solution.listView.dateRangeFieldKey}" must be a Date field.`);
  }
  const sortField = byKey[solution.listView.defaultSortFieldKey];
  if (!sortField || !sortField.sortable) {
    problems.push(`listView.defaultSortFieldKey "${solution.listView.defaultSortFieldKey}" must be a sortable field.`);
  }
  if (solution.library.documentTypes.indexOf(solution.library.photoDocumentType) < 0) {
    problems.push(`library.photoDocumentType "${solution.library.photoDocumentType}" must be one of library.documentTypes.`);
  }

  return problems;
}

/** Settings from the web part property pane that override solution.json. */
export interface IConfigOverrides {
  siteUrl: string;
  listTitle?: string;
  libraryTitle?: string;
  dateWindowDays?: number;
  pageSize?: number;
}

export function buildAppConfig(
  overrides: IConfigOverrides,
  fieldsFile: IFieldsConfigFile = defaultFieldsConfig,
  solutionFile: ISolutionConfig = defaultSolutionConfig
): IAppConfig {
  const problems = validateConfig(fieldsFile, solutionFile);
  if (problems.length > 0) {
    throw new AppError(
      'ConfigurationMismatch',
      'The field configuration (fields.json) has errors. Ask your administrator to fix it and redeploy the app.',
      problems.join('\n')
    );
  }

  const solution: ISolutionConfig = {
    ...solutionFile,
    list: { ...solutionFile.list, title: overrides.listTitle || solutionFile.list.title },
    library: { ...solutionFile.library, title: overrides.libraryTitle || solutionFile.library.title },
    listView: {
      ...solutionFile.listView,
      defaultDateWindowDays:
        overrides.dateWindowDays && overrides.dateWindowDays > 0
          ? overrides.dateWindowDays
          : solutionFile.listView.defaultDateWindowDays,
      pageSize: overrides.pageSize && overrides.pageSize > 0 ? overrides.pageSize : solutionFile.listView.pageSize
    }
  };

  const fields = fieldsFile.fields.slice().sort((a, b) => a.order - b.order);
  return {
    sections: fieldsFile.sections,
    fields,
    solution,
    siteUrl: overrides.siteUrl.replace(/\/+$/, '')
  };
}

export function getField(config: IAppConfig, key: string): IFieldConfig {
  for (const field of config.fields) {
    if (field.key === key) return field;
  }
  throw new AppError('ConfigurationMismatch', `Field "${key}" is not configured.`);
}

/** Fields the user can edit in the new/edit form. */
export function getFormFields(config: IAppConfig): IFieldConfig[] {
  return config.fields.filter((f) => !f.readOnly && !f.builtIn && f.showInForm !== false);
}
