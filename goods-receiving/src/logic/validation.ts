import { IFieldConfig } from '../models/IFieldConfig';
import { FieldValue, FieldValues, isReferenceValue } from '../models/IReceivingRecord';

/**
 * draft  = "Save as Draft": only fields marked requiredForDraft must be filled.
 * submit = "Submit for Approval" (and approval actions): all required and
 *          conditionally required fields must be filled.
 * Format rules (number ranges, length, pattern) are checked in both modes.
 */
export type ValidationMode = 'draft' | 'submit';

/** Error messages keyed by field key. */
export type ValidationErrors = Record<string, string>;

const DEFAULT_TEXT_MAX_LENGTH = 255;

export function isEmptyValue(value: FieldValue): boolean {
  if (value === undefined) return true;
  if (typeof value === 'string') return value.trim() === '';
  if (typeof value === 'number') return false;
  if (value instanceof Date) return isNaN(value.getTime());
  return !isReferenceValue(value) || !(value.id > 0);
}

function valueAsText(value: FieldValue): string {
  if (typeof value === 'string') return value;
  if (typeof value === 'number') return String(value);
  return '';
}

/** Whether the field is required for the given values and mode. */
export function isFieldRequired(field: IFieldConfig, values: FieldValues, mode: ValidationMode): boolean {
  if (mode === 'draft') {
    return !!field.requiredForDraft;
  }
  if (field.required || field.requiredForDraft) {
    return true;
  }
  if (field.requiredWhen) {
    const other = valueAsText(values[field.requiredWhen.field]);
    return field.requiredWhen.in.indexOf(other) >= 0;
  }
  return false;
}

function requiredMessage(field: IFieldConfig, values: FieldValues, mode: ValidationMode): string {
  if (mode === 'submit' && !field.required && field.requiredWhen && field.requiredWhen.message) {
    return field.requiredWhen.message;
  }
  return `${field.displayName} is required.`;
}

function formatNumber(value: number): string {
  return String(value);
}

function validateFormat(field: IFieldConfig, value: FieldValue): string | undefined {
  const rules = field.validation || {};
  const name = field.displayName;

  switch (field.type) {
    case 'Number': {
      if (typeof value !== 'number' || !isFinite(value)) {
        return `${name} must be a number.`;
      }
      if (rules.integer && Math.floor(value) !== value) {
        return `${name} must be a whole number.`;
      }
      if (rules.minExclusive !== undefined && !(value > rules.minExclusive)) {
        return `${name} must be greater than ${formatNumber(rules.minExclusive)}.`;
      }
      if (rules.min !== undefined && value < rules.min) {
        return `${name} must be at least ${formatNumber(rules.min)}.`;
      }
      if (rules.maxExclusive !== undefined && !(value < rules.maxExclusive)) {
        return `${name} must be less than ${formatNumber(rules.maxExclusive)}.`;
      }
      if (rules.max !== undefined && value > rules.max) {
        return `${name} must be at most ${formatNumber(rules.max)}.`;
      }
      if (field.decimals !== undefined) {
        const factor = Math.pow(10, field.decimals);
        if (Math.abs(Math.round(value * factor) - value * factor) > 1e-6) {
          return field.decimals === 0
            ? `${name} must be a whole number.`
            : `${name} can have at most ${field.decimals} decimal places.`;
        }
      }
      return undefined;
    }
    case 'Text':
    case 'Note': {
      const text = typeof value === 'string' ? value : '';
      const maxLength = rules.maxLength || (field.type === 'Text' ? DEFAULT_TEXT_MAX_LENGTH : undefined);
      if (maxLength !== undefined && text.length > maxLength) {
        return `${name} can be at most ${maxLength} characters (currently ${text.length}).`;
      }
      // The pattern comes from fields.json, which only administrators can change.
      // eslint-disable-next-line @rushstack/security/no-unsafe-regexp
      if (rules.pattern && !new RegExp(rules.pattern).test(text)) {
        return rules.patternMessage || `${name} has an invalid format.`;
      }
      return undefined;
    }
    case 'Choice': {
      const text = typeof value === 'string' ? value : '';
      if (field.choices && field.choices.indexOf(text) < 0) {
        return `${name} must be one of: ${field.choices.join(', ')}.`;
      }
      return undefined;
    }
    case 'Date':
    case 'DateTime':
      return value instanceof Date && !isNaN(value.getTime()) ? undefined : `${name} must be a valid date.`;
    case 'User':
    case 'Lookup':
      return isReferenceValue(value) && value.id > 0 ? undefined : `${name} is not a valid selection.`;
    default:
      return undefined;
  }
}

/** Validates one field. Returns the error message or undefined. */
export function validateField(field: IFieldConfig, values: FieldValues, mode: ValidationMode): string | undefined {
  const value = values[field.key];
  if (isEmptyValue(value)) {
    return isFieldRequired(field, values, mode) ? requiredMessage(field, values, mode) : undefined;
  }
  return validateFormat(field, value);
}

/** Validates the given fields and returns the errors (empty object when valid). */
export function validateValues(fields: IFieldConfig[], values: FieldValues, mode: ValidationMode): ValidationErrors {
  const errors: ValidationErrors = {};
  fields.forEach((field) => {
    const message = validateField(field, values, mode);
    if (message) {
      errors[field.key] = message;
    }
  });
  return errors;
}

export function hasErrors(errors: ValidationErrors): boolean {
  return Object.keys(errors).length > 0;
}
