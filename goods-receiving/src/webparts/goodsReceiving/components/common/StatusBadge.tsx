import * as React from 'react';
import { BadgeTone, IFieldConfig } from '../../../../models/IFieldConfig';
import styles from '../App.module.scss';

const TONE_CLASS: Record<BadgeTone, string> = {
  neutral: styles.toneNeutral,
  info: styles.toneInfo,
  success: styles.toneSuccess,
  warning: styles.toneWarning,
  danger: styles.toneDanger
};

export interface IStatusBadgeProps {
  field: IFieldConfig;
  value: string | undefined;
  /** Prefix the value with the field name, e.g. "Inspection Status: Accepted". */
  showLabel?: boolean;
}

/** Coloured badge for choice fields that define "badgeTones" in fields.json. */
export const StatusBadge: React.FC<IStatusBadgeProps> = ({ field, value, showLabel }) => {
  if (!value) return null;
  const tone: BadgeTone = (field.badgeTones && field.badgeTones[value]) || 'neutral';
  return (
    <span className={`${styles.badge} ${TONE_CLASS[tone]}`} title={`${field.displayName}: ${value}`}>
      {showLabel ? `${field.displayName}: ${value}` : value}
    </span>
  );
};

export function hasBadge(field: IFieldConfig): boolean {
  return field.type === 'Choice' && !!field.badgeTones;
}
