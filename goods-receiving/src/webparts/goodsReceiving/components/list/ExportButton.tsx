import * as React from 'react';
import { IRecordQuery } from '../../../../logic/caml';

export interface IExportButtonProps {
  query: IRecordQuery;
  disabled?: boolean;
}

/** Placeholder until stage 7 (Excel export). */
export const ExportButton: React.FC<IExportButtonProps> = () => null;
