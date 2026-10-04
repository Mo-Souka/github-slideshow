import * as React from 'react';
import { RecordForm } from '../form/RecordForm';

export interface IRecordDetailProps {
  recordId: number;
}

/** Placeholder until stage 6 (detail view): opens the form. */
export const RecordDetail: React.FC<IRecordDetailProps> = ({ recordId }) => <RecordForm recordId={recordId} />;
